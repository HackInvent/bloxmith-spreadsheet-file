"""Bounded spreadsheet reading and writing: CSV through the standard library, XLSX through openpyxl.

Everything here is pure block logic with explicit limits. Third-party packages are
never imported at module level; ``workbook`` loads openpyxl lazily, only for XLSX.
"""

from __future__ import annotations

from collections.abc import Mapping
import csv
import datetime as dt
import hashlib
import io
import json
import math
import os
from pathlib import Path
import re
import stat
import time
import uuid

DEFAULTS = {
    "operation": "read", "on_error": "fail",
    "allowed_root": "", "sheet": "", "sheet_index": 1, "header_row": 1,
    "skip_empty_rows": True, "on_limit": "fail", "allow_macro_workbooks": False,
    "format": "xlsx", "output_directory": "", "file_name": "rows-{date}-{unique}",
    "overwrite": False, "columns": [], "header": True, "formula_guard": "escape",
    "delimiter": "comma", "encoding": "utf-8-sig",
    "max_rows": 10000, "max_columns": 256, "max_bytes": 20971520, "timeout_sec": 60,
}
# Framework-authored execution metadata may accompany the block settings.
FRAMEWORK_KEYS = {"execution", "runtime_path", "runtime_path_label"}
CHOICES = {"operation": ("read", "write"), "on_error": ("fail", "route"), "on_limit": ("fail", "truncate"),
           "format": ("xlsx", "csv"), "formula_guard": ("escape", "reject", "off"),
           "delimiter": ("comma", "semicolon", "tab", "pipe"),
           "encoding": ("utf-8-sig", "utf-8", "latin-1", "cp1252")}
BOUNDS = {"sheet_index": (1, 1000), "header_row": (0, 1000), "max_rows": (1, 100000),
          "max_columns": (1, 1024), "max_bytes": (1024, 52428800), "timeout_sec": (1, 240)}
FLAGS = ("skip_empty_rows", "allow_macro_workbooks", "overwrite", "header")
DELIMITERS = {"comma": ",", "semicolon": ";", "tab": "\t", "pipe": "|"}
READ_SUFFIXES = {".xlsx": "xlsx", ".xlsm": "xlsx", ".csv": "csv", ".tsv": "csv", ".txt": "csv"}
UNSUPPORTED_SUFFIXES = {".xls": "Legacy .xls (BIFF) workbooks are not supported; save the file as .xlsx.",
                        ".xlsb": "Binary .xlsb workbooks are not supported; save the file as .xlsx.",
                        ".ods": "OpenDocument .ods spreadsheets are not supported; save the file as .xlsx or .csv.",
                        ".xltx": "Excel templates are not read; save the workbook as .xlsx.",
                        ".xltm": "Excel templates are not read; save the workbook as .xlsx.",
                        ".numbers": "Apple Numbers files are not supported; export them as .xlsx or .csv."}
SPREADSHEET_SUFFIXES = {".xlsx", ".xlsm", ".xls", ".xlsb", ".xltx", ".xltm", ".csv", ".tsv", ".txt", ".ods"}
MAX_OUTPUT_BYTES = 8 * 1024 * 1024
XLSX_MAX_ROWS = 1048576
XLSX_CELL_CHARS = 32767
SAFE_INTEGER = 2 ** 53 - 1
SHEET_FORBIDDEN = re.compile(r"[\\/?*\[\]:]")
CONTROL = re.compile(r"[\x00-\x1f\x7f]")
XLSX_ILLEGAL = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")
FILE_NAME = re.compile(r"[A-Za-z0-9][A-Za-z0-9 ._()-]*")
PLACEHOLDER = re.compile(r"\{([^{}]*)\}")
PLACEHOLDERS = {"date", "time", "unique"}
# OWASP CSV injection prefixes; plain signed numbers stay untouched because they cannot run a formula.
FORMULA_PREFIX = ("=", "+", "-", "@", "\t", "\r")
PLAIN_NUMBER = re.compile(r"[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?")


class SpreadsheetError(ValueError):
    """Expected, content-free failure with a stable machine code."""

    def __init__(self, code, message):
        super().__init__(message)
        self.code = code
        self.message = message

    def envelope(self, operation):
        """Return the JSON document published on the error port when routing is enabled."""
        return {"code": self.code, "message": self.message, "operation": operation}


def fail(code, message):
    """Raise one expected error; kept short because most checks are one-liners."""
    raise SpreadsheetError(code, message)


def configuration(raw):
    """Validate every setting before any file is opened, read or written."""
    if not isinstance(raw, Mapping):
        fail("invalid_settings", "Settings must be an object.")
    unknown = sorted(str(key) for key in raw if key not in DEFAULTS and key not in FRAMEWORK_KEYS)
    if unknown:
        fail("invalid_settings", f"Unknown setting: {unknown[0][:64]}.")
    config = {**DEFAULTS, **{key: raw[key] for key in raw if key in DEFAULTS}}
    for key, options in CHOICES.items():
        if config[key] not in options:
            fail("invalid_settings", f"{key} must be one of: {', '.join(options)}.")
    for key, (low, high) in BOUNDS.items():
        value = config[key]
        if type(value) is not int or not low <= value <= high:
            fail("invalid_settings", f"{key} must be an integer between {low} and {high}.")
    for key in FLAGS:
        if not isinstance(config[key], bool):
            fail("invalid_settings", f"{key} must be true or false.")
    for key in ("allowed_root", "output_directory"):
        value = config[key]
        if not isinstance(value, str) or len(value) > 4096 or "\x00" in value:
            fail("invalid_settings", f"{key} must be a directory path of at most 4096 characters.")
    sheet = config["sheet"]
    if not isinstance(sheet, str) or len(sheet) > 31 or CONTROL.search(sheet):
        fail("invalid_settings", "sheet must be a worksheet name of at most 31 characters.")
    if sheet and (SHEET_FORBIDDEN.search(sheet) or sheet.startswith("'") or sheet.endswith("'")):
        fail("invalid_settings", "sheet cannot contain \\ / ? * [ ] : or start or end with an apostrophe.")
    columns = config["columns"]
    if isinstance(columns, str):
        try:
            columns = json.loads(columns) if columns.strip() else []
        except (ValueError, RecursionError):
            fail("invalid_settings", "columns must be a JSON array of column names.")
    if isinstance(columns, tuple):
        columns = list(columns)
    if not isinstance(columns, list) or len(columns) > config["max_columns"] or any(
            not isinstance(name, str) or not name or len(name) > 256 for name in columns) or len(set(columns)) != len(columns):
        fail("invalid_settings", "columns must list distinct, non-empty names (at most max_columns, 256 characters each).")
    config["columns"] = columns
    template = config["file_name"]
    if not isinstance(template, str) or not 1 <= len(template) <= 100:
        fail("invalid_settings", "file_name must contain 1 to 100 characters.")
    for name in PLACEHOLDER.findall(template):
        if name not in PLACEHOLDERS:
            fail("invalid_settings", "file_name accepts only the {date}, {time} and {unique} placeholders.")
    literal = PLACEHOLDER.sub("x", template)
    if "{" in literal or "}" in literal or not FILE_NAME.fullmatch(literal):
        fail("invalid_settings", "file_name may use letters, digits, spaces, . _ - ( ) and must start with a letter or digit.")
    suffix = Path(literal).suffix.lower()
    if config["operation"] == "write" and suffix in SPREADSHEET_SUFFIXES and suffix != "." + config["format"]:
        fail("invalid_settings", "The file_name extension must match format, or be omitted.")
    return config


def checker(timeout_sec, cancel):
    """Return a cheap deadline/cancellation probe; the cancel service is polled at most every 100 ms."""
    deadline = time.monotonic() + timeout_sec
    state = {"next": 0.0}

    def check():
        now = time.monotonic()
        if now >= state["next"]:
            state["next"] = now + 0.1
            if cancel():
                fail("cancelled", "Spreadsheet operation cancelled.")
        if now > deadline:
            fail("timeout", "The spreadsheet operation exceeded timeout_sec.")
    return check


def reference(value):
    """Accept a plain path, a JSON string or a JSON {"path": ...} reference; never a list of files."""
    if isinstance(value, str) and value.strip()[:1] in {"{", '"'}:
        try:
            value = json.loads(value)
        except (ValueError, RecursionError):
            fail("invalid_input", "Malformed file reference; send a path or {\"path\": \"...\"}.")
    if isinstance(value, Mapping):
        if set(value) != {"path"}:
            fail("invalid_input", "A file reference accepts only the path field.")
        value = value["path"]
    if not isinstance(value, str) or not value.strip() or len(value) > 4096 or "\x00" in value:
        fail("invalid_input", "A spreadsheet file path is required.")
    return value.strip()


def resolve_source(value, config, root_dir):
    """Resolve a regular file inside allowed_root, including symlink targets, and detect its format."""
    base = Path(root_dir).resolve()
    try:
        root = Path(config["allowed_root"]).expanduser() if config["allowed_root"] else base
        root = (root if root.is_absolute() else base / root).resolve(strict=True)
    except (OSError, RuntimeError):
        fail("invalid_path", "allowed_root does not exist.")
    if not root.is_dir():
        fail("invalid_path", "allowed_root must be a directory.")
    reference_path = reference(value)
    try:
        path = Path(reference_path).expanduser()
        path = (path if path.is_absolute() else root / path).resolve(strict=True)
    except FileNotFoundError:
        fail("not_found", "The spreadsheet file does not exist.")
    except (OSError, RuntimeError):
        fail("invalid_path", "The spreadsheet path cannot be resolved.")
    if not path.is_relative_to(root):
        fail("invalid_path", "The spreadsheet file must be inside allowed_root.")
    if not path.is_file():
        fail("invalid_path", "The source must be a regular file.")
    suffix = path.suffix.lower()
    if suffix in UNSUPPORTED_SUFFIXES:
        fail("unsupported_format", UNSUPPORTED_SUFFIXES[suffix])
    if suffix not in READ_SUFFIXES:
        fail("unsupported_format", "Only .xlsx, .xlsm, .csv, .tsv and .txt files are read.")
    if suffix == ".xlsm" and not config["allow_macro_workbooks"]:
        fail("macro_workbook", "Macro-enabled .xlsm workbooks are refused unless allow_macro_workbooks is enabled.")
    return path, READ_SUFFIXES[suffix]


def read_bytes(path, limit, check):
    """Read one stable bounded snapshot without following a swapped symlink or blocking on a FIFO."""
    flags = os.O_RDONLY | getattr(os, "O_NOFOLLOW", 0) | getattr(os, "O_NONBLOCK", 0) | getattr(os, "O_CLOEXEC", 0)
    try:
        descriptor = os.open(path, flags)
    except OSError:
        fail("invalid_path", "The spreadsheet file cannot be opened.")
    chunks, size = [], 0
    try:
        before = os.fstat(descriptor)
        if not stat.S_ISREG(before.st_mode):
            fail("invalid_path", "The source must be a regular file.")
        if before.st_size > limit:
            fail("too_large", "The spreadsheet file exceeds max_bytes.")
        while True:
            check()
            chunk = os.read(descriptor, 1048576)
            if not chunk:
                break
            size += len(chunk)
            if size > limit:
                fail("too_large", "The spreadsheet file exceeds max_bytes.")
            chunks.append(chunk)
        after = os.fstat(descriptor)
        if (before.st_size, before.st_mtime_ns) != (after.st_size, after.st_mtime_ns) or size != after.st_size:
            fail("source_changed", "The file changed while it was read; retry with the completed file.")
    finally:
        os.close(descriptor)
    return b"".join(chunks)


def effective_length(values):
    """Ignore trailing empty cells, so formatting or trailing separators never widen a table."""
    length = len(values)
    while length and (values[length - 1] is None or values[length - 1] == ""):
        length -= 1
    return length


def lower_bound(values):
    """Cheap lower bound of the JSON size of a row, to stop huge tables before building them."""
    total = 0
    for value in values:
        total += 2 + (len(value) if isinstance(value, str) else 2 if value is None or value is True or value is False else 0)
    return total


def header_text(value):
    """Turn a header cell into a trimmed key; numbers and dates become their text form."""
    if value is None:
        return ""
    if value is True or value is False:
        return "true" if value else "false"
    if isinstance(value, float):
        return repr(value)
    return str(value).strip()


def column_names(header, width):
    """Name every column deterministically and report each generated or renamed key."""
    texts = [header_text(header[index]) if index < len(header) else "" for index in range(width)]
    names, used = [None] * width, set()
    for index, text in enumerate(texts):
        if text and text not in used:
            names[index] = text
            used.add(text)
    changes = []
    for index, text in enumerate(texts):
        if names[index] is not None:
            continue
        base = text or f"column_{index + 1}"
        candidate, suffix = base, 1
        while candidate in used:
            suffix += 1
            candidate = f"{base}_{suffix}"
        names[index] = candidate
        used.add(candidate)
        changes.append({"column": index + 1, "original": text, "name": candidate,
                        "reason": "duplicate" if text else "empty"})
    return names, changes


def tabulate(records, config, check, row_limit=None):
    """Turn raw rows into JSON text: objects keyed by the header row, or arrays without a header."""
    header_row, max_rows, max_columns = config["header_row"], config["max_rows"], config["max_columns"]
    truncate = config["on_limit"] == "truncate"
    header, data, skipped, width, budget, seen, content = None, [], 0, 0, 2, 0, False
    truncated = {"rows": False, "columns": False}
    for number, record in enumerate(records, 1):
        check()
        seen = number
        if row_limit and number > row_limit:
            fail("malformed_file", "The worksheet declares rows beyond the XLSX limit of 1,048,576.")
        values = list(record)
        length = effective_length(values)
        content = content or bool(length)
        if header_row and number < header_row:
            continue
        if length > max_columns:
            if not truncate:
                fail("limit_exceeded", f"Row {number} has values beyond max_columns ({max_columns}).")
            truncated["columns"] = True
            values = values[:max_columns]
            length = effective_length(values)
        if header_row and number == header_row:
            header = values[:length]
            continue
        if not length and config["skip_empty_rows"]:
            skipped += 1
            continue
        if len(data) >= max_rows:
            if not truncate:
                fail("limit_exceeded", f"The source has more than max_rows ({max_rows}) data rows.")
            truncated["rows"] = True
            break
        while values and values[-1] is None:
            values.pop()
        budget += lower_bound(values)
        if budget > MAX_OUTPUT_BYTES:
            fail("too_large", "The extracted rows exceed the 8 MiB output limit; reduce max_rows or the selected columns.")
        width = max(width, length)
        data.append(values)
    names, changes = None, []
    if not content:
        data, skipped = [], 0
    if header_row:
        if content and header is None:
            fail("invalid_input", f"header_row {header_row} is beyond the last row ({seen}).")
        if content and not header:
            fail("invalid_input", f"Header row {header_row} is empty; select the row that holds the column names.")
        width = max(width, len(header or []))
        names, changes = column_names(header or [], width)
    encoded, size = [], 2
    for values in data:
        check()
        padded = values[:width] + [None] * (width - min(len(values), width))
        text = json.dumps(dict(zip(names, padded)) if names is not None else padded, ensure_ascii=False, allow_nan=False)
        size += len(text.encode("utf-8")) + 1
        if size > MAX_OUTPUT_BYTES:
            fail("too_large", "The extracted rows exceed the 8 MiB output limit; reduce max_rows or the selected columns.")
        encoded.append(text)
    shape = {"header_row": header_row, "columns": names if names is not None else width, "column_count": width,
             "rows": len(data), "skipped_empty_rows": skipped, "truncated": truncated,
             "header_changes": changes, "empty": not content}
    return "[" + ",".join(encoded) + "]", shape


def read_csv(data, config, check):
    """Parse delimiter-separated text with an explicit dialect; every value stays a string."""
    if data.startswith(b"PK\x03\x04"):
        fail("unsupported_format", "The file content is a ZIP package (for example XLSX), not CSV text.")
    try:
        text = data.decode(config["encoding"])
    except UnicodeDecodeError:
        fail("malformed_file", f"The file is not valid {config['encoding']} text; select the matching encoding.")
    if "\x00" in text:
        fail("malformed_file", "The CSV file contains NUL characters (binary content).")
    reader = csv.reader(io.StringIO(text, newline=""), delimiter=DELIMITERS[config["delimiter"]], strict=True)

    def records():
        try:
            yield from reader
        except csv.Error:
            fail("malformed_file", f"CSV parsing failed near line {reader.line_num}; check the delimiter, quoting "
                                   "and field length (at most 131,072 characters).")
    rows, shape = tabulate(records(), config, check)
    return rows, {**shape, "delimiter": config["delimiter"], "encoding": config["encoding"]}


def read(value, config, root_dir, cancel=lambda: False):
    """Read one complete source into JSON rows and a report; nothing partial is published."""
    config = configuration(config)
    check = checker(config["timeout_sec"], cancel)
    path, kind = resolve_source(value, config, root_dir)
    data = read_bytes(path, config["max_bytes"], check)
    if kind == "xlsx":
        from . import workbook
        rows, details = workbook.read_xlsx(data, config, check, path.suffix.lower() == ".xlsm")
    else:
        rows, details = read_csv(data, config, check)
    check()
    report = {"operation": "read", "format": kind,
              "source": {"path": str(path), "name": path.name, "bytes": len(data), "sha256": hashlib.sha256(data).hexdigest()},
              **details}
    return {"rows": rows, "report": report}


def strict_json(text):
    """Decode JSON without NaN/Infinity, which spreadsheets and JSON consumers cannot represent."""
    def constant(_):
        raise ValueError("non-finite number")
    try:
        return json.loads(text, parse_constant=constant)
    except (ValueError, RecursionError):
        fail("invalid_input", "rows must be finite JSON: an array of objects or arrays, or {\"rows\": [...]}.")


def input_rows(value, config):
    """Accept an array of objects or of arrays (optionally wrapped in {"rows": ...}), never a mix."""
    if isinstance(value, (bytes, bytearray)):
        value = bytes(value).decode("utf-8", errors="strict")
    if isinstance(value, str):
        if len(value.encode("utf-8")) > config["max_bytes"]:
            fail("too_large", "The rows input exceeds max_bytes.")
        value = strict_json(value)
    else:
        try:
            size = len(json.dumps(value, ensure_ascii=False, allow_nan=False).encode("utf-8"))
        except (TypeError, ValueError, RecursionError):
            fail("invalid_input", "rows must be finite JSON data.")
        if size > config["max_bytes"]:
            fail("too_large", "The rows input exceeds max_bytes.")
    if isinstance(value, Mapping) and set(value) == {"rows"}:
        value = value["rows"]
    if not isinstance(value, list):
        fail("invalid_input", "rows must be a JSON array of objects or arrays, or {\"rows\": [...]}.")
    if len(value) > config["max_rows"]:
        fail("limit_exceeded", f"The rows input has more than max_rows ({config['max_rows']}) rows.")
    kinds = {"object" if isinstance(row, Mapping) else "array" if isinstance(row, list) else "other" for row in value}
    if "other" in kinds:
        index = next(i for i, row in enumerate(value, 1) if not isinstance(row, (Mapping, list)))
        fail("invalid_input", f"Row {index} must be a JSON object or array.")
    if len(kinds) > 1:
        fail("invalid_input", "rows cannot mix objects and arrays.")
    return value, (kinds.pop() if kinds else "object")


def table(value, config, check):
    """Lay out header names and row values in the explicit (or first-seen) column order."""
    rows, kind = input_rows(value, config)
    ignored, ignored_count = [], 0
    if kind == "object":
        names = list(config["columns"])
        if not names:
            seen = set()
            for row in rows:
                for key in row:
                    if key not in seen:
                        seen.add(key)
                        names.append(key)
                        if len(names) > config["max_columns"]:
                            fail("limit_exceeded", f"The rows contain more than max_columns ({config['max_columns']}) keys.")
        known = set(names)
        matrix = []
        for row in rows:
            check()
            extra = [key for key in row if key not in known]
            if extra:
                ignored_count += len(extra)
                for key in extra:
                    if key not in ignored and len(ignored) < 100:
                        ignored.append(key)
            matrix.append([row.get(name) for name in names])
        header = names
    else:
        width = max((len(row) for row in rows), default=0)
        if width > config["max_columns"]:
            fail("limit_exceeded", f"A row has more than max_columns ({config['max_columns']}) values.")
        if config["columns"] and width > len(config["columns"]):
            fail("invalid_input", "An array row has more values than the configured columns.")
        header, matrix = list(config["columns"]), [list(row) for row in rows]
    return (header if config["header"] and header else None), matrix, {
        "input": kind, "columns": header if header else max((len(row) for row in matrix), default=0),
        "ignored_keys": ignored, "ignored_key_count": ignored_count}


def scalar(value, counters):
    """Map one JSON value to a spreadsheet value; nested values become compact JSON text."""
    if isinstance(value, (Mapping, list)):
        counters["json_cells"] += 1
        try:
            return json.dumps(value, ensure_ascii=False, allow_nan=False, separators=(",", ":"))
        except (ValueError, TypeError, RecursionError):
            fail("invalid_input", "A nested value cannot be serialized as JSON text.")
    if isinstance(value, float) and not math.isfinite(value):
        fail("invalid_input", "Numbers must be finite.")
    if value is None or isinstance(value, (str, bool, int, float)):
        return value
    fail("invalid_input", "Cells accept JSON strings, numbers, booleans, null and nested JSON only.")


def csv_text(value, position, config, counters):
    """Render a CSV field; the formula guard neutralizes spreadsheet formula prefixes."""
    if value is None:
        return ""
    if value is True or value is False:
        return "true" if value else "false"
    if isinstance(value, int):
        return str(value)
    if isinstance(value, float):
        return repr(value)
    if "\x00" in value:
        fail("invalid_value", f"{position}: text contains a NUL character.")
    if value.startswith(FORMULA_PREFIX) and not PLAIN_NUMBER.fullmatch(value):
        if config["formula_guard"] == "reject":
            fail("invalid_value", f"{position}: text starts like a spreadsheet formula and formula_guard is reject.")
        if config["formula_guard"] == "escape":
            counters["escaped_formula_cells"] += 1
            return "'" + value
    return value


def write_csv(header, matrix, config, check):
    """Serialize RFC 4180 CSV (CRLF records, minimal quoting) in the selected encoding."""
    counters = {"json_cells": 0, "escaped_formula_cells": 0}
    buffer = io.StringIO()
    writer = csv.writer(buffer, delimiter=DELIMITERS[config["delimiter"]], lineterminator="\r\n", quoting=csv.QUOTE_MINIMAL)
    if header is not None:
        writer.writerow([csv_text(name, f"Header column {index}", config, counters)
                         for index, name in enumerate(header, 1)])
    for number, row in enumerate(matrix, 1):
        check()
        writer.writerow([csv_text(scalar(value, counters), f"Row {number}, column {index}", config, counters)
                         for index, value in enumerate(row, 1)])
    try:
        data = buffer.getvalue().encode(config["encoding"])
    except UnicodeEncodeError:
        fail("invalid_value", f"A value cannot be encoded in {config['encoding']}; choose utf-8-sig or utf-8.")
    return data, {"delimiter": config["delimiter"], "encoding": config["encoding"],
                  "formula_guard": config["formula_guard"], **counters}


def target_name(config, now=None):
    """Expand the fixed placeholders (UTC date/time, random suffix) and add the format extension."""
    now = now or dt.datetime.now(dt.timezone.utc)
    values = {"date": now.strftime("%Y-%m-%d"), "time": now.strftime("%H%M%S"), "unique": uuid.uuid4().hex[:12]}
    name = PLACEHOLDER.sub(lambda match: values[match.group(1)], config["file_name"])
    extension = "." + config["format"]
    if not name.lower().endswith(extension):
        name += extension
    if len(name) > 160 or not FILE_NAME.fullmatch(name):
        fail("invalid_settings", "The expanded file_name is not a valid file name.")
    return name


def write(value, config, storage, root_dir, cancel=lambda: False):
    """Write one complete file atomically, then report it; existing files are never replaced by default."""
    config = configuration(config)
    check = checker(config["timeout_sec"], cancel)
    header, matrix, layout = table(value, config, check)
    if config["format"] == "xlsx":
        from . import workbook
        data, details = workbook.write_xlsx(header, matrix, config, check, scalar)
    else:
        data, details = write_csv(header, matrix, config, check)
    if len(data) > config["max_bytes"]:
        fail("too_large", "The generated file exceeds max_bytes.")
    check()
    from . import storage as publication
    path, overwritten = publication.publish(data, target_name(config), config, storage, root_dir, check)
    report = {"operation": "write", "format": config["format"], "path": path, "file_name": Path(path).name,
              "bytes": len(data), "sha256": hashlib.sha256(data).hexdigest(), "rows": len(matrix),
              "header": header is not None, "overwritten": overwritten, **layout, **details}
    return {"file": path, "report": report}
