"""FB1/FB2/FB3/FB4: direct read/write logic, safety limits, error routing and the missing-openpyxl path."""

from pathlib import Path
from types import SimpleNamespace
import csv
import io
import json
import os
import sys
import tempfile
import threading
import time

ROOT = Path(__file__).resolve().parents[3]
sys.path[:0] = [str(ROOT), str(Path(__file__).parent)]
# FB4 - Without openpyxl the package must still import and CSV must keep working.
sys.modules["openpyxl"] = None
from blocs.spreadsheet_file import logic, workbook  # noqa: E402
from blocs.spreadsheet_file.block import SpreadsheetFileBlock  # noqa: E402
from blocs.spreadsheet_file.logic import SpreadsheetError  # noqa: E402
from bloxsmith_app.block_api import BlockRuntimeContext  # noqa: E402
from bloxsmith_app.block_runtime import BlockInputEvent  # noqa: E402
import spreadsheet_fixtures as fx  # noqa: E402

PORTS_IN = (SimpleNamespace(id=1, name="file"), SimpleNamespace(id=2, name="rows"))
PORTS_OUT = tuple(SimpleNamespace(id=i, name=n) for i, n in enumerate(("rows", "file", "report", "error"), 1))


def fails(code, fn):
    """Require one explicit error code, never a permissive fallback."""
    try:
        fn()
    except SpreadsheetError as exc:
        assert exc.code == code, (code, exc.code, exc.message)
        return exc
    raise AssertionError(f"Expected {code}")


def context(config, root, *, inputs=None, events=(), services=None):
    """Build a direct runtime context with the block's real ports."""
    return BlockRuntimeContext(run_id="unit", node_id="sheet", kind="spreadsheet_file", title="Sheet", config=config,
                               inputs=inputs or {}, input_ports=PORTS_IN, output_ports=PORTS_OUT, root_dir=root,
                               input_events=tuple(events), services=services or {})


def outputs(result):
    """Map emitted port names to values."""
    return {output.port_name: output for output in result.outputs}


def missing_dependency(root):
    """FB4: CSV works, XLSX fails cleanly at prepare and execute time while openpyxl is unavailable."""
    block = SpreadsheetFileBlock()
    assert not workbook.openpyxl_available()
    (root / "plain.csv").write_text("a,b\n1,2\n", encoding="utf-8")
    (root / "book.xlsx").write_bytes(fx.minimal_xlsx({"S": "<row r=\"1\">" + fx.inline("A1", "a") + "</row>"}))
    assert json.loads(logic.read("plain.csv", {}, root)["rows"]) == [{"a": "1", "b": "2"}]
    fails("missing_dependency", lambda: logic.read("book.xlsx", {}, root))
    store = root / "store"
    store.mkdir()
    fails("missing_dependency", lambda: logic.write([{"a": 1}], {"operation": "write"}, lambda: store, root))
    assert not list(store.rglob("*.xlsx")), "No partial file may exist"
    fails("missing_dependency", lambda: block.prepare_runtime(SimpleNamespace(config={"operation": "write", "format": "xlsx"})))
    block.prepare_runtime(SimpleNamespace(config={"operation": "write", "format": "csv"}))
    failed = block.execute_runtime(context({}, root, inputs={"file": "book.xlsx"}))
    assert failed.status == "failed" and not failed.outputs and "openpyxl" in failed.error
    routed = block.execute_runtime(context({"on_error": "route"}, root, inputs={"file": "book.xlsx"}))
    assert routed.status == "success" and json.loads(outputs(routed)["error"].value)["code"] == "missing_dependency"


def settings():
    """FB3: every setting is validated before files are touched."""
    assert logic.configuration({})["operation"] == "read"
    assert logic.configuration({"columns": "[\"a\",\"b\"]", "runtime_path": [], "execution": {}})["columns"] == ["a", "b"]
    assert logic.configuration({"columns": ("a",)})["columns"] == ["a"]
    for bad in ({"operation": "delete"}, {"unknown": 1}, {"max_rows": 0}, {"max_rows": True}, {"max_bytes": 10 ** 9},
                {"timeout_sec": 241}, {"header_row": -1}, {"delimiter": ","}, {"encoding": "utf-16"},
                {"overwrite": "yes"}, {"sheet": "a/b"}, {"sheet": "x" * 32}, {"sheet": "'quoted'"},
                {"columns": ["a", "a"]}, {"columns": [""]}, {"columns": "not json"}, {"max_columns": 1, "columns": ["a", "b"]},
                {"file_name": "../escape"}, {"file_name": "{user}"}, {"file_name": ".hidden"}, {"file_name": "a/b"},
                {"file_name": ""}, {"allowed_root": "a\x00b"}, {"operation": "write", "format": "csv", "file_name": "x.xlsx"}):
        fails("invalid_settings", lambda b=bad: logic.configuration(b))
    assert logic.configuration({"file_name": "x.xlsx", "format": "csv"})  # extension only checked for write
    name = logic.target_name(logic.configuration({"file_name": "report-{date}-{time}-{unique}"}))
    assert name.startswith("report-") and name.endswith(".xlsx") and len(name) == len("report-2024-01-31-120000-") + 12 + 5
    assert logic.target_name(logic.configuration({"file_name": "data.CSV", "format": "csv", "operation": "write"})) == "data.CSV"


def csv_reading(root):
    """FB1: explicit dialect, header naming, empty rows, encodings and strings kept as text."""
    (root / "people.csv").write_text('id;name;;name\n1;"Smith; Ann";x;dup\n\n2;"multi\nline"\n;;;\n', encoding="utf-8")
    result = logic.read("people.csv", {"delimiter": "semicolon"}, root)
    rows, report = json.loads(result["rows"]), result["report"]
    assert rows == [{"id": "1", "name": "Smith; Ann", "column_3": "x", "name_2": "dup"},
                    {"id": "2", "name": "multi\nline", "column_3": None, "name_2": None}], rows
    assert report["skipped_empty_rows"] == 2 and report["columns"] == ["id", "name", "column_3", "name_2"]
    assert [change["reason"] for change in report["header_changes"]] == ["empty", "duplicate"]
    assert report["source"]["name"] == "people.csv" and len(report["source"]["sha256"]) == 64
    kept = json.loads(logic.read("people.csv", {"delimiter": "semicolon", "skip_empty_rows": False}, root)["rows"])
    assert len(kept) == 4 and kept[1] == {"id": None, "name": None, "column_3": None, "name_2": None}
    arrays = json.loads(logic.read("people.csv", {"delimiter": "semicolon", "header_row": 0}, root)["rows"])
    assert arrays[0] == ["id", "name", "", "name"] and arrays[1] == ["1", "Smith; Ann", "x", "dup"]
    (root / "bom.csv").write_bytes("﻿code,label\nE1,café\n".encode("utf-8"))
    assert json.loads(logic.read("bom.csv", {}, root)["rows"]) == [{"code": "E1", "label": "café"}]
    (root / "latin.tsv").write_bytes("code\tlabel\nE1\tcafé\n".encode("cp1252"))
    fails("malformed_file", lambda: logic.read("latin.tsv", {"delimiter": "tab"}, root))
    assert json.loads(logic.read("latin.tsv", {"delimiter": "tab", "encoding": "cp1252"}, root)["rows"])[0]["label"] == "café"
    (root / "title.txt").write_text("Export of the day\n\nsku,qty\nA,3\n", encoding="utf-8")
    assert json.loads(logic.read("title.txt", {"header_row": 3}, root)["rows"]) == [{"sku": "A", "qty": "3"}]
    (root / "empty.csv").write_text("\n\n", encoding="utf-8")
    empty = logic.read("empty.csv", {}, root)
    assert empty["rows"] == "[]" and empty["report"]["empty"]
    (root / "header_only.csv").write_text("a,b\n", encoding="utf-8")
    assert logic.read("header_only.csv", {}, root)["report"]["columns"] == ["a", "b"]
    fails("invalid_input", lambda: logic.read("title.txt", {"header_row": 2}, root))
    fails("invalid_input", lambda: logic.read("title.txt", {"header_row": 9}, root))
    (root / "broken.csv").write_text('a,b\n"open,1\n', encoding="utf-8")
    fails("malformed_file", lambda: logic.read("broken.csv", {}, root))
    (root / "binary.csv").write_bytes(b"a,b\n1,\x00\n")
    fails("malformed_file", lambda: logic.read("binary.csv", {}, root))
    (root / "zip.csv").write_bytes(fx.minimal_xlsx({"S": ""}))
    fails("unsupported_format", lambda: logic.read("zip.csv", {}, root))


def limits_and_paths(root):
    """FB3: bounded rows/columns/bytes/output, path boundary and refused formats."""
    (root / "wide.csv").write_text("a,b,c\n1,2,3\n4,5,6\n7,8,9\n", encoding="utf-8")
    fails("limit_exceeded", lambda: logic.read("wide.csv", {"max_rows": 2}, root))
    fails("limit_exceeded", lambda: logic.read("wide.csv", {"max_columns": 2}, root))
    cut = logic.read("wide.csv", {"max_rows": 2, "max_columns": 2, "on_limit": "truncate"}, root)
    assert json.loads(cut["rows"]) == [{"a": "1", "b": "2"}, {"a": "4", "b": "5"}]
    assert cut["report"]["truncated"] == {"rows": True, "columns": True}
    (root / "long.csv").write_text("a\n" + "x\n" * 600, encoding="utf-8")
    fails("too_large", lambda: logic.read("long.csv", {"max_bytes": 1024}, root))
    (root / "huge.csv").write_text("v\n" + ("y" * 100000 + "\n") * 90, encoding="utf-8")
    error = fails("too_large", lambda: logic.read("huge.csv", {"max_bytes": 52428800}, root))
    assert "y" * 20 not in error.message
    inside = root / "inside"
    inside.mkdir()
    (inside / "ok.csv").write_text("a\n1\n", encoding="utf-8")
    assert logic.read("ok.csv", {"allowed_root": str(inside)}, root)["report"]["rows"] == 1
    assert logic.read(str(inside / "ok.csv"), {"allowed_root": "inside"}, root)["report"]["rows"] == 1
    assert logic.read('{"path": "ok.csv"}', {"allowed_root": "inside"}, root)["report"]["rows"] == 1
    fails("invalid_path", lambda: logic.read("../wide.csv", {"allowed_root": "inside"}, root))
    os.symlink(root / "wide.csv", inside / "link.csv")
    fails("invalid_path", lambda: logic.read("link.csv", {"allowed_root": "inside"}, root))
    fails("invalid_path", lambda: logic.read("inside", {}, root))
    fails("not_found", lambda: logic.read("absent.csv", {}, root))
    fails("invalid_path", lambda: logic.read("ok.csv", {"allowed_root": "missing-dir"}, root))
    fails("invalid_input", lambda: logic.read("", {}, root))
    fails("invalid_input", lambda: logic.read('{"path": "a", "mode": "x"}', {}, root))
    for name, code in (("old.xls", "unsupported_format"), ("sheet.ods", "unsupported_format"),
                       ("notes.pdf", "unsupported_format"), ("macro.xlsm", "macro_workbook")):
        (root / name).write_bytes(b"data")
        fails(code, lambda n=name: logic.read(n, {}, root))


def xlsx_reading(root):
    """FB1/FB3: sheets, cached formulas, dates, error cells and hostile package structures."""
    fx.rich_workbook(root / "rich.xlsx")
    result = logic.read("rich.xlsx", {"sheet": "Data", "header_row": 3}, root)
    rows, report = json.loads(result["rows"]), result["report"]
    assert rows[0] == {"id": 1, "day": "2024-01-31", "moment": "2024-01-31T08:30:05", "clock": "08:30:00",
                       "span": "P1DT2H3M", "ok": True, "ratio": 0.25, "column_8": None, "id_2": "first", "2024": "x"}, rows
    assert rows[1]["moment"] == "2024-02-01" and rows[1]["ok"] is False and len(rows) == 2
    assert report["sheets"] == ["Notes", "Data", "Archive"] and report["sheet_index"] == 2
    assert report["formulas"] == "cached_values" and report["skipped_empty_rows"] == 1
    hidden = logic.read("rich.xlsx", {"sheet_index": 3, "header_row": 0}, root)
    assert json.loads(hidden["rows"]) == [["archived"], ["old row"]] and hidden["report"]["sheet_state"] == "hidden"
    assert json.loads(logic.read("rich.xlsx", {}, root)["rows"]) == []
    fails("invalid_input", lambda: logic.read("rich.xlsx", {"sheet": "Chart"}, root))
    fails("invalid_input", lambda: logic.read("rich.xlsx", {"sheet": "Missing"}, root))
    fails("invalid_input", lambda: logic.read("rich.xlsx", {"sheet_index": 4}, root))
    sheet = ('<row r="1">' + fx.inline("A1", "a") + fx.inline("B1", "b") + fx.inline("C1", "c") + "</row>"
             '<row r="2"><c r="A2"><v>2</v></c><c r="B2"><f>A2*2</f><v>4</v></c><c r="C2" t="e"><v>#DIV/0!</v></c></row>'
             '<row r="3"><c r="A3" t="s"><v>0</v></c><c r="B3"><f>A3*2</f></c><c r="C3" t="b"><v>1</v></c></row>')
    (root / "calc.xlsx").write_bytes(fx.minimal_xlsx({"Calc": sheet}, shared=["shared text"]))
    assert json.loads(logic.read("calc.xlsx", {}, root)["rows"]) == [
        {"a": 2, "b": 4, "c": "#DIV/0!"}, {"a": "shared text", "b": None, "c": True}]
    forged = '<row r="1">' + fx.inline("A1", "only") + "</row>"
    package = fx.minimal_xlsx({"S": forged}).replace(b"<sheetData>", b'<dimension ref="A1:XFD1048576"/><sheetData>')
    (root / "forged.xlsx").write_bytes(package)
    assert logic.read("forged.xlsx", {"header_row": 0}, root)["rows"] == '[["only"]]'
    cases = {
        "beyond.xlsx": (fx.minimal_xlsx({"S": '<row r="1048577"><c r="A1048577"><v>1</v></c></row>'}), "malformed_file"),
        "badindex.xlsx": (fx.minimal_xlsx({"S": '<row r="1"><c r="A1" t="s"><v>7</v></c></row>'}, shared=["x"]), "malformed_file"),
        "infinite.xlsx": (fx.minimal_xlsx({"S": '<row r="1"><c r="A1"><v>1e999</v></c></row>'}), "malformed_file"),
        "dtd.xlsx": (fx.minimal_xlsx({"S": ""}, prolog='<?xml version="1.0"?><!DOCTYPE x [<!ENTITY a "aaaa">]>'), "malformed_file"),
        "utf16.xlsx": (fx.minimal_xlsx({"S": ""}, extra={"xl/extra.xml": '﻿<!DOCTYPE x [<!ENTITY a "b">]><x/>'.encode("utf-16-le")}), "malformed_file"),
        "hidden_dtd.xlsx": (fx.minimal_xlsx({"S": ""}, extra={"xl/padded.xml": '<?xml version="1.0"?><?pad ' + "x" * 70000
                                                                + '?><!DOCTYPE x [<!ENTITY a "b">]><x/>'}), "malformed_file"),
        "macro.xlsx": (fx.minimal_xlsx({"S": ""}, extra={"xl/vbaProject.bin": b"\x00macro"}), "macro_workbook"),
        "text.xlsx": (b"id,name\n1,a\n", "malformed_file"),
        "protected.xlsx": (b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1" + b"\x00" * 600, "unsupported_format"),
        "bomb.xlsx": (fx.forge_entry(fx.minimal_xlsx({"S": ""}), "xl/worksheets/sheet1.xml", size=600 * 1024 * 1024), "too_large"),
        "crypt.xlsx": (fx.forge_entry(fx.minimal_xlsx({"S": ""}), "xl/workbook.xml", flags=0x1), "unsupported_format"),
    }
    for name, (payload, code) in cases.items():
        (root / name).write_bytes(payload)
        fails(code, lambda n=name: logic.read(n, {"header_row": 0}, root))
    allowed = logic.read("macro.xlsx", {"allow_macro_workbooks": True}, root)
    assert allowed["report"]["macro_workbook"] and allowed["rows"] == "[]"
    links = fx.minimal_xlsx({"S": '<row r="1">' + fx.inline("A1", "v") + "</row>"},
                            extra={"xl/externalLinks/externalLink1.xml": "<externalLink/>"})
    (root / "links.xlsx").write_bytes(links)
    assert logic.read("links.xlsx", {"header_row": 0}, root)["report"]["external_links"] == 1
    many = '<row r="1">' + fx.inline("A1", "n") + "</row>" + "".join(f'<row r="{i}"><c r="A{i}"><v>{i}</v></c></row>' for i in range(2, 30))
    (root / "many.xlsx").write_bytes(fx.minimal_xlsx({"S": many}))
    fails("limit_exceeded", lambda: logic.read("many.xlsx", {"max_rows": 10}, root))
    assert logic.read("many.xlsx", {"max_rows": 10, "on_limit": "truncate"}, root)["report"]["rows"] == 10


def writing(root):
    """FB2: explicit column order, literal text, formula guard, encodings and atomic no-overwrite publication."""
    store = root / "storage"
    store.mkdir()
    storage = lambda: store  # noqa: E731 - the framework injects a callable
    rows = [{"id": 1, "name": "=HYPERLINK(\"x\")", "note": "#N/A", "big": 2 ** 60, "tags": ["a", 1], "ok": True},
            {"id": 2, "name": "-5", "extra": "left out", "ratio": 0.1}]
    written = logic.write(json.dumps(rows), {"operation": "write", "columns": ["id", "name", "note", "big", "tags", "ok", "ratio"]},
                          storage, root)
    report = written["report"]
    path = Path(written["file"])
    assert path.parent == store / "spreadsheets" and path.suffix == ".xlsx" and report["path"] == str(path)
    assert report["ignored_keys"] == ["extra"] and report["numbers_as_text"] == 1 and report["json_cells"] == 1
    assert (path.stat().st_mode & 0o077) == 0, "Generated files are private to the Studio user"
    back = json.loads(logic.read(str(path), {}, root)["rows"])
    assert back == [{"id": 1, "name": "=HYPERLINK(\"x\")", "note": "#N/A", "big": str(2 ** 60), "tags": "[\"a\",1]", "ok": True, "ratio": None},
                    {"id": 2, "name": "-5", "note": None, "big": None, "tags": None, "ok": None, "ratio": 0.1}], back
    import openpyxl
    cells = list(openpyxl.load_workbook(path).active.iter_rows(min_row=2, max_row=2))[0]
    assert cells[1].data_type == "s" and cells[2].data_type == "s", "Text must never become a formula or error code"
    plain = {"operation": "write", "format": "csv", "file_name": "export", "columns": ["a", "b", "c", "d", "e", "f"]}
    data = [["=1+2", "@SUM(A1)", "-3.5e2", "+33 1 23", "\tlead", None], [1.5, False, {"k": "v"}, "café", "line\nbreak", ""]]
    first = logic.write(data, plain, storage, root)
    raw = Path(first["file"]).read_bytes()
    assert raw.startswith(b"\xef\xbb\xbfa,b,c,d,e,f\r\n'=1+2,'@SUM(A1),-3.5e2,'+33 1 23,'\tlead,\r\n")
    assert first["report"]["escaped_formula_cells"] == 4 and Path(first["file"]).name == "export.csv"
    parsed = list(csv.reader(io.StringIO(raw.decode("utf-8-sig"), newline="")))
    assert parsed[2] == ["1.5", "false", '{"k":"v"}', "café", "line\nbreak", ""]
    fails("target_exists", lambda: logic.write(data, plain, storage, root))
    assert Path(first["file"]).read_bytes() == raw, "An existing file is never replaced without overwrite"
    replaced = logic.write([["new"]], {**plain, "overwrite": True, "columns": []}, storage, root)
    assert replaced["report"]["overwritten"] and Path(replaced["file"]).read_bytes() == b"\xef\xbb\xbfnew\r\n"
    fails("invalid_value", lambda: logic.write(data, {**plain, "file_name": "r", "formula_guard": "reject"}, storage, root))
    raw_off = Path(logic.write(data, {**plain, "file_name": "raw", "formula_guard": "off"}, storage, root)["file"]).read_bytes()
    assert b"\r\n=1+2,@SUM(A1)," in raw_off
    fails("invalid_value", lambda: logic.write([["€"]], {**plain, "file_name": "l1", "encoding": "latin-1"}, storage, root))
    assert Path(logic.write([["€"]], {**plain, "file_name": "w1", "encoding": "cp1252", "columns": []}, storage, root)["file"]).read_bytes() == b"\x80\r\n"
    (store / "spreadsheets" / "folder.csv").mkdir()
    fails("target_exists", lambda: logic.write([["x"]], {**plain, "file_name": "folder", "overwrite": True, "columns": []}, storage, root))
    xlsx = {"operation": "write", "file_name": "t"}
    fails("invalid_value", lambda: logic.write([["x" * 32768]], xlsx, storage, root))
    fails("invalid_value", lambda: logic.write([["bell\x07"]], xlsx, storage, root))
    fails("invalid_input", lambda: logic.write('[{"a": NaN}]', xlsx, storage, root))
    fails("invalid_input", lambda: logic.write([{"a": 1}, [1]], xlsx, storage, root))
    fails("invalid_input", lambda: logic.write([1, 2], xlsx, storage, root))
    fails("invalid_input", lambda: logic.write('{"data": []}', xlsx, storage, root))
    fails("invalid_input", lambda: logic.write([[1, 2, 3]], {**xlsx, "columns": ["a"]}, storage, root))
    fails("limit_exceeded", lambda: logic.write([{"a": 1}] * 3, {**xlsx, "max_rows": 2}, storage, root))
    fails("limit_exceeded", lambda: logic.write([{"a": 1, "b": 2}], {**xlsx, "max_columns": 1}, storage, root))
    fails("too_large", lambda: logic.write([{"a": "z" * 900}] * 3, {**xlsx, "max_bytes": 1024}, storage, root))
    arrays = logic.write({"rows": [[1, "a"], [2]]}, {**xlsx, "file_name": "arrays", "columns": ["n", "s"]}, storage, root)
    assert json.loads(logic.read(arrays["file"], {}, root)["rows"]) == [{"n": 1, "s": "a"}, {"n": 2, "s": None}]
    bare = logic.write([], {**xlsx, "file_name": "nothing", "header": False}, storage, root)
    assert bare["report"]["rows"] == 0 and logic.read(bare["file"], {}, root)["report"]["empty"]
    named = logic.write([{"x": 1}], {**xlsx, "file_name": "named", "sheet": "Résumé"}, storage, root)
    assert logic.read(named["file"], {"sheet": "Résumé"}, root)["report"]["sheet"] == "Résumé"
    exports = root / "exports"
    exports.mkdir()
    direct = logic.write([{"x": 1}], {**xlsx, "file_name": "direct", "output_directory": "exports"}, None, root)
    assert Path(direct["file"]).parent == exports
    fails("storage_unavailable", lambda: logic.write([{"x": 1}], xlsx, None, root))
    fails("storage_unavailable", lambda: logic.write([{"x": 1}], {**xlsx, "output_directory": "absent"}, None, root))
    leftovers = [p.name for p in (store / "spreadsheets").iterdir() if p.name.startswith(".spreadsheet-")]
    assert not leftovers, leftovers
    # Concurrent writers to one name: exactly one wins, the other refuses instead of overwriting.
    outcomes = []

    def race():
        try:
            outcomes.append(logic.write([{"v": threading.get_ident()}], {**xlsx, "file_name": "race"}, storage, root)["file"])
        except SpreadsheetError as exc:
            outcomes.append(exc.code)
    workers = [threading.Thread(target=race) for _ in range(4)]
    for worker in workers:
        worker.start()
    for worker in workers:
        worker.join()
    assert sorted(outcome == "target_exists" for outcome in outcomes) == [False, True, True, True], outcomes


def deadlines(root):
    """FB3: timeout and cancellation stop long work without output."""
    (root / "slow.csv").write_bytes(b"\n" * 6000000)
    started = time.monotonic()
    fails("timeout", lambda: logic.read("slow.csv", {"timeout_sec": 1, "max_bytes": 52428800}, root))
    assert time.monotonic() - started < 3
    begin = time.monotonic()
    fails("cancelled", lambda: logic.read("slow.csv", {"max_bytes": 52428800}, root, lambda: time.monotonic() - begin > 0.3))
    store = root / "cancel-store"
    store.mkdir()
    fails("cancelled", lambda: logic.write([{"a": 1}] * 50, {"operation": "write", "format": "csv"}, lambda: store, root, lambda: True))
    assert not list(store.rglob("*.csv")), "A cancelled write publishes nothing"


def block_contract(root):
    """FB4: input routing, skipped activations, failed versus routed errors and UI validation."""
    block = SpreadsheetFileBlock()
    (root / "in.csv").write_text("k,v\na,1\n", encoding="utf-8")
    skipped = block.execute_runtime(context({}, root))
    assert skipped.status == "skipped" and not skipped.outputs
    done = block.execute_runtime(context({}, root, inputs={"file": "in.csv"}))
    out = outputs(done)
    assert done.status == "success" and set(out) == {"rows", "report"}
    assert out["rows"].content_type == "application/json" and json.loads(out["rows"].value) == [{"k": "a", "v": "1"}]
    store = root / "block-store"
    store.mkdir()
    services = {"get_block_storage_dir": lambda: store, "cancel_requested": lambda: False}
    event = BlockInputEvent(edge_id="e", input_port_id=2, input_port_name="rows", source_node_id="s", source_port_id=1,
                            value='[{"k":"b"}]', content_type="application/json")
    wrote = block.execute_runtime(context({"operation": "write", "format": "csv"}, root, events=[event], services=services))
    out = outputs(wrote)
    assert wrote.status == "success" and set(out) == {"file", "report"} and out["file"].content_type == "file/path"
    assert Path(out["file"].value).read_bytes() == b"\xef\xbb\xbfk\r\nb\r\n"
    wrong = block.execute_runtime(context({}, root, events=[event]))
    assert wrong.status == "failed" and "rows input" in wrong.error and not wrong.outputs
    (root.parent / "outside.csv").write_text("k\nsecret-value\n", encoding="utf-8")
    routed = block.execute_runtime(context({"on_error": "route"}, root, inputs={"file": "../outside.csv"}))
    document = json.loads(outputs(routed)["error"].value)
    assert routed.status == "success" and set(outputs(routed)) == {"error"}
    assert document == {"code": "invalid_path", "message": document["message"], "operation": "read"}
    assert "outside" not in document["message"] and "secret-value" not in routed.outputs[0].value
    cancelled = block.execute_runtime(context({"on_error": "route"}, root, inputs={"file": "in.csv"},
                                              services={"cancel_requested": lambda: True}))
    assert cancelled.status == "skipped" and not cancelled.outputs, "Cancellation is never routed as data"
    invalid = block.execute_runtime(context({"operation": "copy", "on_error": "route"}, root, inputs={"file": "in.csv"}))
    assert invalid.status == "failed" and not invalid.outputs
    fails("invalid_settings", lambda: block.prepare_runtime(SimpleNamespace(config={"max_rows": -1})))
    node = block.build_node_payload(node_id="sheet", position={"x": 0, "y": 0})
    refused = block.handle_ui_action(node=node, action="save_settings", values={"node_patch": {"config": {"max_rows": 0}}})
    assert "max_rows" in refused["error"]
    saved = block.handle_ui_action(node=node, action="save_settings", values={"node_patch": {"config": {"columns": ["a"]}}})
    operation = saved["graph_operations"][0]
    assert operation["replace"] and operation["config"]["columns"] == ["a"]
    assert "config" not in (saved.get("node_patch") or {}), "JSON settings are replaced, never deep-merged"


def main():
    with tempfile.TemporaryDirectory() as temporary:
        base = Path(temporary)
        for name in ("missing", "csv", "limits", "xlsx", "write", "deadline", "block"):
            (base / name).mkdir()
        missing_dependency(base / "missing")
        del sys.modules["openpyxl"]
        try:
            import openpyxl  # noqa: F401 - test dependency from requirements.txt
        except ImportError as exc:
            raise AssertionError("Install the block requirements (openpyxl==3.1.5) to run the XLSX suites.") from exc
        assert workbook.openpyxl_available()
        settings()
        csv_reading(base / "csv")
        limits_and_paths(base / "limits")
        xlsx_reading(base / "xlsx")
        writing(base / "write")
        deadlines(base / "deadline")
        block_contract(base / "block")
    print("[ok] Spreadsheet File logic, limits and error contracts")


if __name__ == "__main__":
    main()
