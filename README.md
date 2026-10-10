# Spreadsheet File

Read XLSX/CSV files into JSON rows and write JSON rows to XLSX/CSV. Use it at the
edge of a workflow, where tabular files come in or go out: import a supplier price
list, turn an approved table into a file for a colleague, export a nightly report.
Values are read and written as they are; nothing is evaluated, guessed or sent anywhere.

<!-- block-metadata:start -->
[![Block version: 0.1.0](https://img.shields.io/badge/block-0.1.0-blue)](model.json)
[![BloxSmith compatibility: 1.0.9](https://img.shields.io/badge/BloxSmith-1.0.9-brightgreen)](compatibility.json)
[![License: Apache-2.0](https://img.shields.io/badge/license-Apache--2.0-blue)](LICENSE)

Verified BloxSmith versions: **1.0.9** (bundled-block tests; see [test evidence](compatibility.json)).
<!-- block-metadata:end -->

[![Spreadsheet File — XLSX/CSV files to and from JSON rows](media/thumbnail.webp)](media/cover.png)

## Ports

| Direction | Port | Value |
| --- | --- | --- |
| Input | `file` | `read`: a file path, a JSON string or `{"path": "..."}` |
| Input | `rows` | `write`: a JSON array of objects or of arrays, or `{"rows": [...]}` |
| Output | `rows` | `read`: JSON array of objects (with a header row) or of arrays (without) |
| Output | `file` | `write`: absolute path of the completed file (`file/path`) |
| Output | `report` | JSON summary: format, sheet, columns, counts, truncation, hashes |
| Output | `error` | Only when `on_error=route`: `{"code", "message", "operation"}` |

`operation` selects which input is used. Each value that arrives on that input
triggers one activation: one file read, or one file written. A value on the other
input is a wiring error and is refused, never silently ignored. Nothing is
published until the whole file has been read, or written and made visible.

## Reading

- **Formats** come from the file extension: `.xlsx` (and `.xlsm`, see below) through
  [openpyxl](https://openpyxl.readthedocs.io/) in read-only mode; `.csv`, `.tsv` and `.txt`
  through the Python standard `csv` module with the configured `delimiter` and
  `encoding`. Nothing is sniffed: a `.tsv` file needs `delimiter=tab`. `.xls`, `.xlsb`,
  `.ods`, templates and Numbers files are refused with an explanation.
- **Worksheet**: `sheet` selects a worksheet by its exact name; otherwise `sheet_index`
  (1 = first) is used. Hidden sheets count and are reported as hidden; chart sheets
  are not worksheets and are not counted. The report lists the worksheet names.
- **Header**: `header_row` (default 1) holds the column names; rows above it are
  skipped (titles, notes). Each row becomes an object keyed by those names. Header
  text is trimmed; an empty header cell becomes `column_<n>` and a repeated name
  becomes `<name>_2`, `<name>_3`… Every generated or renamed key is listed in
  `report.header_changes`. Values to the right of the last header also get
  `column_<n>` keys, so no value is dropped. `header_row=0` emits arrays instead.
- **Values**: XLSX numbers, booleans and text keep their JSON type. Formulas are
  returned as the **cached value** saved by the spreadsheet application; they are
  never calculated. A formula saved without a cached value (some generators do
  this) reads as `null`. Error cells read as their text (`#N/A`, `#DIV/0!`). Dates
  become ISO strings: `2024-01-31` at midnight, otherwise `2024-01-31T08:30:05`;
  times `08:30:00`; durations ISO 8601 (`P1DT2H3M`). Merged cells carry their value
  in the top-left cell only. CSV values are always strings: an empty field is `""`,
  a field missing at the end of a short row is `null`.
- **Empty rows** (no value at all) are skipped and counted unless `skip_empty_rows`
  is off. A source without any value returns `[]` with `report.empty: true`.
- **Limits**: `max_rows` data rows and `max_columns` columns. With `on_limit=fail`
  (default) a larger source fails; with `truncate` the first rows/columns are kept
  and `report.truncated` says so. The JSON output is capped at 8 MiB.
- **Path boundary**: relative paths start at `allowed_root` (empty = the application
  directory). The resolved file, symbolic links included, must stay inside it and
  be a regular file. The file is read once into memory (at most `max_bytes`) and
  rejected if it changes during the read; `report.source` gives its SHA-256.

## Writing

- **Input**: objects or arrays, not a mix. For objects, `columns` sets the column
  order explicitly; keys not listed are left out and reported in
  `report.ignored_keys`. With `columns` empty, the order is the order in which keys
  first appear. For arrays, `columns` (optional) provides the header names.
  `header` writes the names as the first row.
- **Values**: numbers, booleans and `null` (an empty cell) keep their type. Nested
  objects/arrays are written as compact JSON text. JSON has no date type: ISO date
  strings stay text. XLSX text cells are always typed as literal text, so a value
  such as `=SUM(A1)` or `#N/A` can never become a formula or an error. Integers
  beyond ±9,007,199,254,740,991 are written as text so no digit is lost (XLSX stores
  numbers as 64-bit floats with 16 significant digits). XLSX text over 32,767
  characters or with control characters fails instead of being truncated.
- **CSV**: RFC 4180 records (CRLF, minimal quoting) in the selected encoding;
  `utf-8-sig` writes the byte-order mark expected by spreadsheet software. A value
  that cannot be encoded fails. `formula_guard` protects against CSV injection:
  `escape` (default) prefixes text starting with `=`, `+`, `-`, `@`, tab or carriage
  return with an apostrophe (plain numbers such as `-5` are left alone), `reject`
  fails, `off` writes it unchanged for machine consumers only.
- **Destination**: empty `output_directory` uses this node's persistent storage
  (`get_block_storage_dir()/spreadsheets/`), which survives runs and is included in
  a full project export. Otherwise the directory must already exist; relative
  paths start at the application directory. The application's configured exports
  folder is not exposed to blocks, so point `output_directory` at it explicitly if
  that is where files belong.
- **File name**: `file_name` (default `rows-{date}-{unique}`) accepts `{date}`
  (`YYYY-MM-DD`), `{time}` (`HHMMSS`), both UTC, and `{unique}` (12 random hex
  digits). The format extension is added when missing; another spreadsheet
  extension is refused.
- **Atomic publication**: the file is written to a private temporary name in the
  destination, flushed to disk, then published. Without `overwrite` it is hard-linked
  to the final name, which fails if that name exists: an existing file is never
  replaced and concurrent writers cannot both win. With `overwrite`, an existing
  regular file is replaced atomically (never a directory or a link). Files are
  created with mode 0600. Readers never see a partial file.

## Settings

| Setting | Default | Bounds / values |
| --- | --- | --- |
| `operation` | `read` | `read`, `write` |
| `on_error` | `fail` | `fail`, `route` |
| `allowed_root` | empty | read boundary directory |
| `sheet` / `sheet_index` | empty / 1 | 31 characters / 1–1000 |
| `header_row` | 1 | 0–1000 (0 = arrays) |
| `skip_empty_rows` | true | |
| `on_limit` | `fail` | `fail`, `truncate` |
| `allow_macro_workbooks` | false | read `.xlsm` / macro-carrying data |
| `format` | `xlsx` | `xlsx`, `csv` (write) |
| `output_directory` | empty | existing directory, empty = node storage |
| `file_name` | `rows-{date}-{unique}` | 1–100 characters, `{date}` `{time}` `{unique}` |
| `overwrite` | false | |
| `columns` | `[]` | distinct names, at most `max_columns` |
| `header` | true | |
| `formula_guard` | `escape` | `escape`, `reject`, `off` (CSV write) |
| `delimiter` | `comma` | `comma`, `semicolon`, `tab`, `pipe` |
| `encoding` | `utf-8-sig` | `utf-8-sig`, `utf-8`, `latin-1`, `cp1252` |
| `max_rows` | 10,000 | 1–100,000 |
| `max_columns` | 256 | 1–1,024 |
| `max_bytes` | 20 MiB | 1 KiB–50 MiB (source, `rows` input and generated file) |
| `timeout_sec` | 60 | 1–240 |

Settings are validated on Apply and again before a Run starts: an invalid setting,
or `write` + `xlsx` without openpyxl installed, fails Run preparation before any file
is touched.

## Errors and routing

With `on_error=fail` (default) an error fails the node with a translated message and
publishes nothing; the run reports the failure. With `on_error=route` the node
completes and publishes only `error`, so the graph can handle bad files itself.
Wire that port before choosing `route`: an unwired error port hides failures.
Stable codes: `invalid_input`, `wrong_input`, `invalid_path`, `not_found`,
`unsupported_format`, `macro_workbook`, `too_large`, `limit_exceeded`, `malformed_file`,
`source_changed`, `invalid_value`, `missing_dependency`, `target_exists`,
`storage_unavailable`, `write_failed`, `timeout`, `internal`. Messages never contain cell
contents or paths. Stop (or live inhibition) ends the activation without output;
it is not routed, the node reports skipped and the run reports cancelled.

## Example

Read the `Orders` sheet of `imports/orders.xlsx` whose column names are on row 3,
below a title: set `sheet=Orders`, `header_row=3`, and send `imports/orders.xlsx`
to `file`. The `rows` output looks like:

```json
[{"order": 1042, "customer": "ACME", "date": "2024-01-31", "total": 125.5, "paid": true}]
```

To export rows, add a second node with `operation=write`, `format=csv`,
`columns=["order","total"]`, `delimiter=semicolon`, `file_name=orders-{date}`, and
wire the rows to its `rows` input. Its `file` output carries the path of the CSV.

## Spreadsheet File, Table Tools and Document Reader

- **Spreadsheet File** is the file boundary for tables: XLSX and CSV in and out,
  worksheet selection, header handling, dates and cached formula values, and atomic
  file writing. It does not transform rows.
- **[Table Tools](https://github.com/HackInvent/bloxmith-table-tools)** transforms
  rows in memory: types, projection, comparison and joins. It reads CSV with a
  header on the first row only, never reads XLSX and never writes files. Chain them:
  Spreadsheet File (read) → Table Tools → Spreadsheet File (write).
- **[Document Reader](https://github.com/HackInvent/bloxmith-document-reader)** extracts
  prose from PDF, DOCX and text documents with page/paragraph references. It does not
  read spreadsheets or keep a tabular structure.

## Limits and warnings

- Macro-enabled `.xlsm` files, or any workbook carrying a macro project, are refused
  unless `allow_macro_workbooks` is on; even then macros are never run or kept.
- Password-protected workbooks, legacy `.xls` and encrypted ZIP entries are refused.
- Links to other workbooks are never loaded or followed; their cells keep the
  cached values saved in the file and `report.external_links` counts the link parts.
- Before openpyxl opens a workbook, the package is checked: at most 10,000 ZIP entries,
  1 GiB declared expansion in total, 128 MiB per non-worksheet part, compression ratio
  at most 1,000 for large parts, and no XML document type declaration (which blocks
  entity-expansion attacks). The worksheet dimension stored in the file is ignored.
- CSV fields longer than 131,072 characters fail (Python `csv` limit).
- Cancellation and the deadline are checked between rows. Opening a workbook
  (shared strings and styles) is not interruptible; very large shared-string
  tables delay Stop until they are loaded.
- Everything is held in memory: up to `max_bytes` of source plus the parsed rows.
  The limits bound it; they are not an OS memory quota.
- An empty string written to XLSX becomes an empty cell and reads back as `null`.

## Dependencies

- XLSX needs **openpyxl `>=3.1,<4`** (MIT license) in the Python environment used by
  the Studio and its package hosts; install it from `requirements.txt` (tests pin
  3.1.5). It is imported only when an XLSX file is read or written. Without it, CSV
  works, reading an XLSX file fails with `missing_dependency`, and an XLSX writer
  fails Run preparation. openpyxl also uses `defusedxml` automatically when present.
- Writing needs a POSIX filesystem supporting hard links and directory descriptors.

## Security notes

The read boundary and the path checks are application-level rules, not an OS
sandbox against another local process changing files. Written files hold whatever
the rows contain: choose a destination with appropriate access rights. The block
uses no network, no credential and no external program; it never opens a URL found
in a workbook. Outputs, reports and logs contain no secret; error documents carry
no cell contents or paths.

## Execution and UI

One Shot and Active Runtime behave identically, in bundled, managed and linked
installations. The modal and inspector share English/French settings that show only
the fields of the selected operation, explicit Apply/Cancel, associated labels and a
collapsed advanced section. Drafts stay local until Apply; restart a loaded Run to use
changed settings. The node card shows the operation and format only.

## Framing decisions

The functional framing left these choices open; they were settled as follows:

1. One block with an explicit `operation` instead of two blocks, mirroring the
   Extract From File / Convert to File pair; the unused input is refused.
2. Errors fail the node by default; routing to `error` is an opt-in, so an unwired
   error port cannot hide failures.
3. Default destination: the node's persistent storage, because the public block API
   exposes no project exports folder; any existing directory can be configured.
4. Never overwrite by default; unique names via `{unique}`; atomic hard-link publication.
5. XLSX cells are written as literal text; CSV gets an OWASP formula guard.
6. Formulas are read as cached values only; dates as ISO strings; nothing is inferred
   from CSV text.
7. Header problems are repaired deterministically and reported, never fatal, except
   an empty or missing header row.
8. Cancellation publishes nothing and is reported as skipped, so the run reads cancelled.

## Tests and qualification

Three package suites run against a disposable BloxSmith 1.0.9 application:

- `F5.01`: direct logic with real openpyxl 3.1.5 and the standard library — every
  setting bound, CSV dialects and encodings, header repair, rich workbooks (dates,
  durations, hidden and chart sheets, cached formulas, error cells), hostile packages
  (DTD in UTF-8/UTF-16, forged ZIP sizes and flags, forged dimensions, out-of-range
  rows, macro parts, encrypted/legacy signatures), literal-text XLSX cells, the CSV
  formula guard, atomic publication with concurrent writers, timeout, cancellation,
  error routing, and the whole package without openpyxl.
- `F5.02`: real graphs in One Shot and Active Runtime for bundled, managed and linked
  installations — XLSX written to node storage and read back, CSV read, routed
  `not_found`, an unrouted failure, and Stop during a long read (bundled installation).
- `F8.01`: managed and linked browser surfaces in English and French at four widths,
  operation-specific fields, draft discard, Apply/persistence and the inspector.

All files are synthetic and created in temporary test storage. Spreadsheets produced
by Excel, LibreOffice or Google Sheets were not part of this qualification; the
hand-built packages follow the same Office Open XML structures. No real project,
network service or credential is used.
