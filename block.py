"""Read XLSX/CSV files into JSON rows and write JSON rows to XLSX/CSV, with explicit bounds."""

import json

from bloxsmith_app.block_api import (APPLICATION_JSON, FILE_PATH, BlockDefinition, BlockRuntimeOutput,
                                     BlockRuntimePreparation, BlockRuntimeResult)
from . import logic, ui, workbook


# FB1 - Read an XLSX worksheet or a CSV file into JSON rows: header row, sheet choice, cached formula values, ISO dates.
# FB2 - Write JSON object/array rows to XLSX or CSV in an explicit column order, as literal text cells, atomically.
# FB3 - Bound paths, sizes, rows, columns, output, ZIP/XML content, macros, time and cancellation before any output.
# FB4 - Fail explicitly or route content-free errors to the error port; optional openpyxl fails cleanly when absent.
# FB5 - Keep identical contracts in both runtime modes and in bundled, managed and linked installations.
# FB6 - Own translated responsive settings with operation-specific fields, draft/Apply and release-scoped assets.
class SpreadsheetFileBlock(BlockDefinition):
    """File I/O boundary for tabular data; transformations belong to downstream blocks."""

    kind = "spreadsheet_file"

    def ui_assets(self, surface="modal"):
        """Use the manifest's asset list for bundled surfaces as well as hosted releases."""
        return list(self.model.get("ui_assets", {}).get(surface, []))

    def prepare_runtime(self, context):
        """Refuse invalid settings, or an XLSX writer without openpyxl, before the Run starts."""
        try:
            config = logic.configuration(context.config)
        except logic.SpreadsheetError as exc:
            raise logic.SpreadsheetError(exc.code, "Invalid Spreadsheet File settings: " + exc.message) from None
        if config["operation"] == "write" and config["format"] == "xlsx" and not workbook.openpyxl_available():
            raise logic.SpreadsheetError("missing_dependency", workbook.DEPENDENCY)
        return BlockRuntimePreparation()

    def execute_runtime(self, context):
        """One activation reads one file or writes one file; failures never publish partial data."""
        operation, route = "read", False
        try:
            config = logic.configuration(context.config)
            operation, route = config["operation"], config["on_error"] == "route"
            value = self._input(context, "file" if operation == "read" else "rows")
            if value is None:
                return BlockRuntimeResult(status="skipped", outputs=[], logs=["[spreadsheet-file] No input value; nothing to do."])
            cancel = context.services.get("cancel_requested", lambda: False)
            if operation == "read":
                values = logic.read(value, config, context.root_dir, cancel)
            else:
                values = logic.write(value, config, context.services.get("get_block_storage_dir"), context.root_dir, cancel)
        except logic.SpreadsheetError as exc:
            if exc.code == "cancelled":
                # Stop or live inhibition: publish nothing and let the framework record the cancellation.
                return BlockRuntimeResult(status="skipped", outputs=[], last_message=exc.message,
                                          metadata={self.kind: {"operation": operation, "error_code": exc.code}},
                                          logs=["[spreadsheet-file] Cancelled before completion; nothing was published."])
            return self._failure(context, exc, operation, route and exc.code != "invalid_settings")
        except (ValueError, TypeError, OSError, RecursionError, RuntimeError, MemoryError):
            error = logic.SpreadsheetError("internal", "The spreadsheet operation could not complete safely.")
            return self._failure(context, error, operation, route)
        report = values["report"]
        encoded = {"rows": values.get("rows"), "file": values.get("file"),
                   "report": json.dumps(report, ensure_ascii=False, allow_nan=False)}
        types = {"rows": APPLICATION_JSON, "file": FILE_PATH, "report": APPLICATION_JSON}
        outputs = [BlockRuntimeOutput(port_id=int(port.id), port_name=port.name, value=encoded[port.name],
                                      content_type=types[port.name])
                   for port in context.output_ports if encoded.get(port.name) is not None]
        verb = "Read" if operation == "read" else "Wrote"
        return BlockRuntimeResult(status="success", outputs=outputs, content_type=APPLICATION_JSON,
                                  last_message=encoded["report"][:2000], worker_received=str(report["rows"]),
                                  metadata={self.kind: {"operation": operation, "format": report["format"], "rows": report["rows"]}},
                                  logs=[f"[spreadsheet-file] {verb} {report['rows']} row(s) as {report['format']}."])

    @staticmethod
    def _input(context, port):
        """Return the value of the input used by the configured operation, refusing the other one."""
        ident = "1" if port == "file" else "2"
        if context.input_events:
            names = {int(item.id): str(item.name) for item in context.input_ports}
            relevant, others = [], []
            for event in context.input_events:
                (relevant if names.get(int(event.input_port_id)) == port else others).append(event)
            if others:
                other = names.get(int(others[0].input_port_id), "other")
                raise logic.SpreadsheetError("wrong_input", f"The {other} input is not used when operation is "
                                             f"{'read' if port == 'file' else 'write'}; check the wiring.")
            if len(relevant) > 1:
                raise logic.SpreadsheetError("invalid_input", "Send one value per activation.")
            return relevant[0].value if relevant else None
        if not context.has_input_value(port, ident):
            return None
        return context.input_value(port, ident)

    def _failure(self, context, error, operation, route):
        """Either fail the node, or publish a content-free error document when routing was chosen."""
        if route:
            document = json.dumps(error.envelope(operation), ensure_ascii=False)
            outputs = [BlockRuntimeOutput(port_id=int(port.id), port_name=port.name, value=document,
                                          content_type=APPLICATION_JSON)
                       for port in context.output_ports if port.name == "error"]
            return BlockRuntimeResult(status="success", outputs=outputs, content_type=APPLICATION_JSON,
                                      last_message=document, metadata={self.kind: {"operation": operation, "error_code": error.code}},
                                      logs=[f"[spreadsheet-file] Routed {error.code} to the error port."])
        message = self.translate("block.spreadsheet_file.error", {"detail": error.message},
                                 fallback="Spreadsheet operation failed: {detail}")
        return BlockRuntimeResult(status="failed", outputs=[], error=message, exit_code=1,
                                  metadata={self.kind: {"operation": operation, "error_code": error.code}},
                                  logs=[f"[spreadsheet-file-error] {error.code}: {error.message}"])

    def handle_ui_action(self, *, node, action, values, payload=None):
        """Refuse invalid settings before the framework persists a form patch."""
        if action in {"save_settings", "modal_update_fields", "inspector_update_fields"}:
            patch = (values or {}).get("node_patch") or {}
            if "config" in patch:
                try:
                    logic.configuration({**self.default_config(), **(node.get("config") or {}), **patch["config"]})
                except (ValueError, TypeError) as exc:
                    detail = exc.message if isinstance(exc, logic.SpreadsheetError) else "Invalid settings."
                    return {"error": self.translate("block.spreadsheet_file.error", {"detail": detail},
                                                    fallback="Spreadsheet operation failed: {detail}")}
        result = super().handle_ui_action(node=node, action="modal_update_fields" if action == "save_settings" else action,
                                          values=values, payload=payload)
        return ui.replace_config(result, node)

    def render_modal(self, *, node, payload=None):
        """Render the owned settings form inside the opaque modal contract."""
        return ui.modal(self, node, payload)

    def render_inspector_panel(self, *, node, payload=None):
        """Keep the same fields and controls usable in the narrow inspector."""
        return ui.inspector(self, node, payload)

    def render_node_card(self, *, node, payload=None):
        """Show the operation and format only, never file names or cell contents."""
        config = {**self.default_config(), **(node.get("config") or {})}
        preview = "read · xlsx/csv" if config.get("operation") != "write" else f"write · {config.get('format', 'xlsx')}"
        return ui.card(self, node, preview)
