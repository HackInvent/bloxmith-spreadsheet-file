"""FB1/FB2/FB4/FB5: real graph runs in both runtimes, bundled/managed/linked packages, routing, failure and Stop."""

from pathlib import Path
import json
import sys
import time

ROOT = Path(__file__).resolve().parents[3]
sys.path[:0] = [str(ROOT), str(ROOT / "tests")]
from blocs.spreadsheet_file.block import SpreadsheetFileBlock
from block_test_packages import install_test_package, surface_payload
from ui_smoke_common import (create_project_api, create_run_api, data_edge, display_node, get_run_api, graph_payload,
                             isolated_server, stop_run_api, text_node, wait_for_run_terminal)

ROWS = [{"id": 1, "name": "Ada", "score": 9.5, "active": True},
        {"id": 2, "name": "=cmd|' /C calc'!A0", "score": None, "active": False}]


def sheet_node(block, ident, x, y, version, **config):
    """Create one block node with explicit settings and, for installed packages, the exact release."""
    node = block.build_node_payload(node_id=ident, position={"x": x, "y": y}, config_overrides=config)
    if version:
        node["block_version"] = version
    return node


def value(run, node, port):
    """Return one published output value, or an empty string."""
    return (run.get("output_values", {}).get(f"{node}:{port}") or {}).get("value") or ""


def run_graph(server, graph, mode, timeout=60):
    """Run a project-scoped graph so node storage is available, and wait for its terminal state."""
    project = create_project_api(server, document=graph)["project"]
    created = create_run_api(server, graph, project_id=project["graph_id"], runtime_mode=mode)
    return created["run_id"], wait_for_run_terminal(server, created["run_id"], timeout_sec=timeout)


def round_trip(server, block, mode, version):
    """FB1/FB2/FB4: write XLSX into node storage, read it back, read CSV and route a missing file."""
    fixtures = server.root_dir / "fixtures"
    fixtures.mkdir()
    (fixtures / "people.csv").write_text("id;city\n7;Lyon\n8;Nantes\n", encoding="utf-8")
    nodes = [text_node("rows", "Rows", json.dumps(ROWS), 60, 80),
             sheet_node(block, "writer", 320, 80, version, operation="write", file_name="team-{unique}"),
             sheet_node(block, "reader", 580, 80, version),
             display_node("table", "Table", 840, 80),
             text_node("csv", "CSV path", "fixtures/people.csv", 60, 300),
             sheet_node(block, "csvreader", 320, 300, version, delimiter="semicolon"),
             display_node("cities", "Cities", 580, 300),
             text_node("missing", "Missing path", "fixtures/missing.xlsx", 60, 520),
             sheet_node(block, "guarded", 320, 520, version, on_error="route"),
             display_node("issue", "Issue", 580, 520)]
    edges = [data_edge("rows-writer", "rows", 1, "writer", 2), data_edge("writer-reader", "writer", 2, "reader", 1),
             data_edge("reader-table", "reader", 1, "table", 1), data_edge("csv-reader", "csv", 1, "csvreader", 1),
             data_edge("csv-cities", "csvreader", 1, "cities", 1), data_edge("missing-guarded", "missing", 1, "guarded", 1),
             data_edge("guarded-issue", "guarded", 4, "issue", 1)]
    _, done = run_graph(server, graph_payload("Spreadsheet round trip", nodes, edges), mode)
    assert done["status"] == "success", (mode, version, done.get("logs"))
    path = Path(value(done, "writer", 2))
    assert path.is_file() and path.suffix == ".xlsx" and path.parent.name == "spreadsheets"
    assert path.is_relative_to(server.root_dir / "user" / "projects") and "blocs" in path.parts
    report = json.loads(value(done, "writer", 3))
    assert report["rows"] == 2 and report["path"] == str(path) and report["format"] == "xlsx"
    assert json.loads(value(done, "reader", 1)) == ROWS, value(done, "reader", 1)
    assert json.loads(value(done, "reader", 3))["source"]["sha256"] == report["sha256"]
    assert json.loads(value(done, "csvreader", 1)) == [{"id": "7", "city": "Lyon"}, {"id": "8", "city": "Nantes"}]
    issue = json.loads(value(done, "guarded", 4))
    assert issue["code"] == "not_found" and issue["operation"] == "read" and "fixtures" not in issue["message"]
    assert not value(done, "guarded", 1) and not value(done, "guarded", 3)
    assert done["node_statuses"]["issue"] == "success" and done["node_statuses"]["guarded"] == "success"


def failure_and_stop(server, block, mode):
    """FB4/FB5: an unrouted error fails the node; Stop interrupts a long read without publishing rows."""
    nodes = [text_node("bad", "Bad path", "fixtures/absent.csv", 60, 80), sheet_node(block, "strict", 320, 80, None),
             display_node("never", "Never", 580, 80)]
    edges = [data_edge("bad-strict", "bad", 1, "strict", 1), data_edge("strict-never", "strict", 1, "never", 1)]
    _, failed = run_graph(server, graph_payload("Strict failure", nodes, edges), mode)
    assert failed["node_statuses"]["strict"] == "failed", failed.get("node_statuses")
    assert failed["status"] == "failed", (mode, failed.get("status"))
    assert not value(failed, "strict", 1) and not value(failed, "strict", 4)
    assert "does not exist" in json.dumps(failed.get("logs", []))
    (server.root_dir / "fixtures" / "slow.csv").write_bytes(b"\n" * 12000000)
    nodes = [text_node("slow", "Slow path", "fixtures/slow.csv", 60, 80),
             sheet_node(block, "longread", 320, 80, None, max_bytes=52428800, timeout_sec=240),
             display_node("late", "Late", 580, 80)]
    edges = [data_edge("slow-read", "slow", 1, "longread", 1), data_edge("read-late", "longread", 1, "late", 1)]
    graph = graph_payload("Long read", nodes, edges)
    project = create_project_api(server, document=graph)["project"]
    run_id = create_run_api(server, graph, project_id=project["graph_id"], runtime_mode=mode)["run_id"]
    deadline = time.monotonic() + 20
    while time.monotonic() < deadline:
        state = get_run_api(server, run_id)
        if state.get("node_statuses", {}).get("longread") in {"running", "processing"}:
            break
        assert state.get("status") not in {"success", "failed", "cancelled"}, state.get("logs")
        time.sleep(0.05)
    else:
        raise AssertionError("The long read never started")
    time.sleep(0.5)
    stopped_at = time.monotonic()
    stop_run_api(server, run_id)
    stopped = wait_for_run_terminal(server, run_id, timeout_sec=15)
    assert time.monotonic() - stopped_at < 8, "Stop must interrupt the read promptly"
    assert stopped["status"] == "cancelled", (stopped["status"], stopped.get("logs"))
    assert not value(stopped, "longread", 1) and not value(stopped, "longread", 3)
    (server.root_dir / "fixtures" / "slow.csv").unlink()


def main():
    block = SpreadsheetFileBlock()
    for mode in ("centralized", "zeromq_active"):
        for origin in (None, "managed", "linked"):
            with isolated_server() as server:
                version = None
                if origin:
                    model = install_test_package(server, "spreadsheet_file", origin=origin)
                    version = model["version"]
                    probe = sheet_node(block, "probe", 0, 0, version, operation="write")
                    for surface in ("modal", "inspector_panel", "node_card"):
                        html = surface_payload(server, model, probe, surface=surface)["html"]
                        assert "{{" not in html and "__title__" not in html
                        if surface != "node_card":
                            assert 'data-block-config-field="operation"' in html and "data-owned-apply" in html
                            assert 'data-spreadsheet-operation="write"' in html and 'data-spreadsheet-operation="read" hidden' in html
                        else:
                            assert "write · xlsx" in html
                round_trip(server, block, mode, version)
                if origin is None:
                    failure_and_stop(server, block, mode)
            print(f"[ok] Spreadsheet File runtime {mode} {origin or 'bundled'}", flush=True)


if __name__ == "__main__":
    main()
