"""FB6: real release UI, operation-specific fields, responsive geometry, keyboard, drafts, validation and language."""

from pathlib import Path
import json
import sys

ROOT = Path(__file__).resolve().parents[3]
sys.path[:0] = [str(ROOT), str(ROOT / "tests")]
from playwright.sync_api import sync_playwright, expect
from block_test_artifacts import artifact_path
from block_test_packages import install_test_package
from ui_smoke_common import isolated_server, graph_payload, create_project_api, project_editor_url
from blocs.spreadsheet_file.block import SpreadsheetFileBlock


def geometry(modal):
    """Require an opaque panel with reachable actions, an internally scrollable body and labeled controls."""
    measured = modal.evaluate("""el => {
        const r=el.getBoundingClientRect(), body=el.querySelector('.owned-body');
        const actions=[...el.querySelectorAll('[data-close-block-modal], [data-owned-apply]')];
        return {inside:r.left>=-1&&r.right<=innerWidth+1&&r.top>=-1&&r.bottom<=innerHeight+1,
          overflow:el.scrollWidth>el.clientWidth+2||body.scrollWidth>body.clientWidth+2,
          controls:actions.every(a=>{const b=a.getBoundingClientRect();return b.top>=0&&b.bottom<=innerHeight+1}),
          background:getComputedStyle(el).backgroundColor,
          unlabeled:[...el.querySelectorAll('input,select,textarea')].filter(c=>!c.labels?.length&&!c.getAttribute('aria-label')&&!c.getAttribute('aria-labelledby')).length};
    }""")
    assert measured["inside"] and measured["controls"] and not measured["overflow"], measured
    assert not measured["unlabeled"], measured
    assert measured["background"] not in {"transparent", "rgba(0, 0, 0, 0)"}, measured


def visible_fields(surface):
    """Config keys of the fields currently displayed for the selected operation."""
    return set(surface.evaluate("""el => [...el.querySelectorAll('[data-block-config-field]')]
        .filter(f => f.closest('[data-spreadsheet-operation]')?.hidden !== true).map(f => f.dataset.blockConfigField)"""))


def open_modal(page, card):
    """Open the block modal; on a loaded host a first double-click can land before the editor is bound."""
    modal = page.locator('.owned-modal')
    for attempt in range(2):
        card.locator('h3').dblclick()
        try:
            expect(modal).to_be_visible(timeout=20000)
            return modal
        except AssertionError:
            if attempt:
                raise
    return modal


def field(surface, key):
    """Locate one bound settings control."""
    return surface.locator(f'[data-block-config-field="{key}"]')


def main():
    """Exercise current package files in real managed and linked browser hosts."""
    # Managed hosts render surfaces in spawned supervised processes: allow for a busy test host.
    expect.set_options(timeout=15000)
    for origin in ("managed", "linked"):
        with isolated_server() as server, sync_playwright() as playwright:
            model = install_test_package(server, "spreadsheet_file", origin=origin)
            candidate = SpreadsheetFileBlock().build_node_payload(node_id="ui-case", position={"x": 160, "y": 160})
            candidate["block_version"] = model["version"]
            project = create_project_api(server, document=graph_payload("Owned form", [candidate], []))["project"]
            url = project_editor_url(server.base_url, project["project_id"], workspace_project_id=project["workspace_project_id"])
            browser = playwright.chromium.launch(headless=True)
            try:
                page = browser.new_page(viewport={"width": 1440, "height": 900})
                page.add_init_script("window.localStorage.setItem('bloxsmith.inspectorPinned','true')")
                errors = []
                page.on("pageerror", lambda error: errors.append(str(error)))
                page.goto(url)
                card = page.locator('.canvas-node[data-node-id="ui-case"]')
                expect(card.locator('h3')).to_be_visible()
                expect(card.locator('.node-data-preview')).to_have_text("read · xlsx/csv")
                modal = open_modal(page, card)
                apply = modal.locator('[data-owned-apply]')
                expect(apply).to_be_disabled()
                shown = visible_fields(modal)
                assert {"operation", "allowed_root", "sheet", "sheet_index", "header_row", "on_error"} <= shown, shown
                assert not shown & {"file_name", "format", "output_directory", "overwrite", "formula_guard"}, shown
                field(modal, "operation").select_option("write")
                shown = visible_fields(modal)
                assert {"file_name", "format", "output_directory", "columns", "formula_guard", "sheet"} <= shown, shown
                assert not shown & {"sheet_index", "header_row", "allowed_root", "on_limit"}, shown
                expect(apply).to_be_enabled()
                modal.locator('[data-close-block-modal]').last.click()
                open_modal(page, card)
                expect(field(modal, "operation")).to_have_value("read")
                assert "file_name" not in visible_fields(modal), "Cancel must discard the draft and its field layout"
                field(modal, "operation").select_option("write")
                field(modal, "file_name").fill("../escape")
                with page.expect_response(lambda r: r.url.endswith('/ui-action') and r.request.method == 'POST') as refused:
                    apply.click()
                assert refused.value.json().get("error"), refused.value.json()
                expect(modal.locator('.owned-form-status')).to_contain_text("file_name")
                field(modal, "file_name").fill("export-{date}")
                field(modal, "format").select_option("csv")
                field(modal, "columns").fill('["id", "name"]')
                with page.expect_response(lambda r: r.url.endswith('/ui-action') and r.request.method == 'POST') as saved:
                    apply.click()
                assert not saved.value.json().get("error"), saved.value.json()
                if modal.is_visible():
                    modal.locator('[data-close-block-modal]').first.click()
                page.reload()
                expect(card.locator('.node-data-preview')).to_have_text("write · csv")
                open_modal(page, card)
                expect(field(modal, "operation")).to_have_value("write")
                expect(field(modal, "file_name")).to_have_value("export-{date}")
                assert json.loads(field(modal, "columns").input_value()) == ["id", "name"]
                for width, height in ((1440, 900), (800, 700), (390, 740), (320, 568)):
                    page.set_viewport_size({"width": width, "height": height})
                    page.wait_for_timeout(100)
                    geometry(modal)
                    if width in {1440, 320}:
                        page.screenshot(path=artifact_path(f"spreadsheet_file-{origin}-{width}.png"))
                summary = modal.locator('.owned-advanced').first.locator('summary')
                summary.scroll_into_view_if_needed()
                summary.focus()
                summary.press('Enter')
                expect(modal.locator('.owned-advanced').first).to_have_attribute('open', '')
                expect(field(modal, "formula_guard")).to_be_visible()
                expect(field(modal, "on_limit")).to_be_hidden()
                geometry(modal)
                page.screenshot(path=artifact_path(f"spreadsheet_file-{origin}-advanced.png"))
                page.set_viewport_size({"width": 1440, "height": 900})
                modal.locator('[data-close-block-modal]').first.click()
                card.click(position={"x": 20, "y": 20})
                inspector = page.locator('[data-properties-surface="inspector"][data-node-id="ui-case"]:visible')
                expect(inspector).to_be_visible()
                assert inspector.evaluate('el => el.scrollWidth <= el.clientWidth+2')
                assert "file_name" in visible_fields(inspector) and "sheet_index" not in visible_fields(inspector)
                field(inspector, "operation").select_option("read")
                assert "sheet_index" in visible_fields(inspector) and "file_name" not in visible_fields(inspector)
                inspector.locator('[data-owned-inspector-tab="ports"]').click()
                expect(inspector.locator('[data-inspector-panel-tab="ports"]')).to_be_visible()
                expect(inspector.locator('[data-owned-inspector-tab="ports"]')).to_have_attribute("aria-selected", "true")
                inspector.locator('[data-owned-inspector-tab="ports"]').press("ArrowLeft")
                expect(inspector.locator('[data-inspector-panel-tab="general"]')).to_be_visible()
                expect(field(inspector, "operation")).to_have_value("read")
                assert "sheet_index" in visible_fields(inspector), "Switching views must keep the draft"
                page.screenshot(path=artifact_path(f"spreadsheet_file-{origin}-inspector.png"))
                page.goto(server.base_url + "/")
                page.locator("#homeApplicationSettingsButton").click()
                page.locator("#applicationLanguageSelect").select_option("fr")
                page.wait_for_function("window.CWMessages.getLanguage() === 'fr'")
                page.goto(url)
                open_modal(page, card)
                expect(modal.locator('.owned-advanced').first.locator('summary')).to_have_text("Réglages avancés")
                expect(modal.locator('.panel-kicker')).to_have_text("Fichiers tableur")
                expect(modal.locator('[data-i18n="block.spreadsheet_file.help_overwrite"]')).to_contain_text("Remplace un fichier")
                geometry(modal)
                page.screenshot(path=artifact_path(f"spreadsheet_file-{origin}-fr.png"))
                assert not errors, errors
                assert page.evaluate("!window.CWBlockUiBlocks?.spreadsheet_file")
            finally:
                browser.close()
        print("[ok] spreadsheet_file " + origin + " UI", flush=True)


if __name__ == "__main__":
    main()
