"""Package-owned forms, sharing only the framework's public binding helpers."""

from html import escape
import json
from bloxsmith_app.block_api import render_inspector_template, render_node_card_template, render_path_browser_control


def replace_config(result, node):
    """Replace edited JSON settings exactly; generic graph patches otherwise deep-merge objects."""
    patch = result.get("node_patch") if isinstance(result, dict) else None
    if not isinstance(patch, dict) or not isinstance(patch.get("config"), dict):
        return result
    settings = {**(node.get("config") or {}), **patch.pop("config")}
    result.setdefault("graph_operations", []).append({"op": "update_node_config",
        "node_id": str(node.get("id") or ""), "config": settings, "replace": True})
    return result


def control(block, spec, value, ident):
    """Render one bound control with its label and translated help; values are always escaped."""
    key, kind = spec["key"], spec.get("type", "text")
    attrs = f'data-block-config-field="{key}"'
    if kind in {"json", "integer", "boolean"}:
        attrs += f' data-block-value-type="{kind}"'
    hint_key = f"block.{block.kind}.help_{key}"
    hint = escape(block.translate(hint_key, fallback=spec["help"]))
    help_html = f'<p id="{ident}-hint" class="field-hint" data-i18n="{hint_key}">{hint}</p>'
    if kind == "directory":
        return render_path_browser_control(input_id=ident, label=key, value=str(value), placeholder="",
                                           input_attrs=attrs + f' aria-describedby="{ident}-hint"',
                                           select_mode="directory") + help_html
    common = f'id="{ident}" {attrs} aria-describedby="{ident}-hint"'
    if kind == "select":
        options = "".join(f'<option value="{escape(option, quote=True)}"{" selected" if option == value else ""}>'
                          f'{escape(option)}</option>' for option in spec["options"])
        field = f'<select {common}>{options}</select>'
    elif kind == "json":
        text = value if isinstance(value, str) else json.dumps(value, ensure_ascii=False)
        field = f'<textarea {common} rows="{spec.get("rows", 2)}" spellcheck="false">{escape(text)}</textarea>'
    elif kind == "boolean":
        field = f'<input {common} type="checkbox"{" checked" if value else ""}>'
    elif kind == "integer":
        field = (f'<input {common} type="number" min="{spec["min"]}" max="{spec["max"]}" step="1" '
                 f'value="{escape(str(value), quote=True)}">')
    else:
        field = f'<input {common} type="text" autocomplete="off" spellcheck="false" value="{escape(str(value), quote=True)}">'
    return f'<div class="field-group"><label for="{ident}">{key}</label>{field}{help_html}</div>'


def fields(block, node, surface):
    """Render trusted field definitions; operation-specific fields are toggled by settings.js."""
    specs = json.loads((block.directory / "fields.json").read_text(encoding="utf-8"))
    config = {**block.default_config(), **(node.get("config") or {})}
    current = config.get("operation") if config.get("operation") in {"read", "write"} else "read"
    groups = {"main": [], "advanced": []}
    prefix = escape(f"{block.kind}-{node.get('id', '')}-{surface}", quote=True)
    for spec in specs:
        rendered = control(block, spec, config.get(spec["key"], ""), f"{prefix}-{spec['key']}")
        operation = spec.get("operation")
        if operation:
            hidden = "" if operation == current else " hidden"
            rendered = f'<div class="owned-field" data-spreadsheet-operation="{operation}"{hidden}>{rendered}</div>'
        groups["advanced" if spec.get("advanced") else "main"].append(rendered)
    key = f"block.{block.kind}.advanced"
    return ('<div class="owned-fields">' + "".join(groups["main"]) + "</div>"
            f'<details class="owned-advanced"><summary data-i18n="{key}">'
            f'{escape(block.translate(key, fallback="Advanced settings"))}</summary>'
            f'<div class="owned-fields">{"".join(groups["advanced"])}</div></details>')


def modal(block, node, payload):
    """Use the generic draft host; refreshing runtime must not replace pending settings."""
    template = (block.directory / "block_modal.html").read_text(encoding="utf-8")
    html = block._render_generic_modal_template(template=template, node=node, payload=payload or {})
    html = html.replace("{{ settings }}", fields(block, node, "modal"))
    return {"html": html, "context": {"node_id": str(node.get("id", "")), "node_kind": block.kind}}


def inspector(block, node, payload):
    """Render fixed settings with the standard Ports tab and sticky Apply control."""
    template = (block.directory / "inspector_panel.html").read_text(encoding="utf-8")
    html = render_inspector_template(template=template, node={**node, "kind": block.kind, "type": block.kind},
                                     payload=payload, replacements={"settings": fields(block, node, "inspector")})
    return {"html": html, "context": {"node_id": str(node.get("id", "")), "full_panel": True}}


def card(block, node, preview):
    """Keep name, purpose and current mode inside the existing card hierarchy."""
    return render_node_card_template(block=block, node=node, node_classes=[block.kind + "-node"],
                                     replacements={"title": node.get("title") or block.default_title(),
                                                   "preview": preview, "purpose": block.description()})
