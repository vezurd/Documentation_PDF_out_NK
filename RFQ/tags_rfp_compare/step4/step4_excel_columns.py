"""Named Step4 Excel column templates (order, width, groups, visibility).

The built-in template id ``default`` is always derived from
``OUTPUT_COLUMNS_CONFIG`` and is never stored in user JSON.
"""

from __future__ import annotations

import copy
import uuid
from typing import Any

from RFQ.tags_rfp_compare.rfp_tags_utils import load_config, save_config

DEFAULT_TEMPLATE_ID = "default"
MIN_COLUMN_WIDTH = 4
MAX_COLUMN_WIDTH = 80
_STATE_KEY = "step4_excel_columns"


class Step4ColumnTemplateError(ValueError):
    """Invalid template id or mutation of the built-in default template."""


def _writer_module():
    from RFQ.tags_rfp_compare.step4 import step4_6_save_match_result_to_excel as mod

    return mod


def _clamp_width(value: object, fallback: int) -> int:
    try:
        width = int(value)
    except (TypeError, ValueError):
        width = int(fallback)
    return max(MIN_COLUMN_WIDTH, min(MAX_COLUMN_WIDTH, width))


def _as_bool(value: object, default: bool = False) -> bool:
    if value is None:
        return default
    if isinstance(value, str):
        return value.strip().lower() not in ("", "0", "false", "no")
    return bool(value)


def builtin_column_settings() -> list[dict[str, Any]]:
    """Return default column settings derived from ``OUTPUT_COLUMNS_CONFIG``.

    Group ids ``g1``, ``g2``, … are assigned to maximal runs of columns whose
    ``group_level`` is not None and ``>= 1`` and ``output`` is True. Columns
    with ``output=False`` break a run and get ``group_id=""``. ``group_level``
    0 or None also yields an empty group id. ``group_collapsed`` is True if any
    member of that run has ``group_collapsed=True``.

    Returns:
        One dict per code column, in builtin order.
    """
    writer = _writer_module()
    default_width = int(getattr(writer, "DEFAULT_COLUMN_WIDTH", 15))
    rows: list[dict[str, Any]] = []
    for defn in writer.OUTPUT_COLUMNS_CONFIG:
        rows.append(
            {
                "col_name": defn.col_name,
                "output": bool(defn.output),
                "width": int(defn.width or default_width),
                "header_label": str(defn.header_label),
                "group_level": defn.group_level,
                "group_collapsed": bool(defn.group_collapsed),
            }
        )

    group_ids = [""] * len(rows)
    collapsed_by_id: dict[str, bool] = {}
    group_n = 0
    idx = 0
    n = len(rows)
    while idx < n:
        item = rows[idx]
        level = item["group_level"]
        in_group = (
            bool(item["output"])
            and level is not None
            and int(level) >= 1
        )
        if not in_group:
            idx += 1
            continue
        group_n += 1
        gid = f"g{group_n}"
        members: list[int] = []
        cursor = idx
        while cursor < n:
            cur = rows[cursor]
            cur_level = cur["group_level"]
            if not cur["output"]:
                break
            if cur_level is None or int(cur_level) < 1:
                break
            members.append(cursor)
            cursor += 1
        any_collapsed = any(rows[m]["group_collapsed"] for m in members)
        for member in members:
            group_ids[member] = gid
        collapsed_by_id[gid] = any_collapsed
        idx = cursor

    settings: list[dict[str, Any]] = []
    for pos, item in enumerate(rows):
        gid = group_ids[pos]
        settings.append(
            {
                "col_name": item["col_name"],
                "output": bool(item["output"]),
                "width": int(item["width"]),
                "header_label": item["header_label"],
                "group_id": gid,
                "group_collapsed": bool(collapsed_by_id.get(gid, False)) if gid else False,
            }
        )
    return settings


def _builtin_by_name() -> dict[str, dict[str, Any]]:
    return {item["col_name"]: item for item in builtin_column_settings()}


def _normalize_groups(columns: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """Keep only contiguous runs of the same non-empty group id."""
    split_n = 0
    seen: set[str] = set()
    idx = 0
    n = len(columns)
    while idx < n:
        gid = str(columns[idx].get("group_id") or "").strip()
        if not gid:
            columns[idx]["group_id"] = ""
            columns[idx]["group_collapsed"] = False
            idx += 1
            continue
        end = idx
        while end < n and str(columns[end].get("group_id") or "").strip() == gid:
            end += 1
        run_id = gid
        if run_id in seen:
            split_n += 1
            run_id = f"g_split{split_n}"
        seen.add(run_id)
        any_collapsed = any(
            _as_bool(columns[pos].get("group_collapsed"), False)
            for pos in range(idx, end)
        )
        for pos in range(idx, end):
            columns[pos]["group_id"] = run_id
            columns[pos]["group_collapsed"] = any_collapsed
        idx = end
    return columns


def normalize_template_columns(saved: list | None) -> list[dict[str, Any]]:
    """Merge saved template columns onto the builtin set.

    Known ``col_name`` values keep the saved order (first occurrence wins).
    Unknown names are dropped. Code columns missing from *saved* are appended
    in builtin order with builtin settings. Width is clamped to 4..80.

    Args:
        saved: Raw column dicts from a user template, or None.

    Returns:
        Normalized column settings covering every builtin column.
    """
    builtin = builtin_column_settings()
    by_name = {item["col_name"]: item for item in builtin}
    ordered: list[dict[str, Any]] = []
    seen: set[str] = set()
    if isinstance(saved, list):
        for raw in saved:
            if not isinstance(raw, dict):
                continue
            name = str(raw.get("col_name") or "").strip()
            if not name or name not in by_name or name in seen:
                continue
            seen.add(name)
            base = by_name[name]
            label = str(raw.get("header_label", base["header_label"])).strip()
            ordered.append(
                {
                    "col_name": name,
                    "output": _as_bool(raw.get("output"), base["output"]),
                    "width": _clamp_width(raw.get("width"), base["width"]),
                    "header_label": label or base["header_label"],
                    "group_id": str(raw.get("group_id", base["group_id"]) or "").strip(),
                    "group_collapsed": _as_bool(
                        raw.get("group_collapsed"),
                        base["group_collapsed"],
                    ),
                }
            )
    for item in builtin:
        if item["col_name"] in seen:
            continue
        ordered.append(copy.deepcopy(item))
    return _normalize_groups(ordered)


def settings_to_column_defs(settings: list | None):
    """Build writer ``ColumnDef`` objects from normalized template settings.

    Builtin ``section`` is always taken from code. A non-empty ``group_id``
    becomes ``group_level=1``; empty group id becomes ``group_level=None``.

    Args:
        settings: Template column dicts (normalized if needed).

    Returns:
        Column defs in template order, including ``output=False`` rows.
    """
    writer = _writer_module()
    column_def_cls = writer.ColumnDef
    default_width = int(getattr(writer, "DEFAULT_COLUMN_WIDTH", 15))
    base_by_name = {defn.col_name: defn for defn in writer.OUTPUT_COLUMNS_CONFIG}
    normalized = normalize_template_columns(settings)
    out = []
    for item in normalized:
        base = base_by_name.get(item["col_name"])
        if base is None:
            continue
        gid = str(item.get("group_id") or "").strip()
        if gid:
            group_level = 1
            group_collapsed = bool(item.get("group_collapsed"))
        else:
            group_level = None
            group_collapsed = False
        out.append(
            column_def_cls(
                col_name=base.col_name,
                header_label=str(item.get("header_label") or base.header_label),
                section=base.section,
                output=bool(item.get("output")),
                width=int(item.get("width") or default_width),
                group_level=group_level,
                group_collapsed=group_collapsed,
            )
        )
    return out


def outline_options_for_visible(defs) -> dict[int, dict]:
    """xlsxwriter ``set_column`` options keyed by visible-column index.

    Members of a contiguous visible run with ``group_level >= 1`` get
    ``level=1`` and ``hidden`` from the run's collapsed flag. If the run is
    collapsed, the next visible column gets ``collapsed=True`` and ``level=0``.
    ``collapsed`` is never combined with ``level >= 1``. A group that runs to
    the last visible column has no summary column.

    Args:
        defs: Full column-def list (including ``output=False``).

    Returns:
        Options dict for each visible column index.
    """
    visible = [defn for defn in defs if defn.output]
    n = len(visible)
    opts: dict[int, dict] = {}
    summary_at: set[int] = set()
    idx = 0
    while idx < n:
        if idx in summary_at:
            idx += 1
            continue
        defn = visible[idx]
        level = defn.group_level
        if level is not None and int(level) >= 1:
            end = idx
            collapsed = False
            while end < n and end not in summary_at:
                member = visible[end]
                member_level = member.group_level
                if member_level is None or int(member_level) < 1:
                    break
                collapsed = collapsed or bool(member.group_collapsed)
                end += 1
            for pos in range(idx, end):
                opts[pos] = {"level": 1, "hidden": collapsed}
            if collapsed and end < n:
                opts[end] = {"level": 0, "collapsed": True}
                summary_at.add(end)
            idx = end
            continue
        if idx not in opts:
            opts[idx] = {"level": 0}
        else:
            opts[idx]["level"] = 0
            opts[idx].pop("hidden", None)
        idx += 1
    for pos in range(n):
        if pos not in opts:
            opts[pos] = {"level": 0}
        elif opts[pos].get("collapsed") and int(opts[pos].get("level", 0) or 0) >= 1:
            opts[pos] = {"level": 0, "collapsed": True}
    return opts


def _empty_state() -> dict[str, Any]:
    return {"active_id": DEFAULT_TEMPLATE_ID, "templates": []}


def _sanitize_template(raw: object) -> dict[str, Any] | None:
    if not isinstance(raw, dict):
        return None
    tid = str(raw.get("id") or "").strip()
    if not tid or tid == DEFAULT_TEMPLATE_ID:
        return None
    name = str(raw.get("name") or "").strip() or tid
    columns = raw.get("columns")
    if not isinstance(columns, list):
        columns = []
    return {
        "id": tid,
        "name": name,
        "columns": normalize_template_columns(columns),
    }


def load_step4_excel_column_state() -> dict[str, Any]:
    """Load ``step4_excel_columns`` from the merged RFP config.

    Returns:
        Dict with ``active_id`` and ``templates`` (builtin default is absent).
    """
    config = load_config()
    raw = config.get(_STATE_KEY)
    if not isinstance(raw, dict):
        return _empty_state()
    templates: list[dict[str, Any]] = []
    seen: set[str] = set()
    for item in raw.get("templates") or []:
        tmpl = _sanitize_template(item)
        if tmpl is None or tmpl["id"] in seen:
            continue
        seen.add(tmpl["id"])
        templates.append(tmpl)
    active_id = str(raw.get("active_id") or DEFAULT_TEMPLATE_ID).strip()
    if not active_id:
        active_id = DEFAULT_TEMPLATE_ID
    return {"active_id": active_id, "templates": templates}


def save_step4_excel_column_state(state: dict[str, Any] | None) -> None:
    """Persist only the ``step4_excel_columns`` key of the RFP config.

    Args:
        state: ``active_id`` plus user ``templates``. Builtin ``default`` is
            stripped from the template list.
    """
    incoming = state if isinstance(state, dict) else {}
    templates: list[dict[str, Any]] = []
    seen: set[str] = set()
    for item in incoming.get("templates") or []:
        tmpl = _sanitize_template(item)
        if tmpl is None or tmpl["id"] in seen:
            continue
        seen.add(tmpl["id"])
        templates.append(
            {
                "id": tmpl["id"],
                "name": tmpl["name"],
                "columns": tmpl["columns"],
            }
        )
    active_id = str(incoming.get("active_id") or DEFAULT_TEMPLATE_ID).strip()
    if not active_id:
        active_id = DEFAULT_TEMPLATE_ID
    config = load_config()
    config[_STATE_KEY] = {
        "active_id": active_id,
        "templates": templates,
    }
    save_config(config)


def _template_by_id(state: dict[str, Any], template_id: str) -> dict[str, Any] | None:
    for tmpl in state.get("templates") or []:
        if tmpl.get("id") == template_id:
            return tmpl
    return None


def resolve_active_column_defs():
    """Column defs for the remembered template, or builtin if missing.

    Returns:
        Writer ``ColumnDef`` list for the next Step4 Excel export.
    """
    state = load_step4_excel_column_state()
    active_id = str(state.get("active_id") or DEFAULT_TEMPLATE_ID).strip()
    if active_id == DEFAULT_TEMPLATE_ID:
        return settings_to_column_defs(builtin_column_settings())
    tmpl = _template_by_id(state, active_id)
    if tmpl is None:
        return settings_to_column_defs(builtin_column_settings())
    return settings_to_column_defs(tmpl.get("columns"))


def _new_template_id(existing: set[str]) -> str:
    for _ in range(32):
        tid = f"t_{uuid.uuid4().hex[:8]}"
        if tid not in existing and tid != DEFAULT_TEMPLATE_ID:
            return tid
    raise Step4ColumnTemplateError("Could not allocate a template id")


def create_template(name: str, columns: list | None) -> str:
    """Create a user template from *columns* and make it active.

    Args:
        name: Display name in the GUI combo.
        columns: Column settings to copy (normalized on save).

    Returns:
        New template id (``t_<short>``).
    """
    state = load_step4_excel_column_state()
    existing = {str(tmpl.get("id") or "") for tmpl in state["templates"]}
    tid = _new_template_id(existing)
    label = str(name or "").strip() or tid
    state["templates"].append(
        {
            "id": tid,
            "name": label,
            "columns": normalize_template_columns(columns),
        }
    )
    state["active_id"] = tid
    save_step4_excel_column_state(state)
    return tid


def rename_template(template_id: str, name: str) -> None:
    """Rename a user template.

    Args:
        template_id: User template id (not ``default``).
        name: New display name.

    Raises:
        Step4ColumnTemplateError: If *template_id* is ``default`` or missing.
    """
    tid = str(template_id or "").strip()
    if tid == DEFAULT_TEMPLATE_ID:
        raise Step4ColumnTemplateError("Cannot rename the built-in default template")
    state = load_step4_excel_column_state()
    tmpl = _template_by_id(state, tid)
    if tmpl is None:
        raise Step4ColumnTemplateError(f"Unknown template id: {tid}")
    label = str(name or "").strip()
    if not label:
        raise Step4ColumnTemplateError("Template name must not be empty")
    tmpl["name"] = label
    save_step4_excel_column_state(state)


def delete_template(template_id: str) -> None:
    """Delete a user template. Refuses the built-in ``default`` id.

    Args:
        template_id: User template id.

    Raises:
        Step4ColumnTemplateError: If *template_id* is ``default``.
    """
    tid = str(template_id or "").strip()
    if tid == DEFAULT_TEMPLATE_ID:
        raise Step4ColumnTemplateError("Cannot delete the built-in default template")
    state = load_step4_excel_column_state()
    before = len(state["templates"])
    state["templates"] = [
        tmpl for tmpl in state["templates"] if tmpl.get("id") != tid
    ]
    if len(state["templates"]) == before:
        return
    if state.get("active_id") == tid:
        state["active_id"] = DEFAULT_TEMPLATE_ID
    save_step4_excel_column_state(state)


def set_active_template(template_id: str) -> None:
    """Remember *template_id* as the active Step4 column layout.

    Unknown ids are stored as-is; ``resolve_active_column_defs`` still falls
    back to builtin settings when the id is missing from ``templates``.

    Args:
        template_id: ``default`` or a user template id.
    """
    tid = str(template_id or "").strip() or DEFAULT_TEMPLATE_ID
    state = load_step4_excel_column_state()
    state["active_id"] = tid
    save_step4_excel_column_state(state)


def update_template_columns(template_id: str, columns: list | None) -> None:
    """Replace columns of a user template. Refuses ``default``.

    Args:
        template_id: User template id.
        columns: New column settings.

    Raises:
        Step4ColumnTemplateError: If *template_id* is ``default`` or missing.
    """
    tid = str(template_id or "").strip()
    if tid == DEFAULT_TEMPLATE_ID:
        raise Step4ColumnTemplateError("Cannot update the built-in default template")
    state = load_step4_excel_column_state()
    tmpl = _template_by_id(state, tid)
    if tmpl is None:
        raise Step4ColumnTemplateError(f"Unknown template id: {tid}")
    tmpl["columns"] = normalize_template_columns(columns)
    save_step4_excel_column_state(state)
