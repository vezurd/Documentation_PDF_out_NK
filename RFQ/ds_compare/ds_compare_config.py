"""Параметры сравнения ДС vs MTO (опциональный JSON рядом с модулем)."""

from __future__ import annotations

import json
import os
import re
from copy import deepcopy
from dataclasses import dataclass
from typing import Any, Dict, List, NotRequired, TypedDict

from RFQ.ds_compare.ds_vs_mto_excel_columns import (
    default_columns_settings,
    merge_column_settings,
)

_CONFIG_FILENAME = "ds_compare_config.json"


class MtoPathPresetDict(TypedDict):
    """One MTO root folder preset for DS vs MTO."""

    label: str
    path: str


class InCabinetDebugDict(TypedDict):
    """IN_CABINET trace settings (GUI block in ds_mto_path_settings)."""

    enabled: bool
    watch_title: str
    watch_mark: str
    watch_code: str
    watch_cabinet: str


class GroupedCompareDict(TypedDict):
    """Grouped DS vs MTO settings."""

    group_by_code: bool
    group_by_title: bool
    group_by_system: bool
    group_by_ds_name: bool
    group_mto_new_positions: bool
    cache_ds_mto_result: bool


class DsVsMtoOutputDict(TypedDict):
    """DS vs MTO Excel output settings."""

    show_internal_column_names: bool
    expand_aggregated_replacement_codes: bool
    excel_export_mode: NotRequired[str]
    columns: NotRequired[list[dict[str, Any]]]


EXCEL_EXPORT_MODE_BOTH = "both"
EXCEL_EXPORT_MODE_WITH_COMMENTS = "with_comments"
EXCEL_EXPORT_MODE_WITHOUT_COMMENTS = "without_comments"

_VALID_EXCEL_EXPORT_MODES = frozenset(
    {
        EXCEL_EXPORT_MODE_BOTH,
        EXCEL_EXPORT_MODE_WITH_COMMENTS,
        EXCEL_EXPORT_MODE_WITHOUT_COMMENTS,
    }
)


class GuiPathsDict(TypedDict):
    """Last-used paths in ``ds_compare_center`` GUI."""

    last_ds_file: str
    last_rfq_file: str
    last_merge_ds_folder: str
    last_tsd_packing_folder: str
    last_tsd_summary_file: str
    last_upd_folder: str
    last_upd_summary_file: str


class GuiWindowDict(TypedDict):
    """Window size and left/right splitter sizes for ``ds_compare_center``."""

    width: int
    height: int
    maximized: bool
    splitters: dict[str, list[int]]


_WATCH_SPLIT_RE = re.compile(r"[,;\n]+")

# Defaults match CenterWindow / split_layout initial sizes.
_GUI_WINDOW_DEFAULT_W = 1410
_GUI_WINDOW_DEFAULT_H = 980
_GUI_WINDOW_MIN_W = 1200
_GUI_WINDOW_MIN_H = 860
_GUI_SPLITTER_DEFAULT = (520, 770)
_GUI_SPLITTER_KEYS = (
    "run",
    "packing",
    "rfp_run",
    "rfp_parts",
    "rfp_ds_id",
    "rfp_ds_mp",
    "misc_run",
    "upd",
)


def get_default_in_cabinet_debug() -> InCabinetDebugDict:
    return {
        "enabled": False,
        "watch_title": "",
        "watch_mark": "",
        "watch_code": "",
        "watch_cabinet": "",
    }


def get_default_grouped_compare() -> GroupedCompareDict:
    return {
        "group_by_code": True,
        "group_by_title": True,
        "group_by_system": True,
        "group_by_ds_name": False,
        "group_mto_new_positions": True,
        "cache_ds_mto_result": False,
    }


def get_default_ds_vs_mto_output() -> dict[str, Any]:
    return {
        "show_internal_column_names": False,
        "expand_aggregated_replacement_codes": True,
        "excel_export_mode": EXCEL_EXPORT_MODE_BOTH,
        "columns": default_columns_settings(),
    }


def normalize_excel_export_mode(raw: object) -> str:
    """Normalize xlsx comment export mode from config or GUI."""
    text = str(raw or "").strip().lower()
    if text in _VALID_EXCEL_EXPORT_MODES:
        return text
    return EXCEL_EXPORT_MODE_BOTH


def get_default_gui_paths() -> GuiPathsDict:
    return {
        "last_ds_file": "",
        "last_rfq_file": "",
        "last_merge_ds_folder": (
            r"\\bcc\eng\PrDoc\377_НИПИГАЗ\АГХК\КСБ\Спецификации к ДС в Excel"
        ),
        "last_tsd_packing_folder": (
            r"\\bcc\root\CurProjects\di_manegers\Пятницкий_П"
            r"\Амурский ГХК\Поставки\ТСД по всем ДС"
        ),
        "last_tsd_summary_file": "",
        "last_upd_folder": (
            r"\\bcc\root\CurProjects\di_manegers\Пятницкий_П"
            r"\Амурский ГХК\Поставки\Файл закачки УПД по всем ДС"
        ),
        "last_upd_summary_file": "",
    }


def normalize_gui_paths(raw: Any) -> GuiPathsDict:
    base = get_default_gui_paths()
    if not isinstance(raw, dict):
        return dict(base)
    return {
        "last_ds_file": str(raw.get("last_ds_file", "")).strip(),
        "last_rfq_file": str(raw.get("last_rfq_file", "")).strip(),
        "last_merge_ds_folder": str(
            raw.get("last_merge_ds_folder", base["last_merge_ds_folder"])
        ).strip()
        or base["last_merge_ds_folder"],
        "last_tsd_packing_folder": str(
            raw.get("last_tsd_packing_folder", base["last_tsd_packing_folder"])
        ).strip()
        or base["last_tsd_packing_folder"],
        "last_tsd_summary_file": str(
            raw.get("last_tsd_summary_file", "")
        ).strip(),
        "last_upd_folder": str(
            raw.get("last_upd_folder", base["last_upd_folder"])
        ).strip()
        or base["last_upd_folder"],
        "last_upd_summary_file": str(
            raw.get("last_upd_summary_file", "")
        ).strip(),
    }


def get_default_gui_window() -> GuiWindowDict:
    left, right = _GUI_SPLITTER_DEFAULT
    return {
        "width": _GUI_WINDOW_DEFAULT_W,
        "height": _GUI_WINDOW_DEFAULT_H,
        "maximized": False,
        "splitters": {key: [left, right] for key in _GUI_SPLITTER_KEYS},
    }


def _normalize_splitter_pair(raw: Any, default: list[int]) -> list[int]:
    if not isinstance(raw, (list, tuple)) or len(raw) < 2:
        return list(default)
    try:
        left = max(240, int(raw[0]))
        right = max(420, int(raw[1]))
    except (TypeError, ValueError):
        return list(default)
    return [left, right]


def normalize_gui_window(raw: Any) -> GuiWindowDict:
    """Clamp window geometry and per-tab splitter sizes from JSON."""
    base = get_default_gui_window()
    if not isinstance(raw, dict):
        return dict(base)
    try:
        width = int(raw.get("width", base["width"]))
    except (TypeError, ValueError):
        width = base["width"]
    try:
        height = int(raw.get("height", base["height"]))
    except (TypeError, ValueError):
        height = base["height"]
    width = max(_GUI_WINDOW_MIN_W, min(width, 10000))
    height = max(_GUI_WINDOW_MIN_H, min(height, 10000))
    raw_splitters = raw.get("splitters")
    splitters_in = raw_splitters if isinstance(raw_splitters, dict) else {}
    # Legacy flat keys left/right apply to all tabs when splitters missing.
    legacy_pair = _normalize_splitter_pair(
        [raw.get("splitter_left"), raw.get("splitter_right")],
        list(_GUI_SPLITTER_DEFAULT),
    )
    splitters: dict[str, list[int]] = {}
    for key in _GUI_SPLITTER_KEYS:
        if key in splitters_in:
            splitters[key] = _normalize_splitter_pair(
                splitters_in.get(key),
                list(base["splitters"][key]),
            )
        else:
            splitters[key] = list(legacy_pair)
    return {
        "width": width,
        "height": height,
        "maximized": bool(raw.get("maximized", False)),
        "splitters": splitters,
    }


def parse_watch_field(raw: object) -> frozenset[str]:
    """Split comma/semicolon/newline list; empty string → no filter (all)."""
    text = str(raw or "").strip()
    if not text:
        return frozenset()
    return frozenset(p.strip() for p in _WATCH_SPLIT_RE.split(text) if p.strip())


def normalize_in_cabinet_debug(raw: Any) -> InCabinetDebugDict:
    base = get_default_in_cabinet_debug()
    if not isinstance(raw, dict):
        return dict(base)
    return {
        "enabled": bool(raw.get("enabled", base["enabled"])),
        "watch_title": str(raw.get("watch_title", "")).strip(),
        "watch_mark": str(raw.get("watch_mark", "")).strip(),
        "watch_code": str(raw.get("watch_code", "")).strip(),
        "watch_cabinet": str(raw.get("watch_cabinet", "")).strip(),
    }


def normalize_grouped_compare(raw: Any) -> GroupedCompareDict:
    base = get_default_grouped_compare()
    if not isinstance(raw, dict):
        return dict(base)
    out: GroupedCompareDict = {
        "group_by_code": bool(raw.get("group_by_code", base["group_by_code"])),
        "group_by_title": bool(raw.get("group_by_title", base["group_by_title"])),
        "group_by_system": bool(raw.get("group_by_system", base["group_by_system"])),
        "group_by_ds_name": bool(raw.get("group_by_ds_name", base["group_by_ds_name"])),
        "group_mto_new_positions": bool(
            raw.get("group_mto_new_positions", base["group_mto_new_positions"])
        ),
        "cache_ds_mto_result": bool(
            raw.get("cache_ds_mto_result", base["cache_ds_mto_result"])
        ),
    }
    if not any(
        out[k] for k in ("group_by_code", "group_by_title", "group_by_system", "group_by_ds_name")
    ):
        out.update(
            {
                "group_by_code": base["group_by_code"],
                "group_by_title": base["group_by_title"],
                "group_by_system": base["group_by_system"],
                "group_by_ds_name": base["group_by_ds_name"],
            }
        )
    return out


def normalize_ds_vs_mto_output(raw: Any) -> dict[str, Any]:
    base = get_default_ds_vs_mto_output()
    if not isinstance(raw, dict):
        return dict(base)
    columns_raw = raw.get("columns")
    columns = merge_column_settings(columns_raw if isinstance(columns_raw, list) else None)
    return {
        "show_internal_column_names": bool(
            raw.get("show_internal_column_names", base["show_internal_column_names"])
        ),
        "expand_aggregated_replacement_codes": bool(
            raw.get(
                "expand_aggregated_replacement_codes",
                base["expand_aggregated_replacement_codes"],
            )
        ),
        "excel_export_mode": normalize_excel_export_mode(
            raw.get("excel_export_mode", base.get("excel_export_mode"))
        ),
        "columns": columns,
    }


@dataclass(frozen=True)
class InCabinetWatchFilters:
    """Empty set on a dimension = match all (no filter on that axis)."""

    codes: frozenset[str]
    titles: frozenset[str]
    marks: frozenset[str]
    cabinets: frozenset[str]

    @classmethod
    def from_debug_dict(cls, d: InCabinetDebugDict) -> InCabinetWatchFilters:
        return cls(
            codes=parse_watch_field(d.get("watch_code")),
            titles=parse_watch_field(d.get("watch_title")),
            marks=parse_watch_field(d.get("watch_mark")),
            cabinets=parse_watch_field(d.get("watch_cabinet")),
        )

    def describe(self) -> str:
        def _fmt(name: str, values: frozenset[str]) -> str:
            return f"{name}=<{', '.join(sorted(values))}>" if values else f"{name}=<все>"

        parts = (
            _fmt("код", self.codes),
            _fmt("титул", self.titles),
            _fmt("марка", self.marks),
            _fmt("шкаф", self.cabinets),
        )
        if self.marks:
            return ", ".join(parts) + " (марка: TYPE_MARK, DS_SYSTEM, ключ спецификации)"
        return ", ".join(parts)


def compare_debug_log_requested(
    cfg: Dict[str, Any],
    compare_debug_log: bool | str | os.PathLike | None,
) -> bool:
    """True if trace/log should run (explicit arg or ``in_cabinet_debug.enabled``)."""
    if compare_debug_log is True:
        return True
    if isinstance(compare_debug_log, (str, os.PathLike)) and str(compare_debug_log).strip():
        return True
    if compare_debug_log is False:
        return False
    return normalize_in_cabinet_debug(cfg.get("in_cabinet_debug")).get("enabled", False)


def get_ds_compare_config_path() -> str:
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), _CONFIG_FILENAME)


def _default_mto_path_presets() -> List[MtoPathPresetDict]:
    """Built-in presets (former hardcoded paths in ds_start_init)."""
    return [
        {
            "label": "АН_RFQ / МТО для закупки",
            "path": r"\\bcc\eng\PrDoc\377_НИПИГАЗ\АГХК\КСБ\АН_RFQ\МТО для закупки",
        },
        {
            "label": "РД / сравнение с ДС / МТО для закупки",
            "path": r"\\bcc\eng\PrDoc\377_НИПИГАЗ\АГХК\КСБ\РД\Таблички_графики\14_сравнение с ДС\МТО для закупки",
        },
        {
            "label": "RFP_MTO_VO / готовые для робота",
            "path": r"\\bcc\eng\PrDoc\377_НИПИГАЗ\АГХК\КСБ\RFP_MTO_VO\_ГОТОВЫЕ_ДЛЯ_РОБОТА",
        },
    ]


def get_default_ds_compare_config() -> Dict[str, Any]:
    presets = _default_mto_path_presets()
    return {
        "flat_mto_structure": False,
        "mto_use_cache": True,
        "mto_force_update": False,
        "mto_paths": deepcopy(presets),
        "mto_path_selected_index": len(presets) - 1,
        "in_cabinet_debug": get_default_in_cabinet_debug(),
        "grouped_compare": get_default_grouped_compare(),
        "ds_vs_mto_output": get_default_ds_vs_mto_output(),
        "gui_paths": get_default_gui_paths(),
        "gui_window": get_default_gui_window(),
    }


def normalize_mto_path_entries(raw: Any) -> List[MtoPathPresetDict]:
    """Returns non-empty path presets; falls back to built-in list if none valid."""
    if not isinstance(raw, list):
        return deepcopy(_default_mto_path_presets())
    out: List[MtoPathPresetDict] = []
    for item in raw:
        if not isinstance(item, dict):
            continue
        path = str(item.get("path", "")).strip()
        if not path:
            continue
        label = str(item.get("label", "")).strip() or path
        out.append({"label": label, "path": path})
    return out if out else deepcopy(_default_mto_path_presets())


def _finalize_config_dict(out: Dict[str, Any]) -> Dict[str, Any]:
    out["flat_mto_structure"] = bool(out.get("flat_mto_structure", False))
    out["mto_use_cache"] = bool(out.get("mto_use_cache", True))
    out["mto_force_update"] = bool(out.get("mto_force_update", False))
    out["mto_paths"] = normalize_mto_path_entries(out.get("mto_paths"))
    n = len(out["mto_paths"])
    idx = int(out.get("mto_path_selected_index", 0))
    out["mto_path_selected_index"] = max(0, min(idx, n - 1)) if n else 0
    out["in_cabinet_debug"] = normalize_in_cabinet_debug(out.get("in_cabinet_debug"))
    out["grouped_compare"] = normalize_grouped_compare(out.get("grouped_compare"))
    out["ds_vs_mto_output"] = normalize_ds_vs_mto_output(out.get("ds_vs_mto_output"))
    out["gui_paths"] = normalize_gui_paths(out.get("gui_paths"))
    out["gui_window"] = normalize_gui_window(out.get("gui_window"))
    # Drop obsolete keys left from older builds (e.g. temporary timing_log block).
    out.pop("tsd_packing", None)
    return out


def format_mto_path_menu_line(index: int, entry: MtoPathPresetDict) -> str:
    """Line shown in the main-window dropdown (1-based index + short label)."""
    return f"{index + 1}. {entry['label']}"


def menu_lines_and_index(cfg: Dict[str, Any]) -> tuple[list[str], int]:
    """Builds option-menu value strings and clamped selected index."""
    entries: List[MtoPathPresetDict] = normalize_mto_path_entries(cfg.get("mto_paths"))
    idx = int(cfg.get("mto_path_selected_index", 0))
    if entries:
        idx = max(0, min(idx, len(entries) - 1))
    else:
        idx = 0
    lines = [format_mto_path_menu_line(i, e) for i, e in enumerate(entries)]
    return lines, idx


def resolve_mto_path_from_config(cfg: Dict[str, Any] | None = None) -> str | None:
    """Folder path for the currently selected MTO preset, or None if list is empty."""
    if cfg is None:
        cfg = load_ds_compare_config()
    entries = normalize_mto_path_entries(cfg.get("mto_paths"))
    if not entries:
        return None
    idx = int(cfg.get("mto_path_selected_index", 0))
    idx = max(0, min(idx, len(entries) - 1))
    return entries[idx]["path"]


_MENU_LINE_RE = re.compile(r"^(\d+)\.\s+")


def index_from_menu_line(line: str, n_entries: int) -> int | None:
    """Parses leading ``N. `` prefix from dropdown value; returns 0-based index or None."""
    m = _MENU_LINE_RE.match(line.strip())
    if not m or n_entries <= 0:
        return None
    one_based = int(m.group(1))
    idx = one_based - 1
    if 0 <= idx < n_entries:
        return idx
    return None


def load_ds_compare_config() -> Dict[str, Any]:
    """Читает ds_compare_config.json; при отсутствии или ошибке — defaults."""
    base = get_default_ds_compare_config()
    path = get_ds_compare_config_path()
    if not os.path.isfile(path):
        return _finalize_config_dict(dict(base))
    try:
        with open(path, "r", encoding="utf-8") as f:
            loaded = json.load(f)
        if isinstance(loaded, dict):
            out = {**base, **loaded}
            return _finalize_config_dict(out)
    except Exception:
        pass
    return _finalize_config_dict(dict(base))


def save_ds_compare_config(cfg: Dict[str, Any]) -> bool:
    """Writes ds_compare_config.json (UTF-8)."""
    path = get_ds_compare_config_path()
    to_save = _finalize_config_dict(dict(cfg))
    try:
        with open(path, "w", encoding="utf-8") as f:
            json.dump(to_save, f, ensure_ascii=False, indent=2)
        return True
    except Exception as e:
        print(f"Ошибка сохранения ds_compare_config {path}: {e}")
        return False
