"""Configuration helpers for the PDF v2 pipeline and control center.

The JSON file lives in the project root next to ``main.py``.
"""

from __future__ import annotations

import json
import os
from typing import Any

_CONFIG_FILENAME = "pdf_v2_config.json"
PER_PAGE_DOC_TYPES: tuple[str, ...] = (
    "DW",
    "WIR",
    "LAY",
    "CAE",
    "GA",
    "PL",
    "NI",
    "MTO",
    "BOE",
    "BOM",
    "BOQ",
    "OD",
    "CJ",
    "VO",
)


def _project_root() -> str:
    return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def get_v2_config_path() -> str:
    return os.path.join(_project_root(), _CONFIG_FILENAME)


def get_default_excel_config() -> dict[str, Any]:
    """Defaults for Excel writers used by the v2 pipeline (openpyxl)."""
    return {
        # Strip XML-disallowed control characters from string cells (Excel paste-like behavior).
        "sanitize_illegal_chars": True,
    }


def get_default_per_page_config() -> dict[str, Any]:
    """Return defaults for selective per-page extraction parallelism."""
    return {
        "enabled": True,
        "min_pages": 4,
        "max_workers": 0,
        "chunk_size_pages": 0,
        "doc_types": {
            doc_type: (doc_type == "DW")
            for doc_type in PER_PAGE_DOC_TYPES
        },
    }


# Island prefilter for detected cells (`grid_detected_prefilter.py`); keys mirrored in GUI / pdf_v2_config.json.
_DEFAULT_GRID_DETECTED_PREFILTER_FLAT: dict[str, Any] = {
    "grid_detected_prefilter_enabled": True,
    "grid_detected_prefilter_min_cells": 8,
    "grid_detected_prefilter_y_pad_frac": 0.18,
    "grid_detected_prefilter_y_pad_floor_pt": 8.0,
    "grid_detected_prefilter_max_removal_frac": 0.58,
    "grid_detected_prefilter_min_component_cells": 4,
    "grid_detected_prefilter_min_component_area_frac": 0.08,
    "grid_detected_prefilter_adjacency_tol_floor_pt": 2.5,
}

GRID_DETECTED_PREFILTER_KEYS: tuple[str, ...] = tuple(_DEFAULT_GRID_DETECTED_PREFILTER_FLAT.keys())

# PDF v2 Control Center UI themes (Qt ColorScheme + Fusion palettes from main_v2/appearance.py).
MONITOR_UI_THEME_CHOICES: tuple[str, ...] = (
    "system",
    "dark",
    "light",
    "light_high_contrast",
    "dark_high_contrast",
)

MONITOR_UI_THEME_LABELS: dict[str, str] = {
    "system": "как в системе",
    "dark": "тёмная",
    "light": "светлая",
    "light_high_contrast": "светлая контрастная",
    "dark_high_contrast": "тёмная контрастная",
}


def normalize_monitor_ui_theme(value: Any) -> str:
    s = str(value or "").strip().lower()
    return s if s in MONITOR_UI_THEME_CHOICES else "system"


def get_default_tag_analysis_ui_config() -> dict[str, Any]:
    """Persisted layout for the «Анализ тегов» tab (splitters, column widths)."""
    return {
        # [vertical_split_top, vertical_split_bottom] — tag analysis upper vs invalid zone.
        "split_vertical": [520, 260],
        # [checks_list_width, detail_pane_width]
        "split_horizontal": [320, 720],
        # check id -> list of column widths (px)
        "detail_column_widths": {},
        # Лист | Тег | Ошибка | PDF
        "invalid_column_widths": [120, 180, 200, 280],
    }


def get_default_monitor_ui_config() -> dict[str, Any]:
    """Nested GUI state for pdf_v2_monitor (non-pipeline)."""
    return {
        "tag_analysis": get_default_tag_analysis_ui_config(),
    }


def normalize_tag_analysis_ui_config(data: Any) -> dict[str, Any]:
    """Public alias for persisted «Анализ тегов» layout subsection."""
    return _normalize_tag_analysis_ui_config_impl(data)


def _normalize_tag_analysis_ui_config_impl(data: Any) -> dict[str, Any]:
    base = get_default_tag_analysis_ui_config()
    src = data if isinstance(data, dict) else {}
    out = dict(base)

    sv = src.get("split_vertical")
    if isinstance(sv, list) and len(sv) >= 2:
        try:
            out["split_vertical"] = [max(40, int(sv[0])), max(80, int(sv[1]))]
        except (TypeError, ValueError):
            pass

    sh = src.get("split_horizontal")
    if isinstance(sh, list) and len(sh) >= 2:
        try:
            out["split_horizontal"] = [max(120, int(sh[0])), max(200, int(sh[1]))]
        except (TypeError, ValueError):
            pass

    dcw = src.get("detail_column_widths")
    widths: dict[str, list[int]] = {}
    if isinstance(dcw, dict):
        for k, v in dcw.items():
            ck = str(k).strip() or "_default"
            if isinstance(v, list):
                try:
                    widths[ck] = [max(24, int(x)) for x in v]
                except (TypeError, ValueError):
                    continue
    out["detail_column_widths"] = widths

    icw = src.get("invalid_column_widths")
    if isinstance(icw, list) and len(icw) >= 4:
        try:
            out["invalid_column_widths"] = [max(24, int(x)) for x in icw[:4]]
        except (TypeError, ValueError):
            pass

    return out


def _normalize_monitor_ui_config(data: Any) -> dict[str, Any]:
    base = get_default_monitor_ui_config()
    src = data if isinstance(data, dict) else {}
    out = dict(base)
    out["tag_analysis"] = _normalize_tag_analysis_ui_config_impl(src.get("tag_analysis"))
    return out


def get_default_v2_config() -> dict[str, Any]:
    return {
        "enabled": True,
        "templates_dir": "pdf_parsing_v2_engine/templates",
        # Пустая строка = все проекты; иначе имя папки проекта (как для load_project_templates).
        "project": "",
        # Parallel extraction via ProcessPoolExecutor.
        "parallel": True,
        # 0 / None = auto.
        "max_workers": 0,
        # Selective per-page parallel extraction for heavy documents.
        "per_page": get_default_per_page_config(),
        # Parallel tag parsing via per-file ProcessPoolExecutor.
        # 0 / None = auto by CPU count.
        "tags_max_workers": 0,
        # Tag extraction backend: "fitz" (fast) or "pdfminer" (legacy).
        "tags_text_backend": "fitz",
        # Large-page frame search border band margins (mm), applied per side.
        "find_frame_border_band_left_mm": 20.0,
        "find_frame_border_band_right_mm": 50.0,
        "find_frame_border_band_top_mm": 20.0,
        "find_frame_border_band_bottom_mm": 20.0,
        "find_tables_snap_x_tolerance": 2.2,
        "find_tables_snap_y_tolerance": 2.0,
        **_DEFAULT_GRID_DETECTED_PREFILTER_FLAT,
        "grid_tolerance_detected_mm": 0.5,
        "snap_max_distance_mm": 5.0,
        "export_debug_excel": True,
        "debug_visual": False,
        "debug_visual_dir": "debug/v2_visual",
        "debug_visual_overlay": False,
        "debug_verbose_log": False,
        # Подробные тайминги pipeline → v2_timing_log.xlsx в папке результатов.
        "timing_log": False,
        # Text extraction mode.
        # "char_center" — per-character center-point containment (recommended, default).
        #   Each character is assigned to exactly one cell; eliminates adjacent-cell
        #   bleed caused by spans that physically cross cell boundaries.
        # "get_textbox" — legacy fitz get_textbox() (intersection-based, more noise).
        "text_extraction_mode": "char_center",
        # После основного прогона v2_pipeline.py (только CLI main): MTO page1 smoke vs mto_page1.json.
        "mto_regression_after_run": False,
        "mto_regression_pdf_dir": "",
        "mto_regression_template": "",
        "mto_regression_strict": False,
        "mto_regression_max_nl_rot90": 280,
        # PDF v2 Control Center (`pdf_v2_monitor`) main window size (logical px).
        "monitor_window_width": 1350,
        "monitor_window_height": 1060,
        # Run tab: колонка «Разъезд» — целая оценка нестыковки каркаса (см. stamp_grid_metrics).
        "monitor_grid_mismatch_ok_max": 3000,
        "monitor_grid_mismatch_warn_max": 45000,
        "monitor_grid_mismatch_weight_global": 100000.0,
        "monitor_grid_mismatch_weight_walk_pt": 50.0,
        "monitor_grid_mismatch_weight_no_match": 25000.0,
        # Control Center appearance: "system" | "dark" | "light" (Qt ColorScheme).
        "monitor_ui_theme": "system",
        # Control Center tab layouts (splitters, tables); pipeline ignores.
        "monitor_ui": get_default_monitor_ui_config(),
        # openpyxl-based exports (normcontrol xlsx, timing log xlsx).
        "excel": get_default_excel_config(),
    }


def _normalize_excel_config(data: Any) -> dict[str, Any]:
    """Normalize the nested Excel / openpyxl export section."""
    base = get_default_excel_config()
    out = dict(base)
    src = data if isinstance(data, dict) else {}
    out["sanitize_illegal_chars"] = bool(
        src.get("sanitize_illegal_chars", base["sanitize_illegal_chars"])
    )
    return out


def normalize_v2_config(data: dict[str, Any] | None) -> dict[str, Any]:
    """Return config normalized against the current v2 schema."""
    base = get_default_v2_config()
    out = dict(base)
    src = data or {}
    for key in base:
        if key == "per_page":
            continue
        if key == "excel":
            continue
        if key == "monitor_ui":
            continue
        if key in src:
            out[key] = src[key]
    out["per_page"] = _normalize_per_page_config(src.get("per_page"))
    out["excel"] = _normalize_excel_config(src.get("excel"))
    out["monitor_ui"] = _normalize_monitor_ui_config(src.get("monitor_ui"))
    out["monitor_ui_theme"] = normalize_monitor_ui_theme(out.get("monitor_ui_theme"))
    _migrate_templates_dir(out)
    return out


def excel_sanitize_illegal_chars_enabled(cfg: dict[str, Any] | None) -> bool:
    """Whether string cells written by openpyxl in the v2 pipeline should be sanitized."""
    effective = normalize_v2_config(cfg)
    section = effective.get("excel") if isinstance(effective.get("excel"), dict) else {}
    return bool(section.get("sanitize_illegal_chars", True))


def _normalize_per_page_config(data: Any) -> dict[str, Any]:
    """Normalize the nested per-page parallel config section."""
    base = get_default_per_page_config()
    out = dict(base)
    src = data if isinstance(data, dict) else {}
    out["enabled"] = bool(src.get("enabled", base["enabled"]))

    try:
        out["min_pages"] = max(1, int(src.get("min_pages", base["min_pages"]) or 1))
    except (TypeError, ValueError):
        out["min_pages"] = int(base["min_pages"])

    try:
        out["max_workers"] = max(0, int(src.get("max_workers", base["max_workers"]) or 0))
    except (TypeError, ValueError):
        out["max_workers"] = int(base["max_workers"])

    try:
        out["chunk_size_pages"] = max(
            0, int(src.get("chunk_size_pages", base["chunk_size_pages"]) or 0)
        )
    except (TypeError, ValueError):
        out["chunk_size_pages"] = int(base["chunk_size_pages"])

    doc_types = dict(base["doc_types"])
    raw_doc_types = src.get("doc_types")
    if isinstance(raw_doc_types, dict):
        for doc_type in doc_types:
            if doc_type in raw_doc_types:
                doc_types[doc_type] = bool(raw_doc_types[doc_type])
    elif isinstance(raw_doc_types, list):
        enabled = {str(item).strip().upper() for item in raw_doc_types}
        for doc_type in doc_types:
            doc_types[doc_type] = doc_type in enabled
    out["doc_types"] = doc_types
    return out


def resolve_v2_path(path_value: str, *, default: str = "") -> str:
    """Resolve a possibly relative path against the project root."""
    raw = str(path_value or default or "").strip()
    if not raw:
        return _project_root()
    if os.path.isabs(raw):
        return os.path.normpath(raw)
    return os.path.normpath(os.path.join(_project_root(), raw))


def resolve_templates_dir(cfg: dict[str, Any] | None = None) -> str:
    """Return absolute templates directory from v2 config."""
    effective = normalize_v2_config(cfg)
    return resolve_v2_path(
        str(effective.get("templates_dir", "pdf_parsing_v2_engine/templates")),
        default="pdf_parsing_v2_engine/templates",
    )


def _migrate_templates_dir(cfg: dict[str, Any]) -> None:
    """Rewrite legacy templates_dir to engine package path when needed."""
    td = cfg.get("templates_dir")
    if not isinstance(td, str) or not td.strip():
        return
    raw = td.strip()
    norm = raw.replace("\\", "/")
    legacy = "pdf_parsing_v2/templates"
    new_rel = "pdf_parsing_v2_engine/templates"
    if norm == legacy:
        print(
            f"[v2_config] templates_dir: legacy path {legacy!r} → {new_rel!r} "
            "(see pdf_parsing_v2_engine)",
        )
        cfg["templates_dir"] = new_rel
        return
    root = _project_root()
    if os.path.isabs(raw):
        abs_path = os.path.normpath(raw)
    else:
        abs_path = os.path.normpath(os.path.join(root, raw))
    if not os.path.isdir(abs_path) and legacy in norm:
        print(
            f"[v2_config] templates_dir: path not found ({raw!r}), using {new_rel!r}",
        )
        cfg["templates_dir"] = new_rel


def load_v2_config() -> dict[str, Any]:
    path = get_v2_config_path()
    if not os.path.isfile(path):
        return normalize_v2_config(None)
    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        if not isinstance(data, dict):
            return normalize_v2_config(None)
        return normalize_v2_config(data)
    except Exception as e:
        print(f"[v2_config] Ошибка чтения {path}: {e}")
        return normalize_v2_config(None)


def save_v2_config(config: dict[str, Any]) -> None:
    path = get_v2_config_path()
    merged = normalize_v2_config(config)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(merged, f, ensure_ascii=False, indent=2)
