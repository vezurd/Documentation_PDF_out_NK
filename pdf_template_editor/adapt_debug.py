"""Debug snapshot and visual dump for grid adaptation diagnostics."""

from __future__ import annotations

import json
import os
import re
import shutil
import tempfile
import time
from typing import Any

import fitz

from pdf_parsing_v2_engine.coord_transform import (
    SCALE,
    field_to_fitz_rect,
    get_origin_points,
    grid_line_to_fitz_pts,
    pdfminer_to_fitz,
)
from pdf_parsing_v2_engine.frame_detector import find_frame
from pdf_parsing_v2_engine.grid_matcher import (
    AlignmentLineDiag,
    AlignmentReport,
    AlignmentSummary,
    CellAssignmentInfo,
    CellBbox,
    DetectedGridLineInfo,
    _MAX_CELL_AREA_FRACTION,
    _align_grid_lines_ordered,
    _build_detected_grid_lines,
    _extract_grid_lines,
    _fields_have_bindings,
    _filter_cells_to_bbox,
    _make_unique_keys,
    _template_stamp_bbox,
    adapt_by_cell_assignment,
    get_detected_stamp_cells,
)
from pdf_parsing_v2_engine.models import FieldDef, FrameInfo, StampTemplate

from .auto_detect import detect_cells

_PNG_DPI = 150

_GREEN = (0.0, 0.7, 0.0)
_BLUE = (0.1, 0.3, 0.9)
_RED = (0.9, 0.1, 0.1)
_ORANGE = (1.0, 0.55, 0.0)
_PURPLE = (0.62, 0.12, 0.94)
_MAGENTA = (0.82, 0.0, 0.6)
_CYAN = (0.0, 0.62, 0.8)
_GREY = (0.55, 0.55, 0.55)
_BLACK = (0.0, 0.0, 0.0)


def _bbox_to_list(cb: CellBbox) -> list[float]:
    return [round(cb.x0, 3), round(cb.y0, 3), round(cb.x1, 3), round(cb.y1, 3)]


def _rect_to_list(rect: fitz.Rect | None) -> list[float] | None:
    if rect is None:
        return None
    return [round(rect.x0, 3), round(rect.y0, 3), round(rect.x1, 3), round(rect.y1, 3)]


def _tuple_rect_to_list(rect: tuple[float, float, float, float] | None) -> list[float] | None:
    if rect is None:
        return None
    return [round(float(value), 3) for value in rect]


def _compact_dict(payload: dict[str, Any]) -> dict[str, Any]:
    return {
        key: value
        for key, value in payload.items()
        if value not in ("", None, [], {}, ())
    }


def _safe_json_value(value: Any) -> Any:
    if isinstance(value, (str, int, float, bool)) or value is None:
        return value
    if isinstance(value, (list, tuple)):
        return [_safe_json_value(v) for v in value]
    if isinstance(value, dict):
        return {str(k): _safe_json_value(v) for k, v in value.items()}
    return str(value)


def _sanitize_artifact_part(value: str, fallback: str) -> str:
    text = (value or "").strip()
    if not text:
        return fallback
    text = os.path.splitext(os.path.basename(text))[0]
    text = re.sub(r'[<>:"/\\|?*\s]+', "_", text)
    text = re.sub(r"_+", "_", text).strip("._")
    return text or fallback


def _format_rect(rect: list[float] | None) -> str:
    if not rect:
        return "n/a"
    return f"x0={rect[0]:.3f} y0={rect[1]:.3f} x1={rect[2]:.3f} y1={rect[3]:.3f}"


def _serialize_detected_grid_line(line: DetectedGridLineInfo) -> dict[str, Any]:
    return {
        "id": line.id,
        "orientation": line.orientation,
        "pos_pts": round(line.pos_pts, 3),
        "span_lo_pts": round(line.span_lo_pts, 3),
        "span_hi_pts": round(line.span_hi_pts, 3),
        "span_pts": round(line.span_hi_pts - line.span_lo_pts, 3),
        "cell_count": line.cell_count,
    }


def _serialize_template_grid_line(
    line,
    frame: FrameInfo,
    origin: str,
) -> dict[str, Any]:
    orientation, pos_pts, start_pts, end_pts, boundary, line_id = grid_line_to_fitz_pts(
        line,
        frame,
        origin=origin,
    )
    return {
        "id": line_id,
        "orientation": orientation,
        "pos_mm": round(line.pos_mm, 3),
        "start_mm": round(line.start_mm, 3),
        "end_mm": round(line.end_mm, 3),
        "boundary": boundary,
        "display_pos_pts": round(pos_pts, 3),
        "display_start_pts": round(start_pts, 3),
        "display_end_pts": round(end_pts, 3),
        **({"start_line_id": line.start_line_id} if line.start_line_id else {}),
        **({"end_line_id": line.end_line_id} if line.end_line_id else {}),
    }


def _serialize_alignment_diag(diag: AlignmentLineDiag) -> dict[str, Any]:
    return {
        "line_id": diag.line_id,
        "orientation": diag.orientation,
        "pos_interp_pts": round(diag.pos_interp_pts, 3),
        "pos_snapped_pts": round(diag.pos_snapped_pts, 3),
        "snap_method": diag.snap_method,
        "matched_detected_id": diag.matched_detected_id,
        "matched_detected_pos": (
            round(diag.matched_detected_pos, 3)
            if diag.matched_detected_pos is not None
            else None
        ),
        "matched_detected_span": (
            round(diag.matched_detected_span, 3)
            if diag.matched_detected_span is not None
            else None
        ),
        "match_score": round(diag.match_score, 6),
        "span_ratio": round(diag.span_ratio, 6),
        "dp_action": diag.dp_action,
    }


def _extract_probe_line(raw: dict[str, Any], prefix: str) -> dict[str, Any]:
    return _compact_dict({
        "id": raw.get(f"{prefix}_id"),
        "pos_pts": raw.get(f"{prefix}_pos_pts"),
        "span_pts": raw.get(f"{prefix}_span_pts"),
        "span_ratio_to_template": raw.get(f"{prefix}_span_ratio_to_template"),
    })


def _extract_probe_virtual(raw: dict[str, Any], prefix: str) -> dict[str, Any]:
    return _compact_dict({
        "virtual_pos_pts": raw.get(f"{prefix}_virtual_pos_pts"),
        "matched_template_id": raw.get(f"{prefix}_matched_template_id"),
        "matched_template_pos_pts": raw.get(f"{prefix}_matched_template_pos_pts"),
        "source_detected_id": raw.get(f"{prefix}_source_detected_id"),
        "source_detected_pos_pts": raw.get(f"{prefix}_source_detected_pos_pts"),
        "skip_count": raw.get(f"{prefix}_skip_count"),
        "scale_estimate": raw.get(f"{prefix}_scale_estimate"),
        "score": raw.get(f"{prefix}_score"),
        "candidate_count": raw.get(f"{prefix}_candidate_count"),
        "gap_count_used": raw.get(f"{prefix}_gap_count_used"),
    })


def _serialize_extent_probe(raw: dict[str, Any]) -> dict[str, Any]:
    raw_safe = {key: _safe_json_value(value) for key, value in raw.items()}
    inferred_edges = [
        edge
        for edge in str(raw.get("extent_probe_inferred_edges", "")).split(",")
        if edge
    ]
    return {
        "raw": raw_safe,
        "inferred_edges": inferred_edges,
        "line_source": raw.get("extent_probe_line_source", ""),
        "candidate_counts": {
            "top_strong": raw.get("extent_probe_top_candidate_count_strong", 0),
            "top_medium": raw.get("extent_probe_top_candidate_count_medium", 0),
            "left_strong": raw.get("extent_probe_left_candidate_count_strong", 0),
            "left_medium": raw.get("extent_probe_left_candidate_count_medium", 0),
            "bottom_strong": raw.get("extent_probe_bottom_candidate_count_strong", 0),
            "bottom_medium": raw.get("extent_probe_bottom_candidate_count_medium", 0),
            "right_strong": raw.get("extent_probe_right_candidate_count_strong", 0),
            "right_medium": raw.get("extent_probe_right_candidate_count_medium", 0),
            "virtual_top": raw.get("extent_probe_virtual_top_candidate_count", 0),
            "virtual_left": raw.get("extent_probe_virtual_left_candidate_count", 0),
        },
        "selected_boundaries": {
            "top": _extract_probe_line(raw, "extent_probe_top"),
            "left": _extract_probe_line(raw, "extent_probe_left"),
            "bottom": _extract_probe_line(raw, "extent_probe_bottom"),
            "right": _extract_probe_line(raw, "extent_probe_right"),
        },
        "virtual_edges": {
            "top": _extract_probe_virtual(raw, "extent_probe_virtual_top"),
            "left": _extract_probe_virtual(raw, "extent_probe_virtual_left"),
        },
    }


def prealign_dict_from_cell_assignment_info(
    ca_info: CellAssignmentInfo,
    *,
    fitz_page: fitz.Page | None = None,
    template: StampTemplate | None = None,
    cfg: dict | None = None,
    production_detected_cells: dict[str, list[list[float]]] | None = None,
) -> dict[str, Any]:
    """Build the ``prealign`` JSON block (same keys as ``collect_adapt_debug_snapshot``).

    If *production_detected_cells* is passed (from a prior ``_build_production_detected_cells``),
    it is used as-is. Otherwise when *fitz_page* / *template* / *cfg* are set, cells are
    recomputed; else empty lists (overlay-only callers).
    """
    cells_out: dict[str, list[list[float]]]
    if production_detected_cells is not None:
        cells_out = production_detected_cells
    elif fitz_page is not None and template is not None:
        production_detected_raw, production_detected_probe, production_detected_effective = (
            _build_production_detected_cells(
                fitz_page=fitz_page,
                template=template,
                search_bbox=_tuple_rect_to_list(ca_info.search_bbox),
                effective_bbox=_tuple_rect_to_list(ca_info.effective_bbox),
                cfg=cfg,
            )
        )
        cells_out = {
            "raw": [_bbox_to_list(cell) for cell in production_detected_raw],
            "probe_filtered": [_bbox_to_list(cell) for cell in production_detected_probe],
            "effective": [_bbox_to_list(cell) for cell in production_detected_effective],
        }
    else:
        cells_out = {"raw": [], "probe_filtered": [], "effective": []}
    return {
        "scale_threshold": round(ca_info.scale_threshold, 6),
        "prealign_confidence_score": (
            round(ca_info.prealign_confidence_score, 6)
            if ca_info.prealign_confidence_score is not None
            else None
        ),
        "prealign_confidence_tier": ca_info.prealign_confidence_tier,
        "prealign_confidence_reason": ca_info.prealign_confidence_reason,
        "proposed_bbox_applied": ca_info.proposed_bbox_applied,
        "search_bbox": _tuple_rect_to_list(ca_info.search_bbox),
        "base_bbox": _tuple_rect_to_list(ca_info.base_bbox),
        "proposed_bbox": _tuple_rect_to_list(ca_info.proposed_bbox),
        "effective_bbox": _tuple_rect_to_list(ca_info.effective_bbox),
        "detected_union_bbox": _tuple_rect_to_list(ca_info.detected_union_bbox),
        "detected_counts": {
            "raw": ca_info.n_detected_raw,
            "probe": ca_info.n_detected_probe,
            "effective": ca_info.n_detected_effective,
            "oversized_filtered": ca_info.n_filtered_oversized,
            "outside_effective_bbox_filtered": ca_info.n_filtered_outside_effective_bbox,
        },
        "extent_probe": _serialize_extent_probe(ca_info.extent_probe),
        "prealign_detected_grid_lines": [
            _serialize_detected_grid_line(line)
            for line in ca_info.prealign_detected_grid_lines
        ],
        "detected_prefilter": dict(getattr(ca_info, "detected_prefilter", {}) or {}),
        "production_detected_cells": cells_out,
    }


def stamp_prealign_debug_bundle(
    *,
    frame: FrameInfo,
    template: StampTemplate,
    ca_info: CellAssignmentInfo,
    fitz_page: fitz.Page,
    cfg: dict | None,
) -> dict[str, Any]:
    """JSON-safe bundle for ``V2PageResult.metadata['stamp_prealign_debug']`` (editor F5)."""
    from dataclasses import asdict

    prealign = prealign_dict_from_cell_assignment_info(
        ca_info, fitz_page=fitz_page, template=template, cfg=cfg,
    )
    diags = [_safe_json_value(asdict(d)) for d in ca_info.alignment_diagnostics]
    grid_tpl = [
        _serialize_template_grid_line(line, frame, template.origin)
        for line in template.grid_lines
    ]
    det_prod = [_serialize_detected_grid_line(line) for line in ca_info.detected_grid_lines]
    return {
        "prealign": prealign,
        "alignment_diagnostics": diags,
        "grid_lines_template": grid_tpl,
        "detected_grid_lines_production": det_prod,
        "environment": {
            "rotation": frame.rotation,
            "frame": {
                "x0": round(frame.x0, 3),
                "y0": round(frame.y0, 3),
                "x1": round(frame.x1, 3),
                "y1": round(frame.y1, 3),
                "page_width": round(frame.page_width, 3),
                "page_height": round(frame.page_height, 3),
            },
        },
    }


def _frame_rect_to_fitz(frame_payload: dict[str, Any], rotation: int) -> fitz.Rect:
    return pdfminer_to_fitz(
        (
            float(frame_payload["x0"]),
            float(frame_payload["y0"]),
            float(frame_payload["x1"]),
            float(frame_payload["y1"]),
        ),
        float(frame_payload["page_height"]),
        rotation,
    )


def _overlay_doc_for_page(
    fitz_page: fitz.Page,
    dpi: int = _PNG_DPI,
) -> tuple[fitz.Document, fitz.Page, float]:
    pix = fitz_page.get_pixmap(dpi=dpi)
    scale = dpi / 72.0
    doc = fitz.open()
    page = doc.new_page(width=pix.width, height=pix.height)
    page.insert_image(page.rect, pixmap=pix)
    return doc, page, scale


def _draw_rect(
    page: fitz.Page,
    scale: float,
    rect: list[float] | None,
    color: tuple[float, float, float],
    *,
    width: float = 1.5,
    dashes: str | None = None,
) -> None:
    if not rect:
        return
    shape = page.new_shape()
    shape.draw_rect(
        fitz.Rect(
            rect[0] * scale,
            rect[1] * scale,
            rect[2] * scale,
            rect[3] * scale,
        )
    )
    shape.finish(color=color, width=width, dashes=dashes)
    shape.commit()


def _draw_line(
    page: fitz.Page,
    scale: float,
    start: tuple[float, float],
    end: tuple[float, float],
    color: tuple[float, float, float],
    *,
    width: float = 1.2,
    dashes: str | None = None,
) -> None:
    shape = page.new_shape()
    shape.draw_line(
        fitz.Point(start[0] * scale, start[1] * scale),
        fitz.Point(end[0] * scale, end[1] * scale),
    )
    shape.finish(color=color, width=width, dashes=dashes)
    shape.commit()


def _draw_circle(
    page: fitz.Page,
    scale: float,
    point: tuple[float, float],
    radius_pts: float,
    color: tuple[float, float, float],
) -> None:
    shape = page.new_shape()
    shape.draw_circle(
        fitz.Point(point[0] * scale, point[1] * scale),
        radius_pts * scale,
    )
    shape.finish(color=color, fill=color)
    shape.commit()


def _insert_text(
    page: fitz.Page,
    scale: float,
    point: tuple[float, float],
    text: str,
    *,
    fontsize: float = 7.0,
    color: tuple[float, float, float] = _BLACK,
) -> None:
    if not text:
        return
    page.insert_text(
        fitz.Point(point[0] * scale, point[1] * scale),
        text,
        fontsize=fontsize,
        color=color,
    )


def _draw_detected_grid_line(
    page: fitz.Page,
    scale: float,
    line: dict[str, Any],
    color: tuple[float, float, float],
    *,
    width: float = 1.0,
    dashes: str | None = None,
    label: bool = False,
) -> None:
    if line.get("orientation") == "h":
        start = (float(line["span_lo_pts"]), float(line["pos_pts"]))
        end = (float(line["span_hi_pts"]), float(line["pos_pts"]))
        label_pt = (float(line["span_lo_pts"]), float(line["pos_pts"]) - 2.0)
    else:
        start = (float(line["pos_pts"]), float(line["span_lo_pts"]))
        end = (float(line["pos_pts"]), float(line["span_hi_pts"]))
        label_pt = (float(line["pos_pts"]) + 1.5, float(line["span_lo_pts"]) + 4.0)
    _draw_line(page, scale, start, end, color, width=width, dashes=dashes)
    if label:
        _insert_text(page, scale, label_pt, str(line.get("id", "")), fontsize=6.0, color=color)


def _draw_template_grid_line(
    page: fitz.Page,
    scale: float,
    line: dict[str, Any],
    color: tuple[float, float, float],
    *,
    snapped_pos_pts: float | None = None,
    width: float = 1.3,
    dashes: str | None = None,
    label: bool = False,
) -> None:
    pos = float(snapped_pos_pts) if snapped_pos_pts is not None else float(line["display_pos_pts"])
    if line.get("orientation") == "h":
        start = (float(line["display_start_pts"]), pos)
        end = (float(line["display_end_pts"]), pos)
        label_pt = (float(line["display_start_pts"]), pos - 2.0)
    else:
        start = (pos, float(line["display_start_pts"]))
        end = (pos, float(line["display_end_pts"]))
        label_pt = (pos + 1.5, float(line["display_start_pts"]) + 4.0)
    _draw_line(page, scale, start, end, color, width=width, dashes=dashes)
    if label:
        _insert_text(page, scale, label_pt, str(line.get("id", "")), fontsize=6.0, color=color)


def _build_production_detected_cells(
    *,
    fitz_page: fitz.Page,
    template: StampTemplate,
    search_bbox: list[float] | None,
    effective_bbox: list[float] | None,
    cfg: dict | None,
) -> tuple[list[CellBbox], list[CellBbox], list[CellBbox]]:
    bbox_for_search = (
        fitz.Rect(*search_bbox)
        if search_bbox is not None
        else _template_stamp_bbox(template, find_frame(fitz_page, frame_mode=template.frame_mode, template=template if template.frame_mode == "drawing_union" else None)[0])
    )
    if bbox_for_search is None:
        return [], [], []
    detected_raw = get_detected_stamp_cells(
        fitz_page=fitz_page,
        stamp_bbox=bbox_for_search,
        tolerance_mm=template.grid_tolerance_detected_mm,
        cfg=cfg,
        template=template,
    )
    stamp_area = max(bbox_for_search.get_area(), 1e-6)
    detected_probe = [
        cell
        for cell in detected_raw
        if cell.area <= stamp_area * _MAX_CELL_AREA_FRACTION
    ]
    if effective_bbox is None:
        return detected_raw, detected_probe, detected_probe
    detected_effective = _filter_cells_to_bbox(
        detected_probe,
        fitz.Rect(*effective_bbox),
    )
    return detected_raw, detected_probe, detected_effective


def collect_adapt_debug_snapshot(
    *,
    fitz_page: fitz.Page,
    frame: FrameInfo,
    template: StampTemplate,
    scene_items: list,
    fields_from_scene: list[FieldDef],
    pdf_path: str,
    page_num: int,
    dpi_scale: float,
    cfg: dict | None = None,
) -> dict[str, Any]:
    """Collect a comprehensive debug snapshot for adaptation analysis."""
    snap: dict[str, Any] = {
        "timestamp": time.strftime("%Y-%m-%d %H:%M:%S"),
    }
    cfg_eff = dict(cfg or {})
    if pdf_path:
        cfg_eff.setdefault("_pdf_path", pdf_path)
    cfg_eff.setdefault("_page_num", page_num)

    t_union = template if template.frame_mode == "drawing_union" else None
    frame, frame_detection = find_frame(
        fitz_page,
        debug=True,
        frame_mode=template.frame_mode,
        template=t_union,
    )

    snap["environment"] = {
        "pdf_path": pdf_path,
        "pdf_basename": os.path.basename(pdf_path) if pdf_path else "",
        "page_num": page_num,
        "rotation": frame.rotation,
        "frame": {
            "x0": round(frame.x0, 3),
            "y0": round(frame.y0, 3),
            "x1": round(frame.x1, 3),
            "y1": round(frame.y1, 3),
            "page_width": round(frame.page_width, 3),
            "page_height": round(frame.page_height, 3),
            "border_left_mm": round(frame.border_left_mm, 3),
            "border_bottom_mm": round(frame.border_bottom_mm, 3),
            "border_top_mm": round(frame.border_top_mm, 3),
        },
        "dpi_scale": round(dpi_scale, 6),
        "scale_pts_per_mm": round(SCALE, 6),
    }

    origin_points = get_origin_points(frame)
    snap["origin_points_pts"] = {
        key: [round(value[0], 3), round(value[1], 3)]
        for key, value in origin_points.items()
    }
    snap["origin_points_mm"] = {
        key: [round(value[0] / SCALE, 4), round(value[1] / SCALE, 4)]
        for key, value in origin_points.items()
    }

    grid_lines_h = [line for line in template.grid_lines if line.orientation == "h"]
    grid_lines_v = [line for line in template.grid_lines if line.orientation == "v"]
    boundary_lines = [line for line in template.grid_lines if line.boundary]
    snap["template"] = {
        "name": template.name,
        "source_path": template.source_path,
        "source_basename": os.path.basename(template.source_path) if template.source_path else "",
        "doc_types": list(template.doc_types),
        "page_selector": template.page_selector,
        "priority": template.priority,
        "frame_mode": template.frame_mode,
        "origin": template.origin,
        "field_count": len(template.fields),
        "visible_field_count": sum(1 for field in template.fields if not field.document_property),
        "document_property_field_count": sum(1 for field in template.fields if field.document_property),
        "outside_stamp_field_count": sum(1 for field in template.fields if field.outside_stamp),
        "grid_line_count": len(template.grid_lines),
        "grid_line_summary": {
            "h_count": len(grid_lines_h),
            "v_count": len(grid_lines_v),
            "boundary_count": len(boundary_lines),
            "boundary_h_count": sum(1 for line in boundary_lines if line.orientation == "h"),
            "boundary_v_count": sum(1 for line in boundary_lines if line.orientation == "v"),
        },
    }

    snap["settings"] = {
        "grid_adapt": template.grid_adapt,
        "grid_tolerance_template_mm": template.grid_tolerance_template_mm,
        "grid_tolerance_detected_mm": template.grid_tolerance_detected_mm,
        "snap_max_distance_mm": template.snap_max_distance_mm,
        "max_shape_change_ratio": template.max_shape_change_ratio,
        "cascade_score_threshold": template.cascade_score_threshold,
        "padding_mm": template.padding_mm,
        "find_tables_snap_x_tolerance": template.find_tables_snap_x_tolerance,
        "find_tables_snap_y_tolerance": template.find_tables_snap_y_tolerance,
    }

    keys = _make_unique_keys(fields_from_scene)
    fields_before = []
    for idx, (item, field_def) in enumerate(zip(scene_items, fields_from_scene)):
        pos = item.scenePos()
        rect = item.rect()
        fitz_rect = field_to_fitz_rect(field_def, frame, padding_mm=0.0)
        fields_before.append({
            "key": keys[idx],
            "id": field_def.id,
            "label": field_def.label,
            "bbox_mm": [round(value, 3) for value in field_def.bbox_mm],
            "fitz_rect": _rect_to_list(fitz_rect),
            "scene_rect": [
                round(pos.x() + rect.x(), 1),
                round(pos.y() + rect.y(), 1),
                round(rect.width(), 1),
                round(rect.height(), 1),
            ],
            "is_anchor": field_def.is_anchor,
            "outside_stamp": field_def.outside_stamp,
            "document_property": field_def.document_property,
            "field_type": field_def.field_type,
            "expected": field_def.expected,
        })
    snap["fields_before"] = fields_before
    snap["fields_before_count"] = len(fields_before)
    snap["unique_ids"] = len(set(field_def.id for field_def in fields_from_scene))
    snap["duplicate_ids"] = {
        field_id: count
        for field_id, count in __import__("collections").Counter(
            field_def.id for field_def in fields_from_scene
        ).items()
        if count > 1
    }

    try:
        raw_auto_cells = detect_cells(fitz_page, cfg=cfg_eff)
        snap["auto_detect_cells"] = [
            {
                "x0": round(cell[0], 3),
                "y0": round(cell[1], 3),
                "x1": round(cell[2], 3),
                "y1": round(cell[3], 3),
            }
            for cell in raw_auto_cells
        ]
        snap["auto_detect_count"] = len(raw_auto_cells)
    except Exception as exc:
        snap["auto_detect_cells"] = []
        snap["auto_detect_error"] = str(exc)

    stamp_bbox = _template_stamp_bbox(template, frame)
    snap["stamp_search_bbox"] = _rect_to_list(stamp_bbox)
    snap["frame_detection"] = (
        {
            **frame_detection,
            "stamp_search_bbox_pts": snap.get("stamp_search_bbox"),
        }
        if frame_detection is not None
        else None
    )

    detected_cells_base: list[CellBbox] = []
    if stamp_bbox is not None:
        try:
            detected_cells_base = get_detected_stamp_cells(
                fitz_page=fitz_page,
                stamp_bbox=stamp_bbox,
                tolerance_mm=template.grid_tolerance_detected_mm,
                cfg=cfg_eff,
                template=template,
            )
        except Exception as exc:
            snap["detected_cells_error"] = str(exc)
    snap["detected_stamp_cells"] = [_bbox_to_list(cell) for cell in detected_cells_base]
    snap["detected_stamp_cells_count"] = len(detected_cells_base)

    if detected_cells_base:
        x_lines, y_lines = _extract_grid_lines(
            detected_cells_base,
            template.grid_tolerance_detected_mm,
        )
        snap["grid_lines_detected_from_base_bbox"] = {
            "x": [round(value, 3) for value in x_lines],
            "y": [round(value, 3) for value in y_lines],
            "x_count": len(x_lines),
            "y_count": len(y_lines),
        }
        h_det, v_det = _build_detected_grid_lines(
            detected_cells_base,
            template.grid_tolerance_detected_mm,
        )
        snap["detected_grid_lines_enriched"] = {
            "h_lines": [
                {
                    "pos": round(line.pos, 3),
                    "span_lo": round(line.span_lo, 3),
                    "span_hi": round(line.span_hi, 3),
                    "cell_count": line.cell_count,
                    "span_length": round(line.span_length, 3),
                }
                for line in h_det
            ],
            "v_lines": [
                {
                    "pos": round(line.pos, 3),
                    "span_lo": round(line.span_lo, 3),
                    "span_hi": round(line.span_hi, 3),
                    "cell_count": line.cell_count,
                    "span_length": round(line.span_length, 3),
                }
                for line in v_det
            ],
            "h_count": len(h_det),
            "v_count": len(v_det),
        }
    else:
        snap["grid_lines_detected_from_base_bbox"] = {
            "x": [],
            "y": [],
            "x_count": 0,
            "y_count": 0,
        }
        snap["detected_grid_lines_enriched"] = {
            "h_lines": [],
            "v_lines": [],
            "h_count": 0,
            "v_count": 0,
        }

    adapted_list: list[tuple[str, Any]] = []
    ca_info = None
    try:
        adapted_list, ca_info = adapt_by_cell_assignment(
            template=template,
            frame=frame,
            fitz_page=fitz_page,
            cfg=cfg_eff,
        )
    except Exception as exc:
        snap["adaptation_error"] = str(exc)

    production_detected_raw: list[CellBbox] = []
    production_detected_probe: list[CellBbox] = []
    production_detected_effective: list[CellBbox] = []
    if ca_info is not None:
        production_detected_raw, production_detected_probe, production_detected_effective = (
            _build_production_detected_cells(
                fitz_page=fitz_page,
                template=template,
                search_bbox=_tuple_rect_to_list(ca_info.search_bbox),
                effective_bbox=_tuple_rect_to_list(ca_info.effective_bbox),
                cfg=cfg_eff,
            )
        )
        snap["cell_assignment_info"] = {
            "rectangularity": round(ca_info.rectangularity, 6),
            "n_detected": ca_info.n_detected,
            "n_fields": ca_info.n_fields,
            "n_assigned": ca_info.n_assigned,
            "transform_scale_x": round(ca_info.transform_scale_x, 6),
            "transform_scale_y": round(ca_info.transform_scale_y, 6),
            "transform_dx": round(ca_info.transform_dx, 3),
            "transform_dy": round(ca_info.transform_dy, 3),
            "warnings": list(ca_info.warnings),
        }
        snap["prealign"] = prealign_dict_from_cell_assignment_info(
            ca_info,
            production_detected_cells={
                "raw": [_bbox_to_list(cell) for cell in production_detected_raw],
                "probe_filtered": [_bbox_to_list(cell) for cell in production_detected_probe],
                "effective": [_bbox_to_list(cell) for cell in production_detected_effective],
            },
        )
        snap["detected_grid_lines_production"] = [
            _serialize_detected_grid_line(line)
            for line in ca_info.detected_grid_lines
        ]
    else:
        snap["cell_assignment_info"] = None
        snap["prealign"] = None
        snap["detected_grid_lines_production"] = []

    snap["grid_lines_template"] = [
        _serialize_template_grid_line(line, frame, template.origin)
        for line in template.grid_lines
    ]
    snap["grid_lines_template_count"] = len(template.grid_lines)

    align_diags: list[AlignmentLineDiag] = []
    align_summaries: list[AlignmentSummary] = []
    align_reports: list[AlignmentReport] = []
    production_snapped = dict(ca_info.snapped_line_positions) if ca_info is not None else {}
    align_positions = dict(production_snapped)
    if template.grid_lines and ca_info is not None and production_detected_effective:
        try:
            align_positions = _align_grid_lines_ordered(
                template,
                frame,
                ca_info.transform_scale_x,
                ca_info.transform_scale_y,
                ca_info.transform_dx,
                ca_info.transform_dy,
                production_detected_effective,
                diagnostics=align_diags,
                summary_out=align_summaries,
                alignment_report_out=align_reports,
                template_path=template.source_path,
                pdf_path=pdf_path,
                page_num=page_num,
                stamp_bbox_override=(
                    fitz.Rect(*ca_info.effective_bbox)
                    if ca_info.proposed_bbox_applied and ca_info.effective_bbox is not None
                    else None
                ),
            )
        except Exception as exc:
            snap["grid_lines_alignment_error"] = str(exc)

    diag_by_id = {diag.line_id: diag for diag in align_diags}
    snap["alignment_diagnostics"] = [
        _serialize_alignment_diag(diag)
        for diag in align_diags
    ]
    if align_reports:
        snap["_alignment_log_text"] = align_reports[0].log
    elif ca_info is not None and ca_info.alignment_report is not None:
        snap["_alignment_log_text"] = ca_info.alignment_report.log

    snap["grid_lines_snapped"] = []
    for line_payload in snap["grid_lines_template"]:
        line_id = str(line_payload["id"])
        diag = diag_by_id.get(line_id)
        snapped_pts = production_snapped.get(line_id, align_positions.get(line_id))
        snap["grid_lines_snapped"].append({
            "id": line_id,
            "orientation": line_payload["orientation"],
            "boundary": line_payload.get("boundary"),
            "pos_template_mm": line_payload["pos_mm"],
            "display_pos_template_pts": line_payload["display_pos_pts"],
            "pos_snapped_pts": round(snapped_pts, 3) if snapped_pts is not None else None,
            **(
                {
                    "alignment_method": diag.snap_method,
                    "alignment_interp_pts": round(diag.pos_interp_pts, 3),
                    "alignment_pos_pts": round(diag.pos_snapped_pts, 3),
                    "alignment_match_score": round(diag.match_score, 6),
                    "alignment_span_ratio": round(diag.span_ratio, 6),
                    "alignment_dp_action": diag.dp_action,
                    "matched_to": diag.matched_detected_id,
                    "alignment_detected_pos": (
                        round(diag.matched_detected_pos, 3)
                        if diag.matched_detected_pos is not None
                        else None
                    ),
                    "alignment_detected_span": (
                        round(diag.matched_detected_span, 3)
                        if diag.matched_detected_span is not None
                        else None
                    ),
                }
                if diag is not None
                else {}
            ),
        })

    summary_src = align_summaries[0] if align_summaries else None
    snap["grid_alignment_summary"] = (
        {
            "boundary_matched": summary_src.boundary_matched,
            "alignment_matched": summary_src.alignment_matched,
            "template_skipped": summary_src.template_skipped,
            "detected_skipped": summary_src.detected_skipped,
            "avg_match_score": round(summary_src.avg_match_score, 6),
        }
        if summary_src is not None
        else None
    )

    report_src = (
        align_reports[0]
        if align_reports
        else (ca_info.alignment_report if ca_info is not None else None)
    )
    snap["alignment_report"] = (
        {
            "verdict": report_src.verdict,
            "total_tpl_lines": report_src.total_tpl_lines,
            "matched_count": report_src.matched_count,
            "no_match_count": report_src.no_match_count,
            "skipped_detected_count": report_src.skipped_detected_count,
            "avg_match_score": round(report_src.avg_match_score, 6),
            "problem_lines": list(report_src.problem_lines),
            "message": report_src.message,
        }
        if report_src is not None
        else None
    )

    snap["field_bindings"] = [
        {
            "id": field.id,
            "bound_top": field.bound_top,
            "bound_bottom": field.bound_bottom,
            "bound_left": field.bound_left,
            "bound_right": field.bound_right,
        }
        for field in fields_from_scene
        if field.bound_top or field.bound_bottom or field.bound_left or field.bound_right
    ]
    snap["fields_with_bindings"] = len(snap["field_bindings"])
    snap["grid_binding_active"] = bool(
        template.grid_lines and _fields_have_bindings(list(fields_from_scene))
    )

    key_to_field = {
        keys[idx]: field_def
        for idx, field_def in enumerate(fields_from_scene)
    }
    adaptation_results = []
    for key, adapt_result in adapted_list:
        field_before = key_to_field.get(key)
        fitz_before = None
        if field_before is not None:
            fitz_before = _rect_to_list(field_to_fitz_rect(field_before, frame, padding_mm=0.0))
        adaptation_results.append({
            "key": key,
            "field_id": adapt_result.field_id,
            "status": adapt_result.status,
            "bbox_before_fitz": fitz_before,
            "approx_bbox": _bbox_to_list(adapt_result.approx_bbox),
            "bbox_after": _bbox_to_list(adapt_result.bbox),
            "snap_distance_mm": round(adapt_result.snap_distance_mm, 6),
            "shape_score": round(adapt_result.shape_score, 6),
            "snapped_boundaries": adapt_result.snapped_boundaries,
            "confidence": round(adapt_result.confidence, 6),
            "iteration": adapt_result.iteration,
            "matched_detected": [
                _bbox_to_list(cell)
                for cell in adapt_result.matched_detected
            ],
        })
    snap["adaptation_results"] = adaptation_results

    status_counts: dict[str, int] = {}
    all_results = [adapt_result for _, adapt_result in adapted_list]
    for adapt_result in all_results:
        status_counts[adapt_result.status] = status_counts.get(adapt_result.status, 0) + 1
    snap["adaptation_summary"] = {
        "total_fields": len(adapted_list),
        "status_counts": status_counts,
        "grid_bound_count": status_counts.get("grid_bound", 0),
        "fields_with_score_above_07": sum(
            1 for result in all_results if result.confidence > 0.7
        ),
        "fields_with_score_below_04": sum(
            1
            for result in all_results
            if result.confidence < 0.4 and result.status not in ("excluded",)
        ),
    }

    return snap


def _build_artifact_base_name(snapshot: dict[str, Any]) -> str:
    env = snapshot.get("environment", {})
    template = snapshot.get("template", {})
    pdf_part = _sanitize_artifact_part(env.get("pdf_basename", ""), "pdf")
    template_part = _sanitize_artifact_part(
        template.get("source_basename", "") or template.get("name", ""),
        "template",
    )
    timestamp_part = time.strftime("%Y_%m_%d__%H%M%S")
    return f"adapt_debug_{timestamp_part}__{pdf_part}__{template_part}"


def _build_prealign_log(snapshot: dict[str, Any]) -> str:
    env = snapshot.get("environment", {})
    template = snapshot.get("template", {})
    prealign = snapshot.get("prealign") or {}
    extent_probe = (prealign.get("extent_probe") or {})
    selected = extent_probe.get("selected_boundaries") or {}
    virtual = extent_probe.get("virtual_edges") or {}
    counts = prealign.get("detected_counts") or {}
    lines = [
        "=== ADAPT DEBUG PREALIGN ===",
        f"Template: {template.get('name', '')}",
        f"Template path: {template.get('source_path', '') or 'n/a'}",
        f"PDF: {env.get('pdf_basename', '')} (page {env.get('page_num', 0)})",
        f"Search bbox: {_format_rect(prealign.get('search_bbox'))}",
        f"Base bbox: {_format_rect(prealign.get('base_bbox'))}",
        f"Proposed bbox: {_format_rect(prealign.get('proposed_bbox'))}",
        f"Effective bbox: {_format_rect(prealign.get('effective_bbox'))}",
        f"Detected union bbox: {_format_rect(prealign.get('detected_union_bbox'))}",
        (
            "Counts: "
            f"raw={counts.get('raw', 0)} "
            f"probe={counts.get('probe', 0)} "
            f"effective={counts.get('effective', 0)} "
            f"oversized_filtered={counts.get('oversized_filtered', 0)} "
            f"outside_effective_bbox_filtered={counts.get('outside_effective_bbox_filtered', 0)}"
        ),
        (
            "Proposed bbox applied: "
            f"{bool(prealign.get('proposed_bbox_applied', False))}"
        ),
        (
            "Adaptive threshold: "
            f"{prealign.get('scale_threshold', 'n/a')} "
            f"(tier={prealign.get('prealign_confidence_tier', '') or 'n/a'}, "
            f"score={prealign.get('prealign_confidence_score', 'n/a')})"
        ),
        f"Reason: {prealign.get('prealign_confidence_reason', '') or 'n/a'}",
        "Inferred edges: " + ", ".join(extent_probe.get("inferred_edges", [])) or "n/a",
        "",
        "BOUNDARY-LIKE LINES:",
    ]
    for edge in ("top", "left", "bottom", "right"):
        payload = selected.get(edge) or {}
        if payload:
            lines.append(
                f"  {edge}: id={payload.get('id', '')} "
                f"pos={payload.get('pos_pts', '')} "
                f"span={payload.get('span_pts', '')} "
                f"ratio={payload.get('span_ratio_to_template', '')}"
            )
        else:
            lines.append(f"  {edge}: n/a")
    lines.extend([
        "",
        "VIRTUAL EDGES:",
    ])
    for edge in ("top", "left"):
        payload = virtual.get(edge) or {}
        if payload:
            lines.append(
                f"  {edge}: virtual_pos={payload.get('virtual_pos_pts', '')} "
                f"matched_template={payload.get('matched_template_id', '')} "
                f"source_detected={payload.get('source_detected_id', '')} "
                f"score={payload.get('score', '')}"
            )
        else:
            lines.append(f"  {edge}: n/a")
    return "\n".join(lines).strip() + "\n"


def _open_snapshot_page(snapshot: dict[str, Any]) -> tuple[fitz.Document, fitz.Page] | None:
    env = snapshot.get("environment", {})
    pdf_path = env.get("pdf_path", "")
    page_num = int(env.get("page_num", 0) or 0)
    if not pdf_path or page_num <= 0 or not os.path.isfile(pdf_path):
        return None
    doc = fitz.open(pdf_path)
    if page_num - 1 >= len(doc):
        doc.close()
        return None
    return doc, doc[page_num - 1]


def _render_frame_and_search_bbox(snapshot: dict[str, Any], out_path: str) -> None:
    opened = _open_snapshot_page(snapshot)
    if opened is None:
        return
    doc, fitz_page = opened
    try:
        overlay_doc, overlay_page, scale = _overlay_doc_for_page(fitz_page)
        try:
            env = snapshot["environment"]
            frame_rect = _rect_to_list(
                _frame_rect_to_fitz(env["frame"], int(env["rotation"]))
            )
            prealign = snapshot.get("prealign") or {}
            _draw_rect(overlay_page, scale, frame_rect, _GREEN, width=2.0)
            _draw_rect(overlay_page, scale, prealign.get("base_bbox"), _BLUE, width=1.8)
            _draw_rect(
                overlay_page,
                scale,
                prealign.get("search_bbox"),
                _ORANGE,
                width=1.8,
                dashes="[6 4] 0",
            )
            _draw_rect(overlay_page, scale, prealign.get("effective_bbox"), _MAGENTA, width=1.5)
            _insert_text(overlay_page, scale, (12.0, 16.0), "green=frame", fontsize=8.0, color=_GREEN)
            _insert_text(overlay_page, scale, (12.0, 28.0), "blue=base bbox", fontsize=8.0, color=_BLUE)
            _insert_text(
                overlay_page,
                scale,
                (12.0, 40.0),
                "orange dashed=search bbox",
                fontsize=8.0,
                color=_ORANGE,
            )
            _insert_text(
                overlay_page,
                scale,
                (12.0, 52.0),
                "magenta=effective bbox",
                fontsize=8.0,
                color=_MAGENTA,
            )
            overlay_page.get_pixmap().save(out_path)
        finally:
            overlay_doc.close()
    finally:
        doc.close()


def _render_extent_probe_boundaries(snapshot: dict[str, Any], out_path: str) -> None:
    opened = _open_snapshot_page(snapshot)
    if opened is None:
        return
    doc, fitz_page = opened
    try:
        overlay_doc, overlay_page, scale = _overlay_doc_for_page(fitz_page)
        try:
            prealign = snapshot.get("prealign") or {}
            extent_probe = prealign.get("extent_probe") or {}
            _draw_rect(
                overlay_page,
                scale,
                prealign.get("search_bbox"),
                _ORANGE,
                width=1.5,
                dashes="[6 4] 0",
            )
            for line in prealign.get("prealign_detected_grid_lines", []):
                _draw_detected_grid_line(
                    overlay_page,
                    scale,
                    line,
                    _GREY,
                    width=0.9,
                )
            selected = extent_probe.get("selected_boundaries") or {}
            detected_by_id = {
                line["id"]: line
                for line in prealign.get("prealign_detected_grid_lines", [])
            }
            colors = {
                "top": _RED,
                "left": _BLUE,
                "bottom": _GREEN,
                "right": _PURPLE,
            }
            for edge, color in colors.items():
                payload = dict(selected.get(edge) or {})
                if not payload or not payload.get("id"):
                    continue
                detected_line = detected_by_id.get(str(payload["id"]))
                if detected_line is None:
                    continue
                _draw_detected_grid_line(
                    overlay_page,
                    scale,
                    detected_line,
                    color,
                    width=2.0,
                    label=True,
                )
            virtual = extent_probe.get("virtual_edges") or {}
            top_virtual = virtual.get("top") or {}
            if top_virtual.get("virtual_pos_pts") not in ("", None):
                y = float(top_virtual["virtual_pos_pts"])
                _draw_line(
                    overlay_page,
                    scale,
                    (prealign["search_bbox"][0], y),
                    (prealign["search_bbox"][2], y),
                    _MAGENTA,
                    width=1.6,
                    dashes="[3 3] 0",
                )
                _insert_text(
                    overlay_page,
                    scale,
                    (prealign["search_bbox"][0], y - 2.0),
                    f"virtual top -> {top_virtual.get('source_detected_id', '')}",
                    fontsize=7.0,
                    color=_MAGENTA,
                )
            left_virtual = virtual.get("left") or {}
            if left_virtual.get("virtual_pos_pts") not in ("", None):
                x = float(left_virtual["virtual_pos_pts"])
                _draw_line(
                    overlay_page,
                    scale,
                    (x, prealign["search_bbox"][1]),
                    (x, prealign["search_bbox"][3]),
                    _CYAN,
                    width=1.6,
                    dashes="[3 3] 0",
                )
                _insert_text(
                    overlay_page,
                    scale,
                    (x + 1.5, prealign["search_bbox"][1] + 8.0),
                    f"virtual left -> {left_virtual.get('source_detected_id', '')}",
                    fontsize=7.0,
                    color=_CYAN,
                )
            overlay_page.get_pixmap().save(out_path)
        finally:
            overlay_doc.close()
    finally:
        doc.close()


def _render_proposed_vs_effective_bbox(snapshot: dict[str, Any], out_path: str) -> None:
    opened = _open_snapshot_page(snapshot)
    if opened is None:
        return
    doc, fitz_page = opened
    try:
        overlay_doc, overlay_page, scale = _overlay_doc_for_page(fitz_page)
        try:
            prealign = snapshot.get("prealign") or {}
            _draw_rect(overlay_page, scale, prealign.get("base_bbox"), _BLUE, width=1.7)
            _draw_rect(
                overlay_page,
                scale,
                prealign.get("proposed_bbox"),
                _PURPLE,
                width=1.7,
                dashes="[6 3] 0",
            )
            _draw_rect(overlay_page, scale, prealign.get("effective_bbox"), _GREEN, width=1.9)
            _draw_rect(
                overlay_page,
                scale,
                prealign.get("search_bbox"),
                _ORANGE,
                width=1.2,
                dashes="[3 5] 0",
            )
            _insert_text(overlay_page, scale, (12.0, 16.0), "blue=base", fontsize=8.0, color=_BLUE)
            _insert_text(
                overlay_page,
                scale,
                (12.0, 28.0),
                "purple dashed=proposed",
                fontsize=8.0,
                color=_PURPLE,
            )
            _insert_text(
                overlay_page,
                scale,
                (12.0, 40.0),
                "green=effective",
                fontsize=8.0,
                color=_GREEN,
            )
            overlay_page.get_pixmap().save(out_path)
        finally:
            overlay_doc.close()
    finally:
        doc.close()


def _render_ordered_alignment_overlay(snapshot: dict[str, Any], out_path: str) -> None:
    opened = _open_snapshot_page(snapshot)
    if opened is None:
        return
    doc, fitz_page = opened
    try:
        overlay_doc, overlay_page, scale = _overlay_doc_for_page(fitz_page)
        try:
            prealign = snapshot.get("prealign") or {}
            _draw_rect(overlay_page, scale, prealign.get("effective_bbox"), _GREY, width=1.0)
            template_lines = {
                line["id"]: line
                for line in snapshot.get("grid_lines_template", [])
                if line.get("boundary")
            }
            matched_detected_ids = {
                line.get("matched_to")
                for line in snapshot.get("grid_lines_snapped", [])
                if line.get("matched_to")
            }
            detected_by_id = {
                line["id"]: line
                for line in snapshot.get("detected_grid_lines_production", [])
            }
            for detected_id in matched_detected_ids:
                detected_line = detected_by_id.get(detected_id)
                if detected_line is None:
                    continue
                _draw_detected_grid_line(
                    overlay_page,
                    scale,
                    detected_line,
                    _MAGENTA,
                    width=2.0,
                    label=True,
                )
            for line in snapshot.get("grid_lines_snapped", []):
                if not line.get("boundary"):
                    continue
                template_line = template_lines.get(line["id"])
                if template_line is None:
                    continue
                snapped_pos = line.get("pos_snapped_pts")
                _draw_template_grid_line(
                    overlay_page,
                    scale,
                    template_line,
                    _GREEN,
                    snapped_pos_pts=snapped_pos,
                    width=1.8,
                    label=True,
                )
                if snapped_pos is not None:
                    if template_line["orientation"] == "h":
                        marker = (template_line["display_start_pts"], float(snapped_pos))
                    else:
                        marker = (float(snapped_pos), template_line["display_start_pts"])
                    _draw_circle(overlay_page, scale, marker, 1.5, _GREEN)
            _insert_text(
                overlay_page,
                scale,
                (12.0, 16.0),
                "green=template boundary snapped",
                fontsize=8.0,
                color=_GREEN,
            )
            _insert_text(
                overlay_page,
                scale,
                (12.0, 28.0),
                "magenta=matched detected boundary",
                fontsize=8.0,
                color=_MAGENTA,
            )
            overlay_page.get_pixmap().save(out_path)
        finally:
            overlay_doc.close()
    finally:
        doc.close()


def _render_snapshot_pngs(
    snapshot: dict[str, Any],
    output_dir: str,
    base_name: str,
) -> dict[str, str]:
    renderers = {
        "frame_and_search_bbox": _render_frame_and_search_bbox,
        "extent_probe_boundaries": _render_extent_probe_boundaries,
        "proposed_vs_effective_bbox": _render_proposed_vs_effective_bbox,
        "ordered_alignment_overlay": _render_ordered_alignment_overlay,
    }
    created: dict[str, str] = {}
    for key, renderer in renderers.items():
        file_name = f"{base_name}__{key}.png"
        out_path = os.path.join(output_dir, file_name)
        fd, tmp_path = tempfile.mkstemp(suffix=".png")
        os.close(fd)
        try:
            renderer(snapshot, tmp_path)
            if os.path.isfile(tmp_path):
                shutil.copyfile(tmp_path, out_path)
                created[key] = file_name
        except Exception:
            pass
        finally:
            try:
                os.remove(tmp_path)
            except OSError:
                pass
    return created


def save_adapt_debug_snapshot(
    snapshot: dict[str, Any],
    output_dir: str | None = None,
    alignment_log: str = "",
) -> str:
    """Write a timestamped JSON/TXT/PNG dump and return the JSON path."""
    if output_dir is None:
        import tempfile
        output_dir = tempfile.gettempdir()
    os.makedirs(output_dir, exist_ok=True)

    base_name = _build_artifact_base_name(snapshot)
    snapshot["artifact_basename"] = base_name

    prealign_log = _build_prealign_log(snapshot)
    if not alignment_log:
        alignment_log = snapshot.get("_alignment_log_text", "")
    snapshot.pop("_alignment_log_text", None)
    combined_log = prealign_log
    if alignment_log:
        combined_log += "\n=== GRID ALIGNMENT LOG ===\n"
        if alignment_log.startswith("=== GRID ALIGNMENT LOG ==="):
            combined_log += alignment_log.split("=== GRID ALIGNMENT LOG ===", 1)[1].lstrip("\n")
        else:
            combined_log += alignment_log

    txt_file_name = f"{base_name}.txt"
    txt_path = os.path.join(output_dir, txt_file_name)
    with open(txt_path, "w", encoding="utf-8") as handle:
        handle.write(combined_log)

    png_files = _render_snapshot_pngs(snapshot, output_dir, base_name)
    json_file_name = f"{base_name}.json"
    json_path = os.path.join(output_dir, json_file_name)

    snapshot["artifacts"] = {
        "json": json_file_name,
        "txt": txt_file_name,
        "png": png_files,
    }
    with open(json_path, "w", encoding="utf-8") as handle:
        json.dump(snapshot, handle, ensure_ascii=False, indent=2)
    return json_path
