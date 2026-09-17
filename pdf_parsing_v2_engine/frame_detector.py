"""
Унифицированная детекция рамки по fitz.get_drawings().

Портирован из pdf_parsing/find_function_MTO_BBB.py::BiggestElement.find_big_element_fitz.
Один путь для всех типов документов (DWG, MTO, BBB, OD).
Не использует глобальные переменные — результат инкапсулирован в FrameInfo.

Опционально (шаблон): ``frame_mode=drawing_union`` — union AABB отфильтрованных rect
внутри ROI (печатное поле ГОСТ ∩ раздутый bbox полей штампа относительно GOST-frame).
"""

from __future__ import annotations

from dataclasses import dataclass
import time

import fitz

from pdf_parsing_v2_engine.coord_transform import (
    field_to_fitz_rect,
    fitz_rect_transform_by_matrix,
    unrotated_fitz_rect_to_pdfminer_bbox,
)
from pdf_parsing_v2_engine.models import FrameInfo, StampTemplate, normalize_frame_mode

SCALE: float = 2.83444  # pts/mm

_MIN_FRAME_FRACTION = 0.20  # frame must span ≥ 20 % of page in each dimension

# GOST 2.301-68 standard margins (mm)
_GOST_LEFT_MARGIN_MM = 20.0
_GOST_OTHER_MARGIN_MM = 5.0

# Large-sheet optimization: for pages larger than A3 by 50% on either axis,
# only keep drawings that touch the outer border band.
_LARGE_PAGE_BORDER_FILTER_MIN_W_MM = 420.0 * 1.5
_LARGE_PAGE_BORDER_FILTER_MIN_H_MM = 297.0 * 1.5
_LARGE_PAGE_BORDER_FILTER_BAND_MM = 50.0

# Extra margin around template inner fields when building ROI for drawing_union (mm)
_DRAWING_UNION_FIELD_MARGIN_MM = 25.0


def _timing_add(timing: dict[str, float] | None, key: str, elapsed_sec: float) -> None:
    if timing is None:
        return
    timing[key] = round(float(timing.get(key, 0.0)) + elapsed_sec, 6)


@dataclass(frozen=True)
class _FrameRectCandidate:
    """Cached frame candidate with precomputed bbox/span in pdfminer space."""

    fitz_rect: fitz.Rect
    pm_bbox: tuple[float, float, float, float]
    x_range: float
    y_range: float
    area_pts2: float


@dataclass(frozen=True)
class _FrameCandidateStats:
    """Stats for candidate collection before frame selection."""

    total_drawings: int
    kept_after_edge_filter: int
    border_band_enabled: bool
    dropped_by_border_band: int
    kept_after_border_band: int


def _serialize_fitz_rect(r: fitz.Rect | None) -> list[float] | None:
    if r is None:
        return None
    return [round(r.x0, 4), round(r.y0, 4), round(r.x1, 4), round(r.y1, 4)]


def _fitz_rects_equal(a: fitz.Rect | None, b: fitz.Rect | None, eps: float = 1e-3) -> bool:
    if a is None or b is None:
        return a is None and b is None
    return (
        abs(a.x0 - b.x0) < eps
        and abs(a.y0 - b.y0) < eps
        and abs(a.x1 - b.x1) < eps
        and abs(a.y1 - b.y1) < eps
    )


def _rect_area_pdfminer_unrot(r: fitz.Rect, fitz_page: fitz.Page) -> float:
    x0, y0, x1, y1 = unrotated_fitz_rect_to_pdfminer_bbox(r, fitz_page)
    return max(0.0, x1 - x0) * max(0.0, y1 - y0)


def _serialize_pm_bbox_unrot(
    r: fitz.Rect | None,
    fitz_page: fitz.Page,
) -> list[float] | None:
    if r is None:
        return None
    t = unrotated_fitz_rect_to_pdfminer_bbox(r, fitz_page)
    return [round(t[0], 4), round(t[1], 4), round(t[2], 4), round(t[3], 4)]


def _serialize_pm_bbox(pm_bbox: tuple[float, float, float, float] | None) -> list[float] | None:
    if pm_bbox is None:
        return None
    return [round(pm_bbox[0], 4), round(pm_bbox[1], 4), round(pm_bbox[2], 4), round(pm_bbox[3], 4)]


def _use_large_page_border_filter(page_w: float, page_h: float) -> bool:
    """Return True for sheets larger than A3 by 50% on either axis."""
    width_mm = page_w / SCALE
    height_mm = page_h / SCALE
    return (
        width_mm > _LARGE_PAGE_BORDER_FILTER_MIN_W_MM
        or height_mm > _LARGE_PAGE_BORDER_FILTER_MIN_H_MM
    )


def _border_band_margins_mm(cfg: dict[str, float] | None) -> tuple[float, float, float, float]:
    """Return per-side border band margins in mm (left, right, top, bottom)."""
    src = cfg or {}
    left_mm = float(src.get("find_frame_border_band_left_mm", 20.0))
    right_mm = float(src.get("find_frame_border_band_right_mm", 50.0))
    top_mm = float(src.get("find_frame_border_band_top_mm", 20.0))
    bottom_mm = float(src.get("find_frame_border_band_bottom_mm", 20.0))
    return left_mm, right_mm, top_mm, bottom_mm


def _touches_border_band(
    displayed_rect: fitz.Rect,
    page_w: float,
    page_h: float,
    *,
    left_mm: float,
    right_mm: float,
    top_mm: float,
    bottom_mm: float,
) -> bool:
    """Keep objects that intersect the outer border band on any side."""
    left_pts = left_mm * SCALE
    right_pts = right_mm * SCALE
    top_pts = top_mm * SCALE
    bottom_pts = bottom_mm * SCALE
    return (
        displayed_rect.x0 <= left_pts
        or displayed_rect.y0 <= bottom_pts
        or displayed_rect.x1 >= (page_w - right_pts)
        or displayed_rect.y1 >= (page_h - top_pts)
    )


def _collect_filtered_frame_candidates(
    fitz_page: fitz.Page,
    cfg: dict[str, float] | None = None,
) -> tuple[list[_FrameRectCandidate], _FrameCandidateStats]:
    """Collect frame candidates once with cached bbox/span/area in pdfminer space."""
    page_w: float = fitz_page.rect.width
    page_h: float = fitz_page.rect.height
    rotation: int = fitz_page.rotation % 360
    use_border_band = _use_large_page_border_filter(page_w, page_h)
    left_mm, right_mm, top_mm, bottom_mm = _border_band_margins_mm(cfg)
    out: list[_FrameRectCandidate] = []
    total_drawings = 0
    kept_after_edge_filter = 0
    dropped_by_border_band = 0
    for drawing in fitz_page.get_drawings():
        total_drawings += 1
        raw_rect = fitz.Rect(drawing["rect"])
        if rotation in (90, 270):
            displayed_rect = fitz_rect_transform_by_matrix(raw_rect, fitz_page.rotation_matrix)
        else:
            displayed_rect = raw_rect
        if displayed_rect.x1 >= (page_w - SCALE) or displayed_rect.y0 <= SCALE:
            continue
        kept_after_edge_filter += 1
        if use_border_band and not _touches_border_band(
            displayed_rect,
            page_w,
            page_h,
            left_mm=left_mm,
            right_mm=right_mm,
            top_mm=top_mm,
            bottom_mm=bottom_mm,
        ):
            dropped_by_border_band += 1
            continue
        pm_bbox = unrotated_fitz_rect_to_pdfminer_bbox(raw_rect, fitz_page)
        x_range = pm_bbox[2] - pm_bbox[0]
        y_range = pm_bbox[3] - pm_bbox[1]
        area_pts2 = max(0.0, x_range) * max(0.0, y_range)
        out.append(
            _FrameRectCandidate(
                fitz_rect=fitz.Rect(raw_rect),
                pm_bbox=pm_bbox,
                x_range=x_range,
                y_range=y_range,
                area_pts2=area_pts2,
            )
        )
    stats = _FrameCandidateStats(
        total_drawings=total_drawings,
        kept_after_edge_filter=kept_after_edge_filter,
        border_band_enabled=use_border_band,
        dropped_by_border_band=dropped_by_border_band,
        kept_after_border_band=len(out),
    )
    return out, stats


def collect_filtered_drawings_rects(fitz_page: fitz.Page) -> list[fitz.Rect]:
    """Rects from ``get_drawings()`` passing the same edge filter as ``find_frame`` (gost)."""
    candidates, _stats = _collect_filtered_frame_candidates(fitz_page)
    return [fitz.Rect(c.fitz_rect) for c in candidates]


def _fitz_displayed_rect_to_pdfminer_bbox(rect: fitz.Rect, page_h: float) -> tuple[float, float, float, float]:
    return (rect.x0, page_h - rect.y1, rect.x1, page_h - rect.y0)


def _intersect_pm_boxes(
    a: tuple[float, float, float, float],
    b: tuple[float, float, float, float],
) -> tuple[float, float, float, float] | None:
    x0 = max(a[0], b[0])
    y0 = max(a[1], b[1])
    x1 = min(a[2], b[2])
    y1 = min(a[3], b[3])
    if x0 >= x1 - 1e-9 or y0 >= y1 - 1e-9:
        return None
    return (x0, y0, x1, y1)


def _pm_intersects(a: tuple[float, float, float, float], b: tuple[float, float, float, float]) -> bool:
    return _intersect_pm_boxes(a, b) is not None


def _drawing_union_roi_pdfminer(
    fitz_page: fitz.Page,
    gost: FrameInfo,
    template: StampTemplate | None,
) -> tuple[tuple[float, float, float, float], str]:
    """ROI in pdfminer coords and a short source tag for diagnostics."""
    page_w = fitz_page.rect.width
    page_h = fitz_page.rect.height
    gost_pm = (gost.x0, gost.y0, gost.x1, gost.y1)
    if template is None:
        return gost_pm, "gost_print_area_only"
    inner_rects: list[fitz.Rect] = []
    for f in template.fields:
        if f.outside_stamp or f.document_property:
            continue
        r = field_to_fitz_rect(f, gost, padding_mm=0.0)
        if not r.is_empty and not r.is_infinite:
            inner_rects.append(r)
    if not inner_rects:
        return gost_pm, "gost_print_area_no_inner_fields"
    m = _DRAWING_UNION_FIELD_MARGIN_MM * SCALE
    sx0 = max(0.0, min(r.x0 for r in inner_rects) - m)
    sy0 = max(0.0, min(r.y0 for r in inner_rects) - m)
    sx1 = min(page_w, max(r.x1 for r in inner_rects) + m)
    sy1 = min(page_h, max(r.y1 for r in inner_rects) + m)
    stamp_pm = _fitz_displayed_rect_to_pdfminer_bbox(fitz.Rect(sx0, sy0, sx1, sy1), page_h)
    tight = _intersect_pm_boxes(gost_pm, stamp_pm)
    if tight is None:
        return gost_pm, "gost_print_area_stamp_intersect_failed"
    return tight, "gost_x_stamp_fields_expanded"


def _frame_info_from_union_pm(
    u0: float,
    u1: float,
    v0: float,
    v1: float,
    page_w: float,
    page_h: float,
    rotation: int,
) -> FrameInfo:
    border_left_mm = (page_w - u1) / SCALE
    border_bottom_mm = v0 / SCALE
    border_top_mm = (page_h - v1) / SCALE
    return FrameInfo(
        x0=u0,
        y0=v0,
        x1=u1,
        y1=v1,
        page_width=page_w,
        page_height=page_h,
        rotation=rotation,
        border_left_mm=border_left_mm,
        border_bottom_mm=border_bottom_mm,
        border_top_mm=border_top_mm,
    )


def _find_frame_drawing_union(
    fitz_page: fitz.Page,
    debug: bool,
    template: StampTemplate | None,
    timing: dict[str, float] | None = None,
    cfg: dict[str, float] | None = None,
) -> tuple[FrameInfo, dict | None]:
    t_total = time.perf_counter()
    page_w = fitz_page.rect.width
    page_h = fitz_page.rect.height
    rotation = fitz_page.rotation % 360
    gost = _gost_fallback(page_w, page_h, rotation)
    t_roi = time.perf_counter()
    roi_pm, roi_src = _drawing_union_roi_pdfminer(fitz_page, gost, template)
    t_roi_elapsed = time.perf_counter() - t_roi
    _timing_add(timing, "find_frame_roi_build", t_roi_elapsed)
    t_filter = time.perf_counter()
    filtered, candidate_stats = _collect_filtered_frame_candidates(fitz_page, cfg=cfg)
    t_filter_elapsed = time.perf_counter() - t_filter
    _timing_add(timing, "find_frame_filter_rects", t_filter_elapsed)
    if timing is not None:
        timing["find_frame_total_drawings"] = round(
            float(timing.get("find_frame_total_drawings", 0.0)) + candidate_stats.total_drawings,
            3,
        )
        timing["find_frame_after_edge_filter"] = round(
            float(timing.get("find_frame_after_edge_filter", 0.0)) + candidate_stats.kept_after_edge_filter,
            3,
        )
        timing["find_frame_border_band_dropped"] = round(
            float(timing.get("find_frame_border_band_dropped", 0.0)) + candidate_stats.dropped_by_border_band,
            3,
        )
        timing["find_frame_after_border_band"] = round(
            float(timing.get("find_frame_after_border_band", 0.0)) + candidate_stats.kept_after_border_band,
            3,
        )
        if candidate_stats.border_band_enabled:
            timing["find_frame_border_band_enabled_pages"] = round(
                float(timing.get("find_frame_border_band_enabled_pages", 0.0)) + 1.0,
                3,
            )
    rect_after_filter = len(filtered)
    kept: list[_FrameRectCandidate] = []
    t_roi_filter = time.perf_counter()
    for candidate in filtered:
        if _pm_intersects(candidate.pm_bbox, roi_pm):
            kept.append(candidate)
    t_roi_filter_elapsed = time.perf_counter() - t_roi_filter
    _timing_add(timing, "find_frame_roi_filter", t_roi_filter_elapsed)
    rect_after_roi = len(kept)
    fallback_gost = rect_after_roi == 0
    if fallback_gost:
        frame = gost
        union_pm: tuple[float, float, float, float] | None = None
        t_union_elapsed = 0.0
    else:
        t_union = time.perf_counter()
        u0 = v0 = float("inf")
        u1 = v1 = float("-inf")
        for candidate in kept:
            px0, py0, px1, py1 = candidate.pm_bbox
            u0 = min(u0, px0)
            v0 = min(v0, py0)
            u1 = max(u1, px1)
            v1 = max(v1, py1)
        union_pm = (u0, v0, u1, v1)
        frame = _frame_info_from_union_pm(u0, u1, v0, v1, page_w, page_h, rotation)
        t_union_elapsed = time.perf_counter() - t_union
    _timing_add(timing, "find_frame_union_build", t_union_elapsed)
    t_total_elapsed = time.perf_counter() - t_total

    diag: dict | None = None
    if debug:
        diag = {
            "frame_mode": "drawing_union",
            "rotation_handled": rotation,
            "were_axes_swapped": rotation in (90, 270),
            "roi_pdfminer": [round(x, 4) for x in roi_pm],
            "roi_source": roi_src,
            "rect_count_after_filter": rect_after_filter,
            "rect_count_total_drawings": candidate_stats.total_drawings,
            "rect_count_after_edge_filter": candidate_stats.kept_after_edge_filter,
            "border_band_enabled": candidate_stats.border_band_enabled,
            "border_band_left_mm": _border_band_margins_mm(cfg)[0],
            "border_band_right_mm": _border_band_margins_mm(cfg)[1],
            "border_band_top_mm": _border_band_margins_mm(cfg)[2],
            "border_band_bottom_mm": _border_band_margins_mm(cfg)[3],
            "rect_count_dropped_by_border_band": candidate_stats.dropped_by_border_band,
            "rect_count_after_border_band": candidate_stats.kept_after_border_band,
            "rect_count_after_roi": rect_after_roi,
            "timing_total_sec": round(t_total_elapsed, 6),
            "timing_roi_build_sec": round(t_roi_elapsed, 6),
            "timing_filter_rects_sec": round(t_filter_elapsed, 6),
            "timing_roi_filter_sec": round(t_roi_filter_elapsed, 6),
            "timing_union_build_sec": round(t_union_elapsed, 6),
            "drawing_union_fallback_gost": fallback_gost,
            "union_bbox_pdfminer": [round(x, 4) for x in union_pm] if union_pm else None,
            "page_width_pts": round(page_w, 2),
            "page_height_pts": round(page_h, 2),
            "gost_fallback_used": fallback_gost,
            # gost heuristic fields absent in this mode
            "winning_x_rect_fitz": None,
            "winning_y_rect_fitz": None,
            "best_pm_x_span_pdfminer": None,
            "best_pm_y_span_pdfminer": None,
            "min_size_fraction_passed": None,
            "composite_area_pts2": None,
            "max_single_rect_area_pts2": None,
            "same_source_rect_for_both_axes": None,
            "min_frame_fraction": None,
            "top_n_rects_by_area": [],
        }
    return frame, diag


def find_frame(
    fitz_page: fitz.Page,
    debug: bool = False,
    *,
    frame_mode: str = "gost",
    template: StampTemplate | None = None,
    timing: dict[str, float] | None = None,
    cfg: dict[str, float] | None = None,
) -> tuple[FrameInfo, dict | None]:
    """Detect the drawing frame (рамка) on *fitz_page* via ``get_drawings()``.

    The returned :class:`FrameInfo` stores the frame bbox in **pdfminer displayed
    coordinates** (x → right, y → up, origin at bottom-left of the displayed page).

    Drawing rects from ``get_drawings()`` are in **unrotated** fitz space; for
    rotation 90°/270° they are converted to pdfminer via ``rotation_matrix`` +
    y-flip (same as :func:`pdf_parsing_v2.coord_transform.unrotated_fitz_rect_to_pdfminer_bbox`).

    **frame_mode** (default ``gost``): unchanged legacy behaviour — max span X/Y with
    20% page threshold, else GOST margins.

    **frame_mode** ``drawing_union``: union AABB of filtered rects that intersect ROI
    (see module docstring). Empty after ROI → GOST fallback. *template* is used for
    ROI when it lists inner (non-``outside_stamp``) fields.

    Parameters
    ----------
    debug
        If True, also return a JSON-serializable dict with detection diagnostics
        (second tuple element). If False, second element is None.
    frame_mode
        ``gost`` | ``drawing_union`` (invalid values treated as ``gost``).
    template
        For ``drawing_union`` only: used to tighten ROI via inner field bboxes vs GOST.
    """
    mode = normalize_frame_mode(frame_mode)
    if mode == "drawing_union":
        return _find_frame_drawing_union(fitz_page, debug, template, timing=timing, cfg=cfg)

    t_total = time.perf_counter()
    page_w: float = fitz_page.rect.width   # displayed width (pts)
    page_h: float = fitz_page.rect.height  # displayed height (pts)
    rotation: int = fitz_page.rotation % 360

    best_pm_x0 = best_pm_x1 = 0.0
    best_pm_y0 = best_pm_y1 = 0.0
    best_x_range = 0.0
    best_y_range = 0.0

    win_x: _FrameRectCandidate | None = None
    win_y: _FrameRectCandidate | None = None
    filtered_candidates: list[_FrameRectCandidate] = []
    rect_count_after_filter = 0
    max_single_rect_area_pts2 = 0.0

    t_filter = time.perf_counter()
    frame_candidates, candidate_stats = _collect_filtered_frame_candidates(fitz_page, cfg=cfg)
    t_filter_elapsed = time.perf_counter() - t_filter
    _timing_add(timing, "find_frame_filter_rects", t_filter_elapsed)
    if timing is not None:
        timing["find_frame_total_drawings"] = round(
            float(timing.get("find_frame_total_drawings", 0.0)) + candidate_stats.total_drawings,
            3,
        )
        timing["find_frame_after_edge_filter"] = round(
            float(timing.get("find_frame_after_edge_filter", 0.0)) + candidate_stats.kept_after_edge_filter,
            3,
        )
        timing["find_frame_border_band_dropped"] = round(
            float(timing.get("find_frame_border_band_dropped", 0.0)) + candidate_stats.dropped_by_border_band,
            3,
        )
        timing["find_frame_after_border_band"] = round(
            float(timing.get("find_frame_after_border_band", 0.0)) + candidate_stats.kept_after_border_band,
            3,
        )
        if candidate_stats.border_band_enabled:
            timing["find_frame_border_band_enabled_pages"] = round(
                float(timing.get("find_frame_border_band_enabled_pages", 0.0)) + 1.0,
                3,
            )
    t_span_scan = time.perf_counter()
    for candidate in frame_candidates:
        rect_count_after_filter += 1
        if debug:
            filtered_candidates.append(candidate)
        max_single_rect_area_pts2 = max(max_single_rect_area_pts2, candidate.area_pts2)
        pm_x0, pm_y0, pm_x1, pm_y1 = candidate.pm_bbox
        if candidate.x_range > best_x_range:
            best_pm_x0 = pm_x0
            best_pm_x1 = pm_x1
            best_x_range = candidate.x_range
            win_x = candidate
        if candidate.y_range > best_y_range:
            best_pm_y0 = pm_y0
            best_pm_y1 = pm_y1
            best_y_range = candidate.y_range
            win_y = candidate
    t_span_scan_elapsed = time.perf_counter() - t_span_scan
    _timing_add(timing, "find_frame_span_scan", t_span_scan_elapsed)

    found = (
        best_x_range >= page_w * _MIN_FRAME_FRACTION
        and best_y_range >= page_h * _MIN_FRAME_FRACTION
    )
    min_size_fraction_passed = found

    composite_area_pts2: float | None = None
    if found:
        composite_area_pts2 = max(0.0, best_pm_x1 - best_pm_x0) * max(
            0.0, best_pm_y1 - best_pm_y0
        )

    top_n_by_area: list[dict[str, object]] = []
    if debug and filtered_candidates:
        t_debug_top = time.perf_counter()
        scored = sorted(filtered_candidates, key=lambda item: item.area_pts2, reverse=True)
        for candidate in scored[:5]:
            top_n_by_area.append({
                "area_pts2": round(candidate.area_pts2, 2),
                "fitz_rect": _serialize_fitz_rect(candidate.fitz_rect),
                "pdfminer_bbox": _serialize_pm_bbox(candidate.pm_bbox),
            })
        t_debug_top_elapsed = time.perf_counter() - t_debug_top
        _timing_add(timing, "find_frame_debug_top_rects", t_debug_top_elapsed)
    else:
        t_debug_top_elapsed = 0.0
    t_total_elapsed = time.perf_counter() - t_total

    diag: dict | None = None
    if debug:
        win_x_fitz = fitz.Rect(win_x.fitz_rect) if win_x is not None else None
        win_y_fitz = fitz.Rect(win_y.fitz_rect) if win_y is not None else None
        same_src = _fitz_rects_equal(win_x_fitz, win_y_fitz)
        diag = {
            "frame_mode": "gost",
            "rotation_handled": rotation,
            "were_axes_swapped": rotation in (90, 270),
            "timing_total_sec": round(t_total_elapsed, 6),
            "timing_filter_rects_sec": round(t_filter_elapsed, 6),
            "timing_span_scan_sec": round(t_span_scan_elapsed, 6),
            "timing_debug_top_rects_sec": round(t_debug_top_elapsed, 6),
            "winning_x_rect_fitz": _serialize_fitz_rect(win_x_fitz),
            "winning_y_rect_fitz": _serialize_fitz_rect(win_y_fitz),
            "best_pm_x_span_pdfminer": _serialize_pm_bbox(win_x.pm_bbox if win_x is not None else None),
            "best_pm_y_span_pdfminer": _serialize_pm_bbox(win_y.pm_bbox if win_y is not None else None),
            "gost_fallback_used": not found,
            "min_size_fraction_passed": min_size_fraction_passed,
            "rect_count_after_filter": rect_count_after_filter,
            "rect_count_total_drawings": candidate_stats.total_drawings,
            "rect_count_after_edge_filter": candidate_stats.kept_after_edge_filter,
            "border_band_enabled": candidate_stats.border_band_enabled,
            "border_band_left_mm": _border_band_margins_mm(cfg)[0],
            "border_band_right_mm": _border_band_margins_mm(cfg)[1],
            "border_band_top_mm": _border_band_margins_mm(cfg)[2],
            "border_band_bottom_mm": _border_band_margins_mm(cfg)[3],
            "rect_count_dropped_by_border_band": candidate_stats.dropped_by_border_band,
            "rect_count_after_border_band": candidate_stats.kept_after_border_band,
            "composite_area_pts2": round(composite_area_pts2, 2) if composite_area_pts2 is not None else None,
            "max_single_rect_area_pts2": round(max_single_rect_area_pts2, 2),
            "same_source_rect_for_both_axes": same_src,
            "page_width_pts": round(page_w, 2),
            "page_height_pts": round(page_h, 2),
            "min_frame_fraction": _MIN_FRAME_FRACTION,
            "top_n_rects_by_area": top_n_by_area,
        }

    if not found:
        frame = _gost_fallback(page_w, page_h, rotation)
        if diag is not None:
            diag["gost_fallback_used"] = True
            diag["min_size_fraction_passed"] = False
        return frame, diag

    border_left_mm = (page_w - best_pm_x1) / SCALE
    border_bottom_mm = best_pm_y0 / SCALE
    border_top_mm = (page_h - best_pm_y1) / SCALE

    frame = FrameInfo(
        x0=best_pm_x0,
        y0=best_pm_y0,
        x1=best_pm_x1,
        y1=best_pm_y1,
        page_width=page_w,
        page_height=page_h,
        rotation=rotation,
        border_left_mm=border_left_mm,
        border_bottom_mm=border_bottom_mm,
        border_top_mm=border_top_mm,
    )
    return frame, diag


def _gost_fallback(page_w: float, page_h: float, rotation: int) -> FrameInfo:
    """GOST 2.301-68 default margins: 20 mm left, 5 mm other sides."""
    x0 = _GOST_LEFT_MARGIN_MM * SCALE
    x1 = page_w - _GOST_OTHER_MARGIN_MM * SCALE
    y0 = _GOST_OTHER_MARGIN_MM * SCALE
    y1 = page_h - _GOST_OTHER_MARGIN_MM * SCALE

    return FrameInfo(
        x0=x0,
        y0=y0,
        x1=x1,
        y1=y1,
        page_width=page_w,
        page_height=page_h,
        rotation=rotation,
        border_left_mm=_GOST_OTHER_MARGIN_MM,
        border_bottom_mm=_GOST_OTHER_MARGIN_MM,
        border_top_mm=_GOST_OTHER_MARGIN_MM,
    )
