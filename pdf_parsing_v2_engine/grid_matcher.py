"""Grid adaptation engine (shared by pipeline and template editor).

Phase A approach (boundary-snap):
  1. Extract X/Y grid lines from detected cells
  2. Match explicit anchors
  3. For each non-anchor field (single pass):
     - approximate_field_bbox from anchors (interpolation)
     - snap 4 boundaries independently to grid lines
     - score with shape_score + position_score + boundary_consistency
     - accept or fallback to interpolated position
"""

from __future__ import annotations

import bisect
import math
import time
from dataclasses import dataclass, field as dc_field, replace

import fitz

from pdf_parsing_v2_engine.coord_transform import (
    SCALE,
    field_to_fitz_rect,
    grid_line_to_fitz_pts,
    snap_outside_stamp_vertical_to_frame,
)
from pdf_parsing_v2_engine.find_tables_settings import (
    build_find_tables_clip,
    call_find_tables,
    merge_cfg_for_find_tables,
)
from pdf_parsing_v2_engine.grid_detected_prefilter import prefilter_detected_cells_for_stamp_table
from pdf_parsing_v2_engine.models import FieldDef, FrameInfo, StampTemplate

_EPS = 1e-9
_CACHED_V2_CFG: dict | None = None


def _load_v2_config() -> dict:
    """Lazy import: avoids circular import pdf_parsing_v2.__init__ ↔ engine."""
    from pdf_parsing_v2.v2_config import load_v2_config as _lc

    return _lc()

_MIN_FIELD_PTS = 3.0
_MAX_CELL_AREA_FRACTION = 0.5
_LEGACY_SCALE_THRESHOLD = 0.05
_PROPOSED_BBOX_OVERFLOW_X_FRAC = 0.08
_PROPOSED_BBOX_OVERFLOW_Y_FRAC = 0.20
_PROPOSED_BBOX_SCALE_THRESHOLD = 0.04
_PROPOSED_BBOX_ADAPTIVE_MIN_SCALE_THRESHOLD = 0.035
_PROPOSED_BBOX_ADAPTIVE_MEDIUM_SCORE = 0.55
_PROPOSED_BBOX_ADAPTIVE_HIGH_SCORE = 0.80


def _clamp(value: float, lo: float, hi: float) -> float:
    """Clamp *value* to the inclusive [`lo`, `hi`] interval."""
    return max(lo, min(hi, value))


def _field_in_stamp_geometry_union(field: FieldDef) -> bool:
    """True if the field bbox belongs to the stamp union / grid geometry.

    Document-property and outside-stamp fields are excluded from detection,
    Hungarian matching, and adjacency.
    """
    return not field.outside_stamp and not field.document_property


def _zero_cell_bbox() -> CellBbox:
    return CellBbox(0.0, 0.0, 0.0, 0.0)


_LARGE_FIELD_THRESHOLD_PTS = 15.0 * SCALE


@dataclass
class CellBbox:
    """One cell bbox in fitz displayed coordinates (pts)."""

    x0: float
    y0: float
    x1: float
    y1: float

    @property
    def center(self) -> tuple[float, float]:
        return ((self.x0 + self.x1) * 0.5, (self.y0 + self.y1) * 0.5)

    @property
    def width(self) -> float:
        return max(0.0, self.x1 - self.x0)

    @property
    def height(self) -> float:
        return max(0.0, self.y1 - self.y0)

    @property
    def area(self) -> float:
        return self.width * self.height

    def to_rect(self) -> fitz.Rect:
        return fitz.Rect(self.x0, self.y0, self.x1, self.y1)

    @classmethod
    def from_rect(cls, rect: fitz.Rect) -> "CellBbox":
        return cls(rect.x0, rect.y0, rect.x1, rect.y1)


@dataclass
class AnchorMatch:
    """Pair template anchor cell <-> detected cell."""

    template_field_id: str
    template_bbox: CellBbox
    detected_bbox: CellBbox
    confidence: float


@dataclass
class SnapResult:
    status: str  # matched | split_merged | no_match
    bbox: CellBbox
    matched_detected: list[CellBbox]


@dataclass
class AdaptResult:
    field_id: str
    bbox: CellBbox
    status: str  # matched | split_merged | no_match | anchor | excluded | derived
    approx_bbox: CellBbox
    matched_detected: list[CellBbox]
    snap_distance_mm: float
    shape_score: float = 1.0
    snapped_boundaries: int = 0
    confidence: float = 0.0
    iteration: int = 0


def cluster_lines(coords: list[float], tolerance: float) -> list[float]:
    """Merge nearby coordinates into logical lines."""
    if not coords:
        return []
    if tolerance < 0:
        raise ValueError("tolerance must be >= 0")

    values = sorted(float(v) for v in coords)
    clusters: list[list[float]] = [[values[0]]]
    for val in values[1:]:
        if abs(val - clusters[-1][-1]) <= tolerance:
            clusters[-1].append(val)
        else:
            clusters.append([val])
    return sorted(sum(c) / len(c) for c in clusters)


def get_detected_stamp_cells(
    fitz_page: fitz.Page,
    stamp_bbox: fitz.Rect,
    tolerance_mm: float = 0.5,
    cfg: dict | None = None,
    template: StampTemplate | None = None,
    timing: dict[str, float] | None = None,
) -> list[CellBbox]:
    """find_tables() -> stamp filter -> line consolidation -> dedup."""
    t_find_tables = time.perf_counter()
    find_tables_clip = build_find_tables_clip(fitz_page, stamp_bbox)
    try:
        tables_obj = call_find_tables(
            fitz_page,
            cfg=merge_cfg_for_find_tables(cfg, template),
            clip=find_tables_clip,
        )
    except Exception:
        return []
    finally:
        if timing is not None:
            timing["find_tables"] = round(
                float(timing.get("find_tables", 0.0))
                + (time.perf_counter() - t_find_tables),
                6,
            )

    tables = getattr(tables_obj, "tables", tables_obj)
    raw_cells: list[CellBbox] = []
    for t in tables:
        for c in getattr(t, "cells", []):
            if not c or len(c) != 4:
                continue
            cb = CellBbox(float(c[0]), float(c[1]), float(c[2]), float(c[3]))
            if cb.width < 2.0 or cb.height < 2.0:
                continue
            if (cb.to_rect() & stamp_bbox).is_empty:
                continue
            raw_cells.append(cb)

    if not raw_cells:
        return []

    tol_pts = tolerance_mm * SCALE
    x_lines = cluster_lines([p for c in raw_cells for p in (c.x0, c.x1)], tol_pts)
    y_lines = cluster_lines([p for c in raw_cells for p in (c.y0, c.y1)], tol_pts)

    def snap(val: float, lines: list[float]) -> float:
        nearest = min(lines, key=lambda x: abs(x - val))
        return nearest if abs(nearest - val) <= tol_pts else val

    snapped: list[CellBbox] = []
    dedup: dict[tuple[int, int, int, int], CellBbox] = {}
    for c in raw_cells:
        s = CellBbox(
            x0=snap(c.x0, x_lines),
            y0=snap(c.y0, y_lines),
            x1=snap(c.x1, x_lines),
            y1=snap(c.y1, y_lines),
        )
        if s.width < 1.0 or s.height < 1.0:
            continue
        k = (round(s.x0, 1), round(s.y0, 1), round(s.x1, 1), round(s.y1, 1))
        dedup[k] = s
    snapped.extend(dedup.values())
    return snapped


def match_anchor_cells(
    template: StampTemplate,
    frame: FrameInfo,
    detected_cells: list[CellBbox],
    stamp_bbox: fitz.Rect,
) -> list[AnchorMatch]:
    """Match manual/auto template anchors to detected cells."""
    if not detected_cells:
        return []

    template_anchor_defs = [
        f for f in template.fields if f.is_anchor and _field_in_stamp_geometry_union(f)
    ]
    if not template_anchor_defs:
        # Safer default for editor UX: without explicit anchors, do not build
        # approximate transforms from auto-selected large cells.
        return []

    matches: list[AnchorMatch] = []
    used_detected: set[int] = set()
    stamp_w = max(stamp_bbox.width, _EPS)
    stamp_h = max(stamp_bbox.height, _EPS)

    for f in template_anchor_defs:
        tb = _template_field_bbox(f, frame)
        best_idx = -1
        best_score = 1e18
        for i, dc in enumerate(detected_cells):
            if i in used_detected:
                continue
            score = _anchor_score(tb, dc, stamp_w, stamp_h)
            if score < best_score:
                best_score = score
                best_idx = i
        if best_idx < 0:
            continue
        used_detected.add(best_idx)
        conf = 1.0 / (1.0 + best_score)
        matches.append(
            AnchorMatch(
                template_field_id=f.id,
                template_bbox=tb,
                detected_bbox=detected_cells[best_idx],
                confidence=max(0.0, min(1.0, conf)),
            )
        )
    return matches


def _match_anchor_cells_indexed(
    template: StampTemplate,
    frame: FrameInfo,
    detected_cells: list[CellBbox],
    stamp_bbox: fitz.Rect,
    keys: list[str],
) -> list[AnchorMatch]:
    """Like ``match_anchor_cells`` but uses positional *keys* instead of
    ``field.id``, so duplicate IDs get separate entries."""
    if not detected_cells:
        return []

    anchor_indices = [
        i for i, f in enumerate(template.fields)
        if f.is_anchor and _field_in_stamp_geometry_union(f)
    ]
    if not anchor_indices:
        return []

    matches: list[AnchorMatch] = []
    used_detected: set[int] = set()
    stamp_w = max(stamp_bbox.width, _EPS)
    stamp_h = max(stamp_bbox.height, _EPS)

    for fi in anchor_indices:
        f = template.fields[fi]
        tb = _template_field_bbox(f, frame)
        best_idx = -1
        best_score = 1e18
        for i, dc in enumerate(detected_cells):
            if i in used_detected:
                continue
            score = _anchor_score(tb, dc, stamp_w, stamp_h)
            if score < best_score:
                best_score = score
                best_idx = i
        if best_idx < 0:
            continue
        used_detected.add(best_idx)
        conf = 1.0 / (1.0 + best_score)
        matches.append(
            AnchorMatch(
                template_field_id=keys[fi],
                template_bbox=tb,
                detected_bbox=detected_cells[best_idx],
                confidence=max(0.0, min(1.0, conf)),
            )
        )
    return matches


def approximate_field_bbox(
    field: FieldDef,
    frame: FrameInfo,
    anchors: list[AnchorMatch],
    stamp_bbox: fitz.Rect,  # reserved for future constraints
) -> CellBbox:
    """Interpolate field position/size from matched anchors."""
    template_bbox = _template_field_bbox(field, frame)
    if not anchors:
        return template_bbox

    fcx, fcy = template_bbox.center
    fw, fh = template_bbox.width, template_bbox.height

    weighted = []
    for a in anchors:
        tcx, tcy = a.template_bbox.center
        dcx, dcy = a.detected_bbox.center
        rx = a.detected_bbox.width / max(a.template_bbox.width, _EPS)
        ry = a.detected_bbox.height / max(a.template_bbox.height, _EPS)

        pred_cx = dcx + (fcx - tcx) * rx
        pred_cy = dcy + (fcy - tcy) * ry
        pred_w = fw * rx
        pred_h = fh * ry

        dist = ((fcx - tcx) ** 2 + (fcy - tcy) ** 2) ** 0.5
        w = a.confidence / max(dist, 4.0)
        weighted.append((w, pred_cx, pred_cy, pred_w, pred_h))

    sw = sum(x[0] for x in weighted)
    if sw <= 0:
        return template_bbox

    cx = sum(w * x for w, x, _, _, _ in weighted) / sw
    cy = sum(w * y for w, _, y, _, _ in weighted) / sw
    ww = max(2.0, sum(w * ww for w, _, _, ww, _ in weighted) / sw)
    hh = max(2.0, sum(w * hh for w, _, _, _, hh in weighted) / sw)
    return CellBbox(cx - ww / 2.0, cy - hh / 2.0, cx + ww / 2.0, cy + hh / 2.0)


def snap_to_detected(
    approx_bbox: CellBbox,
    detected_cells: list[CellBbox],
    size_tolerance: float = 0.4,
    max_distance_mm: float = 5.0,
    iou_threshold: float = 0.3,
) -> SnapResult:
    """Snap approximate bbox to nearest detected cell; handle split/merge."""
    if not detected_cells:
        return SnapResult(status="no_match", bbox=approx_bbox, matched_detected=[])

    acx, acy = approx_bbox.center
    max_dist_pts = max_distance_mm * SCALE

    nearby = []
    for dc in detected_cells:
        dcx, dcy = dc.center
        dist = ((acx - dcx) ** 2 + (acy - dcy) ** 2) ** 0.5
        if dist <= max_dist_pts:
            nearby.append(dc)
    # Do not jump to far-away cells when no local candidates exist.
    if not nearby:
        return SnapResult(status="no_match", bbox=approx_bbox, matched_detected=[])
    candidates = nearby

    best = max(candidates, key=lambda d: _iou(approx_bbox, d))
    best_iou = _iou(approx_bbox, best)
    area_ratio = best.area / max(approx_bbox.area, _EPS)
    size_similar = size_tolerance <= area_ratio <= (1.0 / max(size_tolerance, _EPS))

    if best_iou >= iou_threshold and size_similar:
        return SnapResult(status="matched", bbox=best, matched_detected=[best])

    overlaps = [d for d in candidates if _overlap_area(approx_bbox, d) > 0.0]
    if overlaps and area_ratio < size_tolerance:
        merged = _merge_cells(overlaps)
        if _iou(approx_bbox, merged) >= max(iou_threshold * 0.5, 0.1):
            return SnapResult(status="split_merged", bbox=merged, matched_detected=overlaps)

    return SnapResult(status="no_match", bbox=approx_bbox, matched_detected=[])


def _extract_grid_lines(
    detected: list[CellBbox],
    tolerance_mm: float,
) -> tuple[list[float], list[float]]:
    """Extract consolidated X and Y grid lines from detected cell boundaries."""
    tol_pts = tolerance_mm * SCALE
    x_coords = [p for c in detected for p in (c.x0, c.x1)]
    y_coords = [p for c in detected for p in (c.y0, c.y1)]
    return cluster_lines(x_coords, tol_pts), cluster_lines(y_coords, tol_pts)


def _build_adjacency_graph(
    fields: list[FieldDef],
    frame: FrameInfo,
    tol_mm: float = 2.0,
) -> dict[str, list[tuple[str, str]]]:
    """Precompute which template fields share a boundary.

    Returns {field_id: [(neighbor_id, shared_edge), ...]}.
    shared_edge is "left", "right", "top", "bottom".
    Two fields are adjacent if distance between boundary pair < tol
    AND projection overlap along the perpendicular axis > 50% of the smaller span.
    """
    tol_pts = tol_mm * SCALE
    bboxes: dict[str, CellBbox] = {}
    for f in fields:
        if not _field_in_stamp_geometry_union(f):
            continue
        bboxes[f.id] = _template_field_bbox(f, frame)

    graph: dict[str, list[tuple[str, str]]] = {fid: [] for fid in bboxes}
    ids = list(bboxes.keys())

    for i in range(len(ids)):
        for j in range(i + 1, len(ids)):
            a_id, b_id = ids[i], ids[j]
            a, b = bboxes[a_id], bboxes[b_id]
            _check_edge_pair(a_id, b_id, a, b, "right", "left",
                             a.x1, b.x0, a.y0, a.y1, b.y0, b.y1,
                             tol_pts, graph)
            _check_edge_pair(a_id, b_id, a, b, "left", "right",
                             a.x0, b.x1, a.y0, a.y1, b.y0, b.y1,
                             tol_pts, graph)
            _check_edge_pair(a_id, b_id, a, b, "bottom", "top",
                             a.y1, b.y0, a.x0, a.x1, b.x0, b.x1,
                             tol_pts, graph)
            _check_edge_pair(a_id, b_id, a, b, "top", "bottom",
                             a.y0, b.y1, a.x0, a.x1, b.x0, b.x1,
                             tol_pts, graph)
    return graph


def _check_edge_pair(
    a_id: str, b_id: str,
    a: CellBbox, b: CellBbox,
    a_edge: str, b_edge: str,
    a_coord: float, b_coord: float,
    a_perp_lo: float, a_perp_hi: float,
    b_perp_lo: float, b_perp_hi: float,
    tol: float,
    graph: dict[str, list[tuple[str, str]]],
) -> None:
    if abs(a_coord - b_coord) > tol:
        return
    overlap_lo = max(a_perp_lo, b_perp_lo)
    overlap_hi = min(a_perp_hi, b_perp_hi)
    overlap = max(0.0, overlap_hi - overlap_lo)
    min_span = min(a_perp_hi - a_perp_lo, b_perp_hi - b_perp_lo)
    if min_span <= 0 or overlap / min_span < 0.5:
        return
    graph[a_id].append((b_id, a_edge))
    graph[b_id].append((a_id, b_edge))


def _build_adjacency_graph_keyed(
    fields: list[FieldDef],
    keys: list[str],
    frame: FrameInfo,
    tol_mm: float = 2.0,
) -> dict[str, list[tuple[str, str]]]:
    """Like ``_build_adjacency_graph`` but uses positional *keys*."""
    tol_pts = tol_mm * SCALE
    items: list[tuple[str, CellBbox]] = []
    for idx, f in enumerate(fields):
        if not _field_in_stamp_geometry_union(f):
            continue
        items.append((keys[idx], _template_field_bbox(f, frame)))

    graph: dict[str, list[tuple[str, str]]] = {k: [] for k, _ in items}
    for i in range(len(items)):
        for j in range(i + 1, len(items)):
            a_id, a = items[i]
            b_id, b = items[j]
            _check_edge_pair(a_id, b_id, a, b, "right", "left",
                             a.x1, b.x0, a.y0, a.y1, b.y0, b.y1,
                             tol_pts, graph)
            _check_edge_pair(a_id, b_id, a, b, "left", "right",
                             a.x0, b.x1, a.y0, a.y1, b.y0, b.y1,
                             tol_pts, graph)
            _check_edge_pair(a_id, b_id, a, b, "bottom", "top",
                             a.y1, b.y0, a.x0, a.x1, b.x0, b.x1,
                             tol_pts, graph)
            _check_edge_pair(a_id, b_id, a, b, "top", "bottom",
                             a.y0, b.y1, a.x0, a.x1, b.x0, b.x1,
                             tol_pts, graph)
    return graph


def _snap_4_boundaries(
    bbox: CellBbox,
    x_lines: list[float],
    y_lines: list[float],
    max_snap_pts: float,
) -> tuple[CellBbox, int]:
    """Snap each of 4 edges independently to nearest grid line.

    Returns (snapped_bbox, n_snapped) where n_snapped is 0..4.
    After snap: normalization + min-size check.
    """
    def _nearest(val: float, lines: list[float]) -> float | None:
        if not lines:
            return None
        best = min(lines, key=lambda ln: abs(ln - val))
        return best if abs(best - val) <= max_snap_pts else None

    new_x0 = _nearest(bbox.x0, x_lines)
    new_x1 = _nearest(bbox.x1, x_lines)
    new_y0 = _nearest(bbox.y0, y_lines)
    new_y1 = _nearest(bbox.y1, y_lines)

    n_snapped = sum(1 for v in (new_x0, new_x1, new_y0, new_y1) if v is not None)
    x0 = new_x0 if new_x0 is not None else bbox.x0
    x1 = new_x1 if new_x1 is not None else bbox.x1
    y0 = new_y0 if new_y0 is not None else bbox.y0
    y1 = new_y1 if new_y1 is not None else bbox.y1

    if x0 > x1:
        x0, x1 = x1, x0
    if y0 > y1:
        y0, y1 = y1, y0

    if (x1 - x0) < _MIN_FIELD_PTS or (y1 - y0) < _MIN_FIELD_PTS:
        return bbox, 0

    max_change = getattr(_snap_4_boundaries, '_max_shape_change', 3.0)
    orig_w, orig_h = bbox.width, bbox.height
    new_w, new_h = x1 - x0, y1 - y0
    if orig_w > _EPS and orig_h > _EPS:
        wr = new_w / orig_w
        hr = new_h / orig_h
        if max(wr, 1.0 / max(wr, _EPS)) > max_change or max(hr, 1.0 / max(hr, _EPS)) > max_change:
            return bbox, 0

    return CellBbox(x0, y0, x1, y1), n_snapped


def _shape_score(original: CellBbox, adapted: CellBbox) -> float:
    """Measure how much shape changed: 1.0 = identical, 0.0 = extreme deformation."""
    w_ratio = adapted.width / max(original.width, _EPS)
    h_ratio = adapted.height / max(original.height, _EPS)
    return min(w_ratio, 1.0 / max(w_ratio, _EPS)) * min(h_ratio, 1.0 / max(h_ratio, _EPS))


def _boundary_consistency(
    field_id: str,
    candidate: CellBbox,
    matched: dict[str, tuple[CellBbox, float]],
    adjacency: dict[str, list[tuple[str, str]]],
) -> float:
    """Score how well candidate boundaries align with already-matched neighbors.

    Returns average boundary error score (1.0 = perfect alignment, 0.0 = 5+ mm off).
    If no matched neighbors, returns 1.0 (neutral).
    """
    neighbors = adjacency.get(field_id, [])
    if not neighbors:
        return 1.0

    scores = []
    for nbr_id, shared_edge in neighbors:
        entry = matched.get(nbr_id)
        if entry is None:
            continue
        nbr_bbox, _conf = entry

        if shared_edge == "right":
            error = abs(candidate.x1 - nbr_bbox.x0)
        elif shared_edge == "left":
            error = abs(candidate.x0 - nbr_bbox.x1)
        elif shared_edge == "bottom":
            error = abs(candidate.y1 - nbr_bbox.y0)
        elif shared_edge == "top":
            error = abs(candidate.y0 - nbr_bbox.y1)
        else:
            continue
        error_mm = error / SCALE
        scores.append(1.0 / (1.0 + error_mm))

    return sum(scores) / len(scores) if scores else 1.0


_W_SHAPE = 0.3
_W_POS = 0.3
_W_BOUND = 0.4


def _score_candidate(
    field_id: str,
    candidate: CellBbox,
    approx: CellBbox,
    template_bbox: CellBbox,
    matched: dict[str, tuple[CellBbox, float]],
    adjacency: dict[str, list[tuple[str, str]]],
) -> float:
    """Combined score: shape + position + boundary consistency.

    Normalized so that missing components don't penalize the total.
    """
    shape = _shape_score(template_bbox, candidate)

    dist_mm = _center_distance_mm(candidate, approx)
    position = 1.0 / (1.0 + dist_mm)

    has_matched_neighbors = any(
        matched.get(nid) is not None
        for nid, _ in adjacency.get(field_id, [])
    )
    boundary = _boundary_consistency(field_id, candidate, matched, adjacency)

    if has_matched_neighbors:
        total = shape * _W_SHAPE + position * _W_POS + boundary * _W_BOUND
    else:
        weight_sum = _W_SHAPE + _W_POS
        total = (shape * _W_SHAPE + position * _W_POS) / weight_sum
    return total


def _make_unique_keys(fields: list[FieldDef]) -> list[str]:
    """Assign positional keys so duplicate field IDs don't collide.

    Unique IDs keep their name; duplicates get ``id#N`` suffix (0-based).
    """
    from collections import Counter
    counts = Counter(f.id for f in fields)
    seen: dict[str, int] = {}
    keys: list[str] = []
    for f in fields:
        if counts[f.id] == 1:
            keys.append(f.id)
        else:
            idx = seen.get(f.id, 0)
            keys.append(f"{f.id}#{idx}")
            seen[f.id] = idx + 1
    return keys


def _expand_bbox_for_anchor(
    bbox: fitz.Rect,
    *,
    origin: str,
    x_frac: float,
    y_frac: float,
    page_rect: fitz.Rect,
) -> fitz.Rect:
    """Expand bbox asymmetrically opposite to the anchor corner."""
    if x_frac <= 0 and y_frac <= 0:
        return fitz.Rect(bbox)

    expanded = fitz.Rect(bbox)
    dx = bbox.width * max(0.0, x_frac)
    dy = bbox.height * max(0.0, y_frac)

    if origin.endswith("bottom_right"):
        expanded.x0 -= dx
        expanded.y0 -= dy
    elif origin.endswith("bottom_left"):
        expanded.x1 += dx
        expanded.y0 -= dy
    elif origin.endswith("top_right"):
        expanded.x0 -= dx
        expanded.y1 += dy
    else:
        expanded.x1 += dx
        expanded.y1 += dy

    expanded &= page_rect
    return expanded if not expanded.is_empty else fitz.Rect(bbox)


def _filter_cells_to_bbox(cells: list[CellBbox], bbox: fitz.Rect) -> list[CellBbox]:
    """Keep only cells that intersect *bbox*."""
    return [cell for cell in cells if not (cell.to_rect() & bbox).is_empty]


def _clip_cells_to_bbox(cells: list[CellBbox], bbox: fitz.Rect) -> list[CellBbox]:
    """Clip cells to *bbox* for stable bbox union calculation."""
    stamp_cb = CellBbox.from_rect(bbox)
    clipped: list[CellBbox] = []
    for cell in cells:
        cx0 = max(cell.x0, stamp_cb.x0)
        cy0 = max(cell.y0, stamp_cb.y0)
        cx1 = min(cell.x1, stamp_cb.x1)
        cy1 = min(cell.y1, stamp_cb.y1)
        if cx1 > cx0 and cy1 > cy0:
            clipped.append(CellBbox(cx0, cy0, cx1, cy1))
    return clipped


def _union_bbox_or_fallback(cells: list[CellBbox], fallback: CellBbox) -> CellBbox:
    """Return union bbox for cells or *fallback* when the list is empty."""
    if not cells:
        return fallback
    return CellBbox(
        min(cell.x0 for cell in cells),
        min(cell.y0 for cell in cells),
        max(cell.x1 for cell in cells),
        max(cell.y1 for cell in cells),
    )


def _grid_prealign_settings(cfg: dict | None) -> dict[str, float | bool]:
    """Return production defaults for proposed-bbox pre-alignment."""
    source = cfg or {}
    return {
        "enabled": bool(source.get("grid_prealign_enabled", True)),
        "bbox_overflow_x_frac": float(
            source.get(
                "grid_prealign_bbox_overflow_x_frac",
                _PROPOSED_BBOX_OVERFLOW_X_FRAC,
            )
        ),
        "bbox_overflow_y_frac": float(
            source.get(
                "grid_prealign_bbox_overflow_y_frac",
                _PROPOSED_BBOX_OVERFLOW_Y_FRAC,
            )
        ),
        "scale_threshold": float(
            source.get(
                "grid_prealign_proposed_bbox_scale_threshold",
                _PROPOSED_BBOX_SCALE_THRESHOLD,
            )
        ),
        "adaptive_scale_enabled": bool(
            source.get("grid_prealign_adaptive_scale_enabled", True)
        ),
        "adaptive_scale_min_threshold": float(
            source.get(
                "grid_prealign_adaptive_scale_min_threshold",
                _PROPOSED_BBOX_ADAPTIVE_MIN_SCALE_THRESHOLD,
            )
        ),
    }


def _pick_boundary_like_line(
    *,
    detected_lines: list[DetectedGridLineInfo],
    orientation: str,
    expected_span_pts: float,
    edge: str,
    span_ratio_min: float = 0.8,
    fallback_ratio_min: float = 0.6,
) -> tuple[DetectedGridLineInfo | None, list[DetectedGridLineInfo], list[DetectedGridLineInfo]]:
    """Pick the most extreme boundary-like detected line for one edge."""
    orient_lines = [line for line in detected_lines if line.orientation == orientation]
    if not orient_lines:
        return None, [], []

    def _span_ratio(line: DetectedGridLineInfo) -> float:
        line_span = max(line.span_hi_pts - line.span_lo_pts, _EPS)
        return min(line_span, expected_span_pts) / max(line_span, expected_span_pts, _EPS)

    def _sort_key(line: DetectedGridLineInfo) -> tuple[float, float, str]:
        line_span = line.span_hi_pts - line.span_lo_pts
        edge_pos = line.pos_pts if edge in ("top", "left") else -line.pos_pts
        return (edge_pos, -line_span, line.id)

    strong = [line for line in orient_lines if _span_ratio(line) >= span_ratio_min]
    medium = [line for line in orient_lines if _span_ratio(line) >= fallback_ratio_min]
    chosen_pool = strong or medium
    chosen = sorted(chosen_pool, key=_sort_key)[0] if chosen_pool else None
    return chosen, strong, medium


def _build_merged_detected_line_infos(
    *,
    detected_cells: list[CellBbox],
    template: StampTemplate,
) -> list[DetectedGridLineInfo]:
    """Build merged detected grid lines without stamp-position filtering."""
    h_det_raw, v_det_raw = _build_detected_grid_lines(
        detected_cells, template.grid_tolerance_detected_mm
    ) if detected_cells else ([], [])
    border_merge_pts = _compute_border_merge_dist(detected_cells)
    h_det_merged = _merge_border_pairs(h_det_raw, border_merge_pts)
    v_det_merged = _merge_border_pairs(v_det_raw, border_merge_pts)
    h_id_map, v_id_map = _assign_detected_ids(h_det_merged, v_det_merged)

    merged: list[DetectedGridLineInfo] = []
    for orient, merged_lines, id_map in (
        ("h", h_det_merged, h_id_map),
        ("v", v_det_merged, v_id_map),
    ):
        for line in merged_lines:
            merged.append(
                DetectedGridLineInfo(
                    id=id_map[id(line)],
                    orientation=orient,
                    pos_pts=line.pos,
                    span_lo_pts=line.span_lo,
                    span_hi_pts=line.span_hi,
                    cell_count=line.cell_count,
                )
            )
    return merged


def _collect_template_boundary_positions(
    *,
    template: StampTemplate,
    frame: FrameInfo,
    orientation: str,
) -> list[tuple[str, float]]:
    """Return ordered template boundary positions in fitz coordinates."""
    positions: list[tuple[str, float]] = []
    for grid_line in template.grid_lines:
        line_orientation, pos_fitz, _start_fitz, _end_fitz, boundary, line_id = (
            grid_line_to_fitz_pts(grid_line, frame, template.origin)
        )
        if line_orientation != orientation or boundary is None:
            continue
        positions.append((line_id, pos_fitz))
    positions.sort(key=lambda item: item[1])
    return positions


def _collect_boundary_like_detected_positions(
    *,
    detected_lines: list[DetectedGridLineInfo],
    orientation: str,
    expected_span_pts: float,
    edge: str,
    span_ratio_min: float = 0.55,
) -> list[DetectedGridLineInfo]:
    """Return ordered detected lines that are plausible boundary carriers."""
    orient_lines = [line for line in detected_lines if line.orientation == orientation]
    if not orient_lines:
        return []

    def _span_ratio(line: DetectedGridLineInfo) -> float:
        span_pts = max(line.span_hi_pts - line.span_lo_pts, _EPS)
        return min(span_pts, expected_span_pts) / max(span_pts, expected_span_pts, _EPS)

    filtered = [line for line in orient_lines if _span_ratio(line) >= span_ratio_min]
    filtered.sort(key=lambda line: (line.pos_pts, -(line.span_hi_pts - line.span_lo_pts), line.id))
    if edge in ("right", "bottom"):
        filtered.reverse()
    return filtered


def _infer_virtual_outer_edge(
    *,
    template_positions: list[tuple[str, float]],
    detected_candidates: list[DetectedGridLineInfo],
    max_skip: int = 3,
) -> dict[str, float | int | str]:
    """Infer a virtual outer edge from early detected lines and template offsets."""
    empty: dict[str, float | int | str] = {
        "virtual_pos_pts": "",
        "matched_template_id": "",
        "matched_template_pos_pts": "",
        "source_detected_id": "",
        "source_detected_pos_pts": "",
        "skip_count": "",
        "scale_estimate": "",
        "score": "",
        "candidate_count": 0,
        "gap_count_used": 0,
    }
    if not template_positions or not detected_candidates:
        return empty

    tpl_positions_only = [pos for _line_id, pos in template_positions]
    usable_detected = detected_candidates[:4]
    best: dict[str, float | int | str] | None = None

    for skip_count, (matched_template_id, matched_template_pos) in enumerate(
        template_positions[: max_skip + 1]
    ):
        compare_count = min(len(usable_detected), len(template_positions) - skip_count)
        if compare_count <= 0:
            continue
        tpl_slice = tpl_positions_only[skip_count : skip_count + compare_count]
        det_slice = [line.pos_pts for line in usable_detected[:compare_count]]
        gap_count = max(0, compare_count - 1)
        scale_estimate = 1.0
        gap_error = 0.0
        if gap_count:
            ratios: list[float] = []
            residuals: list[float] = []
            for index in range(gap_count):
                tpl_gap = tpl_slice[index + 1] - tpl_slice[index]
                det_gap = det_slice[index + 1] - det_slice[index]
                if abs(tpl_gap) <= _EPS:
                    continue
                ratios.append(det_gap / tpl_gap)
            if ratios:
                ratios_sorted = sorted(ratios)
                scale_estimate = ratios_sorted[len(ratios_sorted) // 2]
                for index in range(gap_count):
                    tpl_gap = tpl_slice[index + 1] - tpl_slice[index]
                    det_gap = det_slice[index + 1] - det_slice[index]
                    residuals.append(abs(det_gap - tpl_gap * scale_estimate))
                gap_error = sum(residuals) / len(residuals)

        virtual_pos = det_slice[0] - (matched_template_pos - tpl_positions_only[0]) * scale_estimate
        candidate: dict[str, float | int | str] = {
            "virtual_pos_pts": round(virtual_pos, 3),
            "matched_template_id": matched_template_id,
            "matched_template_pos_pts": round(matched_template_pos, 3),
            "source_detected_id": usable_detected[0].id,
            "source_detected_pos_pts": round(usable_detected[0].pos_pts, 3),
            "skip_count": skip_count,
            "scale_estimate": round(scale_estimate, 6),
            "score": round(gap_error + skip_count * 0.75, 6),
            "candidate_count": len(usable_detected),
            "gap_count_used": gap_count,
        }
        if best is None or float(candidate["score"]) < float(best["score"]):
            best = candidate

    return best or empty


def _extent_probe_proposed_bbox(extent_probe: dict[str, object]) -> fitz.Rect | None:
    """Convert extent-probe scalar fields into a rect when available."""
    if not extent_probe.get("extent_probe_has_proposed_bbox"):
        return None
    keys = (
        "extent_probe_proposed_bbox_x0_pts",
        "extent_probe_proposed_bbox_y0_pts",
        "extent_probe_proposed_bbox_x1_pts",
        "extent_probe_proposed_bbox_y1_pts",
    )
    if any(extent_probe.get(key, "") == "" for key in keys):
        return None
    return fitz.Rect(*(float(extent_probe[key]) for key in keys))


def _choose_adaptive_prealign_scale_threshold(
    *,
    prealign_cfg: dict[str, float | bool],
    extent_probe: dict[str, object],
    rect_score: float,
    kept_detected_count: int,
    probe_detected_count: int,
) -> tuple[float, float, str, str]:
    """Select proposed-bbox scale threshold from pre-alignment confidence."""
    medium_threshold = float(prealign_cfg["scale_threshold"])
    if not bool(prealign_cfg.get("adaptive_scale_enabled", True)):
        return (
            medium_threshold,
            0.0,
            "fixed",
            f"adaptive=off, threshold={medium_threshold:.1%}",
        )

    min_threshold = _clamp(
        float(
            prealign_cfg.get(
                "adaptive_scale_min_threshold",
                _PROPOSED_BBOX_ADAPTIVE_MIN_SCALE_THRESHOLD,
            )
        ),
        0.0,
        _LEGACY_SCALE_THRESHOLD,
    )
    medium_threshold = _clamp(
        medium_threshold,
        min_threshold,
        _LEGACY_SCALE_THRESHOLD,
    )

    inferred_edges = {
        edge
        for edge in str(extent_probe.get("extent_probe_inferred_edges", "")).split(",")
        if edge
    }
    virtual_edge_count = int("top_virtual" in inferred_edges) + int(
        "left_virtual" in inferred_edges
    )
    edge_conf = 1.0 if virtual_edge_count == 2 else 0.55 if virtual_edge_count == 1 else 0.0

    top_strong = float(extent_probe.get("extent_probe_top_candidate_count_strong", 0) or 0)
    left_strong = float(extent_probe.get("extent_probe_left_candidate_count_strong", 0) or 0)
    candidate_conf = 0.5 * (
        _clamp(top_strong / 8.0, 0.0, 1.0)
        + _clamp(left_strong / 3.0, 0.0, 1.0)
    )

    span_values = [
        float(value)
        for value in (
            extent_probe.get("extent_probe_top_span_ratio_to_template", 0.0),
            extent_probe.get("extent_probe_left_span_ratio_to_template", 0.0),
            extent_probe.get("extent_probe_bottom_span_ratio_to_template", 0.0),
            extent_probe.get("extent_probe_right_span_ratio_to_template", 0.0),
        )
        if value not in ("", None)
    ]
    span_conf = (
        sum(_clamp((value - 0.75) / 0.20, 0.0, 1.0) for value in span_values)
        / len(span_values)
        if span_values
        else 0.0
    )
    avg_span = sum(span_values) / len(span_values) if span_values else 0.0

    pool_ratio = kept_detected_count / max(probe_detected_count, 1)
    pool_conf = _clamp((pool_ratio - 0.5) / 0.5, 0.0, 1.0)
    rect_conf = _clamp((rect_score - 0.5) / 0.4, 0.0, 1.0)

    confidence_score = _clamp(
        0.35 * edge_conf
        + 0.20 * candidate_conf
        + 0.20 * span_conf
        + 0.15 * pool_conf
        + 0.10 * rect_conf,
        0.0,
        1.0,
    )

    if confidence_score >= _PROPOSED_BBOX_ADAPTIVE_HIGH_SCORE:
        ratio = (
            confidence_score - _PROPOSED_BBOX_ADAPTIVE_HIGH_SCORE
        ) / max(1.0 - _PROPOSED_BBOX_ADAPTIVE_HIGH_SCORE, _EPS)
        threshold = medium_threshold - ratio * (medium_threshold - min_threshold)
        tier = "high"
    elif confidence_score >= _PROPOSED_BBOX_ADAPTIVE_MEDIUM_SCORE:
        ratio = (
            confidence_score - _PROPOSED_BBOX_ADAPTIVE_MEDIUM_SCORE
        ) / max(
            _PROPOSED_BBOX_ADAPTIVE_HIGH_SCORE
            - _PROPOSED_BBOX_ADAPTIVE_MEDIUM_SCORE,
            _EPS,
        )
        threshold = _LEGACY_SCALE_THRESHOLD - ratio * (
            _LEGACY_SCALE_THRESHOLD - medium_threshold
        )
        tier = "medium"
    else:
        threshold = _LEGACY_SCALE_THRESHOLD
        tier = "low"

    threshold = _clamp(threshold, min_threshold, _LEGACY_SCALE_THRESHOLD)
    reason = (
        f"tier={tier}; score={confidence_score:.2f}; virtual_edges={virtual_edge_count}/2; "
        f"strong=({int(top_strong)},{int(left_strong)}); span={avg_span:.2f}; "
        f"pool={pool_ratio:.2f}; rect={rect_score:.2f}"
    )
    return threshold, confidence_score, tier, reason


def _probe_extent_candidates(
    *,
    detected_lines: list[DetectedGridLineInfo],
    stamp_bbox_base: fitz.Rect,
    origin: str,
    template_h_boundaries: list[tuple[str, float]],
    template_v_boundaries: list[tuple[str, float]],
) -> dict[str, object]:
    """Infer a deformed stamp bbox from boundary-like detected lines."""
    expected_h_span = max(stamp_bbox_base.width, _EPS)
    expected_v_span = max(stamp_bbox_base.height, _EPS)

    top_line, top_strong, top_medium = _pick_boundary_like_line(
        detected_lines=detected_lines,
        orientation="h",
        expected_span_pts=expected_h_span,
        edge="top",
    )
    left_line, left_strong, left_medium = _pick_boundary_like_line(
        detected_lines=detected_lines,
        orientation="v",
        expected_span_pts=expected_v_span,
        edge="left",
    )
    right_line, right_strong, right_medium = _pick_boundary_like_line(
        detected_lines=detected_lines,
        orientation="v",
        expected_span_pts=expected_v_span,
        edge="right",
    )
    bottom_line, bottom_strong, bottom_medium = _pick_boundary_like_line(
        detected_lines=detected_lines,
        orientation="h",
        expected_span_pts=expected_h_span,
        edge="bottom",
    )
    top_pos = top_line.pos_pts if top_line is not None else stamp_bbox_base.y0
    bottom_pos = bottom_line.pos_pts if bottom_line is not None else stamp_bbox_base.y1
    vertical_cover_tol = max(expected_v_span * 0.05, 8.0)

    def _pick_covering_vertical(
        preferred: DetectedGridLineInfo | None,
        strong_pool: list[DetectedGridLineInfo],
        medium_pool: list[DetectedGridLineInfo],
        edge_name: str,
    ) -> DetectedGridLineInfo | None:
        def _covers(line: DetectedGridLineInfo) -> bool:
            return (
                line.span_lo_pts <= top_pos + vertical_cover_tol
                and line.span_hi_pts >= bottom_pos - vertical_cover_tol
            )

        def _sort_key(line: DetectedGridLineInfo) -> tuple[float, float, str]:
            span_pts = line.span_hi_pts - line.span_lo_pts
            edge_pos = line.pos_pts if edge_name == "left" else -line.pos_pts
            return (edge_pos, -span_pts, line.id)

        if preferred is not None and _covers(preferred):
            return preferred
        for pool in (strong_pool, medium_pool):
            covering = [line for line in pool if _covers(line)]
            if covering:
                return sorted(covering, key=_sort_key)[0]
        return preferred

    left_line = _pick_covering_vertical(left_line, left_strong, left_medium, "left")
    right_line = _pick_covering_vertical(right_line, right_strong, right_medium, "right")

    virtual_h_candidates = _collect_boundary_like_detected_positions(
        detected_lines=detected_lines,
        orientation="h",
        expected_span_pts=expected_h_span,
        edge="top",
    )
    virtual_v_candidates = _collect_boundary_like_detected_positions(
        detected_lines=detected_lines,
        orientation="v",
        expected_span_pts=expected_v_span,
        edge="left",
    )
    virtual_v_covering = [
        line for line in virtual_v_candidates
        if (
            line.span_lo_pts <= top_pos + vertical_cover_tol
            and line.span_hi_pts >= bottom_pos - vertical_cover_tol
        )
    ]
    if virtual_v_covering:
        virtual_v_candidates = virtual_v_covering

    virtual_top = _infer_virtual_outer_edge(
        template_positions=template_h_boundaries,
        detected_candidates=virtual_h_candidates,
    )
    virtual_left = _infer_virtual_outer_edge(
        template_positions=template_v_boundaries,
        detected_candidates=virtual_v_candidates,
    )

    x0_candidates: list[float] = []
    x1_candidates: list[float] = []
    y0_candidates: list[float] = []
    y1_candidates: list[float] = []
    if top_line is not None:
        x0_candidates.append(top_line.span_lo_pts)
        x1_candidates.append(top_line.span_hi_pts)
        y0_candidates.append(top_line.pos_pts)
    if bottom_line is not None:
        x0_candidates.append(bottom_line.span_lo_pts)
        x1_candidates.append(bottom_line.span_hi_pts)
        y1_candidates.append(bottom_line.pos_pts)
    if left_line is not None:
        x0_candidates.append(left_line.pos_pts)
        y0_candidates.append(left_line.span_lo_pts)
        y1_candidates.append(left_line.span_hi_pts)
    if right_line is not None:
        x1_candidates.append(right_line.pos_pts)
        y0_candidates.append(right_line.span_lo_pts)
        y1_candidates.append(right_line.span_hi_pts)

    proposed_bbox = fitz.Rect(stamp_bbox_base)
    inferred_edges: list[str] = []
    if origin.endswith("bottom_right"):
        if virtual_top["virtual_pos_pts"] != "":
            proposed_bbox.y0 = float(virtual_top["virtual_pos_pts"])
            inferred_edges.append("top_virtual")
        elif y0_candidates:
            proposed_bbox.y0 = min(y0_candidates)
            inferred_edges.append("top")
        if virtual_left["virtual_pos_pts"] != "":
            proposed_bbox.x0 = float(virtual_left["virtual_pos_pts"])
            inferred_edges.append("left_virtual")
        elif x0_candidates:
            proposed_bbox.x0 = min(x0_candidates)
            inferred_edges.append("left")
    elif origin.endswith("bottom_left"):
        if virtual_top["virtual_pos_pts"] != "":
            proposed_bbox.y0 = float(virtual_top["virtual_pos_pts"])
            inferred_edges.append("top_virtual")
        elif y0_candidates:
            proposed_bbox.y0 = min(y0_candidates)
            inferred_edges.append("top")
        if x1_candidates:
            proposed_bbox.x1 = max(x1_candidates)
            inferred_edges.append("right")
    elif origin.endswith("top_right"):
        if y1_candidates:
            proposed_bbox.y1 = max(y1_candidates)
            inferred_edges.append("bottom")
        if virtual_left["virtual_pos_pts"] != "":
            proposed_bbox.x0 = float(virtual_left["virtual_pos_pts"])
            inferred_edges.append("left_virtual")
        elif x0_candidates:
            proposed_bbox.x0 = min(x0_candidates)
            inferred_edges.append("left")
    else:
        if y1_candidates:
            proposed_bbox.y1 = max(y1_candidates)
            inferred_edges.append("bottom")
        if x1_candidates:
            proposed_bbox.x1 = max(x1_candidates)
            inferred_edges.append("right")

    def _line_payload(
        line: DetectedGridLineInfo | None,
        expected_span: float,
    ) -> dict[str, float | str]:
        if line is None:
            return {
                "id": "",
                "pos_pts": "",
                "span_pts": "",
                "span_ratio_to_template": "",
            }
        span_pts = line.span_hi_pts - line.span_lo_pts
        span_ratio = min(span_pts, expected_span) / max(span_pts, expected_span, _EPS)
        return {
            "id": line.id,
            "pos_pts": round(line.pos_pts, 3),
            "span_pts": round(span_pts, 3),
            "span_ratio_to_template": round(span_ratio, 6),
        }

    result: dict[str, object] = {
        "extent_probe_top_candidate_count_strong": len(top_strong),
        "extent_probe_top_candidate_count_medium": len(top_medium),
        "extent_probe_left_candidate_count_strong": len(left_strong),
        "extent_probe_left_candidate_count_medium": len(left_medium),
        "extent_probe_virtual_top_candidate_count": len(virtual_h_candidates),
        "extent_probe_virtual_left_candidate_count": len(virtual_v_candidates),
        "extent_probe_bottom_candidate_count_strong": len(bottom_strong),
        "extent_probe_bottom_candidate_count_medium": len(bottom_medium),
        "extent_probe_right_candidate_count_strong": len(right_strong),
        "extent_probe_right_candidate_count_medium": len(right_medium),
        "extent_probe_inferred_edges": ",".join(inferred_edges),
        "extent_probe_has_proposed_bbox": bool(inferred_edges),
        "extent_probe_line_source": "merged_prefilter",
        "extent_probe_proposed_bbox_x0_pts": round(proposed_bbox.x0, 3),
        "extent_probe_proposed_bbox_y0_pts": round(proposed_bbox.y0, 3),
        "extent_probe_proposed_bbox_x1_pts": round(proposed_bbox.x1, 3),
        "extent_probe_proposed_bbox_y1_pts": round(proposed_bbox.y1, 3),
        "extent_probe_base_bbox_x0_pts": round(stamp_bbox_base.x0, 3),
        "extent_probe_base_bbox_y0_pts": round(stamp_bbox_base.y0, 3),
        "extent_probe_base_bbox_x1_pts": round(stamp_bbox_base.x1, 3),
        "extent_probe_base_bbox_y1_pts": round(stamp_bbox_base.y1, 3),
        "extent_probe_proposed_dx0_pts": round(proposed_bbox.x0 - stamp_bbox_base.x0, 3),
        "extent_probe_proposed_dy0_pts": round(proposed_bbox.y0 - stamp_bbox_base.y0, 3),
        "extent_probe_proposed_dx1_pts": round(proposed_bbox.x1 - stamp_bbox_base.x1, 3),
        "extent_probe_proposed_dy1_pts": round(proposed_bbox.y1 - stamp_bbox_base.y1, 3),
    }
    top_payload = _line_payload(top_line, expected_h_span)
    left_payload = _line_payload(left_line, expected_v_span)
    bottom_payload = _line_payload(bottom_line, expected_h_span)
    right_payload = _line_payload(right_line, expected_v_span)
    for prefix, payload in (
        ("extent_probe_top", top_payload),
        ("extent_probe_left", left_payload),
        ("extent_probe_bottom", bottom_payload),
        ("extent_probe_right", right_payload),
    ):
        for key, value in payload.items():
            result[f"{prefix}_{key}"] = value
    for key, value in virtual_top.items():
        result[f"extent_probe_virtual_top_{key}"] = value
    for key, value in virtual_left.items():
        result[f"extent_probe_virtual_left_{key}"] = value
    return result


def adapt_all_fields(
    template: StampTemplate,
    frame: FrameInfo,
    fitz_page: fitz.Page,
    cfg: dict | None = None,
) -> dict[str, AdaptResult]:
    """Unified adaptation: boundary-snap with shape/position/adjacency scoring.

    Returns dict keyed by positional key (``field_id`` for unique IDs,
    ``field_id#N`` for duplicates). Use ``adapt_all_fields_indexed`` if you
    need results aligned with ``template.fields`` by position.
    """
    return dict(adapt_all_fields_indexed(template, frame, fitz_page, cfg))


def adapt_all_fields_indexed(
    template: StampTemplate,
    frame: FrameInfo,
    fitz_page: fitz.Page,
    cfg: dict | None = None,
) -> list[tuple[str, AdaptResult]]:
    """Same as ``adapt_all_fields`` but returns an ordered list of (key, result)
    pairs aligned 1-to-1 with ``template.fields``.

    Each field gets a unique positional key (``field_id`` or ``field_id#N``).
    Anchors, adjacency, and snapping all operate per position — no collisions.
    """
    global _CACHED_V2_CFG
    if cfg is not None:
        cfg_eff = cfg
    else:
        if _CACHED_V2_CFG is None:
            _CACHED_V2_CFG = _load_v2_config()
        cfg_eff = _CACHED_V2_CFG
    stamp_bbox = _template_stamp_bbox(template, frame)
    if stamp_bbox is None:
        return []

    detected = get_detected_stamp_cells(
        fitz_page=fitz_page,
        stamp_bbox=stamp_bbox,
        tolerance_mm=template.grid_tolerance_detected_mm,
        cfg=cfg_eff,
        template=template,
    )

    keys = _make_unique_keys(template.fields)

    anchors = _match_anchor_cells_indexed(
        template=template,
        frame=frame,
        detected_cells=detected,
        stamp_bbox=stamp_bbox,
        keys=keys,
    )
    anchor_key_set = {a.template_field_id for a in anchors}

    x_lines, y_lines = _extract_grid_lines(detected, template.grid_tolerance_detected_mm)

    max_snap_pts = template.snap_max_distance_mm * SCALE
    score_threshold = getattr(template, 'cascade_score_threshold', 0.4)
    max_shape_ratio = getattr(template, 'max_shape_change_ratio', 2.5)

    _snap_4_boundaries._max_shape_change = max_shape_ratio

    adjacency = _build_adjacency_graph_keyed(
        fields=template.fields,
        keys=keys,
        frame=frame,
        tol_mm=template.grid_tolerance_template_mm,
    )

    matched_so_far: dict[str, tuple[CellBbox, float]] = {}

    for a in anchors:
        snapped, _n = _snap_4_boundaries(a.detected_bbox, x_lines, y_lines, max_snap_pts)
        matched_so_far[a.template_field_id] = (snapped, a.confidence)

    results: list[tuple[str, AdaptResult]] = []

    for idx, field in enumerate(template.fields):
        key = keys[idx]

        if field.document_property:
            zb = _zero_cell_bbox()
            results.append((key, AdaptResult(
                field_id=key,
                bbox=zb,
                status="excluded",
                approx_bbox=zb,
                matched_detected=[],
                snap_distance_mm=0.0,
            )))
            continue

        if field.outside_stamp:
            template_bbox = _template_field_bbox(field, frame)
            results.append((key, AdaptResult(
                field_id=key,
                bbox=template_bbox,
                status="excluded",
                approx_bbox=template_bbox,
                matched_detected=[],
                snap_distance_mm=0.0,
            )))
            continue

        template_bbox = _template_field_bbox(field, frame)
        approx = approximate_field_bbox(
            field=field,
            frame=frame,
            anchors=anchors,
            stamp_bbox=stamp_bbox,
        )

        if key in anchor_key_set:
            entry = matched_so_far.get(key)
            if entry:
                bbox, conf = entry
                shape = _shape_score(template_bbox, bbox)
                results.append((key, AdaptResult(
                    field_id=key,
                    bbox=bbox,
                    status="anchor",
                    approx_bbox=approx,
                    matched_detected=[bbox],
                    snap_distance_mm=_center_distance_mm(approx, bbox),
                    shape_score=shape,
                    snapped_boundaries=4,
                    confidence=conf,
                    iteration=0,
                )))
                continue

        snapped, n_snapped = _snap_4_boundaries(approx, x_lines, y_lines, max_snap_pts)
        shape = _shape_score(template_bbox, snapped)
        score = _score_candidate(
            key, snapped, approx, template_bbox,
            matched_so_far, adjacency,
        )

        if n_snapped >= 2 and score >= score_threshold and shape >= (1.0 / max(max_shape_ratio, _EPS)):
            matched_so_far[key] = (snapped, score)
            results.append((key, AdaptResult(
                field_id=key,
                bbox=snapped,
                status="matched",
                approx_bbox=approx,
                matched_detected=_find_overlapping_detected(snapped, detected),
                snap_distance_mm=_center_distance_mm(approx, snapped),
                shape_score=shape,
                snapped_boundaries=n_snapped,
                confidence=score,
                iteration=1,
            )))
        else:
            results.append((key, AdaptResult(
                field_id=key,
                bbox=approx,
                status="derived",
                approx_bbox=approx,
                matched_detected=[],
                snap_distance_mm=0.0,
                shape_score=_shape_score(template_bbox, approx),
                snapped_boundaries=0,
                confidence=0.0,
                iteration=1,
            )))

    return results


def _find_overlapping_detected(bbox: CellBbox, detected: list[CellBbox]) -> list[CellBbox]:
    """Find detected cells that overlap with bbox."""
    result = []
    for dc in detected:
        if _overlap_area(bbox, dc) > 0:
            result.append(dc)
    return result


def _template_field_bbox(field: FieldDef, frame: FrameInfo) -> CellBbox:
    return CellBbox.from_rect(field_to_fitz_rect(field, frame, padding_mm=0.0))


def _template_stamp_bbox(template: StampTemplate, frame: FrameInfo) -> fitz.Rect | None:
    rects = [
        field_to_fitz_rect(f, frame, padding_mm=0.0)
        for f in template.fields
        if _field_in_stamp_geometry_union(f)
    ]
    rects = [r for r in rects if not r.is_empty and not r.is_infinite]
    if not rects:
        return None
    x0 = min(r.x0 for r in rects)
    y0 = min(r.y0 for r in rects)
    x1 = max(r.x1 for r in rects)
    y1 = max(r.y1 for r in rects)
    return fitz.Rect(x0, y0, x1, y1)


def _anchor_score(tb: CellBbox, dc: CellBbox, stamp_w: float, stamp_h: float) -> float:
    tcx, tcy = tb.center
    dcx, dcy = dc.center
    center_dx = abs(tcx - dcx) / stamp_w
    center_dy = abs(tcy - dcy) / stamp_h
    size_dw = abs(tb.width - dc.width) / max(tb.width, _EPS)
    size_dh = abs(tb.height - dc.height) / max(tb.height, _EPS)
    return center_dx + center_dy + 0.5 * (size_dw + size_dh)


def _overlap_area(a: CellBbox, b: CellBbox) -> float:
    inter = a.to_rect() & b.to_rect()
    if inter.is_empty:
        return 0.0
    return inter.get_area()


def _iou(a: CellBbox, b: CellBbox) -> float:
    inter = _overlap_area(a, b)
    if inter <= 0:
        return 0.0
    union = a.area + b.area - inter
    if union <= 0:
        return 0.0
    return inter / union


def _merge_cells(cells: list[CellBbox]) -> CellBbox:
    x0 = min(c.x0 for c in cells)
    y0 = min(c.y0 for c in cells)
    x1 = max(c.x1 for c in cells)
    y1 = max(c.y1 for c in cells)
    return CellBbox(x0, y0, x1, y1)


def _center_distance_mm(a: CellBbox, b: CellBbox) -> float:
    ax, ay = a.center
    bx, by = b.center
    return (((ax - bx) ** 2 + (ay - by) ** 2) ** 0.5) / SCALE


# ---------------------------------------------------------------------------
#  Cell-assignment adaptation
# ---------------------------------------------------------------------------

@dataclass
class DetectedGridLineInfo:
    """Detected grid line after border-pair merge — for editor visualisation."""
    id: str              # "PDF:h_1", "PDF:v_3" etc.
    orientation: str     # "h" | "v"
    pos_pts: float
    span_lo_pts: float
    span_hi_pts: float
    cell_count: int


@dataclass
class AlignmentReport:
    """Structured feedback from the grid alignment engine."""
    verdict: str = "ok"          # "ok" | "warning" | "fail"
    total_tpl_lines: int = 0
    matched_count: int = 0
    no_match_count: int = 0
    skipped_detected_count: int = 0
    avg_match_score: float = 0.0
    problem_lines: list[str] = dc_field(default_factory=list)
    message: str = ""
    log: str = ""                # full text log for .txt file


@dataclass
class CellAssignmentInfo:
    """Diagnostics for the cell-assignment algorithm."""
    rectangularity: float
    n_detected: int
    n_fields: int
    n_assigned: int
    transform_scale_x: float
    transform_scale_y: float
    transform_dx: float
    transform_dy: float
    warnings: list[str] = dc_field(default_factory=list)
    detected_grid_lines: list[DetectedGridLineInfo] = dc_field(default_factory=list)
    snapped_line_positions: dict[str, float] = dc_field(default_factory=dict)
    alignment_report: AlignmentReport | None = None
    alignment_diagnostics: list[AlignmentLineDiag] = dc_field(default_factory=list)
    scale_threshold: float = _LEGACY_SCALE_THRESHOLD
    prealign_confidence_score: float | None = None
    prealign_confidence_tier: str = ""
    prealign_confidence_reason: str = ""
    prealign_detected_grid_lines: list[DetectedGridLineInfo] = dc_field(default_factory=list)
    extent_probe: dict[str, object] = dc_field(default_factory=dict)
    search_bbox: tuple[float, float, float, float] | None = None
    base_bbox: tuple[float, float, float, float] | None = None
    proposed_bbox: tuple[float, float, float, float] | None = None
    effective_bbox: tuple[float, float, float, float] | None = None
    detected_union_bbox: tuple[float, float, float, float] | None = None
    proposed_bbox_applied: bool = False
    n_detected_raw: int = 0
    n_detected_probe: int = 0
    n_detected_effective: int = 0
    n_filtered_oversized: int = 0
    n_filtered_outside_effective_bbox: int = 0
    detected_prefilter: dict[str, object] = dc_field(default_factory=dict)


def _check_rectangularity(
    detected: list[CellBbox],
    tol_mm: float = 0.5,
) -> float:
    """Fraction of detected cells whose 4 boundaries all align to grid lines.

    Returns 0.0–1.0. Values above ~0.7 indicate a well-formed table grid.
    """
    if not detected:
        return 0.0
    x_lines, y_lines = _extract_grid_lines(detected, tol_mm)
    tol_pts = tol_mm * SCALE
    aligned = 0
    for dc in detected:
        ok = 0
        if any(abs(dc.x0 - xl) <= tol_pts for xl in x_lines):
            ok += 1
        if any(abs(dc.x1 - xl) <= tol_pts for xl in x_lines):
            ok += 1
        if any(abs(dc.y0 - yl) <= tol_pts for yl in y_lines):
            ok += 1
        if any(abs(dc.y1 - yl) <= tol_pts for yl in y_lines):
            ok += 1
        if ok == 4:
            aligned += 1
    return aligned / len(detected)


def _compute_global_transform(
    template_bbox: CellBbox,
    detected_bbox: CellBbox,
) -> tuple[float, float, float, float]:
    """Compute (scale_x, scale_y, dx, dy) from template stamp bbox to detected stamp bbox.

    Anchors the top-left corner of template to top-left of detected,
    then scales to match bottom-right.
    """
    tw = max(template_bbox.width, _EPS)
    th = max(template_bbox.height, _EPS)
    dw = max(detected_bbox.width, _EPS)
    dh = max(detected_bbox.height, _EPS)
    sx = dw / tw
    sy = dh / th
    dx = detected_bbox.x0 - template_bbox.x0 * sx
    dy = detected_bbox.y0 - template_bbox.y0 * sy
    return sx, sy, dx, dy


def _apply_transform(
    bbox: CellBbox, sx: float, sy: float, dx: float, dy: float,
) -> CellBbox:
    return CellBbox(
        x0=bbox.x0 * sx + dx,
        y0=bbox.y0 * sy + dy,
        x1=bbox.x1 * sx + dx,
        y1=bbox.y1 * sy + dy,
    )


def _cell_match_cost(expected: CellBbox, candidate: CellBbox) -> float:
    """Cost of assigning a template field (at expected position) to a detected cell.

    Lower = better. Combines center distance and size similarity.
    """
    ecx, ecy = expected.center
    ccx, ccy = candidate.center
    dist_pts = ((ecx - ccx) ** 2 + (ecy - ccy) ** 2) ** 0.5

    area_e = max(expected.area, _EPS)
    area_c = max(candidate.area, _EPS)
    log_ratio = abs(math.log(area_c / area_e))

    return dist_pts / SCALE + 5.0 * log_ratio


_INF_ASSIGN = 1e12


def _solve_cell_assignment_order_greedy(
    field_data: list[tuple[int, str, CellBbox, CellBbox]],
    detected: list[CellBbox],
    max_dist_pts: float,
) -> dict[int, tuple[int, float, float]]:
    """Greedy: one best cell per field in template order (first field wins ties)."""
    assignments: dict[int, tuple[int, float, float]] = {}
    used_cells: set[int] = set()
    for fi, (_idx, _key, _tb, expected) in enumerate(field_data):
        best: tuple[float, int, float] | None = None
        ecx, ecy = expected.center
        for ci, dc in enumerate(detected):
            if ci in used_cells:
                continue
            ccx, ccy = dc.center
            d = ((ecx - ccx) ** 2 + (ecy - ccy) ** 2) ** 0.5
            if d > max_dist_pts:
                continue
            c = _cell_match_cost(expected, dc)
            if best is None or c < best[0]:
                best = (c, ci, d)
        if best is not None:
            c, ci, d = best
            assignments[fi] = (ci, c, d)
            used_cells.add(ci)
    return assignments


def _solve_cell_assignment_linear_sum(
    field_data: list[tuple[int, str, CellBbox, CellBbox]],
    detected: list[CellBbox],
    max_dist_pts: float,
) -> dict[int, tuple[int, float, float]]:
    """Optimal min-cost matching (fields ↔ cells) via Hungarian / linear assignment.

    Rectangular scipy: each column (cell) at most one field; some fields unmatched.
    Columns with no finite edge are dropped. Rows with no candidate are skipped.
    Falls back to :func:`_solve_cell_assignment_order_greedy` if scipy is missing.
    """
    n_f = len(field_data)
    n_c = len(detected)
    if n_f == 0 or n_c == 0:
        return {}

    cost: list[list[float]] = [[_INF_ASSIGN] * n_c for _ in range(n_f)]
    dist_cache: dict[tuple[int, int], float] = {}

    for fi, (_idx, _key, _tb, expected) in enumerate(field_data):
        ecx, ecy = expected.center
        for ci, dc in enumerate(detected):
            ccx, ccy = dc.center
            d = ((ecx - ccx) ** 2 + (ecy - ccy) ** 2) ** 0.5
            if d > max_dist_pts:
                continue
            c = _cell_match_cost(expected, dc)
            cost[fi][ci] = c
            dist_cache[(fi, ci)] = d

    col_keep = [
        j for j in range(n_c)
        if any(cost[i][j] < _INF_ASSIGN * 0.5 for i in range(n_f))
    ]
    if not col_keep:
        return {}

    row_fi = [
        fi for fi in range(n_f)
        if any(cost[fi][j] < _INF_ASSIGN * 0.5 for j in col_keep)
    ]
    if not row_fi:
        return {}

    n_r = len(row_fi)
    n_k = len(col_keep)
    cost_sub = [
        [cost[fi][col_keep[j]] for j in range(n_k)]
        for fi in row_fi
    ]

    try:
        import numpy as np
        from scipy.optimize import linear_sum_assignment
    except ImportError:
        return _solve_cell_assignment_order_greedy(field_data, detected, max_dist_pts)

    c_arr = np.asarray(cost_sub, dtype=np.float64)
    row_ind, col_ind = linear_sum_assignment(c_arr)

    assignments: dict[int, tuple[int, float, float]] = {}
    for ri, cj in zip(row_ind, col_ind):
        fi = row_fi[ri]
        orig_c = col_keep[cj]
        cval = cost[fi][orig_c]
        if cval >= _INF_ASSIGN * 0.5:
            continue
        d = dist_cache.get((fi, orig_c), 0.0)
        assignments[fi] = (orig_c, cval, d)

    return assignments


def adapt_by_cell_assignment(
    template: StampTemplate,
    frame: FrameInfo,
    fitz_page: fitz.Page,
    cfg: dict | None = None,
    max_match_distance_mm: float = 10.0,
    timing: dict[str, float] | None = None,
) -> tuple[list[tuple[str, AdaptResult]], CellAssignmentInfo]:
    """Cell-assignment adaptation: match each template field to the nearest
    detected cell, using a global top-left-anchored transform.

    Returns (results aligned with template.fields, diagnostics).
    """
    global _CACHED_V2_CFG
    cfg_eff = cfg if cfg is not None else (_CACHED_V2_CFG or _load_v2_config())
    if cfg is None and _CACHED_V2_CFG is None:
        _CACHED_V2_CFG = cfg_eff

    stamp_bbox = _template_stamp_bbox(template, frame)
    if stamp_bbox is None:
        info = CellAssignmentInfo(0.0, 0, 0, 0, 1.0, 1.0, 0.0, 0.0,
                                  warnings=["No template fields"])
        return [], info

    prealign_cfg = _grid_prealign_settings(cfg_eff)
    page_rect = fitz.Rect(0.0, 0.0, fitz_page.rect.width, fitz_page.rect.height)
    search_bbox = (
        _expand_bbox_for_anchor(
            stamp_bbox,
            origin=template.origin,
            x_frac=float(prealign_cfg["bbox_overflow_x_frac"]),
            y_frac=float(prealign_cfg["bbox_overflow_y_frac"]),
            page_rect=page_rect,
        )
        if bool(prealign_cfg["enabled"])
        else fitz.Rect(stamp_bbox)
    )
    detected_raw = get_detected_stamp_cells(
        fitz_page=fitz_page,
        stamp_bbox=search_bbox,
        tolerance_mm=template.grid_tolerance_detected_mm,
        cfg=cfg_eff,
        template=template,
        timing=timing,
    )

    # Filter out frame-level artifacts: cells whose area exceeds a
    # significant fraction of the stamp area are not real stamp cells.
    stamp_area = max(search_bbox.get_area(), _EPS)
    detected_for_probe = [
        dc for dc in detected_raw
        if dc.area <= stamp_area * _MAX_CELL_AREA_FRACTION
    ]

    keys = _make_unique_keys(template.fields)
    warnings: list[str] = []

    n_filtered = len(detected_raw) - len(detected_for_probe)
    if n_filtered:
        warnings.append(f"Filtered {n_filtered} oversized cell(s) (frame artifacts)")

    # --- template bounding box ---
    t_bboxes = []
    for f in template.fields:
        if _field_in_stamp_geometry_union(f):
            t_bboxes.append(_template_field_bbox(f, frame))
    if not t_bboxes:
        info = CellAssignmentInfo(0.0, len(detected_for_probe), 0, 0,
                                  1.0, 1.0, 0.0, 0.0, warnings)
        return [], info

    t_all = CellBbox(
        min(b.x0 for b in t_bboxes), min(b.y0 for b in t_bboxes),
        max(b.x1 for b in t_bboxes), max(b.y1 for b in t_bboxes),
    )

    prefilter_diag: dict[str, object] = {}
    detected_for_probe, prefilter_diag = prefilter_detected_cells_for_stamp_table(
        detected_for_probe,
        stamp_bbox_base=(stamp_bbox.x0, stamp_bbox.y0, stamp_bbox.x1, stamp_bbox.y1),
        t_all=t_all,
        tol_mm=float(template.grid_tolerance_detected_mm),
        scale_pts_per_mm=float(SCALE),
        cfg=cfg_eff,
    )
    if bool(prefilter_diag.get("applied")):
        warnings.append(
            "Detected-cell island prefilter: "
            f"{int(prefilter_diag['n_output'])}/{int(prefilter_diag['n_input'])} cells kept"
        )
    else:
        reason = prefilter_diag.get("reason")
        if isinstance(reason, str) and reason.startswith("prefilter skipped"):
            warnings.append(reason)

    effective_bbox = fitz.Rect(stamp_bbox)
    proposed_bbox_applied = False
    extent_probe: dict[str, object] = {}
    proposed_bbox: fitz.Rect | None = None
    prealign_detected_grid_lines: list[DetectedGridLineInfo] = []
    probe_detected_count = len(detected_for_probe)
    if bool(prealign_cfg["enabled"]) and detected_for_probe and template.grid_lines:
        template_h_boundaries = _collect_template_boundary_positions(
            template=template,
            frame=frame,
            orientation="h",
        )
        template_v_boundaries = _collect_template_boundary_positions(
            template=template,
            frame=frame,
            orientation="v",
        )
        extent_detected_lines = _build_merged_detected_line_infos(
            detected_cells=detected_for_probe,
            template=template,
        )
        prealign_detected_grid_lines = list(extent_detected_lines)
        extent_probe = _probe_extent_candidates(
            detected_lines=extent_detected_lines,
            stamp_bbox_base=stamp_bbox,
            origin=template.origin,
            template_h_boundaries=template_h_boundaries,
            template_v_boundaries=template_v_boundaries,
        )
        proposed_bbox = _extent_probe_proposed_bbox(extent_probe)
        if proposed_bbox is not None:
            proposed_bbox = fitz.Rect(proposed_bbox) & page_rect
            if (
                not proposed_bbox.is_empty
                and not proposed_bbox.is_infinite
                and proposed_bbox.width > _EPS
                and proposed_bbox.height > _EPS
            ):
                proposed_detected = _filter_cells_to_bbox(detected_for_probe, proposed_bbox)
                if proposed_detected:
                    effective_bbox = proposed_bbox
                    detected_for_probe = proposed_detected
                    proposed_bbox_applied = True
                    warnings.append("Applied proposed-bbox pre-alignment before matching")
                else:
                    warnings.append(
                        "Skipped proposed-bbox pre-alignment: inferred bbox captured no cells"
                    )

    detected = _filter_cells_to_bbox(detected_for_probe, effective_bbox)
    n_bbox_filtered = len(detected_raw) - n_filtered - len(detected)
    if n_bbox_filtered > 0:
        warnings.append(f"Filtered {n_bbox_filtered} cell(s) outside effective stamp bbox")

    # --- rectangularity check ---
    rect_score = _check_rectangularity(detected, template.grid_tolerance_detected_mm)
    if rect_score < 0.5:
        warnings.append(
            f"Low rectangularity ({rect_score:.0%}): detected cells may be unreliable"
        )

    if not detected:
        warnings.append("No detected cells (after filtering)")
        info = CellAssignmentInfo(
            rect_score,
            0,
            len(template.fields),
            0,
            1.0,
            1.0,
            0.0,
            0.0,
            warnings,
            detected_prefilter=dict(prefilter_diag),
        )
        results = []
        for idx, field in enumerate(template.fields):
            if field.document_property:
                zb = _zero_cell_bbox()
                results.append((keys[idx], AdaptResult(
                    field_id=keys[idx], bbox=zb, status="excluded",
                    approx_bbox=zb, matched_detected=[], snap_distance_mm=0.0,
                )))
                continue
            snapped_mm = snap_outside_stamp_vertical_to_frame(field, frame)
            if snapped_mm is not None:
                snapped_field = replace(field, bbox_mm=snapped_mm)
                tb = _template_field_bbox(snapped_field, frame)
            else:
                tb = _template_field_bbox(field, frame)
            results.append((keys[idx], AdaptResult(
                field_id=keys[idx], bbox=tb, status="derived",
                approx_bbox=tb, matched_detected=[], snap_distance_mm=0.0,
            )))
        return results, info

    # Proposed-bbox pre-alignment corrects the effective geometry before matching.
    clipped_for_bbox = _clip_cells_to_bbox(detected, effective_bbox)
    d_all = _union_bbox_or_fallback(clipped_for_bbox, t_all)

    # --- global transform ---
    # Default: translation only (same project = same stamp size, different position).
    # Scale applied only when stamps differ significantly. On the proposed-bbox
    # path the threshold is selected adaptively from pre-alignment confidence.
    sx_raw, sy_raw, _, _ = _compute_global_transform(t_all, d_all)
    prealign_confidence_score: float | None = None
    prealign_confidence_tier = ""
    prealign_confidence_reason = ""
    if proposed_bbox_applied:
        (
            scale_threshold,
            prealign_confidence_score,
            prealign_confidence_tier,
            prealign_confidence_reason,
        ) = _choose_adaptive_prealign_scale_threshold(
            prealign_cfg=prealign_cfg,
            extent_probe=extent_probe,
            rect_score=rect_score,
            kept_detected_count=len(detected),
            probe_detected_count=probe_detected_count,
        )
        warnings.append(
            "Adaptive proposed-bbox scale-threshold: "
            f"{scale_threshold:.1%} ({prealign_confidence_reason})"
        )
    else:
        scale_threshold = _LEGACY_SCALE_THRESHOLD
    use_scale = (
        abs(sx_raw - 1.0) > scale_threshold
        or abs(sy_raw - 1.0) > scale_threshold
    )
    if use_scale:
        sx, sy = sx_raw, sy_raw
        dx = d_all.x0 - t_all.x0 * sx
        dy = d_all.y0 - t_all.y0 * sy
        if proposed_bbox_applied:
            warnings.append(
                "Applied scaling "
                f"({sx:.3f}, {sy:.3f}) after proposed-bbox pre-alignment "
                f"(adaptive threshold {scale_threshold:.1%}, "
                f"confidence={prealign_confidence_tier or 'n/a'})"
            )
        else:
            warnings.append(
                f"Applied scaling ({sx:.3f}, {sy:.3f}) — stamp dimensions differ >5%"
            )
    else:
        sx, sy = 1.0, 1.0
        dx = d_all.x0 - t_all.x0
        dy = d_all.y0 - t_all.y0

    # --- cell assignment: linear sum (Hungarian) or greedy fallback ---
    max_dist_pts = max_match_distance_mm * SCALE

    field_data: list[tuple[int, str, CellBbox, CellBbox]] = []
    excluded_indices: set[int] = set()
    for idx, field in enumerate(template.fields):
        key = keys[idx]
        if field.outside_stamp or field.document_property:
            excluded_indices.add(idx)
            continue
        tb = _template_field_bbox(field, frame)
        expected = _apply_transform(tb, sx, sy, dx, dy)
        field_data.append((idx, key, tb, expected))

    # Optimal min-cost bipartite matching (Hungarian); fallback: greedy in template order.
    assignments = _solve_cell_assignment_linear_sum(field_data, detected, max_dist_pts)

    results: list[tuple[str, AdaptResult]] = []
    n_assigned = 0

    for idx, field in enumerate(template.fields):
        key = keys[idx]

        if idx in excluded_indices:
            if field.document_property:
                zb = _zero_cell_bbox()
                results.append((key, AdaptResult(
                    field_id=key, bbox=zb, status="excluded",
                    approx_bbox=zb, matched_detected=[], snap_distance_mm=0.0,
                    shape_score=1.0,
                )))
            else:
                snapped_mm = snap_outside_stamp_vertical_to_frame(field, frame)
                if snapped_mm is not None:
                    snapped_field = replace(field, bbox_mm=snapped_mm)
                    tb = _template_field_bbox(snapped_field, frame)
                else:
                    tb = _template_field_bbox(field, frame)
                bbox_pdf = tb
                results.append((key, AdaptResult(
                    field_id=key, bbox=bbox_pdf, status="excluded",
                    approx_bbox=bbox_pdf, matched_detected=[], snap_distance_mm=0.0,
                    shape_score=_shape_score(tb, bbox_pdf),
                )))
            continue

        fi = next(i for i, (fidx, *_) in enumerate(field_data) if fidx == idx)
        _, _, tb, expected = field_data[fi]

        if fi in assignments:
            ci, cost, dist_pts = assignments[fi]
            dc = detected[ci]
            shape = _shape_score(tb, dc)
            n_assigned += 1
            results.append((key, AdaptResult(
                field_id=key,
                bbox=dc,
                status="matched",
                approx_bbox=expected,
                matched_detected=[dc],
                snap_distance_mm=dist_pts / SCALE,
                shape_score=shape,
                snapped_boundaries=4,
                confidence=1.0 / (1.0 + cost * 0.1),
                iteration=0,
            )))
        else:
            results.append((key, AdaptResult(
                field_id=key,
                bbox=expected,
                status="derived",
                approx_bbox=expected,
                matched_detected=[],
                snap_distance_mm=0.0,
                shape_score=_shape_score(tb, expected),
                confidence=0.0,
                iteration=0,
            )))

    # Post-process: grid-line bindings (deterministic) or heuristic gap-closing.
    det_gl_info: list[DetectedGridLineInfo] = []
    align_report_list: list[AlignmentReport] = []
    align_diags: list[AlignmentLineDiag] = []
    snapped: dict[str, float] = {}
    if template.grid_lines and _fields_have_bindings(template.fields):
        n_boundary = sum(
            1 for gl in template.grid_lines if gl.boundary is not None
        )
        if n_boundary >= 2:
            snapped = _align_grid_lines_ordered(
                template, frame, sx, sy, dx, dy, detected,
                stamp_bbox_override=effective_bbox if proposed_bbox_applied else None,
                detected_merged_out=det_gl_info,
                diagnostics=align_diags,
                alignment_report_out=align_report_list,
                template_path=getattr(template, '_source_path', ''),
                pdf_path=cfg.get('_pdf_path', '') if cfg else '',
                page_num=cfg.get('_page_num', 0) if cfg else 0,
            )
        else:
            snapped = _transform_snap_grid_lines(
                template, frame, sx, sy, dx, dy, detected
            )
            warnings.append("Fallback to walk-snap: < 2 boundary grid lines")
        # --- DISABLED: cell-consensus refine uses cell-assignment edges
        #     that may be split-artifacts (see T9.12 discussion) ---
        # snapped, refine_deltas = _refine_grid_from_cells(
        #     snapped, results, template.fields, keys,
        # )
        # if refine_deltas and align_report_list:
        #     report = align_report_list[0]
        #     refine_lines = ["\nREFINE (cell-consensus, max_delta=2.0pts):"]
        #     for lid, (old_p, new_p) in sorted(refine_deltas.items()):
        #         refine_lines.append(
        #             f"  {lid}: {old_p:.1f} -> {new_p:.1f} ({new_p - old_p:+.1f})"
        #         )
        #     report.log += "\n".join(refine_lines) + "\n"

        exp_keys = expand_fields_by_bindings(results, template.fields, keys, snapped)
        if exp_keys:
            warnings.append(f"Grid-bound {len(exp_keys)} field(s): {exp_keys}")

        # --- DISABLED: cell-assignment validate overrides correct grid-bound
        #     positions with split-artifact cell edges (see T9.12 discussion) ---
        # corr_keys = _validate_grid_against_assignment(
        #     results, template.fields, keys
        # )
        # if corr_keys:
        #     warnings.append(f"Cell-feedback corrected {len(corr_keys)} field(s): {corr_keys}")
    else:
        t_bboxes: dict[str, CellBbox] = {}
        for idx_t, fd in enumerate(template.fields):
            if _field_in_stamp_geometry_union(fd):
                t_bboxes[keys[idx_t]] = _template_field_bbox(fd, frame)
        stamp_cb = CellBbox.from_rect(stamp_bbox) if stamp_bbox is not None else None
        ext_keys = close_fragmented_gaps(results, t_bboxes, stamp_bbox=stamp_cb)
        if ext_keys:
            warnings.append(f"Extended {len(ext_keys)} fragmented field(s): {ext_keys}")

    a_report = align_report_list[0] if align_report_list else None
    if a_report and a_report.verdict != "ok":
        warnings.append(f"Alignment {a_report.verdict}: {a_report.message}")

    info = CellAssignmentInfo(
        rectangularity=rect_score,
        n_detected=len(detected),
        n_fields=len(template.fields),
        n_assigned=n_assigned,
        transform_scale_x=sx,
        transform_scale_y=sy,
        transform_dx=dx,
        transform_dy=dy,
        warnings=warnings,
        detected_grid_lines=det_gl_info,
        snapped_line_positions=snapped,
        alignment_report=a_report,
        scale_threshold=scale_threshold,
        prealign_confidence_score=prealign_confidence_score,
        prealign_confidence_tier=prealign_confidence_tier,
        prealign_confidence_reason=prealign_confidence_reason,
        prealign_detected_grid_lines=prealign_detected_grid_lines,
        extent_probe=dict(extent_probe),
        search_bbox=(search_bbox.x0, search_bbox.y0, search_bbox.x1, search_bbox.y1),
        base_bbox=(stamp_bbox.x0, stamp_bbox.y0, stamp_bbox.x1, stamp_bbox.y1),
        proposed_bbox=(
            (proposed_bbox.x0, proposed_bbox.y0, proposed_bbox.x1, proposed_bbox.y1)
            if proposed_bbox is not None
            else None
        ),
        effective_bbox=(
            effective_bbox.x0,
            effective_bbox.y0,
            effective_bbox.x1,
            effective_bbox.y1,
        ),
        detected_union_bbox=(d_all.x0, d_all.y0, d_all.x1, d_all.y1),
        proposed_bbox_applied=proposed_bbox_applied,
        n_detected_raw=len(detected_raw),
        n_detected_probe=probe_detected_count,
        n_detected_effective=len(detected),
        n_filtered_oversized=n_filtered,
        n_filtered_outside_effective_bbox=max(n_bbox_filtered, 0),
        alignment_diagnostics=list(align_diags),
        detected_prefilter=dict(prefilter_diag),
    )
    return results, info


# ---------------------------------------------------------------------------
#  Grid-line binding: transform, snap, expand
# ---------------------------------------------------------------------------

@dataclass
class _DetectedGridLine:
    """Detected grid line enriched with span and cell-count metadata."""
    pos: float
    span_lo: float
    span_hi: float
    cell_count: int

    @property
    def span_length(self) -> float:
        return max(0.0, self.span_hi - self.span_lo)


@dataclass
class GridLineSnapDiag:
    """Per-line snap diagnostics for debug snapshots."""
    line_id: str
    orientation: str
    pos_transformed: float
    pos_snapped: float
    method: str          # "anchor" | "walk" | "fallback"
    candidate_count: int
    snap_distance_pts: float
    window_lo: float
    window_hi: float


def _build_detected_grid_lines(
    detected: list[CellBbox],
    tolerance_mm: float,
) -> tuple[list[_DetectedGridLine], list[_DetectedGridLine]]:
    """Build enriched detected grid lines from cell boundaries.

    Returns ``(h_lines, v_lines)`` each sorted by ``pos`` ascending.
    H-lines have ``pos`` = y-coordinate, ``span`` along x-axis.
    V-lines have ``pos`` = x-coordinate, ``span`` along y-axis.
    """
    if not detected:
        return [], []
    tol = tolerance_mm * SCALE

    def _build(items: list[tuple[float, float, float]]) -> list[_DetectedGridLine]:
        if not items:
            return []
        items.sort(key=lambda x: x[0])
        clusters: list[list[tuple[float, float, float]]] = [[items[0]]]
        for item in items[1:]:
            if abs(item[0] - clusters[-1][-1][0]) <= tol:
                clusters[-1].append(item)
            else:
                clusters.append([item])
        result: list[_DetectedGridLine] = []
        for cl in clusters:
            pos = sum(c[0] for c in cl) / len(cl)
            span_lo = min(c[1] for c in cl)
            span_hi = max(c[2] for c in cl)
            result.append(_DetectedGridLine(pos, span_lo, span_hi, len(cl)))
        result.sort(key=lambda l: l.pos)
        return result

    h_items: list[tuple[float, float, float]] = []
    v_items: list[tuple[float, float, float]] = []
    for c in detected:
        h_items.append((c.y0, c.x0, c.x1))
        h_items.append((c.y1, c.x0, c.x1))
        v_items.append((c.x0, c.y0, c.y1))
        v_items.append((c.x1, c.y0, c.y1))

    return _build(h_items), _build(v_items)


def _virtual_merge_template_lines(
    grid_lines: list,
    pos_tol_mm: float = 2.0,
    gap_tol_mm: float = 2.0,
) -> dict[str, float]:
    """Identify split template lines and compute their combined span length.

    Two lines of the same orientation whose ``pos_mm`` are within *pos_tol_mm*
    and whose spans are adjacent (gap <= *gap_tol_mm*) are merged virtually.

    Returns ``{line_id: merged_span_length_mm}`` for every input line, where
    the merged span is the union span of the line's merge group.
    """
    from pdf_parsing_v2_engine.models import TemplateGridLine

    by_orient: dict[str, list[TemplateGridLine]] = {"h": [], "v": []}
    for gl in grid_lines:
        by_orient.setdefault(gl.orientation, []).append(gl)

    merged_spans: dict[str, float] = {}

    for orient, lines in by_orient.items():
        lines_sorted = sorted(lines, key=lambda l: l.pos_mm)
        used: set[int] = set()
        for i, a in enumerate(lines_sorted):
            if i in used:
                continue
            group_ids = [a.id]
            s_lo = min(a.start_mm, a.end_mm)
            s_hi = max(a.start_mm, a.end_mm)
            for j in range(i + 1, len(lines_sorted)):
                if j in used:
                    continue
                b = lines_sorted[j]
                if abs(b.pos_mm - a.pos_mm) > pos_tol_mm:
                    break
                b_lo = min(b.start_mm, b.end_mm)
                b_hi = max(b.start_mm, b.end_mm)
                gap = max(0.0, max(b_lo - s_hi, s_lo - b_hi))
                if gap <= gap_tol_mm:
                    s_lo = min(s_lo, b_lo)
                    s_hi = max(s_hi, b_hi)
                    group_ids.append(b.id)
                    used.add(j)
            used.add(i)
            span = s_hi - s_lo
            for gid in group_ids:
                merged_spans[gid] = span

    return merged_spans


def _fields_have_bindings(fields: list[FieldDef]) -> bool:
    """Return True if at least one non-excluded field has a bound_* set."""
    return any(
        f.bound_top or f.bound_bottom or f.bound_left or f.bound_right
        for f in fields
        if _field_in_stamp_geometry_union(f)
    )


def _transform_snap_grid_lines(
    template: StampTemplate,
    frame: FrameInfo,
    sx: float,
    sy: float,
    dx: float,
    dy: float,
    detected: list[CellBbox],
    diagnostics: list[GridLineSnapDiag] | None = None,
) -> dict[str, float]:
    """Anchor-propagate snap with co-located grouping and used-line tracking.

    Returns ``{line_id: snapped_pos_fitz_pts}`` for all template grid lines.

    If *diagnostics* is a list, per-line ``GridLineSnapDiag`` records are appended.

    Algorithm (Anchor-Propagate v3)
    -------------------------------
    1. Build enriched detected grid lines with span metadata.
    2. Convert each template line to fitz pts; apply global transform.
    3. **Group** co-located template lines (pos within 2mm tolerance) so that
       split segments of the same structural line snap together.
    4. **Anchor**: boundary lines snap to best detected line.
    5. **Local transform**: compute scale+offset from boundary anchors.
    6. **Walk**: process groups in order of interpolated position; each group
       claims ONE detected line (shared by all members); used-line tracking
       prevents two groups from grabbing the same detected line.
    """
    _CO_LOCATE_PTS = 2.0 * SCALE

    if not template.grid_lines:
        return {}

    h_det, v_det = _build_detected_grid_lines(
        detected, template.grid_tolerance_detected_mm
    ) if detected else ([], [])

    snap_max_pts = template.snap_max_distance_mm * SCALE

    merged_spans_mm = _virtual_merge_template_lines(template.grid_lines)

    # --- step 1: convert template lines to fitz, apply transform ---
    _TLine = tuple[str, str, float, float, float, str | None, float]
    t_lines: list[_TLine] = []
    for gl in template.grid_lines:
        orientation, pos_fitz, start_fitz, end_fitz, bnd, line_id = grid_line_to_fitz_pts(
            gl, frame, template.origin
        )
        if orientation == "h":
            pos_t = pos_fitz * sy + dy
        else:
            pos_t = pos_fitz * sx + dx
        msm = merged_spans_mm.get(line_id, abs(gl.end_mm - gl.start_mm))
        t_lines.append((line_id, orientation, pos_t, start_fitz, end_fitz, bnd, msm))

    stamp_bbox = _template_stamp_bbox(template, frame)
    if stamp_bbox is not None:
        stamp_w_pts = max(1.0, stamp_bbox.width)
        stamp_h_pts = max(1.0, stamp_bbox.height)
    else:
        stamp_w_pts = max(1.0, abs(frame.x1 - frame.x0))
        stamp_h_pts = max(1.0, abs(frame.y1 - frame.y0))

    snapped: dict[str, float] = {}

    def _claim_idx(pool: list[_DetectedGridLine], used: set[int], pos: float) -> None:
        for i, dl in enumerate(pool):
            if abs(dl.pos - pos) < 0.01:
                used.add(i)
                return

    def _pick_best(
        pool: list[_DetectedGridLine],
        used: set[int],
        pos_t: float,
        window_lo: float,
        window_hi: float,
        t_span_lo: float,
        t_span_hi: float,
        stamp_extent: float,
        merged_span_mm: float,
    ) -> tuple[float | None, int]:
        """Find best unused detected line in ``[window_lo, window_hi]``."""
        margin = 1.5
        candidates = [
            (i, dl) for i, dl in enumerate(pool)
            if i not in used
            and (window_lo - margin) <= dl.pos <= (window_hi + margin)
        ]
        if not candidates:
            return None, 0

        t_span_len = max(1.0, abs(t_span_hi - t_span_lo))
        full_span_threshold = stamp_extent * 0.45

        def _score(idx_dl: tuple[int, _DetectedGridLine]) -> float:
            _idx, dl = idx_dl
            dist = abs(dl.pos - pos_t)
            if dist > snap_max_pts * 2.5:
                return -1e9
            dist_score = 1.0 - min(dist / max(snap_max_pts, 1.0), 1.0)
            span_overlap = max(0.0,
                min(dl.span_hi, t_span_hi) - max(dl.span_lo, t_span_lo))
            span_ratio = min(span_overlap / t_span_len, 1.0)
            importance = min(dl.span_length / max(stamp_extent * 0.3, 1.0), 1.0)
            if merged_span_mm * SCALE > full_span_threshold and dl.span_length > full_span_threshold:
                importance = min(importance + 0.3, 1.0)
            return 0.50 * dist_score + 0.30 * span_ratio + 0.20 * importance

        best_pair = max(candidates, key=_score)
        if _score(best_pair) <= -1e9:
            return None, len(candidates)
        return best_pair[1].pos, len(candidates)

    # --- step 2: anchor boundary lines ---
    h_used: set[int] = set()
    v_used: set[int] = set()

    for orient, pool, used, extent in [
        ("h", h_det, h_used, stamp_w_pts),
        ("v", v_det, v_used, stamp_h_pts),
    ]:
        boundaries = [t for t in t_lines if t[1] == orient and t[5] is not None]
        for lid, _o, pos_t, sp_lo, sp_hi, _bnd, msm in boundaries:
            if not pool:
                snapped[lid] = pos_t
                if diagnostics is not None:
                    diagnostics.append(GridLineSnapDiag(
                        lid, orient, pos_t, pos_t, "fallback", 0, 0.0,
                        pos_t - snap_max_pts, pos_t + snap_max_pts))
                continue
            w_lo = pos_t - snap_max_pts * 2
            w_hi = pos_t + snap_max_pts * 2
            best, n_cand = _pick_best(pool, used, pos_t, w_lo, w_hi,
                                      sp_lo, sp_hi, extent, msm)
            result_pos = best if best is not None else pos_t
            snapped[lid] = result_pos
            if best is not None:
                _claim_idx(pool, used, best)
            if diagnostics is not None:
                diagnostics.append(GridLineSnapDiag(
                    lid, orient, pos_t, result_pos,
                    "anchor" if best is not None else "fallback",
                    n_cand, abs(result_pos - pos_t), w_lo, w_hi))

    # --- step 3: local transform from boundary anchors ---
    def _local_interpolate(orient: str) -> tuple[float, float]:
        bnd_lines = [t for t in t_lines if t[1] == orient and t[5] is not None]
        if len(bnd_lines) < 2:
            return 1.0, 0.0
        pairs = [(t[2], snapped[t[0]]) for t in bnd_lines if t[0] in snapped]
        if len(pairs) < 2:
            return 1.0, 0.0
        pairs.sort(key=lambda p: p[0])
        t_lo, s_lo = pairs[0]
        t_hi, s_hi = pairs[-1]
        dt = t_hi - t_lo
        ds = s_hi - s_lo
        if abs(dt) < _EPS:
            return 1.0, 0.0
        local_s = ds / dt
        local_d = s_lo - t_lo * local_s
        return local_s, local_d

    h_local_s, h_local_d = _local_interpolate("h")
    v_local_s, v_local_d = _local_interpolate("v")

    # --- step 4: group co-located lines, then ordered walk ---
    for orient, pool, used, extent, local_s, local_d in [
        ("h", h_det, h_used, stamp_w_pts, h_local_s, h_local_d),
        ("v", v_det, v_used, stamp_h_pts, v_local_s, v_local_d),
    ]:
        non_boundary = [t for t in t_lines if t[1] == orient and t[5] is None]
        if not non_boundary:
            continue

        # Compute interpolated positions and group co-located lines
        interp_map: dict[str, float] = {}
        for lid, _o, pos_t, *_ in non_boundary:
            interp_map[lid] = pos_t * local_s + local_d

        non_boundary.sort(key=lambda t: interp_map[t[0]])

        # Group lines within _CO_LOCATE_PTS of each other
        groups: list[list[_TLine]] = []
        for t in non_boundary:
            if groups and abs(interp_map[t[0]] - interp_map[groups[-1][0][0]]) <= _CO_LOCATE_PTS:
                groups[-1].append(t)
            else:
                groups.append([t])

        boundary_of = [t for t in t_lines if t[1] == orient and t[5] is not None]
        boundary_snapped = sorted(snapped[t[0]] for t in boundary_of if t[0] in snapped)

        if len(boundary_snapped) >= 2:
            lo_bound = boundary_snapped[0]
            hi_bound = boundary_snapped[-1]
        elif pool:
            lo_bound = min(dl.pos for dl in pool) - snap_max_pts
            hi_bound = max(dl.pos for dl in pool) + snap_max_pts
        else:
            lo_bound = 0.0
            hi_bound = 1e6

        prev_locked = lo_bound

        for group in groups:
            rep = group[0]
            rep_lid = rep[0]
            pos_interp = interp_map[rep_lid]

            # Combine spans from all group members (sum for split segments)
            sp_lo = min(t[3] for t in group)
            sp_hi = max(t[4] for t in group)
            msm = sum(t[6] for t in group)

            next_anchor = hi_bound
            for bp in boundary_snapped:
                if bp > pos_interp + _EPS:
                    next_anchor = bp
                    break

            w_lo = prev_locked
            w_hi = next_anchor

            best, n_cand = _pick_best(pool, used, pos_interp, w_lo, w_hi,
                                      sp_lo, sp_hi, extent, msm)
            if best is not None:
                for t in group:
                    snapped[t[0]] = best
                prev_locked = best
                _claim_idx(pool, used, best)
            else:
                for t in group:
                    snapped[t[0]] = pos_interp
                prev_locked = pos_interp

            if diagnostics is not None:
                result_pos = best if best is not None else pos_interp
                for t in group:
                    diagnostics.append(GridLineSnapDiag(
                        t[0], orient, pos_interp, result_pos,
                        "walk" if best is not None else "fallback",
                        n_cand, abs(result_pos - pos_interp), w_lo, w_hi))

    return snapped


# ---------------------------------------------------------------------------
#  Ordered sequence alignment for grid lines (T9.6)
# ---------------------------------------------------------------------------

@dataclass
class AlignmentLineDiag:
    """Per-line diagnostics for the ordered alignment algorithm."""
    line_id: str
    orientation: str
    pos_interp_pts: float
    pos_snapped_pts: float
    snap_method: str           # "alignment" | "boundary" | "fallback"
    matched_detected_pos: float | None = None
    matched_detected_span: float | None = None
    matched_detected_id: str = ""   # "PDF:h_5" etc.
    match_score: float = 0.0
    span_ratio: float = 0.0
    dp_action: str = ""        # "match" | "skip_detected" | ""


@dataclass
class AlignmentSummary:
    """Summary diagnostics for the full alignment pass."""
    boundary_matched: int = 0
    alignment_matched: int = 0
    template_skipped: int = 0
    detected_skipped: int = 0
    avg_match_score: float = 0.0


def _compute_border_merge_dist(
    detected: list[CellBbox],
    fraction: float = 0.20,
    fallback_pts: float = 2.0,
) -> float:
    """Adaptive border-pair merge distance from detected cell dimensions.

    Returns ``fraction`` (default 20%) of the smallest cell dimension (pts),
    or *fallback_pts* when no cells are available.  Uses the same fraction as
    ``grid_lines_utils.BORDER_MERGE_FRACTION``.
    """
    dims: list[float] = []
    for c in detected:
        w = c.x1 - c.x0
        h = c.y1 - c.y0
        if w > 1.0:
            dims.append(w)
        if h > 1.0:
            dims.append(h)
    if not dims:
        return fallback_pts
    return max(min(dims) * fraction, fallback_pts)


def _merge_border_pairs(
    lines: list[_DetectedGridLine],
    merge_dist_pts: float = 2.0,
) -> list[_DetectedGridLine]:
    """Merge detected lines that are border-pairs (top/bottom of the same table line).

    Lines within *merge_dist_pts* are merged using the median position
    (center of the physical table line) with combined span.
    """
    if not lines:
        return []
    out: list[_DetectedGridLine] = []
    i = 0
    while i < len(lines):
        j = i + 1
        while j < len(lines) and abs(lines[j].pos - lines[i].pos) <= merge_dist_pts:
            j += 1
        group = lines[i:j]
        total_cc = sum(dl.cell_count for dl in group)
        positions = sorted(dl.pos for dl in group)
        n = len(positions)
        median_pos = positions[n // 2] if n % 2 == 1 else (positions[n // 2 - 1] + positions[n // 2]) * 0.5
        span_lo = min(dl.span_lo for dl in group)
        span_hi = max(dl.span_hi for dl in group)
        out.append(_DetectedGridLine(median_pos, span_lo, span_hi, total_cc))
        i = j
    return out


def _filter_outside_stamp(
    lines: list[_DetectedGridLine],
    lo: float,
    hi: float,
    margin_pts: float = 3.0,
) -> list[_DetectedGridLine]:
    """Remove detected lines whose position is outside [lo-margin, hi+margin]."""
    return [dl for dl in lines if (lo - margin_pts) <= dl.pos <= (hi + margin_pts)]


def _assign_detected_ids(
    h_lines: list[_DetectedGridLine],
    v_lines: list[_DetectedGridLine],
) -> tuple[dict[int, str], dict[int, str]]:
    """Assign PDF:h_N / PDF:v_N IDs to detected lines (1-based by position order)."""
    h_ids = {id(dl): f"PDF:h_{i+1}" for i, dl in enumerate(h_lines)}
    v_ids = {id(dl): f"PDF:v_{i+1}" for i, dl in enumerate(v_lines)}
    return h_ids, v_ids


# ---------------------------------------------------------------------------
# Progressive (piecewise-linear) interpolation helpers
# ---------------------------------------------------------------------------

def _progressive_interp(pos_t: float, anchors: list[tuple[float, float]]) -> float:
    """Piecewise-linear interpolation between the nearest bracketing anchors.

    *anchors* is sorted by ``template_pos`` with monotonic ``snapped_pos``.
    Finds the two anchors that bracket *pos_t* and linearly interpolates.
    If *pos_t* is outside all anchors, extrapolates from the nearest pair.
    With < 2 anchors falls back to identity (returns *pos_t*).
    """
    n = len(anchors)
    if n == 0:
        return pos_t
    if n == 1:
        return pos_t + (anchors[0][1] - anchors[0][0])

    # bisect: idx = first anchor with template_pos > pos_t
    tpl_vals = [a[0] for a in anchors]
    idx = bisect.bisect_right(tpl_vals, pos_t)

    if idx == 0:
        lo, hi = anchors[0], anchors[1]
    elif idx >= n:
        lo, hi = anchors[-2], anchors[-1]
    else:
        lo, hi = anchors[idx - 1], anchors[idx]

    dt = hi[0] - lo[0]
    if abs(dt) < _EPS:
        return lo[1]
    ratio = (pos_t - lo[0]) / dt
    return lo[1] + ratio * (hi[1] - lo[1])


def _insert_anchor(
    anchors: list[tuple[float, float]],
    new: tuple[float, float],
) -> bool:
    """Insert anchor maintaining sort by template_pos.

    Returns ``True`` if inserted, ``False`` if skipped (would break
    snapped_pos monotonicity — safety net, should not happen with walk order).
    """
    t_new, s_new = new
    tpl_vals = [a[0] for a in anchors]
    idx = bisect.bisect_left(tpl_vals, t_new)
    if idx > 0 and anchors[idx - 1][1] >= s_new:
        return False
    if idx < len(anchors) and anchors[idx][1] <= s_new:
        return False
    anchors.insert(idx, new)
    return True


def _anchor_bracket_ids(
    pos_t: float,
    anchors: list[tuple[float, float]],
    t_lines: list,
    orient: str,
) -> tuple[str, str]:
    """Return human-readable IDs of the two anchors bracketing *pos_t*.

    Used for alignment_log diagnostics.  Falls back to ``"?"`` if not found.
    """
    n = len(anchors)
    if n < 2:
        return "?", "?"

    tpl_vals = [a[0] for a in anchors]
    idx = bisect.bisect_right(tpl_vals, pos_t)

    if idx == 0:
        lo_t, hi_t = anchors[0][0], anchors[1][0]
    elif idx >= n:
        lo_t, hi_t = anchors[-2][0], anchors[-1][0]
    else:
        lo_t, hi_t = anchors[idx - 1][0], anchors[idx][0]

    def _find_id(target_pos: float) -> str:
        for t in t_lines:
            if t[1] == orient and abs(t[2] - target_pos) < _EPS:
                return t[0]
        return f"@{target_pos:.1f}"

    return _find_id(lo_t), _find_id(hi_t)


def _align_grid_lines_ordered(
    template: StampTemplate,
    frame: FrameInfo,
    sx: float,
    sy: float,
    dx: float,
    dy: float,
    detected: list[CellBbox],
    diagnostics: list[AlignmentLineDiag] | None = None,
    summary_out: list[AlignmentSummary] | None = None,
    detected_merged_out: list[DetectedGridLineInfo] | None = None,
    alignment_report_out: list[AlignmentReport] | None = None,
    template_path: str = "",
    pdf_path: str = "",
    page_num: int = 0,
    stamp_bbox_override: fitz.Rect | None = None,
) -> dict[str, float]:
    """Sequential ordered alignment of template grid lines to detected lines.

    Template is the source of truth for line count.  Starts from boundary
    lines, then walks interior lines sequentially, matching each to the best
    detected line within a search window.  Extra detected lines between
    matches are classified as split-artifacts and skipped.

    Returns ``{line_id: snapped_pos_fitz_pts}`` for all template grid lines.
    """
    _CO_LOCATE_PTS = 2.0 * SCALE
    MATCH_THRESHOLD = 0.35

    if not template.grid_lines:
        return {}

    h_det_raw, v_det_raw = _build_detected_grid_lines(
        detected, template.grid_tolerance_detected_mm
    ) if detected else ([], [])
    h_raw_count = len(h_det_raw)
    v_raw_count = len(v_det_raw)

    merged_spans_mm = _virtual_merge_template_lines(template.grid_lines)

    # --- convert template lines to fitz, apply transform ---
    _TLine = tuple[str, str, float, float, float, str | None, float]
    t_lines: list[_TLine] = []
    for gl in template.grid_lines:
        orientation, pos_fitz, start_fitz, end_fitz, bnd, line_id = grid_line_to_fitz_pts(
            gl, frame, template.origin
        )
        if orientation == "h":
            pos_t = pos_fitz * sy + dy
        else:
            pos_t = pos_fitz * sx + dx
        msm = merged_spans_mm.get(line_id, abs(gl.end_mm - gl.start_mm))
        t_lines.append((line_id, orientation, pos_t, start_fitz, end_fitz, bnd, msm))

    stamp_bbox = (
        fitz.Rect(stamp_bbox_override)
        if stamp_bbox_override is not None
        else _template_stamp_bbox(template, frame)
    )
    if stamp_bbox is not None:
        stamp_w_pts = max(1.0, stamp_bbox.width)
        stamp_h_pts = max(1.0, stamp_bbox.height)
    else:
        stamp_w_pts = max(1.0, abs(frame.x1 - frame.x0))
        stamp_h_pts = max(1.0, abs(frame.y1 - frame.y0))

    snapped: dict[str, float] = {}
    snap_max_pts = template.snap_max_distance_mm * SCALE

    # --- merge border pairs ---
    border_merge_pts = _compute_border_merge_dist(detected)
    h_det_merged = _merge_border_pairs(h_det_raw, border_merge_pts)
    v_det_merged = _merge_border_pairs(v_det_raw, border_merge_pts)
    h_merged_count = len(h_det_merged)
    v_merged_count = len(v_det_merged)

    # --- filter outside stamp bbox ---
    if stamp_bbox is not None:
        margin = snap_max_pts
        h_det_merged = _filter_outside_stamp(h_det_merged, stamp_bbox.y0, stamp_bbox.y1, margin)
        v_det_merged = _filter_outside_stamp(v_det_merged, stamp_bbox.x0, stamp_bbox.x1, margin)

    h_filtered_count = len(h_det_merged)
    v_filtered_count = len(v_det_merged)

    # --- assign PDF IDs ---
    h_id_map, v_id_map = _assign_detected_ids(h_det_merged, v_det_merged)

    if detected_merged_out is not None:
        for dl in h_det_merged:
            detected_merged_out.append(DetectedGridLineInfo(
                h_id_map[id(dl)], "h", dl.pos, dl.span_lo, dl.span_hi, dl.cell_count,
            ))
        for dl in v_det_merged:
            detected_merged_out.append(DetectedGridLineInfo(
                v_id_map[id(dl)], "v", dl.pos, dl.span_lo, dl.span_hi, dl.cell_count,
            ))

    # --- alignment log builder ---
    log_lines: list[str] = []
    log_lines.append("=== GRID ALIGNMENT LOG ===")
    log_lines.append(f"Template: {template_path or template.name} ({len(template.grid_lines)} grid lines)")
    import os
    log_lines.append(f"PDF: {os.path.basename(pdf_path) if pdf_path else '(unknown)'} (page {page_num})")
    log_lines.append(f"Frame: x0={frame.x0:.1f} y0={frame.y0:.1f} x1={frame.x1:.1f} y1={frame.y1:.1f}")
    log_lines.append(f"Scale: sx={sx:.4f} sy={sy:.4f} dx={dx:.1f} dy={dy:.1f}")
    log_lines.append(
        f"Params: MATCH_THRESHOLD={MATCH_THRESHOLD} "
        f"BORDER_MERGE={border_merge_pts:.1f}pts "
        f"CO_LOCATE={_CO_LOCATE_PTS:.1f}pts"
    )
    log_lines.append(f"Log format: .cursor/rules/AI_v2_grid_log.mdc")
    log_lines.append(f"Engine rules: .cursor/rules/AI_v2_grid.mdc")
    log_lines.append(f"Code: pdf_parsing_v2/grid_matcher.py :: _align_grid_lines_ordered()")
    log_lines.append("")

    boundary_diags: list[AlignmentLineDiag] = []
    summary = AlignmentSummary()
    report_problems: list[str] = []
    all_match_scores: list[float] = []
    line_method: dict[str, str] = {}
    line_matched_to: dict[str, str] = {}

    h_tpl_count = sum(1 for t in t_lines if t[1] == "h")
    v_tpl_count = sum(1 for t in t_lines if t[1] == "v")

    # --- snap boundary lines first ---
    for orient, pool, extent, id_map in [
        ("h", h_det_merged, stamp_w_pts, h_id_map),
        ("v", v_det_merged, stamp_h_pts, v_id_map),
    ]:
        boundaries = [t for t in t_lines if t[1] == orient and t[5] is not None]
        for lid, _o, pos_t, sp_lo, sp_hi, _bnd, msm in boundaries:
            if not pool:
                snapped[lid] = pos_t
                line_method[lid] = "boundary-fallback"
                boundary_diags.append(AlignmentLineDiag(
                    lid, orient, pos_t, pos_t, "fallback",
                ))
                report_problems.append(f"\u26a0 TPL:{lid} boundary -> no detected pool, fallback to transform")
                log_lines.append(f"  TPL:{lid} -> NO POOL, fallback={pos_t:.1f} \u26a0")
                continue
            best_dl = None
            best_score = -1.0
            top_pref_dl = None
            top_pref_score = -1.0
            full_span_threshold = extent * 0.3
            for dl in pool:
                dist = abs(dl.pos - pos_t)
                if dist > snap_max_pts * 2.5:
                    continue
                dist_score = 1.0 - min(dist / max(snap_max_pts, 1.0), 1.0)
                span_score = min(dl.span_length / max(extent * 0.3, 1.0), 1.0)
                sc = 0.6 * dist_score + 0.4 * span_score
                if sc > best_score:
                    best_score = sc
                    best_dl = dl
                if (
                    orient == "h"
                    and lid == "h_top"
                    and dl.pos <= pos_t
                    and dist <= snap_max_pts
                    and dl.span_length >= full_span_threshold
                    and (
                        top_pref_dl is None
                        or dl.pos < top_pref_dl.pos
                    )
                ):
                    top_pref_dl = dl
                    top_pref_score = sc
            selected_method = "boundary"
            if (
                orient == "h"
                and lid == "h_top"
                and top_pref_dl is not None
                and best_dl is not None
                and best_dl is not top_pref_dl
                and best_dl.pos > pos_t
                and top_pref_score >= max(MATCH_THRESHOLD, best_score - 0.2)
            ):
                best_dl = top_pref_dl
                best_score = top_pref_score
                selected_method = "boundary-top-pref"
            if best_dl is not None:
                snapped[lid] = best_dl.pos
                summary.boundary_matched += 1
                all_match_scores.append(best_score)
                det_id = id_map.get(id(best_dl), "?")
                line_method[lid] = selected_method
                line_matched_to[lid] = det_id
                boundary_diags.append(AlignmentLineDiag(
                    lid, orient, pos_t, best_dl.pos, selected_method,
                    matched_detected_pos=best_dl.pos,
                    matched_detected_span=best_dl.span_length,
                    matched_detected_id=det_id,
                    match_score=best_score,
                ))
            else:
                snapped[lid] = pos_t
                line_method[lid] = "boundary-fallback"
                boundary_diags.append(AlignmentLineDiag(
                    lid, orient, pos_t, pos_t, "fallback",
                ))
                report_problems.append(f"\u26a0 TPL:{lid} boundary -> no candidate within range, fallback")

    # --- progressive interpolation helpers ---

    def _boundary_affine(orient: str) -> tuple[float, float]:
        """Initial affine from boundary anchors (seed for progressive interp)."""
        bnd_lines = [t for t in t_lines if t[1] == orient and t[5] is not None]
        if len(bnd_lines) < 2:
            return 1.0, 0.0
        pairs = [(t[2], snapped[t[0]]) for t in bnd_lines if t[0] in snapped]
        if len(pairs) < 2:
            return 1.0, 0.0
        pairs.sort(key=lambda p: p[0])
        t_lo, s_lo = pairs[0]
        t_hi, s_hi = pairs[-1]
        dt = t_hi - t_lo
        ds = s_hi - s_lo
        if abs(dt) < _EPS:
            return 1.0, 0.0
        local_s = ds / dt
        local_d = s_lo - t_lo * local_s
        return local_s, local_d

    # --- sequential walk for non-boundary lines ---
    h_skip_artifact = 0
    v_skip_artifact = 0

    for orient, pool_raw, extent, id_map in [
        ("h", h_det_merged, stamp_w_pts, h_id_map),
        ("v", v_det_merged, stamp_h_pts, v_id_map),
    ]:
        non_boundary = [t for t in t_lines if t[1] == orient and t[5] is None]
        if not non_boundary:
            continue

        # Seed: boundary-only affine for initial sort order
        bnd_s, bnd_d = _boundary_affine(orient)

        seed_interp: dict[str, float] = {}
        for lid, _o, pos_t, *_ in non_boundary:
            seed_interp[lid] = pos_t * bnd_s + bnd_d

        non_boundary.sort(key=lambda t: seed_interp[t[0]])

        # Group co-located template lines (border pairs in template)
        groups: list[list[_TLine]] = []
        for t in non_boundary:
            if groups and abs(seed_interp[t[0]] - seed_interp[groups[-1][0][0]]) <= _CO_LOCATE_PTS:
                groups[-1].append(t)
            else:
                groups.append([t])

        # Build progressive anchor list from boundary matches
        # anchors: sorted by template_pos, snapped_pos guaranteed monotonic
        anchors: list[tuple[float, float]] = sorted(
            [
                (t[2], snapped[t[0]])
                for t in t_lines
                if t[1] == orient and t[5] is not None and t[0] in snapped
            ],
            key=lambda p: p[0],
        )

        # Remove boundary-claimed detected lines from pool
        boundary_positions = set()
        for t in t_lines:
            if t[1] == orient and t[5] is not None and t[0] in snapped:
                boundary_positions.add(round(snapped[t[0]], 1))

        pool = [
            dl for dl in pool_raw
            if not any(abs(dl.pos - bp) < 1.0 for bp in boundary_positions)
        ]
        pool.sort(key=lambda dl: dl.pos)

        full_span_threshold = extent * 0.3

        walk_label = "H" if orient == "h" else "V"
        log_lines.append(f"WALK {walk_label} ({len(groups)} TPL groups, {len(pool)} det pool):")

        # Boundary lines in log
        for bd in boundary_diags:
            if bd.orientation == orient:
                if bd.matched_detected_id:
                    log_lines.append(
                        f"  TPL:{bd.line_id} -> {bd.matched_detected_id} "
                        f"({bd.pos_snapped_pts:.1f}) score={bd.match_score:.2f} {bd.snap_method} \u2713"
                    )
                else:
                    log_lines.append(
                        f"  TPL:{bd.line_id} -> fallback ({bd.pos_snapped_pts:.1f}) boundary \u26a0"
                    )

        # Sequential walk: for each template group, find best detected in window
        det_ptr = 0  # next unused detected line index
        prev_matched_pos = -1e18
        recent_deltas: list[float] = []

        for gi, group in enumerate(groups):
            rep = group[0]
            pos_t = rep[2]  # template pos (fitz pts, after global transform)
            msm = sum(t[6] for t in group)

            # Progressive interpolation from nearest bracketing anchors
            pos_interp = _progressive_interp(pos_t, anchors)
            pos_interp = max(pos_interp, prev_matched_pos + 0.5)

            # Anchor IDs for log
            anchor_lo_id, anchor_hi_id = _anchor_bracket_ids(pos_t, anchors, t_lines, orient)

            # Search window: adaptive margin based on recent deltas
            adaptive_margin = max(
                snap_max_pts,
                1.5 * max(recent_deltas[-5:], default=0.0),
            )
            if gi + 1 < len(groups):
                next_pos_t = groups[gi + 1][0][2]
                next_interp = _progressive_interp(next_pos_t, anchors)
                next_interp = max(next_interp, pos_interp + 0.5)
                window_hi = (pos_interp + next_interp) * 0.5 + adaptive_margin
            else:
                window_hi = pos_interp + adaptive_margin * 3
            window_lo = prev_matched_pos + 0.5  # just past previous match

            # Find best candidate in window
            best_dl = None
            best_score = -1.0
            best_di = -1
            skipped_in_window: list[tuple[_DetectedGridLine, int]] = []

            for di in range(det_ptr, len(pool)):
                dl = pool[di]
                if dl.pos < window_lo:
                    continue
                if dl.pos > window_hi:
                    break

                dist = abs(dl.pos - pos_interp)
                pos_score = max(0.0, 1.0 - dist / max(snap_max_pts * 2, 1.0))

                span_t = msm * SCALE
                span_d = dl.span_length
                if span_t < _EPS or span_d < _EPS:
                    sc = 0.3 * pos_score
                else:
                    span_ratio = min(span_t, span_d) / max(span_t, span_d)
                    both_full = (span_t > full_span_threshold) and (span_d > full_span_threshold)
                    category_bonus = 0.2 if both_full else 0.0
                    sc = 0.4 * span_ratio + 0.4 * pos_score + category_bonus

                if sc > best_score:
                    if best_dl is not None:
                        skipped_in_window.append((best_dl, best_di))
                    best_score = sc
                    best_dl = dl
                    best_di = di
                else:
                    skipped_in_window.append((dl, di))

            group_ids_str = ", ".join(f"TPL:{t[0]}" for t in group)

            if best_dl is not None and best_score >= MATCH_THRESHOLD:
                for t in group:
                    snapped[t[0]] = best_dl.pos
                    line_method[t[0]] = "match"
                summary.alignment_matched += 1
                all_match_scores.append(best_score)
                prev_matched_pos = best_dl.pos
                det_ptr = best_di + 1

                delta = best_dl.pos - pos_interp
                recent_deltas.append(abs(delta))

                # Insert new anchor (only real matches)
                _insert_anchor(anchors, (pos_t, best_dl.pos))

                det_id = id_map.get(id(best_dl), "?")
                for t in group:
                    line_matched_to[t[0]] = det_id
                span_t = msm * SCALE
                span_d = best_dl.span_length
                sr = (min(span_t, span_d) / max(span_t, span_d)) if max(span_t, span_d) > _EPS else 0.0

                log_lines.append(
                    f"  {group_ids_str} -> {det_id} ({best_dl.pos:.1f}) "
                    f"score={best_score:.2f} interp={pos_interp:.1f} delta={delta:+.1f} "
                    f"anchors=({anchor_lo_id},{anchor_hi_id}) \u2713"
                )

                if diagnostics is not None:
                    for t in group:
                        diagnostics.append(AlignmentLineDiag(
                            t[0], orient, pos_interp, best_dl.pos, "alignment",
                            matched_detected_pos=best_dl.pos,
                            matched_detected_span=best_dl.span_length,
                            matched_detected_id=det_id,
                            match_score=best_score,
                            span_ratio=sr,
                            dp_action="match",
                        ))

                # Log skipped detected lines as split-artifacts
                for sk_dl, sk_di in skipped_in_window:
                    sk_id = id_map.get(id(sk_dl), "?")
                    summary.detected_skipped += 1
                    if orient == "h":
                        h_skip_artifact += 1
                    else:
                        v_skip_artifact += 1
                    reason = f"split-artifact, span={sk_dl.span_length:.0f}pts"
                    log_lines.append(f"    [skip {sk_id} ({sk_dl.pos:.1f}) {reason}]")
                    report_problems.append(f"\u2298 {sk_id} ({sk_dl.pos:.1f}) skip: {reason}")
            else:
                # No match — use interpolated position
                for t in group:
                    snapped[t[0]] = pos_interp
                    line_method[t[0]] = "interp"
                summary.template_skipped += 1

                nearest_info = ""
                if best_dl is not None:
                    n_id = id_map.get(id(best_dl), "?")
                    nearest_info = f" (nearest {n_id} dist={abs(best_dl.pos - pos_interp):.1f}pts, score={best_score:.2f} < {MATCH_THRESHOLD})"
                log_lines.append(
                    f"  {group_ids_str} -> NO MATCH, interp={pos_interp:.1f}{nearest_info} "
                    f"anchors=({anchor_lo_id},{anchor_hi_id}) \u26a0"
                )
                report_problems.append(
                    f"\u26a0 {group_ids_str} no_match -> interpolated {pos_interp:.1f}{nearest_info}"
                )

                if diagnostics is not None:
                    for t in group:
                        diagnostics.append(AlignmentLineDiag(
                            t[0], orient, pos_interp, pos_interp, "fallback",
                            dp_action="",
                        ))

        log_lines.append("")

    if diagnostics is not None:
        diagnostics.extend(boundary_diags)

    if summary_out is not None:
        summary_out.append(summary)

    # --- build alignment report ---
    total_matched = summary.boundary_matched + summary.alignment_matched
    total_tpl = len(template.grid_lines)
    avg_score = (sum(all_match_scores) / len(all_match_scores)) if all_match_scores else 0.0
    summary.avg_match_score = avg_score

    warn_threshold = max(5, int(total_tpl * 0.15))
    if summary.template_skipped == 0 and avg_score >= 0.5:
        verdict = "ok"
    elif summary.template_skipped <= warn_threshold or avg_score < 0.5:
        verdict = "warning"
    else:
        verdict = "fail"

    message_parts = [f"verdict={verdict}"]
    if summary.template_skipped > 0:
        message_parts.append(f"{summary.template_skipped} TPL lines without match")
    if summary.detected_skipped > 0:
        message_parts.append(f"{summary.detected_skipped} PDF lines skipped (artifacts)")
    message_parts.append(f"avg_score={avg_score:.2f}")
    message = "; ".join(message_parts)

    # Insert counts and verdict at top of log
    counts_lines = [
        f"VERDICT: {verdict.upper()} ({len(report_problems)} problems)" if report_problems
        else f"VERDICT: {verdict.upper()}",
        "",
        "COUNTS:",
        f"  H: {h_tpl_count} TPL | {h_raw_count} det(raw) -> {h_merged_count} det(merged)"
        f" -> {h_filtered_count} det(in-stamp) | {h_skip_artifact} skip-artifact",
        f"  V: {v_tpl_count} TPL | {v_raw_count} det(raw) -> {v_merged_count} det(merged)"
        f" -> {v_filtered_count} det(in-stamp) | {v_skip_artifact} skip-artifact",
        "",
    ]
    log_lines[8:8] = counts_lines

    if report_problems:
        log_lines.append("PROBLEMS:")
        for p in report_problems:
            log_lines.append(f"  {p}")

    log_lines.append("")
    log_lines.append("SNAPPED POSITIONS (final):")
    log_lines.append(f"  {'Line ID':<20} {'Pos pts':>8} {'Method':<18} {'Matched to':<12}")
    log_lines.append(f"  {'-'*20} {'-'*8} {'-'*18} {'-'*12}")
    for lid in sorted(snapped.keys(), key=lambda k: (0 if any(
        t[1] == "h" for t in t_lines if t[0] == k
    ) else 1, snapped[k])):
        pos = snapped[lid]
        method = line_method.get(lid, "?")
        matched = line_matched_to.get(lid, "")
        log_lines.append(f"  {lid:<20} {pos:>8.1f} {method:<18} {matched:<12}")

    full_log = "\n".join(log_lines)

    report = AlignmentReport(
        verdict=verdict,
        total_tpl_lines=total_tpl,
        matched_count=total_matched,
        no_match_count=summary.template_skipped,
        skipped_detected_count=summary.detected_skipped,
        avg_match_score=avg_score,
        problem_lines=report_problems,
        message=message,
        log=full_log,
    )

    if alignment_report_out is not None:
        alignment_report_out.append(report)

    return snapped


def _refine_grid_from_cells(
    snapped: dict[str, float],
    results: list[tuple[str, AdaptResult]],
    fields: list[FieldDef],
    keys: list[str],
    normal_gap_pts: float = 1.5,
    max_delta_pts: float = 2.0,
) -> tuple[dict[str, float], dict[str, tuple[float, float]]]:
    """Refine grid line positions using cell-edge votes from cell-assignment.

    For each grid line, collects edge positions from detected cells matched
    to bound fields, takes the median, and updates the snapped position
    only if the delta is within *max_delta_pts* (~0.7mm).  This prevents
    the refinement from jumping to a different detected line (border-pair
    partner or split-artifact).

    Returns ``(refined_snapped, refine_deltas)`` where *refine_deltas* is
    ``{line_id: (old_pos, new_pos)}`` for lines that were actually changed.
    """
    ng = normal_gap_pts / 2.0
    key_to_field: dict[str, FieldDef] = {keys[i]: f for i, f in enumerate(fields)}
    votes: dict[str, list[float]] = {}

    for key, ar in results:
        if ar.status == "excluded" or not ar.matched_detected:
            continue
        fd = key_to_field.get(key)
        if fd is None:
            continue
        cell = ar.matched_detected[0]

        if fd.bound_top and fd.bound_top in snapped:
            votes.setdefault(fd.bound_top, []).append(cell.y0 - ng)
        if fd.bound_bottom and fd.bound_bottom in snapped:
            votes.setdefault(fd.bound_bottom, []).append(cell.y1 + ng)
        if fd.bound_left and fd.bound_left in snapped:
            votes.setdefault(fd.bound_left, []).append(cell.x0 - ng)
        if fd.bound_right and fd.bound_right in snapped:
            votes.setdefault(fd.bound_right, []).append(cell.x1 + ng)

    refined = dict(snapped)
    refine_deltas: dict[str, tuple[float, float]] = {}
    for line_id, line_votes in votes.items():
        old = snapped[line_id]
        if len(line_votes) >= 2:
            candidate = _median(line_votes)
        elif len(line_votes) == 1:
            candidate = line_votes[0]
        else:
            continue
        if abs(candidate - old) <= max_delta_pts:
            refined[line_id] = candidate
            if abs(candidate - old) > 0.01:
                refine_deltas[line_id] = (old, candidate)

    return refined, refine_deltas


def expand_fields_by_bindings(
    results: list[tuple[str, AdaptResult]],
    fields: list[FieldDef],
    keys: list[str],
    snapped_line_positions: dict[str, float],
    normal_gap_pts: float = 1.5,
) -> list[str]:
    """Set field bbox edges from bound grid-line positions (dict-lookup, no heuristics).

    For each non-excluded field with bound_* set, the corresponding bbox edge
    is placed at ``snapped_line_positions[line_id] ± normal_gap_pts/2``.

    Modifies *results* in-place.  Returns list of keys that were expanded.
    """
    ng = normal_gap_pts
    result_map: dict[str, AdaptResult] = {key: ar for key, ar in results}
    key_to_field: dict[str, FieldDef] = {keys[i]: f for i, f in enumerate(fields)}
    expanded: list[str] = []

    for key, ar in results:
        if ar.status == "excluded":
            continue
        fd = key_to_field.get(key)
        if fd is None:
            continue
        if fd.outside_stamp or fd.document_property:
            continue
        has_any = (fd.bound_top or fd.bound_bottom
                   or fd.bound_left or fd.bound_right)
        if not has_any:
            continue

        bb = ar.bbox
        new_x0, new_y0, new_x1, new_y1 = bb.x0, bb.y0, bb.x1, bb.y1
        changed = False

        if fd.bound_top and fd.bound_top in snapped_line_positions:
            new_y0 = snapped_line_positions[fd.bound_top] + ng / 2
            changed = True
        if fd.bound_bottom and fd.bound_bottom in snapped_line_positions:
            new_y1 = snapped_line_positions[fd.bound_bottom] - ng / 2
            changed = True
        if fd.bound_left and fd.bound_left in snapped_line_positions:
            new_x0 = snapped_line_positions[fd.bound_left] + ng / 2
            changed = True
        if fd.bound_right and fd.bound_right in snapped_line_positions:
            new_x1 = snapped_line_positions[fd.bound_right] - ng / 2
            changed = True

        if changed and new_x1 > new_x0 and new_y1 > new_y0:
            ar_new = result_map[key]
            ar_new.bbox = CellBbox(new_x0, new_y0, new_x1, new_y1)
            ar_new.status = "grid_bound"
            expanded.append(key)

    return expanded


def _validate_grid_against_assignment(
    results: list[tuple[str, AdaptResult]],
    fields: list[FieldDef],
    keys: list[str],
    edge_error_threshold_pts: float = 5.0,
) -> list[str]:
    """Compare grid-bound field edges with cell-assignment edges; correct outliers.

    For each ``grid_bound`` field that also has a ``matched_detected`` cell from
    the cell-assignment stage, compare the grid-derived edge with the cell edge.
    If the error on any bound edge exceeds *edge_error_threshold_pts*, that edge
    is corrected to the detected-cell edge.

    Modifies *results* in-place.  Returns list of keys that were corrected.
    """
    key_to_field: dict[str, FieldDef] = {keys[i]: f for i, f in enumerate(fields)}
    corrected: list[str] = []

    for key, ar in results:
        if ar.status != "grid_bound":
            continue
        if not ar.matched_detected:
            continue
        fd = key_to_field.get(key)
        if fd is None:
            continue

        cell = ar.matched_detected[0]
        gb = ar.bbox
        changed = False

        new_x0, new_y0, new_x1, new_y1 = gb.x0, gb.y0, gb.x1, gb.y1

        if fd.bound_top and abs(gb.y0 - cell.y0) > edge_error_threshold_pts:
            new_y0 = cell.y0
            changed = True
        if fd.bound_bottom and abs(gb.y1 - cell.y1) > edge_error_threshold_pts:
            new_y1 = cell.y1
            changed = True
        if fd.bound_left and abs(gb.x0 - cell.x0) > edge_error_threshold_pts:
            new_x0 = cell.x0
            changed = True
        if fd.bound_right and abs(gb.x1 - cell.x1) > edge_error_threshold_pts:
            new_x1 = cell.x1
            changed = True

        if changed and new_x1 > new_x0 and new_y1 > new_y0:
            ar.bbox = CellBbox(new_x0, new_y0, new_x1, new_y1)
            corrected.append(key)

    return corrected


# ---------------------------------------------------------------------------
#  Post-processing: close gaps left by fragmented cells
# ---------------------------------------------------------------------------

def _v_overlap_frac(a: CellBbox, b: CellBbox) -> float:
    """Vertical overlap as fraction of the smaller height."""
    overlap = max(0.0, min(a.y1, b.y1) - max(a.y0, b.y0))
    min_h = min(a.height, b.height)
    return overlap / max(min_h, _EPS)


def _h_overlap_frac(a: CellBbox, b: CellBbox) -> float:
    """Horizontal overlap as fraction of the smaller width."""
    overlap = max(0.0, min(a.x1, b.x1) - max(a.x0, b.x0))
    min_w = min(a.width, b.width)
    return overlap / max(min_w, _EPS)


def _median(values: list[float]) -> float:
    if not values:
        return 0.0
    s = sorted(values)
    n = len(s)
    if n % 2 == 1:
        return s[n // 2]
    return (s[n // 2 - 1] + s[n // 2]) * 0.5


def close_fragmented_gaps(
    results: list[tuple[str, AdaptResult]],
    template_bboxes: dict[str, CellBbox],
    stamp_bbox: CellBbox | None = None,
    normal_gap_pts: float | None = None,
    *,
    shape_score_threshold: float = 0.5,
    size_ratio_factor: float = 0.6,
    gap_factor: float = 3.0,
    min_gap_threshold_pts: float = 3.0,
    overlap_frac_min: float = 0.5,
    max_small_gap_pts: float = 5.0,
) -> list[str]:
    """Extend fields matched to cell fragments so they abut correctly-matched neighbors.

    Modifies *results* in-place (bbox, status).
    Returns list of keys that were extended.

    Parameters
    ----------
    results : adapted field list from ``adapt_by_cell_assignment``
    template_bboxes : {key: CellBbox} of template field positions (fitz pts)
    stamp_bbox : overall stamp bbox (for boundary extension)
    normal_gap_pts : override for the "line-thickness" gap; auto-computed if None
    shape_score_threshold : fields below this are candidate victims
    size_ratio_factor : fields with width/height ratio < factor*median are candidates
    gap_factor : gap > factor * normal_gap is abnormal
    min_gap_threshold_pts : absolute minimum for gap anomaly threshold
    overlap_frac_min : perpendicular overlap fraction for adjacency
    max_small_gap_pts : gaps above this are excluded from normal-gap computation
    """
    matched = [
        (key, ar) for key, ar in results
        if ar.status in ("matched", "anchor")
    ]
    if len(matched) < 3:
        return []

    # ------ Phase 1: baselines ------
    width_ratios: list[float] = []
    height_ratios: list[float] = []
    for key, ar in matched:
        tb = template_bboxes.get(key)
        if tb is None:
            continue
        wr = ar.bbox.width / max(tb.width, _EPS)
        hr = ar.bbox.height / max(tb.height, _EPS)
        width_ratios.append(wr)
        height_ratios.append(hr)

    median_wr = _median(width_ratios) if width_ratios else 1.0
    median_hr = _median(height_ratios) if height_ratios else 1.0

    # Compute normal gap from all small inter-field gaps (x-axis).
    small_x_gaps: list[float] = []
    small_y_gaps: list[float] = []
    matched_bboxes = [(key, ar.bbox) for key, ar in matched]

    for i, (k_a, bb_a) in enumerate(matched_bboxes):
        for j, (k_b, bb_b) in enumerate(matched_bboxes):
            if i >= j:
                continue
            # Horizontal adjacency: a is left of b.
            if _v_overlap_frac(bb_a, bb_b) >= overlap_frac_min:
                if bb_a.x1 < bb_b.x0:
                    g = bb_b.x0 - bb_a.x1
                    if 0 < g < max_small_gap_pts:
                        small_x_gaps.append(g)
                elif bb_b.x1 < bb_a.x0:
                    g = bb_a.x0 - bb_b.x1
                    if 0 < g < max_small_gap_pts:
                        small_x_gaps.append(g)
            # Vertical adjacency: a is above b.
            if _h_overlap_frac(bb_a, bb_b) >= overlap_frac_min:
                if bb_a.y1 < bb_b.y0:
                    g = bb_b.y0 - bb_a.y1
                    if 0 < g < max_small_gap_pts:
                        small_y_gaps.append(g)
                elif bb_b.y1 < bb_a.y0:
                    g = bb_a.y0 - bb_b.y1
                    if 0 < g < max_small_gap_pts:
                        small_y_gaps.append(g)

    if normal_gap_pts is not None:
        ng_x = ng_y = normal_gap_pts
    else:
        ng_x = _median(small_x_gaps) if small_x_gaps else 1.5
        ng_y = _median(small_y_gaps) if small_y_gaps else 1.5

    gap_thresh_x = max(gap_factor * ng_x, min_gap_threshold_pts)
    gap_thresh_y = max(gap_factor * ng_y, min_gap_threshold_pts)

    # ------ Phase 2: identify victims ------
    # Adaptive shape threshold: if ALL fields have similarly low shape_score
    # (e.g. template bboxes are uniformly larger than detected cells), the
    # absolute threshold would incorrectly flag everyone as "fragmented".
    # Use median shape_score to distinguish genuine fragments (outliers) from
    # globally inflated templates.
    all_shape_scores = [ar.shape_score for _, ar in matched]
    median_shape = _median(all_shape_scores) if all_shape_scores else 1.0
    effective_shape_threshold = max(
        0.3,
        min(shape_score_threshold, median_shape * 0.7),
    )

    victim_keys: set[str] = set()
    for key, ar in matched:
        tb = template_bboxes.get(key)
        if tb is None:
            continue
        wr = ar.bbox.width / max(tb.width, _EPS)
        hr = ar.bbox.height / max(tb.height, _EPS)

        is_victim = False
        if ar.shape_score < effective_shape_threshold:
            is_victim = True
        if wr < size_ratio_factor * median_wr:
            is_victim = True
        if hr < size_ratio_factor * median_hr:
            is_victim = True
        if is_victim:
            victim_keys.add(key)

    if not victim_keys:
        return []

    # ------ Phase 3: extend boundaries ------
    result_map: dict[str, AdaptResult] = {key: ar for key, ar in results}
    extended_keys: list[str] = []

    for vkey in victim_keys:
        ar = result_map[vkey]
        vbox = ar.bbox
        changed = False
        new_x0, new_y0, new_x1, new_y1 = vbox.x0, vbox.y0, vbox.x1, vbox.y1

        # Find nearest neighbor on each side among non-victim matched fields.
        best_left: CellBbox | None = None  # neighbor whose right edge is left of victim
        best_left_x1 = -1e18
        best_right: CellBbox | None = None  # neighbor whose left edge is right of victim
        best_right_x0 = 1e18
        best_top: CellBbox | None = None  # neighbor whose bottom edge is above victim
        best_top_y1 = -1e18
        best_bottom: CellBbox | None = None  # neighbor whose top edge is below victim
        best_bottom_y0 = 1e18

        for nkey, nar in matched:
            if nkey == vkey:
                continue
            if nkey in victim_keys:
                continue
            nb = nar.bbox

            # Left neighbor: nb is to the left, sufficient vertical overlap.
            if nb.x1 <= vbox.x0 + ng_x and _v_overlap_frac(vbox, nb) >= overlap_frac_min:
                if nb.x1 > best_left_x1:
                    best_left_x1 = nb.x1
                    best_left = nb

            # Right neighbor: nb is to the right.
            if nb.x0 >= vbox.x1 - ng_x and _v_overlap_frac(vbox, nb) >= overlap_frac_min:
                if nb.x0 < best_right_x0:
                    best_right_x0 = nb.x0
                    best_right = nb

            # Top neighbor: nb is above, sufficient horizontal overlap.
            if nb.y1 <= vbox.y0 + ng_y and _h_overlap_frac(vbox, nb) >= overlap_frac_min:
                if nb.y1 > best_top_y1:
                    best_top_y1 = nb.y1
                    best_top = nb

            # Bottom neighbor: nb is below.
            if nb.y0 >= vbox.y1 - ng_y and _h_overlap_frac(vbox, nb) >= overlap_frac_min:
                if nb.y0 < best_bottom_y0:
                    best_bottom_y0 = nb.y0
                    best_bottom = nb

        tb_v = template_bboxes.get(vkey)
        # Stamp-edge fallback only when the template field actually lies on that
        # edge. Otherwise (vbox.y0 - stamp.y0) > thresh is true for every row
        # below the top, and middle cells get stretched to the full stamp.
        edge_tol = max(ng_x, ng_y, gap_thresh_x, gap_thresh_y) * 2.0

        # Extend left.
        if best_left is not None:
            gap = vbox.x0 - best_left.x1
            if gap > gap_thresh_x:
                new_x0 = best_left.x1 + ng_x
                changed = True
        elif (
            stamp_bbox is not None
            and tb_v is not None
            and (tb_v.x0 - stamp_bbox.x0) <= edge_tol
            and (vbox.x0 - stamp_bbox.x0) > gap_thresh_x
        ):
            new_x0 = stamp_bbox.x0
            changed = True

        # Extend right.
        if best_right is not None:
            gap = best_right.x0 - vbox.x1
            if gap > gap_thresh_x:
                new_x1 = best_right.x0 - ng_x
                changed = True
        elif (
            stamp_bbox is not None
            and tb_v is not None
            and (stamp_bbox.x1 - tb_v.x1) <= edge_tol
            and (stamp_bbox.x1 - vbox.x1) > gap_thresh_x
        ):
            new_x1 = stamp_bbox.x1
            changed = True

        # Extend top.
        if best_top is not None:
            gap = vbox.y0 - best_top.y1
            if gap > gap_thresh_y:
                new_y0 = best_top.y1 + ng_y
                changed = True
        elif (
            stamp_bbox is not None
            and tb_v is not None
            and (tb_v.y0 - stamp_bbox.y0) <= edge_tol
            and (vbox.y0 - stamp_bbox.y0) > gap_thresh_y
        ):
            new_y0 = stamp_bbox.y0
            changed = True

        # Extend bottom.
        if best_bottom is not None:
            gap = best_bottom.y0 - vbox.y1
            if gap > gap_thresh_y:
                new_y1 = best_bottom.y0 - ng_y
                changed = True
        elif (
            stamp_bbox is not None
            and tb_v is not None
            and (stamp_bbox.y1 - tb_v.y1) <= edge_tol
            and (stamp_bbox.y1 - vbox.y1) > gap_thresh_y
        ):
            new_y1 = stamp_bbox.y1
            changed = True

        if changed and new_x1 > new_x0 and new_y1 > new_y0:
            ar.bbox = CellBbox(new_x0, new_y0, new_x1, new_y1)
            ar.status = "matched_extended"
            extended_keys.append(vkey)

    return extended_keys

