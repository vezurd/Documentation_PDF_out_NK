"""Utilities for generating and binding TemplateGridLine objects.

Phase 3: generate_grid_lines_from_fields — auto-create wireframe lines from field edges.
Phase 4: auto_bind_fields_to_lines — fill field.bound_* by nearest line per side.

Coordinate convention: same as FieldDef.bbox_mm / TemplateGridLine.pos_mm.
Currently implemented for template origin ``frame_bottom_right`` (the default).
For other origins the axis directions differ; manual correction or a future
extension may be needed.
"""

from __future__ import annotations

from dataclasses import replace
from typing import AbstractSet, Sequence

from pdf_parsing_v2_engine.models import FieldDef, TemplateGridLine

_EPS = 1e-9

BORDER_MERGE_FRACTION = 0.20


# ---------------------------------------------------------------------------
# Internal helpers
# ---------------------------------------------------------------------------

def _cluster(values: list[float], tolerance: float) -> list[float]:
    """Cluster nearby float values and return sorted centroid list."""
    if not values:
        return []
    sv = sorted(values)
    clusters: list[list[float]] = [[sv[0]]]
    for v in sv[1:]:
        if abs(v - clusters[-1][-1]) <= tolerance:
            clusters[-1].append(v)
        else:
            clusters.append([v])
    return sorted(sum(c) / len(c) for c in clusters)


def _merge_border_pairs_mm(
    values: list[float], threshold: float,
) -> list[float]:
    """Merge values within *threshold*, taking the median of each group.

    This collapses border pairs (top/bottom edge of the same table line)
    that survive the initial clustering step.
    """
    if not values or threshold <= _EPS:
        return values
    sv = sorted(values)
    groups: list[list[float]] = [[sv[0]]]
    for v in sv[1:]:
        if v - groups[-1][-1] <= threshold:
            groups[-1].append(v)
        else:
            groups.append([v])
    result: list[float] = []
    for g in groups:
        n = len(g)
        if n % 2 == 1:
            result.append(g[n // 2])
        else:
            result.append((g[n // 2 - 1] + g[n // 2]) * 0.5)
    return sorted(result)


def _min_cell_dim_mm(fields: Sequence[FieldDef]) -> float:
    """Return the smallest cell dimension (height or width) among active fields."""
    dims: list[float] = []
    for f in fields:
        if f.outside_stamp:
            continue
        h1, v1, h2, v2 = f.bbox_mm
        dh = abs(h2 - h1)
        dv = abs(v2 - v1)
        if dh > _EPS:
            dims.append(dh)
        if dv > _EPS:
            dims.append(dv)
    return min(dims) if dims else 0.0


def _span_covers(line_start: float, line_end: float,
                 field_lo: float, field_hi: float,
                 overlap_frac: float = 0.5) -> bool:
    """Return True if the line's span covers ≥ overlap_frac of the field's range."""
    lo = min(line_start, line_end)
    hi = max(line_start, line_end)
    overlap = max(0.0, min(hi, field_hi) - max(lo, field_lo))
    return overlap / max(field_hi - field_lo, _EPS) >= overlap_frac


def _id_from_pos(prefix: str, pos_mm: float) -> str:
    """Auto-generate a line ID like 'h_5.3' or 'v_65.5'."""
    rounded = round(pos_mm, 2)
    if rounded == int(rounded):
        return f"{prefix}_{int(rounded)}"
    return f"{prefix}_{rounded}"


# ---------------------------------------------------------------------------
# Phase 3: generate_grid_lines_from_fields
# ---------------------------------------------------------------------------

def generate_grid_lines_from_fields(
    fields: Sequence[FieldDef],
    tolerance_mm: float = 0.5,
) -> list[TemplateGridLine]:
    """Create TemplateGridLine objects from field bounding-box edges.

    Works for template origin ``frame_bottom_right``.
    In this convention ``bbox_mm = [h1, v1, h2, v2]`` where:
      - h-axis increases leftward (h1/h2 are horizontal distances from right)
      - v-axis increases upward   (v1/v2 are vertical distances from bottom)

    Steps
    -----
    1. Collect all v-edge values (v1, v2) → cluster → one H-line per cluster.
    2. Collect all h-edge values (h1, h2) → cluster → one V-line per cluster.
    3. Set each line's span to the extent of touching fields.
    4. Mark boundary lines (smallest/largest v → bottom/top, smallest/largest h → right/left).
    5. Return sorted list: H-lines by pos ascending, then V-lines.
    """
    active = [f for f in fields if not f.outside_stamp]
    if not active:
        return []

    # --- gather all edge values ---
    v_vals: list[float] = []
    h_vals: list[float] = []
    for f in active:
        h1, v1, h2, v2 = f.bbox_mm
        v_vals += [v1, v2]
        h_vals += [h1, h2]

    v_clusters = _cluster(v_vals, tolerance_mm)
    h_clusters = _cluster(h_vals, tolerance_mm)

    if not v_clusters or not h_clusters:
        return []

    # Adaptive border-pair merge: 20% of min cell dimension
    min_dim = _min_cell_dim_mm(active)
    border_merge_mm = min_dim * BORDER_MERGE_FRACTION
    if border_merge_mm > tolerance_mm:
        v_clusters = _merge_border_pairs_mm(v_clusters, border_merge_mm)
        h_clusters = _merge_border_pairs_mm(h_clusters, border_merge_mm)

    v_min, v_max = min(v_clusters), max(v_clusters)
    h_min, h_max = min(h_clusters), max(h_clusters)

    def _boundary_v(pos: float) -> str | None:
        if abs(pos - v_min) <= tolerance_mm:
            return "bottom"
        if abs(pos - v_max) <= tolerance_mm:
            return "top"
        return None

    def _boundary_h(pos: float) -> str | None:
        if abs(pos - h_min) <= tolerance_mm:
            return "right"
        if abs(pos - h_max) <= tolerance_mm:
            return "left"
        return None

    # --- build H-lines (for each v-cluster centroid) ---
    h_lines: list[TemplateGridLine] = []
    used_h_ids: set[str] = set()
    for v_pos in v_clusters:
        # Find h-extent of all fields whose v1 or v2 is close to this v-cluster.
        touching_h: list[float] = []
        for f in active:
            h1, v1, h2, v2 = f.bbox_mm
            if abs(v1 - v_pos) <= tolerance_mm or abs(v2 - v_pos) <= tolerance_mm:
                touching_h += [h1, h2]
        if not touching_h:
            touching_h = [h_min, h_max]
        span_start = min(touching_h)
        span_end = max(touching_h)

        boundary = _boundary_v(v_pos)
        raw_id = _id_from_pos("h", v_pos)
        lid = raw_id
        if boundary == "bottom":
            lid = "h_bottom"
        elif boundary == "top":
            lid = "h_top"
        # ensure uniqueness
        if lid in used_h_ids:
            lid = f"{raw_id}_a"
        used_h_ids.add(lid)

        h_lines.append(TemplateGridLine(
            id=lid,
            orientation="h",
            pos_mm=round(v_pos, 3),
            start_mm=round(span_start, 3),
            end_mm=round(span_end, 3),
            boundary=boundary,
        ))

    # --- build V-lines (for each h-cluster centroid) ---
    v_lines: list[TemplateGridLine] = []
    used_v_ids: set[str] = set()
    for h_pos in h_clusters:
        touching_v: list[float] = []
        for f in active:
            h1, v1, h2, v2 = f.bbox_mm
            if abs(h1 - h_pos) <= tolerance_mm or abs(h2 - h_pos) <= tolerance_mm:
                touching_v += [v1, v2]
        if not touching_v:
            touching_v = [v_min, v_max]
        span_start = min(touching_v)
        span_end = max(touching_v)

        boundary = _boundary_h(h_pos)
        raw_id = _id_from_pos("v", h_pos)
        lid = raw_id
        if boundary == "right":
            lid = "v_right"
        elif boundary == "left":
            lid = "v_left"
        if lid in used_v_ids:
            lid = f"{raw_id}_a"
        used_v_ids.add(lid)

        v_lines.append(TemplateGridLine(
            id=lid,
            orientation="v",
            pos_mm=round(h_pos, 3),
            start_mm=round(span_start, 3),
            end_mm=round(span_end, 3),
            boundary=boundary,
        ))

    # H-lines sorted by pos ascending (bottom → top), then V-lines.
    h_lines.sort(key=lambda l: l.pos_mm)
    v_lines.sort(key=lambda l: l.pos_mm)
    all_lines = h_lines + v_lines

    all_lines = snap_endpoints_to_refs(auto_assign_line_refs(all_lines, tolerance_mm))

    return all_lines


# ---------------------------------------------------------------------------
# Phase 4: auto_bind_fields_to_lines
# ---------------------------------------------------------------------------

def auto_bind_fields_to_lines(
    fields: Sequence[FieldDef],
    grid_lines: Sequence[TemplateGridLine],
    tolerance_mm: float = 1.0,
    overlap_frac: float = 0.5,
) -> list[FieldDef]:
    """Assign bound_top/bottom/left/right IDs to each field by nearest line.

    Works for template origin ``frame_bottom_right``.

    In this convention:
      - H-line at ``pos_mm`` represents a physical V-axis position (mm from bottom).
        - ``bound_top`` = H-line with pos_mm ≥ max(v1,v2) (above field → smaller fitz y0)
        - ``bound_bottom`` = H-line with pos_mm ≤ min(v1,v2) (below field → larger fitz y1)
      - V-line at ``pos_mm`` represents a physical H-axis position (mm from right).
        - ``bound_right`` = V-line with pos_mm ≤ min(h1,h2) (right side → larger fitz x1)
        - ``bound_left``  = V-line with pos_mm ≥ max(h1,h2) (left side → smaller fitz x0)

    Returns a new list with updated FieldDef objects.  Original objects are not mutated.
    """
    h_lines = [gl for gl in grid_lines if gl.orientation == "h"]
    v_lines = [gl for gl in grid_lines if gl.orientation == "v"]

    result: list[FieldDef] = []
    for f in fields:
        if f.outside_stamp:
            result.append(f)
            continue

        h1, v1, h2, v2 = f.bbox_mm
        v_lo, v_hi = min(v1, v2), max(v1, v2)
        h_lo, h_hi = min(h1, h2), max(h1, h2)

        # --- bound_top: nearest H-line at pos_mm >= v_hi ---
        candidates_top = [
            gl for gl in h_lines
            if gl.pos_mm >= v_hi - tolerance_mm
            and _span_covers(gl.start_mm, gl.end_mm, h_lo, h_hi, overlap_frac)
        ]
        bound_top: str | None = None
        if candidates_top:
            best = min(candidates_top, key=lambda gl: abs(gl.pos_mm - v_hi))
            if abs(best.pos_mm - v_hi) <= tolerance_mm * 3:
                bound_top = best.id

        # --- bound_bottom: nearest H-line at pos_mm <= v_lo ---
        candidates_bottom = [
            gl for gl in h_lines
            if gl.pos_mm <= v_lo + tolerance_mm
            and _span_covers(gl.start_mm, gl.end_mm, h_lo, h_hi, overlap_frac)
        ]
        bound_bottom: str | None = None
        if candidates_bottom:
            best = min(candidates_bottom, key=lambda gl: abs(gl.pos_mm - v_lo))
            if abs(best.pos_mm - v_lo) <= tolerance_mm * 3:
                bound_bottom = best.id

        # --- bound_right: nearest V-line at pos_mm <= h_lo ---
        candidates_right = [
            gl for gl in v_lines
            if gl.pos_mm <= h_lo + tolerance_mm
            and _span_covers(gl.start_mm, gl.end_mm, v_lo, v_hi, overlap_frac)
        ]
        bound_right: str | None = None
        if candidates_right:
            best = min(candidates_right, key=lambda gl: abs(gl.pos_mm - h_lo))
            if abs(best.pos_mm - h_lo) <= tolerance_mm * 3:
                bound_right = best.id

        # --- bound_left: nearest V-line at pos_mm >= h_hi ---
        candidates_left = [
            gl for gl in v_lines
            if gl.pos_mm >= h_hi - tolerance_mm
            and _span_covers(gl.start_mm, gl.end_mm, v_lo, v_hi, overlap_frac)
        ]
        bound_left: str | None = None
        if candidates_left:
            best = min(candidates_left, key=lambda gl: abs(gl.pos_mm - h_hi))
            if abs(best.pos_mm - h_hi) <= tolerance_mm * 3:
                bound_left = best.id

        result.append(replace(
            f,
            bound_top=bound_top,
            bound_bottom=bound_bottom,
            bound_left=bound_left,
            bound_right=bound_right,
        ))

    return result


def snap_fields_to_bindings(
    fields: Sequence[FieldDef],
    grid_lines: Sequence[TemplateGridLine],
) -> list[FieldDef]:
    """Resize field bbox_mm edges to match bound grid-line positions.

    For each field with ``bound_top/bottom/left/right`` set, the corresponding
    edge of ``bbox_mm`` is moved to the grid line's ``pos_mm``.
    This mirrors what ``expand_fields_by_bindings`` does in the pipeline
    (in fitz pts) but operates directly in template mm coordinates.

    Returns a new list; original FieldDef objects are not mutated.
    """
    line_map: dict[str, float] = {gl.id: gl.pos_mm for gl in grid_lines}
    result: list[FieldDef] = []

    for f in fields:
        if f.outside_stamp or not (
            f.bound_top or f.bound_bottom or f.bound_left or f.bound_right
        ):
            result.append(f)
            continue

        h1, v1, h2, v2 = f.bbox_mm
        v_lo, v_hi = min(v1, v2), max(v1, v2)
        h_lo, h_hi = min(h1, h2), max(h1, h2)
        changed = False

        if f.bound_top and f.bound_top in line_map:
            v_hi = line_map[f.bound_top]
            changed = True
        if f.bound_bottom and f.bound_bottom in line_map:
            v_lo = line_map[f.bound_bottom]
            changed = True
        if f.bound_left and f.bound_left in line_map:
            h_hi = line_map[f.bound_left]
            changed = True
        if f.bound_right and f.bound_right in line_map:
            h_lo = line_map[f.bound_right]
            changed = True

        if not changed or h_hi <= h_lo or v_hi <= v_lo:
            result.append(f)
            continue

        new_v1, new_v2 = (v_lo, v_hi) if v1 <= v2 else (v_hi, v_lo)
        new_h1, new_h2 = (h_lo, h_hi) if h1 <= h2 else (h_hi, h_lo)

        result.append(replace(f, bbox_mm=(new_h1, new_v1, new_h2, new_v2)))

    return result


def sanitize_grid_line_refs(grid_lines: list[TemplateGridLine]) -> list[TemplateGridLine]:
    """Clear ``start_line_id`` / ``end_line_id`` that reference missing line IDs.

    Call after deleting (or renaming) lines so refs do not point outside the list.
    """
    valid = {gl.id for gl in grid_lines}
    out: list[TemplateGridLine] = []
    for gl in grid_lines:
        sid = gl.start_line_id if gl.start_line_id in valid else None
        eid = gl.end_line_id if gl.end_line_id in valid else None
        if sid != gl.start_line_id or eid != gl.end_line_id:
            out.append(replace(gl, start_line_id=sid, end_line_id=eid))
        else:
            out.append(gl)
    return out


def sanitize_field_line_bindings(
    fd: FieldDef,
    valid_line_ids: AbstractSet[str],
) -> FieldDef:
    """Clear ``bound_*`` that reference grid line IDs not in *valid_line_ids*.

    Returns *fd* unchanged if nothing to strip. Used when lines are removed from
    the template so field wireframe colors and combos stay consistent.
    """
    if fd.outside_stamp:
        return fd

    def _keep(b: str | None) -> str | None:
        return b if b and b in valid_line_ids else None

    bt = _keep(fd.bound_top)
    bb = _keep(fd.bound_bottom)
    bl = _keep(fd.bound_left)
    br = _keep(fd.bound_right)
    if (bt, bb, bl, br) == (fd.bound_top, fd.bound_bottom, fd.bound_left, fd.bound_right):
        return fd
    return replace(fd, bound_top=bt, bound_bottom=bb, bound_left=bl, bound_right=br)


# ---------------------------------------------------------------------------
# Phase 4b: auto_assign_line_refs
# ---------------------------------------------------------------------------

def auto_assign_line_refs(
    grid_lines: list[TemplateGridLine],
    tolerance_mm: float = 1.0,
) -> list[TemplateGridLine]:
    """For each grid line, find perpendicular lines at start_mm and end_mm.

    A perpendicular line qualifies if its span covers the current line's
    ``pos_mm`` (within *tolerance_mm*).  The closest such perpendicular line
    to ``start_mm`` / ``end_mm`` is always chosen — no distance upper-bound
    is applied.  The cover check already guarantees semantic correctness:
    the assigned ref line passes through the region spanned by this line.

    This means refs are correctly re-assigned even after a line has been
    dragged in snapped-wireframe mode, where ``start_mm``/``end_mm`` may
    reflect PDF-snapped positions that differ from template values by more
    than a small tolerance.

    Returns a new list; originals are not mutated.
    """
    result: list[TemplateGridLine] = []
    for gl in grid_lines:
        perp = [p for p in grid_lines if p.orientation != gl.orientation]
        if not perp:
            result.append(gl)
            continue

        best_start_id: str | None = None
        best_start_d = float("inf")
        best_end_id: str | None = None
        best_end_d = float("inf")

        for p in perp:
            lo = min(p.start_mm, p.end_mm) - tolerance_mm
            hi = max(p.start_mm, p.end_mm) + tolerance_mm
            if not (lo <= gl.pos_mm <= hi):
                continue
            d_s = abs(p.pos_mm - gl.start_mm)
            d_e = abs(p.pos_mm - gl.end_mm)
            if d_s < best_start_d:
                best_start_d = d_s
                best_start_id = p.id
            if d_e < best_end_d:
                best_end_d = d_e
                best_end_id = p.id

        result.append(replace(
            gl,
            start_line_id=best_start_id,
            end_line_id=best_end_id,
        ))
    return result


# ---------------------------------------------------------------------------
# Phase 4c: snap_endpoints_to_refs
# ---------------------------------------------------------------------------

def snap_endpoints_to_refs(
    grid_lines: list[TemplateGridLine],
) -> list[TemplateGridLine]:
    """Set start_mm / end_mm to the pos_mm of the referenced perpendicular line.

    If ``start_line_id`` points to a line in the list, ``start_mm`` is set to
    that line's ``pos_mm`` (and analogously for ``end_line_id`` / ``end_mm``).
    Lines without refs keep their original coordinates.

    Returns a new list; originals are not mutated.
    """
    by_id = {gl.id: gl for gl in grid_lines}
    result: list[TemplateGridLine] = []
    for gl in grid_lines:
        s = gl.start_mm
        e = gl.end_mm
        if gl.start_line_id and gl.start_line_id in by_id:
            s = round(by_id[gl.start_line_id].pos_mm, 3)
        if gl.end_line_id and gl.end_line_id in by_id:
            e = round(by_id[gl.end_line_id].pos_mm, 3)
        if s != gl.start_mm or e != gl.end_mm:
            result.append(replace(gl, start_mm=s, end_mm=e))
        else:
            result.append(gl)
    return result


# ---------------------------------------------------------------------------
# Phase 5b: resolve_line_endpoints_from_snapped
# ---------------------------------------------------------------------------

def resolve_line_endpoints_from_snapped(
    grid_lines: Sequence[TemplateGridLine],
    snapped_positions: dict[str, float],
) -> dict[str, tuple[float, float]]:
    """Resolve each line's start/end from snapped perpendicular line positions.

    If a line has ``start_line_id`` / ``end_line_id`` and the referenced line
    exists in *snapped_positions*, the start/end is set to that snapped value.
    Otherwise the original (transformed) value is kept.

    Returns ``{line_id: (start_pts, end_pts)}`` — resolved span for each line
    in the same coordinate system as *snapped_positions* values (fitz pts).
    """
    result: dict[str, tuple[float, float]] = {}
    for gl in grid_lines:
        start_val = snapped_positions.get(gl.start_line_id, None) if gl.start_line_id else None
        end_val = snapped_positions.get(gl.end_line_id, None) if gl.end_line_id else None
        result[gl.id] = (start_val, end_val)
    return result


# ---------------------------------------------------------------------------
# Doubled-line detection and collapse
# ---------------------------------------------------------------------------

def _span_overlap_ratio(a: TemplateGridLine, b: TemplateGridLine) -> float:
    """Overlap of two line spans as a fraction of the shorter span."""
    lo = max(a.start_mm, b.start_mm)
    hi = min(a.end_mm, b.end_mm)
    overlap = max(0.0, hi - lo)
    min_span = min(abs(a.end_mm - a.start_mm), abs(b.end_mm - b.start_mm))
    if min_span <= _EPS:
        return 0.0
    return overlap / min_span


def find_doubled_lines(
    grid_lines: Sequence[TemplateGridLine],
    fields: Sequence[FieldDef],
    fraction: float = BORDER_MERGE_FRACTION,
    min_overlap: float = 0.5,
) -> list[tuple[TemplateGridLine, TemplateGridLine]]:
    """Find pairs of grid lines that are border-pair duplicates.

    Returns a list of ``(line_a, line_b)`` where both are of the same
    orientation, their ``pos_mm`` difference is less than *fraction*
    of the minimum cell dimension, **and** their spans overlap by at least
    *min_overlap* of the shorter span.  Pairs with non-overlapping spans
    are split segments and are NOT considered duplicates.
    """
    min_dim = _min_cell_dim_mm(fields)
    threshold = min_dim * fraction
    if threshold <= _EPS:
        return []

    pairs: list[tuple[TemplateGridLine, TemplateGridLine]] = []
    for orient in ("h", "v"):
        lines = sorted(
            [gl for gl in grid_lines if gl.orientation == orient],
            key=lambda gl: gl.pos_mm,
        )
        for i in range(len(lines) - 1):
            if abs(lines[i + 1].pos_mm - lines[i].pos_mm) <= threshold:
                if _span_overlap_ratio(lines[i], lines[i + 1]) >= min_overlap:
                    pairs.append((lines[i], lines[i + 1]))
    return pairs


def find_split_segments(
    grid_lines: Sequence[TemplateGridLine],
    fields: Sequence[FieldDef],
    fraction: float = BORDER_MERGE_FRACTION,
    min_overlap: float = 0.5,
) -> list[tuple[TemplateGridLine, TemplateGridLine]]:
    """Find pairs of close lines that are split segments (NOT border pairs).

    Same proximity check as ``find_doubled_lines`` but returns pairs whose
    spans do NOT overlap enough — i.e. they cover different parts of the stamp.
    """
    min_dim = _min_cell_dim_mm(fields)
    threshold = min_dim * fraction
    if threshold <= _EPS:
        return []

    pairs: list[tuple[TemplateGridLine, TemplateGridLine]] = []
    for orient in ("h", "v"):
        lines = sorted(
            [gl for gl in grid_lines if gl.orientation == orient],
            key=lambda gl: gl.pos_mm,
        )
        for i in range(len(lines) - 1):
            if abs(lines[i + 1].pos_mm - lines[i].pos_mm) <= threshold:
                if _span_overlap_ratio(lines[i], lines[i + 1]) < min_overlap:
                    pairs.append((lines[i], lines[i + 1]))
    return pairs


def collapse_doubled_lines(
    grid_lines: list[TemplateGridLine],
    fields: Sequence[FieldDef],
    fraction: float = BORDER_MERGE_FRACTION,
) -> tuple[list[TemplateGridLine], list[FieldDef], list[str]]:
    """Collapse border-pair duplicate lines and rebind affected fields.

    Returns ``(merged_lines, updated_fields, messages)`` where *messages*
    lists the IDs of lines that were merged (for user feedback).
    """
    pairs = find_doubled_lines(grid_lines, fields, fraction)
    if not pairs:
        return list(grid_lines), list(fields), []

    # Build a mapping: removed_id → surviving_id
    remap: dict[str, str] = {}
    to_remove: set[str] = set()
    merged_lines = list(grid_lines)
    messages: list[str] = []

    for a, b in pairs:
        if a.id in to_remove or b.id in to_remove:
            continue
        median_pos = round((a.pos_mm + b.pos_mm) / 2.0, 3)
        span_start = min(a.start_mm, b.start_mm)
        span_end = max(a.end_mm, b.end_mm)
        # Keep whichever has a boundary marker, or the first one
        if b.boundary and not a.boundary:
            keep, drop = b, a
        else:
            keep, drop = a, b
        remap[drop.id] = keep.id
        to_remove.add(drop.id)
        # Update the surviving line position to median
        idx = next(i for i, gl in enumerate(merged_lines) if gl.id == keep.id)
        merged_lines[idx] = replace(
            keep,
            pos_mm=median_pos,
            start_mm=round(span_start, 3),
            end_mm=round(span_end, 3),
        )
        messages.append(
            f"{drop.id} ({drop.pos_mm:.2f}mm) → {keep.id} ({median_pos:.2f}mm)"
        )

    merged_lines = [gl for gl in merged_lines if gl.id not in to_remove]

    # Rebind fields
    updated_fields: list[FieldDef] = []
    for f in fields:
        bt = remap.get(f.bound_top, f.bound_top) if f.bound_top else None
        bb = remap.get(f.bound_bottom, f.bound_bottom) if f.bound_bottom else None
        bl = remap.get(f.bound_left, f.bound_left) if f.bound_left else None
        br = remap.get(f.bound_right, f.bound_right) if f.bound_right else None
        if (bt, bb, bl, br) != (f.bound_top, f.bound_bottom, f.bound_left, f.bound_right):
            updated_fields.append(replace(
                f, bound_top=bt, bound_bottom=bb, bound_left=bl, bound_right=br,
            ))
        else:
            updated_fields.append(f)

    # Re-assign refs since some referenced lines may have been removed
    merged_lines = auto_assign_line_refs(merged_lines)
    merged_lines = snap_endpoints_to_refs(merged_lines)

    return merged_lines, updated_fields, messages
