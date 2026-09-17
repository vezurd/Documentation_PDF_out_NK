"""Post-pass: expand field rectangles to abut horizontal/vertical neighbors.

Used by the template editor after cell-assignment when PDF find_tables splits one
logical cell into several small cells: the matched bbox is only one fragment.
This pass snaps each side to the nearest neighbor edge (same row/column band)
so gaps inside the stamp are absorbed by the field that lies between neighbors.

Pure float geometry (scene or fitz pts); no Qt dependency."""

from __future__ import annotations

from typing import Any

Rect = tuple[float, float, float, float]  # x0, y0, x1, y1


def _v_overlap(a: Rect, b: Rect) -> float:
    return max(0.0, min(a[3], b[3]) - max(a[1], b[1]))


def _h_overlap(a: Rect, b: Rect) -> float:
    return max(0.0, min(a[2], b[2]) - max(a[0], b[0]))


def _min_w(a: Rect, b: Rect) -> float:
    wa = max(0.0, a[2] - a[0])
    wb = max(0.0, b[2] - b[0])
    return max(wa, wb, 1e-9)


def _find_left_neighbor(
    f: Rect, others: list[Rect], overlap_frac: float, eps: float,
) -> Rect | None:
    fh = f[3] - f[1]
    best: Rect | None = None
    best_x1 = -1e300
    for o in others:
        if o is f:
            continue
        ov = _v_overlap(f, o)
        if ov < overlap_frac * min(fh, o[3] - o[1]):
            continue
        if o[2] <= f[0] + eps:
            if o[2] > best_x1:
                best_x1 = o[2]
                best = o
    return best


def _find_right_neighbor(
    f: Rect, others: list[Rect], overlap_frac: float, eps: float,
) -> Rect | None:
    fh = f[3] - f[1]
    best: Rect | None = None
    best_x0 = 1e300
    for o in others:
        if o is f:
            continue
        ov = _v_overlap(f, o)
        if ov < overlap_frac * min(fh, o[3] - o[1]):
            continue
        if o[0] >= f[2] - eps:
            if o[0] < best_x0:
                best_x0 = o[0]
                best = o
    return best


def _find_top_neighbor(
    f: Rect, others: list[Rect], overlap_frac: float, eps: float,
) -> Rect | None:
    fw = f[2] - f[0]
    best: Rect | None = None
    best_y1 = -1e300
    for o in others:
        if o is f:
            continue
        oh = _h_overlap(f, o)
        if oh < overlap_frac * _min_w(f, o):
            continue
        if o[3] <= f[1] + eps:
            if o[3] > best_y1:
                best_y1 = o[3]
                best = o
    return best


def _find_bottom_neighbor(
    f: Rect, others: list[Rect], overlap_frac: float, eps: float,
) -> Rect | None:
    fw = f[2] - f[0]
    best: Rect | None = None
    best_y0 = 1e300
    for o in others:
        if o is f:
            continue
        oh = _h_overlap(f, o)
        if oh < overlap_frac * _min_w(f, o):
            continue
        if o[1] >= f[3] - eps:
            if o[1] < best_y0:
                best_y0 = o[1]
                best = o
    return best


def _clip_to_stamp(r: Rect, stamp: Rect) -> Rect:
    x0 = max(r[0], stamp[0])
    y0 = max(r[1], stamp[1])
    x1 = min(r[2], stamp[2])
    y1 = min(r[3], stamp[3])
    return (x0, y0, x1, y1)


def fill_stamp_gaps_for_rects(
    items: list[tuple[Any, Rect]],
    stamp: Rect,
    *,
    overlap_frac: float = 0.22,
    eps_scene: float = 1.0,
    expand_horizontal: bool = True,
    expand_vertical: bool = True,
    min_side_scene: float = 2.0,
) -> dict[Any, Rect]:
    """Return *changed* rectangles only: expand each field to meet neighbors inside stamp.

    Neighbors are detected by vertical overlap (for left/right) or horizontal
    overlap (for top/bottom). Uses a single pass over *initial* geometry (no
    sequential drift).
    """
    if not items:
        return {}

    out: dict[Any, Rect] = {}

    for i, (key, f) in enumerate(items):
        others = [items[j][1] for j in range(len(items)) if j != i]
        x0, y0, x1, y1 = f
        nx0, ny0, nx1, ny1 = x0, y0, x1, y1

        # Horizontal: only when BOTH neighbors exist (sandwiched column).
        # Otherwise edge fields would absorb gaps toward the middle (wrong).
        if expand_horizontal:
            left = _find_left_neighbor(f, others, overlap_frac, eps_scene)
            right = _find_right_neighbor(f, others, overlap_frac, eps_scene)
            if left is not None and right is not None and left[2] < right[0]:
                nx0 = max(stamp[0], left[2])
                nx1 = min(stamp[2], right[0])

        if expand_vertical:
            top = _find_top_neighbor(f, others, overlap_frac, eps_scene)
            bottom = _find_bottom_neighbor(f, others, overlap_frac, eps_scene)
            if top is not None and bottom is not None and top[3] < bottom[1]:
                ny0 = max(stamp[1], top[3])
                ny1 = min(stamp[3], bottom[1])

        new_r = _clip_to_stamp((nx0, ny0, nx1, ny1), stamp)
        w = new_r[2] - new_r[0]
        h = new_r[3] - new_r[1]
        if w < min_side_scene or h < min_side_scene:
            continue
        if (
            abs(new_r[0] - x0) > 0.02
            or abs(new_r[1] - y0) > 0.02
            or abs(new_r[2] - x1) > 0.02
            or abs(new_r[3] - y1) > 0.02
        ):
            out[key] = new_r

    return out
