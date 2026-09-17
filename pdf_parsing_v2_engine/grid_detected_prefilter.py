"""Prefilter detected table cells to one stamp-connected component before extent / alignment.

Islands are built from touch/adjacency in displayed space; components outside the Y-anchor
band around the stamp or far from template union ``t_all`` are dropped when safe.
"""

from __future__ import annotations

import math
from typing import Any, Protocol, Sequence


class _CellBBox(Protocol):
    x0: float
    y0: float
    x1: float
    y1: float


def _grid_detected_prefilter_settings(cfg: dict | None) -> dict[str, float | bool | int]:
    src = cfg or {}
    return {
        "enabled": bool(src.get("grid_detected_prefilter_enabled", True)),
        "min_cells": int(src.get("grid_detected_prefilter_min_cells", 8)),
        "y_pad_frac": float(src.get("grid_detected_prefilter_y_pad_frac", 0.18)),
        "y_pad_floor_pt": float(src.get("grid_detected_prefilter_y_pad_floor_pt", 8.0)),
        "max_removal_frac": float(src.get("grid_detected_prefilter_max_removal_frac", 0.58)),
        "min_component_cells": int(src.get("grid_detected_prefilter_min_component_cells", 4)),
        "min_component_area_frac": float(
            src.get("grid_detected_prefilter_min_component_area_frac", 0.08)
        ),
        "adjacency_tol_floor_pt": float(
            src.get("grid_detected_prefilter_adjacency_tol_floor_pt", 2.5)
        ),
    }


def _cell_area(c: _CellBBox) -> float:
    return max(0.0, c.x1 - c.x0) * max(0.0, c.y1 - c.y0)


def _cell_center(c: _CellBBox) -> tuple[float, float]:
    return ((c.x0 + c.x1) * 0.5, (c.y0 + c.y1) * 0.5)


def _intersects_open_rect(c: _CellBBox, x0: float, y0: float, x1: float, y1: float) -> bool:
    ix0 = max(c.x0, x0)
    iy0 = max(c.y0, y0)
    ix1 = min(c.x1, x1)
    iy1 = min(c.y1, y1)
    return ix1 > ix0 and iy1 > iy0


def _union_bbox(cells: Sequence[_CellBBox]) -> tuple[float, float, float, float]:
    return (
        min(c.x0 for c in cells),
        min(c.y0 for c in cells),
        max(c.x1 for c in cells),
        max(c.y1 for c in cells),
    )


def _overlap_area(a: _CellBBox, b: _CellBBox) -> float:
    x0 = max(a.x0, b.x0)
    y0 = max(a.y0, b.y0)
    x1 = min(a.x1, b.x1)
    y1 = min(a.y1, b.y1)
    if x1 <= x0 or y1 <= y0:
        return 0.0
    return (x1 - x0) * (y1 - y0)


def _cells_adjacent(a: _CellBBox, b: _CellBBox, tol: float) -> bool:
    x_overlap = min(a.x1, b.x1) - max(a.x0, b.x0)
    y_overlap = min(a.y1, b.y1) - max(a.y0, b.y0)
    h_gap = max(0.0, max(a.x0, b.x0) - min(a.x1, b.x1))
    v_gap = max(0.0, max(a.y0, b.y0) - min(a.y1, b.y1))
    if h_gap <= tol and x_overlap > -tol and y_overlap > -tol * 0.5:
        return True
    if v_gap <= tol and y_overlap > -tol and x_overlap > -tol * 0.5:
        return True
    return False


def _build_components(
    cells: list[_CellBBox],
    tol: float,
) -> list[list[int]]:
    """Return list of component id lists (indices into *cells*)."""
    n = len(cells)
    if n == 0:
        return []
    # Bucket by rounded y0 for locality (coarse grid).
    bucket_h = max(24.0, tol * 6.0)
    buckets: dict[int, list[int]] = {}
    for i, c in enumerate(cells):
        key = int(math.floor(c.y0 / bucket_h))
        buckets.setdefault(key, []).append(i)

    parent = list(range(n))

    def find(x: int) -> int:
        while parent[x] != x:
            parent[x] = parent[parent[x]]
            x = parent[x]
        return x

    def union(a: int, b: int) -> None:
        ra, rb = find(a), find(b)
        if ra != rb:
            parent[rb] = ra

    for i, c in enumerate(cells):
        bi0 = int(math.floor(c.y0 / bucket_h))
        for bk in (bi0 - 1, bi0, bi0 + 1):
            for j in buckets.get(bk, ()):
                if j <= i:
                    continue
                if _cells_adjacent(c, cells[j], tol):
                    union(i, j)

    comp_map: dict[int, list[int]] = {}
    for i in range(n):
        r = find(i)
        comp_map.setdefault(r, []).append(i)
    return list(comp_map.values())


def _bbox_like(x0: float, y0: float, x1: float, y1: float) -> Any:
    class _B:
        __slots__ = ("x0", "y0", "x1", "y1")

        def __init__(self) -> None:
            self.x0 = x0
            self.y0 = y0
            self.x1 = x1
            self.y1 = y1

    b = _B()
    return b


def prefilter_detected_cells_for_stamp_table(
    cells: list[_CellBBox],
    *,
    stamp_bbox_base: tuple[float, float, float, float],
    t_all: _CellBBox,
    tol_mm: float,
    scale_pts_per_mm: float,
    cfg: dict | None = None,
) -> tuple[list[_CellBBox], dict[str, Any]]:
    """Return filtered cells and JSON-safe diagnostics.

    On unsafe conditions returns the original *cells* list unchanged.
    """
    diag: dict[str, Any] = {
        "applied": False,
        "skipped": True,
        "reason": "disabled_or_trivial",
        "y_anchor_band_rect": None,
        "prior_rect": [t_all.x0, t_all.y0, t_all.x1, t_all.y1],
        "dropped_component_bboxes": [],
        "kept_component_bbox": None,
        "selected_component_ids": [],
        "n_input": len(cells),
        "n_output": len(cells),
    }
    settings = _grid_detected_prefilter_settings(cfg)
    if not settings["enabled"] or len(cells) < int(settings["min_cells"]):
        diag["reason"] = "disabled_or_trivial"
        return cells, diag

    sx0, sy0, sx1, sy1 = stamp_bbox_base
    sh = max(1e-6, sy1 - sy0)
    pad_y = max(float(settings["y_pad_frac"]) * sh, float(settings["y_pad_floor_pt"]))
    band_x0, band_y0, band_x1, band_y1 = sx0, sy0 - pad_y, sx1, sy1 + pad_y
    diag["y_anchor_band_rect"] = [band_x0, band_y0, band_x1, band_y1]

    tol_pts = max(
        float(tol_mm) * float(scale_pts_per_mm),
        float(settings["adjacency_tol_floor_pt"]),
    )

    comps = _build_components(cells, tol_pts)
    if len(comps) <= 1:
        diag["reason"] = "single_component"
        diag["skipped"] = True
        diag["n_output"] = len(cells)
        return cells, diag

    # Components with at least one cell intersecting the Y-anchor band (stamp X, expanded Y).
    touching: list[list[int]] = []
    for comp in comps:
        if any(
            _intersects_open_rect(cells[i], band_x0, band_y0, band_x1, band_y1)
            for i in comp
        ):
            touching.append(comp)

    if not touching:
        diag["reason"] = "prefilter skipped (unsafe): no component in Y-anchor band"
        diag["skipped"] = True
        return cells, diag

    t_like = t_all
    scored: list[tuple[float, float, list[int]]] = []
    for comp in touching:
        union_cells = [cells[i] for i in comp]
        ub = _bbox_like(*_union_bbox(union_cells))
        ov = _overlap_area(ub, t_like)
        cy = (ub.y0 + ub.y1) * 0.5
        _, ty = _cell_center(t_like)
        scored.append((-ov, abs(cy - ty), comp))

    scored.sort()
    winner = scored[0][2]
    kept_indices = set(winner)
    kept = [cells[i] for i in sorted(kept_indices)]

    n_orig = len(cells)
    n_kept = len(kept)
    removed_frac = (n_orig - n_kept) / max(1, n_orig)
    max_rm = float(settings["max_removal_frac"])
    if removed_frac > max_rm:
        diag["reason"] = (
            f"prefilter skipped (unsafe): removal {removed_frac:.0%} > {max_rm:.0%}"
        )
        diag["skipped"] = True
        diag["n_output"] = len(cells)
        return cells, diag

    min_cells = int(settings["min_component_cells"])
    if n_kept < min_cells:
        diag["reason"] = f"prefilter skipped (unsafe): kept {n_kept} < min_component_cells {min_cells}"
        diag["skipped"] = True
        diag["n_output"] = len(cells)
        return cells, diag

    orig_union = _union_bbox(cells)
    orig_area = max(
        1e-6,
        (orig_union[2] - orig_union[0]) * (orig_union[3] - orig_union[1]),
    )
    kept_union = _union_bbox(kept)
    kept_area = max(
        0.0,
        (kept_union[2] - kept_union[0]) * (kept_union[3] - kept_union[1]),
    )
    frac_area = kept_area / orig_area
    min_af = float(settings["min_component_area_frac"])
    if frac_area < min_af:
        diag["reason"] = (
            f"prefilter skipped (unsafe): kept union area {frac_area:.0%} < {min_af:.0%}"
        )
        diag["skipped"] = True
        diag["n_output"] = len(cells)
        return cells, diag

    dropped_boxes: list[list[float]] = []
    for comp in comps:
        if set(comp) <= kept_indices:
            continue
        if not set(comp) & kept_indices:
            u = _union_bbox([cells[i] for i in comp])
            dropped_boxes.append([u[0], u[1], u[2], u[3]])

    ku = _union_bbox(kept)
    diag["applied"] = True
    diag["skipped"] = False
    diag["reason"] = "selected_best_component_by_y_band_and_prior"
    diag["dropped_component_bboxes"] = dropped_boxes
    diag["kept_component_bbox"] = [ku[0], ku[1], ku[2], ku[3]]
    diag["selected_component_ids"] = [int(i) for i in sorted(kept_indices)]
    diag["n_output"] = n_kept
    return kept, diag
