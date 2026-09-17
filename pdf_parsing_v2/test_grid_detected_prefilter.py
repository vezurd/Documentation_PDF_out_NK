"""Unit tests for detected-cell island prefilter (two-table handoff)."""

from __future__ import annotations

from pdf_parsing_v2_engine.coord_transform import SCALE
from pdf_parsing_v2_engine.grid_detected_prefilter import prefilter_detected_cells_for_stamp_table
from pdf_parsing_v2_engine.grid_matcher import CellBbox


def test_prefilter_keeps_stamp_island_drops_side_table() -> None:
    # Stamp body 100–200 x 100–200; stray "second table" to the right, not in Y-band overlap with stamp X.
    stamp = (100.0, 100.0, 200.0, 200.0)
    t_all = CellBbox(110.0, 120.0, 190.0, 180.0)

    island_a: list[CellBbox] = []
    for row in range(4):
        y0 = 110.0 + row * 18.0
        for col in range(4):
            x0 = 115.0 + col * 18.0
            island_a.append(CellBbox(x0, y0, x0 + 16.0, y0 + 16.0))

    island_b: list[CellBbox] = []
    for row in range(4):
        y0 = 105.0 + row * 20.0
        for col in range(3):
            x0 = 320.0 + col * 22.0
            island_b.append(CellBbox(x0, y0, x0 + 20.0, y0 + 18.0))

    cells = island_a + island_b
    cfg = {"grid_detected_prefilter_min_cells": 4}
    out, diag = prefilter_detected_cells_for_stamp_table(
        cells,
        stamp_bbox_base=stamp,
        t_all=t_all,
        tol_mm=0.5,
        scale_pts_per_mm=float(SCALE),
        cfg=cfg,
    )
    assert diag.get("applied") is True
    assert len(out) == len(island_a)
    assert all(c.x0 < 250.0 for c in out)


def test_prefilter_noop_single_component() -> None:
    stamp = (0.0, 0.0, 100.0, 100.0)
    t_all = CellBbox(10.0, 10.0, 90.0, 90.0)
    cells = [CellBbox(20.0, 20.0, 40.0, 40.0), CellBbox(40.0, 20.0, 60.0, 40.0)]
    out, diag = prefilter_detected_cells_for_stamp_table(
        cells,
        stamp_bbox_base=stamp,
        t_all=t_all,
        tol_mm=0.5,
        scale_pts_per_mm=float(SCALE),
        cfg={"grid_detected_prefilter_min_cells": 2},
    )
    assert out is cells or len(out) == len(cells)
    assert diag.get("reason") == "single_component"
