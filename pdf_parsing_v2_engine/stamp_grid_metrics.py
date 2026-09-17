"""Целочисленная оценка «разъезда» каркаса штампа относительно PDF (монитор Run)."""

from __future__ import annotations

from typing import Any

from pdf_parsing_v2_engine.grid_matcher import AlignmentLineDiag, CellAssignmentInfo


def _normalize_weights(cfg: dict[str, Any] | None) -> tuple[float, float, float]:
    src = cfg or {}
    wg = float(src.get("monitor_grid_mismatch_weight_global", 100_000.0))
    wd = float(src.get("monitor_grid_mismatch_weight_walk_pt", 50.0))
    wn = float(src.get("monitor_grid_mismatch_weight_no_match", 25_000.0))
    return max(wg, 0.0), max(wd, 0.0), max(wn, 0.0)


def compute_stamp_grid_mismatch(
    ca_info: CellAssignmentInfo,
    cfg: dict[str, Any] | None = None,
) -> tuple[int, dict[str, Any]]:
    """Возвращает (score int >= 0, breakdown для tooltip / JSON).

    Score = округление суммы трёх вкладов:

    - глобальный разный масштаб по осям: ``weight_global * |sx − sy|`` (безразмерное × коэффициент → крупное целое);
    - сумма по линиям каркаса: ``weight_walk * Σ|snap − interp|`` в pt (ожидание прогрессивной интерполяции vs фактический snap);
    - строки без ребра PDF: ``weight_no_match * no_match_count`` из отчёта alignment.
    """
    wg, wd, wn = _normalize_weights(cfg)
    sx = float(ca_info.transform_scale_x)
    sy = float(ca_info.transform_scale_y)
    global_part = wg * abs(sx - sy)

    walk_sum_pts = 0.0
    for d in ca_info.alignment_diagnostics:
        if isinstance(d, AlignmentLineDiag):
            walk_sum_pts += abs(float(d.pos_snapped_pts) - float(d.pos_interp_pts))

    walk_part = wd * walk_sum_pts

    no_match = 0
    if ca_info.alignment_report is not None:
        no_match = int(ca_info.alignment_report.no_match_count)
    nm_part = wn * float(no_match)

    total = int(round(global_part + walk_part + nm_part))
    breakdown: dict[str, Any] = {
        "global": int(round(global_part)),
        "walk_weighted": int(round(walk_part)),
        "no_match_weighted": int(round(nm_part)),
        "walk_sum_pts": round(walk_sum_pts, 3),
        "no_match_lines": no_match,
        "abs_sx_minus_sy": round(abs(sx - sy), 6),
    }
    return max(0, total), breakdown
