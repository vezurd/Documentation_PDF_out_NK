"""Агрегация оценки «Разъезд сетки» по страницам для timing и Control Center."""

from __future__ import annotations

import json
from typing import Any

from pdf_parsing_v2_engine.models import V2PageResult


def stamp_grid_mismatch_file_aggregate(v2_results: list[V2PageResult]) -> dict[str, Any]:
    """Поля для ``TimingCollector.record_file`` (JSON-safe).

    Худшая страница по ``stamp_grid_mismatch_score`` задаёт цвет ячейки в Run.
    """
    rows: list[dict[str, Any]] = []
    worst: int | None = None
    for v2r in v2_results:
        meta = v2r.metadata or {}
        sc = meta.get("stamp_grid_mismatch_score")
        if sc is None:
            continue
        score_i = int(sc)
        bd_raw = meta.get("stamp_grid_mismatch_breakdown_json")
        breakdown: dict[str, Any] = {}
        if isinstance(bd_raw, str) and bd_raw.strip():
            try:
                breakdown = json.loads(bd_raw)
            except json.JSONDecodeError:
                breakdown = {}
        rows.append(
            {
                "page": int(v2r.page_num),
                "score": score_i,
                "breakdown": breakdown,
            }
        )
        if worst is None or score_i > worst:
            worst = score_i
    out: dict[str, Any] = {}
    if rows and worst is not None:
        out["stamp_grid_mismatch_worst_score"] = int(worst)
        out["stamp_grid_mismatch_by_page_json"] = json.dumps(rows, ensure_ascii=False)
    return out


def stamp_grid_mismatch_detail_for_ui(v2_results: list[V2PageResult]) -> dict[str, Any]:
    """Payload для ``on_file_done`` (список страниц с разбором)."""
    flat = stamp_grid_mismatch_file_aggregate(v2_results)
    if not flat:
        return {}
    try:
        pages = json.loads(str(flat["stamp_grid_mismatch_by_page_json"]))
    except (json.JSONDecodeError, KeyError):
        pages = []
    return {
        "stamp_grid_mismatch_worst_score": flat.get("stamp_grid_mismatch_worst_score"),
        "stamp_grid_mismatch_by_page": pages,
    }


# Совместимость имён для старых импортов внутри пакета (перенаправление).
def stamp_affine_uniformity_from_pages(v2_results: list[V2PageResult]) -> dict[str, Any]:
    return stamp_grid_mismatch_file_aggregate(v2_results)


def stamp_affine_detail_for_ui(v2_results: list[V2PageResult]) -> dict[str, Any]:
    return stamp_grid_mismatch_detail_for_ui(v2_results)
