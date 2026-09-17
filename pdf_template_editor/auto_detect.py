"""fitz.find_tables → bbox ячеек для начальной авто-разметки."""

from __future__ import annotations

from typing import TYPE_CHECKING

import fitz

from pdf_parsing_v2_engine.find_tables_settings import call_find_tables, merge_cfg_for_find_tables

if TYPE_CHECKING:
    from pdf_parsing_v2_engine.models import StampTemplate

_MIN_CELL_PTS = 5.0
_MAX_CELL_FRACTION = 0.80


def detect_cells(
    fitz_page: fitz.Page,
    cfg: dict | None = None,
    template: StampTemplate | None = None,
) -> list[tuple[float, float, float, float]]:
    """Find table cells on *fitz_page* via ``find_tables()`` (PyMuPDF >= 1.23).

    Returns list of ``(x0, y0, x1, y1)`` in **fitz displayed coords** (pts).
    Filters out cells smaller than *_MIN_CELL_PTS* or larger than 80 % of page.

    Args:
        fitz_page: Page to scan.
        cfg: v2 config dict (``find_tables_snap_*`` defaults).
        template: When set, non-None ``find_tables_snap_*`` on the template override *cfg*.
    """
    page_w = fitz_page.rect.width
    page_h = fitz_page.rect.height
    max_w = page_w * _MAX_CELL_FRACTION
    max_h = page_h * _MAX_CELL_FRACTION

    result: list[tuple[float, float, float, float]] = []

    try:
        tables = call_find_tables(
            fitz_page, cfg=merge_cfg_for_find_tables(cfg, template)
        )
    except Exception:
        return result

    for table in tables:
        for cell in table.cells:
            x0, y0, x1, y1 = cell
            w = x1 - x0
            h = y1 - y0
            if w < _MIN_CELL_PTS or h < _MIN_CELL_PTS:
                continue
            if w > max_w and h > max_h:
                continue
            result.append((x0, y0, x1, y1))

    seen: set[tuple[int, int, int, int]] = set()
    deduped: list[tuple[float, float, float, float]] = []
    for bbox in result:
        key = (round(bbox[0]), round(bbox[1]), round(bbox[2]), round(bbox[3]))
        if key not in seen:
            seen.add(key)
            deduped.append(bbox)
    return deduped
