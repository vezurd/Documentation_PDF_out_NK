"""Geometry helpers for OD manifest table clipping (engine-aligned, future path).

Callers must supply rectangles in the **same coordinate system** as
``fitz.Page.find_tables`` on the page (displayed MuPDF space for this project).
"""

from __future__ import annotations

from typing import TYPE_CHECKING, Any

import fitz

from pdf_parsing_v2_engine.find_tables_settings import call_find_tables, merge_cfg_for_find_tables

if TYPE_CHECKING:
    from pdf_parsing_v2_engine.models import StampTemplate


def build_manifest_clip_frame_minus_stamp(
    *,
    frame_rect: fitz.Rect,
    stamp_rect: fitz.Rect,
    page_rect: fitz.Rect,
    origin: str = "frame_bottom_right",
) -> fitz.Rect | None:
    """Return a clip inside the drawing frame with the stamp corner cleared.

    Heuristic for typical GOST layouts: the large manifest grid sits **above**
    a bottom-right stamp. Clip = intersection(frame, page) capped at ``stamp.y0``
    on the Y axis (full frame width). If that band is too thin, fall back to the
    band **left** of the stamp (capped at ``stamp.x0``).

    Args:
        frame_rect: Detected drawing frame (displayed pts).
        stamp_rect: Stamp effective bbox (displayed pts), e.g. from
            ``V2PageResult.metadata['stamp_effective_bbox']``.
        page_rect: Page size rect (usually ``fitz_page.rect``).
        origin: Template origin hint (reserved for finer corner logic).

    Returns:
        Non-empty clip rect or ``None`` if no safe region remains.
    """
    del origin  # Reserved for non-bottom-right stamp layouts.
    base = frame_rect & page_rect
    if base.is_empty or base.is_infinite:
        return None
    stamp = stamp_rect & page_rect
    if stamp.is_empty:
        return base if base.width > 1.0 and base.height > 1.0 else None

    y_top = min(base.y1, stamp.y0)
    above = fitz.Rect(base.x0, base.y0, base.x1, y_top)
    if above.width > 1.0 and above.height > 1.0:
        return above

    x_right = min(base.x1, stamp.x0)
    left = fitz.Rect(base.x0, base.y0, x_right, base.y1)
    if left.width > 1.0 and left.height > 1.0:
        return left
    return None


def find_tables_od_manifest_region(
    fitz_page: fitz.Page,
    clip: fitz.Rect,
    cfg: dict[str, Any] | None,
    template: StampTemplate | None = None,
):
    """Run ``find_tables`` on *clip* using merged snap settings (cf. engine)."""
    merged = merge_cfg_for_find_tables(cfg, template)
    return call_find_tables(fitz_page, cfg=merged, clip=clip)
