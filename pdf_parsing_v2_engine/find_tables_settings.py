"""Shared settings for PyMuPDF find_tables() in v2."""

from __future__ import annotations

from typing import TYPE_CHECKING, Any

import fitz

if TYPE_CHECKING:
    from pdf_parsing_v2_engine.models import StampTemplate


_DEFAULT_FIND_TABLES_CLIP_SCALE = 1.5


def merge_cfg_for_find_tables(
    cfg: dict[str, Any] | None,
    template: StampTemplate | None = None,
) -> dict[str, Any]:
    """Shallow merge: template non-None snap fields override *cfg* for find_tables."""
    out: dict[str, Any] = dict(cfg or {})
    if template is None:
        return out
    if template.find_tables_snap_x_tolerance is not None:
        out["find_tables_snap_x_tolerance"] = float(template.find_tables_snap_x_tolerance)
    if template.find_tables_snap_y_tolerance is not None:
        out["find_tables_snap_y_tolerance"] = float(template.find_tables_snap_y_tolerance)
    return out


def get_find_tables_kwargs(cfg: dict[str, Any] | None) -> dict[str, float]:
    """Build find_tables kwargs from v2 config."""
    src = cfg or {}
    x_tol = float(src.get("find_tables_snap_x_tolerance", 2.2))
    y_tol = float(src.get("find_tables_snap_y_tolerance", 2.0))
    return {
        "snap_x_tolerance": x_tol,
        "snap_y_tolerance": y_tol,
    }


def build_find_tables_clip(
    fitz_page: fitz.Page,
    stamp_bbox: fitz.Rect | None,
    scale: float = _DEFAULT_FIND_TABLES_CLIP_SCALE,
) -> fitz.Rect | None:
    """Build a page-clamped clip rect around the stamp area."""
    if stamp_bbox is None:
        return None
    clip = fitz.Rect(stamp_bbox)
    if clip.is_empty or clip.is_infinite:
        return None
    if scale <= 0:
        raise ValueError("scale must be > 0")

    page_rect = fitz.Rect(0.0, 0.0, fitz_page.rect.width, fitz_page.rect.height)
    if page_rect.is_empty:
        return None

    cx = (clip.x0 + clip.x1) * 0.5
    cy = (clip.y0 + clip.y1) * 0.5
    half_w = clip.width * scale * 0.5
    half_h = clip.height * scale * 0.5
    expanded = fitz.Rect(cx - half_w, cy - half_h, cx + half_w, cy + half_h) & page_rect
    if expanded.is_empty:
        return None
    return expanded


def call_find_tables(
    fitz_page,
    cfg: dict[str, Any] | None = None,
    clip: fitz.Rect | None = None,
):
    """Call fitz_page.find_tables with config kwargs and safe fallback."""
    kwargs = get_find_tables_kwargs(cfg)
    if clip is not None:
        kwargs["clip"] = clip
    try:
        return fitz_page.find_tables(**kwargs)
    except TypeError:
        # Older/different PyMuPDF builds may not support x/y kwargs.
        fallback_kwargs = dict(kwargs)
        fallback_kwargs.pop("snap_x_tolerance", None)
        fallback_kwargs.pop("snap_y_tolerance", None)
        try:
            return fitz_page.find_tables(**fallback_kwargs)
        except TypeError:
            return fitz_page.find_tables()

