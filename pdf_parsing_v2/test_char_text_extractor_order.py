"""Reading-order fix for rotation 90/270 (displayed-space sort in char_text_extractor)."""

from __future__ import annotations

import fitz

from pdf_parsing_v2_engine.char_text_extractor import PageCharIndex, _chars_to_text
from pdf_parsing_v2_engine.models import FrameInfo


def _dummy_frame(*, rotation: int, page_width: float, page_height: float) -> FrameInfo:
    return FrameInfo(
        x0=0.0,
        y0=0.0,
        x1=page_width,
        y1=page_height,
        page_width=page_width,
        page_height=page_height,
        rotation=rotation,
        border_left_mm=0.0,
        border_bottom_mm=0.0,
        border_top_mm=0.0,
    )


def _page_rotated_90() -> tuple[fitz.Document, fitz.Page]:
    doc = fitz.open()
    page = doc.new_page(width=1000, height=800)
    page.set_rotation(90)
    return doc, page


def test_displayed_order_90_reverses_vs_unrotated_strip() -> None:
    """One on-screen row: unrotated y-spacing breaks legacy line clustering; displayed fixes LTR."""
    ch_a = {"c": "A", "x0": 98.0, "y0": 198.0, "x1": 102.0, "y1": 202.0}
    ch_b = {"c": "B", "x0": 98.0, "y0": 218.0, "x1": 102.0, "y1": 222.0}
    chars = [ch_a, ch_b]
    doc, page = _page_rotated_90()
    try:
        legacy = _chars_to_text(list(chars))
        fixed = _chars_to_text(list(chars), fitz_page=page)

        assert legacy == "A\nB", f"expected legacy two lines, got {legacy!r}"
        assert fixed == "BA", f"expected displayed LTR BA, got {fixed!r}"
    finally:
        doc.close()


def test_rotation_0_unchanged_path() -> None:
    ch_a = {"c": "A", "x0": 10.0, "y0": 100.0, "x1": 20.0, "y1": 110.0}
    ch_b = {"c": "B", "x0": 30.0, "y0": 100.0, "x1": 40.0, "y1": 110.0}
    chars = [ch_b, ch_a]
    doc = fitz.open()
    page = doc.new_page(width=600, height=800)
    try:
        frame = _dummy_frame(rotation=0, page_width=600, page_height=800)

        no_frame = _chars_to_text(chars)
        with_page = _chars_to_text(chars, fitz_page=page)

        assert no_frame == "AB"
        assert with_page == "AB"
    finally:
        doc.close()


def test_page_char_index_extract_text_passes_fitz_page() -> None:
    """Smoke: extract_text(..., fitz_page) runs without error on synthetic index."""
    doc, page = _page_rotated_90()
    try:
        r = page.rect
        frame = _dummy_frame(rotation=90, page_width=r.width, page_height=r.height)
        ch = {"c": "X", "x0": 0.0, "y0": 0.0, "x1": 10.0, "y1": 10.0}
        idx = PageCharIndex([ch])
        rect = fitz.Rect(-1, -1, 20, 20)
        s = idx.extract_text(rect, frame=frame, fitz_page=page)
        assert s == "X"
    finally:
        doc.close()
