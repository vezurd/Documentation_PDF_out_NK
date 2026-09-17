"""
E.2: PNG-оверлей — рамка, bbox полей, подписи, цветовая индикация.

Два режима:
1. **Аннотация результатов** (``render_debug_page``) — зелёная рамка, цветные bbox полей,
   подписи ``field_id = cleaned_value``, score и warnings в углу.
2. **Overlay** (``render_debug_page_overlay``) — bbox ВСЕХ шаблонов-кандидатов, score каждого.

Интеграция: вызывается из ``v2_pipeline`` при ``debug_visual=True`` в конфиге.
"""

from __future__ import annotations

import os
from typing import TYPE_CHECKING

import fitz

from pdf_parsing_v2_engine.coord_transform import fitz_unrotated_to_displayed, pdfminer_to_fitz
from pdf_parsing_v2_engine.models import FieldResult, FrameInfo, StampTemplate, V2PageResult

if TYPE_CHECKING:
    pass

_DPI = 150
_SCALE = _DPI / 72.0

_GREEN = (0, 0.7, 0)
_BLUE = (0.1, 0.3, 0.9)
_RED = (0.9, 0.1, 0.1)
_YELLOW = (0.85, 0.75, 0)
_GREY = (0.5, 0.5, 0.5)
_ORANGE = (1.0, 0.55, 0)
_WHITE = (1, 1, 1)
_BLACK = (0, 0, 0)


def _frame_to_fitz_rect(frame: FrameInfo) -> fitz.Rect:
    """Convert FrameInfo (pdfminer coords) → fitz.Rect (displayed coords)."""
    return pdfminer_to_fitz(
        (frame.x0, frame.y0, frame.x1, frame.y1),
        frame.page_height,
        frame.rotation,
    )


def _field_color(fr: FieldResult, expected: str) -> tuple[float, float, float]:
    has_value = bool(fr.cleaned_value and fr.cleaned_value.strip())
    if expected == "required":
        if not has_value:
            return _RED
        if fr.is_valid is False:
            return _RED
        return _GREEN
    if expected == "absent":
        if has_value:
            return _ORANGE
        return _GREY
    if has_value:
        return _BLUE
    return _YELLOW


def _draw_rect(
    page: fitz.Page,
    rect: fitz.Rect,
    color: tuple[float, float, float],
    width: float = 1.5,
    dashes: str | None = None,
) -> None:
    annot_rect = fitz.Rect(
        rect.x0 * _SCALE, rect.y0 * _SCALE,
        rect.x1 * _SCALE, rect.y1 * _SCALE,
    )
    if annot_rect.is_empty or annot_rect.is_infinite:
        return
    shape = page.new_shape()
    shape.draw_rect(annot_rect)
    shape.finish(color=color, width=width, dashes=dashes)
    shape.commit()


def _insert_text(
    page: fitz.Page,
    point: fitz.Point,
    text: str,
    fontsize: float = 7,
    color: tuple[float, float, float] = _BLACK,
    bg_color: tuple[float, float, float] | None = None,
) -> None:
    if not text:
        return
    scaled_pt = fitz.Point(point.x * _SCALE, point.y * _SCALE)
    if bg_color:
        tw = fitz.TextWriter(page.rect)
        tw.append(scaled_pt, text, fontsize=fontsize)
        bg_rect = tw.text_rect
        bg_rect = bg_rect + (-1, -1, 1, 1)
        shape = page.new_shape()
        shape.draw_rect(bg_rect)
        shape.finish(color=None, fill=bg_color)
        shape.commit()
        tw.write_text(page, color=color)
    else:
        page.insert_text(scaled_pt, text, fontsize=fontsize, color=color)


_CYAN = (0.0, 0.65, 0.75)
_CYAN_TEXT = (0.0, 0.45, 0.55)


def _draw_pdf_text_spans(
    fitz_page: fitz.Page,
    tmp_page: fitz.Page,
    frame: FrameInfo | None = None,
) -> int:
    """Draw all text spans from fitz_page onto tmp_page (already pixel-scaled).

    Returns the number of spans drawn.
    Uses ``get_text("dict")`` to obtain span ``text`` and ``bbox``.

    For pages with rotation=90/270 PyMuPDF returns span bboxes in the
    *unrotated* PDF coordinate space.  If *frame* is provided, each span bbox
    is converted to *displayed* (fitz_page.rect) coordinates via
    ``fitz_unrotated_to_displayed(..., fitz_page)`` before drawing.

    Each span gets a cyan outline + its text content in tiny font above it.
    """
    text_dict = fitz_page.get_text("dict")
    count = 0
    for block in text_dict.get("blocks", []):
        if block.get("type") != 0:  # skip image blocks
            continue
        for line in block.get("lines", []):
            for span in line.get("spans", []):
                span_bbox_raw = fitz.Rect(span["bbox"])
                if span_bbox_raw.is_empty or span_bbox_raw.is_infinite:
                    continue
                # Convert unrotated → displayed for rotation 90/270
                if frame is not None:
                    span_bbox = fitz_unrotated_to_displayed(span_bbox_raw, fitz_page)
                else:
                    span_bbox = span_bbox_raw
                if span_bbox.is_empty or span_bbox.is_infinite:
                    continue
                _draw_rect_on_pixpage(tmp_page, span_bbox, _CYAN, width=0.8)
                text = span.get("text", "").strip()
                if text:
                    label_pt = fitz.Point(span_bbox.x0, span_bbox.y0 - 1)
                    _insert_text_on_pixpage(
                        tmp_page, label_pt, text[:40], fontsize=5, color=_CYAN_TEXT
                    )
                count += 1
    return count


def render_debug_page(
    fitz_page: fitz.Page,
    v2_result: V2PageResult,
    output_path: str,
    field_expected_map: dict[str, str] | None = None,
    show_text_spans: bool = False,
) -> None:
    """Render annotated PNG for a single page.

    Parameters
    ----------
    fitz_page :
        Original PDF page (used for background rendering).
    v2_result :
        Extraction result for this page.
    output_path :
        Where to write the PNG.
    field_expected_map :
        ``{field_id: expected}`` — if not given, all fields treated as ``"optional"``.
    show_text_spans :
        If True, draw all PDF text spans (cyan bboxes + text) BEFORE field bboxes.
        Useful to diagnose coordinate mismatches between template and actual text.
    """
    if field_expected_map is None:
        field_expected_map = {}

    pix = fitz_page.get_pixmap(dpi=_DPI)
    tmp_pdf = fitz.open()
    tmp_page = tmp_pdf.new_page(width=pix.width, height=pix.height)
    tmp_page.insert_image(tmp_page.rect, pixmap=pix)

    # Layer 1 (optional): cyan text spans from PDF structure
    span_count = 0
    if show_text_spans:
        span_count = _draw_pdf_text_spans(fitz_page, tmp_page, frame=v2_result.frame)

    # Layer 2: frame (green outline)
    frame_rect = _frame_to_fitz_rect(v2_result.frame)
    _draw_rect_on_pixpage(tmp_page, frame_rect, _GREEN, width=2.0)

    # Layer 3: template field bboxes
    for field_id, fr in v2_result.fields.items():
        expected = field_expected_map.get(field_id, "optional")
        color = _field_color(fr, expected)
        fr_rect = fitz.Rect(*fr.bbox_pts)

        _draw_rect_on_pixpage(tmp_page, fr_rect, color, width=1.2)

        cleaned_short = (fr.cleaned_value or "")[:50]
        label = f"{field_id} = {cleaned_short}"
        label_pt = fitz.Point(fr_rect.x0, fr_rect.y0 - 2)
        _insert_text_on_pixpage(tmp_page, label_pt, label, fontsize=6, color=color)

    info_lines = [
        f"Template: {v2_result.template_name}",
        f"Score: {v2_result.template_score:.2f}",
    ]
    if show_text_spans:
        info_lines.append(f"Cyan = PDF text spans ({span_count} spans) | Colors = template fields")
    for w in v2_result.warnings[:5]:
        info_lines.append(f"  ! {w[:80]}")

    y_offset = 10.0
    for line in info_lines:
        tmp_page.insert_text(
            fitz.Point(10, y_offset + 8),
            line,
            fontsize=8,
            color=_BLACK,
        )
        y_offset += 12.0

    os.makedirs(os.path.dirname(os.path.abspath(output_path)) or ".", exist_ok=True)
    final_pix = tmp_page.get_pixmap()
    final_pix.save(output_path)
    tmp_pdf.close()


def _draw_rect_on_pixpage(
    page: fitz.Page,
    rect_fitz: fitz.Rect,
    color: tuple[float, float, float],
    width: float = 1.5,
    dashes: str | None = None,
) -> None:
    """Draw a rectangle on a page whose size is already in pixel-space (DPI-scaled)."""
    annot_rect = fitz.Rect(
        rect_fitz.x0 * _SCALE,
        rect_fitz.y0 * _SCALE,
        rect_fitz.x1 * _SCALE,
        rect_fitz.y1 * _SCALE,
    )
    if annot_rect.is_empty or annot_rect.is_infinite:
        return
    shape = page.new_shape()
    shape.draw_rect(annot_rect)
    shape.finish(color=color, width=width, dashes=dashes)
    shape.commit()


def _insert_text_on_pixpage(
    page: fitz.Page,
    point_fitz: fitz.Point,
    text: str,
    fontsize: float = 7,
    color: tuple[float, float, float] = _BLACK,
) -> None:
    if not text:
        return
    page.insert_text(
        fitz.Point(point_fitz.x * _SCALE, point_fitz.y * _SCALE),
        text,
        fontsize=fontsize,
        color=color,
    )


# ---------------------------------------------------------------------------
# Overlay mode — bbox ВСЕХ шаблонов-кандидатов
# ---------------------------------------------------------------------------

def render_debug_page_overlay(
    fitz_page: fitz.Page,
    v2_result: V2PageResult,
    all_candidates: list[tuple[StampTemplate, float]],
    output_path: str,
) -> None:
    """Render overlay PNG showing bbox from ALL candidate templates.

    Parameters
    ----------
    all_candidates :
        ``[(template, score), ...]`` — all templates tried for this page,
        sorted by score desc (winner first).
    """
    pix = fitz_page.get_pixmap(dpi=_DPI)
    tmp_pdf = fitz.open()
    tmp_page = tmp_pdf.new_page(width=pix.width, height=pix.height)
    tmp_page.insert_image(tmp_page.rect, pixmap=pix)

    frame_rect = _frame_to_fitz_rect(v2_result.frame)
    _draw_rect_on_pixpage(tmp_page, frame_rect, _GREEN, width=2.0)

    for rank, (tmpl, score) in enumerate(all_candidates):
        is_winner = rank == 0
        color = _BLUE if is_winner else _GREY
        line_width = 1.5 if is_winner else 0.8

        from pdf_parsing_v2_engine.coord_transform import field_to_fitz_rect

        for fd in tmpl.fields:
            eff_pad = fd.padding_mm if fd.padding_mm is not None else tmpl.padding_mm
            fr = field_to_fitz_rect(fd, v2_result.frame, padding_mm=eff_pad)
            _draw_rect_on_pixpage(tmp_page, fr, color, width=line_width)

        label_y = 10.0 + rank * 14.0
        prefix = ">> " if is_winner else "   "
        tmp_page.insert_text(
            fitz.Point(10, label_y + 8),
            f"{prefix}{tmpl.name} — score: {score:.2f}",
            fontsize=8,
            color=color,
        )

    os.makedirs(os.path.dirname(os.path.abspath(output_path)) or ".", exist_ok=True)
    final_pix = tmp_page.get_pixmap()
    final_pix.save(output_path)
    tmp_pdf.close()


# ---------------------------------------------------------------------------
# Pipeline integration helper
# ---------------------------------------------------------------------------

def render_debug_for_file(
    file_path: str,
    v2_results: list[V2PageResult],
    templates: list[StampTemplate],
    output_dir: str,
    overlay: bool = False,
) -> list[str]:
    """Render debug PNGs for all pages of one file.

    Returns list of created PNG paths.
    """
    import fitz as fitz_lib

    base_name = os.path.splitext(os.path.basename(file_path))[0]
    created: list[str] = []
    fitz_doc = fitz_lib.open(file_path)

    try:
        for v2r in v2_results:
            page_idx = v2r.page_num - 1
            if page_idx < 0 or page_idx >= len(fitz_doc):
                continue
            fitz_page = fitz_doc[page_idx]

            png_name = f"{base_name}_p{v2r.page_num}.png"
            png_path = os.path.join(output_dir, png_name)

            field_expected: dict[str, str] = {}
            from pdf_parsing_v2_engine.stamp_extractor import select_templates
            candidates = select_templates(templates, v2r.doc_type, v2r.page_num)
            for tmpl in candidates:
                if tmpl.name == v2r.template_name:
                    field_expected = {fd.id: fd.expected for fd in tmpl.fields}
                    break

            render_debug_page(
                fitz_page, v2r, png_path,
                field_expected_map=field_expected,
                show_text_spans=True,
            )
            created.append(png_path)

            if overlay or v2r.template_score < 0.7:
                from pdf_parsing_v2_engine.stamp_extractor import _score_template, _extract_with_template
                overlay_name = f"{base_name}_p{v2r.page_num}_overlay.png"
                overlay_path = os.path.join(output_dir, overlay_name)
                all_cands: list[tuple[StampTemplate, float]] = []
                for tmpl in candidates:
                    fields, _stamp_meta = _extract_with_template(
                        fitz_page, tmpl, v2r.frame, v2r.doc_type, v2r.page_num,
                    )
                    ts = _score_template(tmpl, fields)
                    all_cands.append((tmpl, ts.score))
                all_cands.sort(key=lambda x: x[1], reverse=True)
                render_debug_page_overlay(fitz_page, v2r, all_cands, overlay_path)
                created.append(overlay_path)
    finally:
        fitz_doc.close()

    return created
