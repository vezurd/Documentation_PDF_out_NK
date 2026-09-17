"""Extract document / page properties via PyMuPDF (logic aligned with v1, standalone)."""

from __future__ import annotations

from typing import Any

import fitz

from pdf_parsing_v2_engine.document_properties.registry import (
    KEY_FILE_NAME,
    KEY_PAGE_ANNOTATIONS,
    KEY_PAGE_HEIGHT_MM,
    KEY_PAGE_LAYERS,
    KEY_PAGE_REAL_FORMAT,
    KEY_PAGE_WIDTH_MM,
    SCALE_POINTS_PER_MM,
)
from utils.formats_A import get_Real_Page_Format


def extract_page_layers_flag(fitz_doc: fitz.Document) -> list[int]:
    """Return ``[-1]`` if no optional-content UI layers, else ``[1]`` (v1 semantics)."""
    try:
        configs = fitz_doc.layer_ui_configs()
    except Exception:
        return [-1]
    if not configs:
        return [-1]
    return [1]


def extract_text_annotations_flag(fitz_page: fitz.Page) -> list[int]:
    """Return ``[-1]`` if no text annotations with non-empty info, else ``[1]``."""
    flag = [-1]
    try:
        for annot in fitz_page.annots(types=(fitz.PDF_ANNOT_TEXT,)):
            if annot and annot.info:
                flag = [1]
                break
    except Exception:
        pass
    return flag


def extract_page_width_height_mm(fitz_page: fitz.Page) -> tuple[list[int], list[int]]:
    """Displayed page size in mm (int), same formula as v1 pdfplumber path."""
    rect = fitz_page.rect
    w_mm = int(rect.width / SCALE_POINTS_PER_MM)
    h_mm = int(rect.height / SCALE_POINTS_PER_MM)
    return [w_mm], [h_mm]


def build_document_properties_metadata(
    fitz_page: fitz.Page,
    *,
    file_basename: str,
) -> dict[str, Any]:
    """Build the mandatory metadata map for one page.

    Args:
        fitz_page: Opened page.
        file_basename: PDF file name only (e.g. ``doc.MTO.pdf``).

    Returns:
        Dict with keys ``61_Page_Layers``, ``62_...``, ``63_...``, ``64_...``,
        ``65_Page_Real_Format``, ``file_name``. Lists use int elements where
        v1 / ``rules_check`` expect ints (61, 62, 63, 64).
    """
    fitz_doc = fitz_page.parent
    w_list, h_list = extract_page_width_height_mm(fitz_page)
    w0, h0 = w_list[0], h_list[0]
    real_fmt = get_Real_Page_Format(w0, h0)
    return {
        KEY_PAGE_LAYERS: extract_page_layers_flag(fitz_doc),
        KEY_PAGE_ANNOTATIONS: extract_text_annotations_flag(fitz_page),
        KEY_PAGE_WIDTH_MM: w_list,
        KEY_PAGE_HEIGHT_MM: h_list,
        KEY_PAGE_REAL_FORMAT: [real_fmt],
        KEY_FILE_NAME: [file_basename or ""],
    }
