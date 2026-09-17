"""
Registry of field ``clean`` keys → string cleaners for v2 stamp extraction.

Domain logic lives in ``pdf_parsing_v2.stamp_text``; this module exposes
``CLEANERS`` (str API) and ``CLEANER_DESCRIPTIONS`` (editor tooltips).
"""

from __future__ import annotations

from typing import Callable

from pdf_parsing_v2_engine.stamp_text.context import StampTextContext
from pdf_parsing_v2_engine.stamp_text.pipelines import FIELD_CLEAN_PIPELINES, run_field_clean

CleanerFn = Callable[..., str]


def _make_str_cleaner(key: str) -> CleanerFn:
    def fn(text: str, doc_type: str = "", page_num: int = 1) -> str:
        return run_field_clean(key, text, StampTextContext(doc_type, page_num)).value

    return fn


CLEANERS: dict[str, CleanerFn] = {k: _make_str_cleaner(k) for k in FIELD_CLEAN_PIPELINES}

# Editor tooltips (properties panel Clean combo).
CLEANER_DESCRIPTIONS: dict[str, str] = {
    "doc_title": "AGCC→first newline, junk-filter; if slice looks over-inclusive, relaxed cut + DOC_TITLE_RELAXED warning.",
    "facility_name": "Text after first AGCC line; junk filter with sheet labels removed.",
    "unit_title": "Replace newlines with spaces (minimal).",
    "document_name": "Remove date tokens (DD.MM.YY), join lines with spaces.",
    "documentation_type": "Strip newlines, junk filter (default labels).",
    "sheet_number_6_1": "BBB: page 1 — digit regex; page 2+ — «л. X из Y листов» → X, else legacy л./из.",
    "sheet_number_6_2": "BBB: page 1 — на…л.; page 2+ — «л. X из Y листов» → Y, else legacy из…листов.",
    "total_number_of_sheets": "Strip newlines, junk filter including Лист/Листов.",
    "revision": "Junk filter; first token of first remaining line.",
    "file_name_stamp": "First junk-cleaned line from AGCC; duplicate AGCC → warning text.",
    "page_format": "Junk filter; prefer line containing «Формат».",
    "current_revision_18_1": "18-col revision code: junk, flatten words, drop IFC/IFR/IFU-like tokens.",
    "revision_date_18_2": "Date in DD.MM.YYYY from junk-cleaned block (18-col table).",
    "purpose_of_issue_18_3": "Text after first numeric triple in junk-cleaned block; first comma segment.",
    "signatures_date_10": "DD.MM.YY date from junk-cleaned block (column 10 style).",
}
