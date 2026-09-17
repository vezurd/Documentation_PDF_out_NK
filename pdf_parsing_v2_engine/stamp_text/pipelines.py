"""Cleaner pipelines: map ``clean`` keys to :class:`CleanOutcome`.

Each key maps to :func:`stamp_field_cleaners.pipeline_*` (``CleanOutcome``).
``doc_title`` uses strict→relaxed with :class:`ParseWarning` (``tier=\"relaxed\"``).
External ``FieldResult.cleaned_value`` stays ``str``.
"""

from __future__ import annotations

from collections.abc import Callable

from pdf_parsing_v2_engine.stamp_text import stamp_field_cleaners as sf
from pdf_parsing_v2_engine.stamp_text.context import StampTextContext
from pdf_parsing_v2_engine.stamp_text.outcome import CleanOutcome, ParseWarning
# Re-export for stamp_extractor / tests
UNKNOWN_CLEANER_CODE = "UNKNOWN_CLEANER"

CleanerPipeline = Callable[[str, StampTextContext], CleanOutcome]


def run_field_clean(clean_key: str, text: str, ctx: StampTextContext) -> CleanOutcome:
    """Run the registered pipeline for *clean_key* (must be a known key)."""
    fn = FIELD_CLEAN_PIPELINES[clean_key]
    return fn(text, ctx)


# Keys mirror ``field_cleaners.CLEANERS`` / template JSON ``clean`` values.
FIELD_CLEAN_PIPELINES: dict[str, CleanerPipeline] = {
    "doc_title": sf.pipeline_doc_title,
    "facility_name": sf.pipeline_facility_name,
    "unit_title": sf.pipeline_unit_title,
    "document_name": sf.pipeline_document_name,
    "documentation_type": sf.pipeline_documentation_type,
    "sheet_number_6_1": sf.pipeline_sheet_number_6_1,
    "sheet_number_6_2": sf.pipeline_sheet_number_6_2,
    "total_number_of_sheets": sf.pipeline_total_number_of_sheets,
    "revision": sf.pipeline_revision,
    "file_name_stamp": sf.pipeline_file_name_stamp,
    "page_format": sf.pipeline_page_format,
    "current_revision_18_1": sf.pipeline_current_revision_18_1,
    "revision_date_18_2": sf.pipeline_revision_date_18_2,
    "purpose_of_issue_18_3": sf.pipeline_purpose_of_issue_18_3,
    "signatures_date_10": sf.pipeline_signatures_date_10,
}


def unknown_cleaner_outcome(clean_key: str, raw: str) -> CleanOutcome:
    """Outcome when ``clean`` is set but not in the registry."""
    return CleanOutcome(
        value=raw,
        raw_input=raw,
        warnings=[
            ParseWarning(
                code=UNKNOWN_CLEANER_CODE,
                message=f"Unknown clean key {clean_key!r} (not in FIELD_CLEAN_PIPELINES)",
            )
        ],
        via="unknown",
        clean_tier=None,
    )


def cleaner_exception_outcome(raw: str, exc: BaseException) -> CleanOutcome:
    """Fallback when a cleaner raises (legacy: use raw text)."""
    return CleanOutcome(
        value=raw,
        raw_input=raw,
        warnings=[
            ParseWarning(
                code="CLEANER_EXCEPTION",
                message=f"{type(exc).__name__}: {exc}",
            )
        ],
        via="exception",
        clean_tier=None,
    )
