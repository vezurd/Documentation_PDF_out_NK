"""Context passed through stamp text cleaners and pipelines."""

from __future__ import annotations

from dataclasses import dataclass


@dataclass(frozen=True)
class StampTextContext:
    """Per-field extraction context for domain cleaners and pipelines.

    Attributes:
        doc_type: Document type code (e.g. ``MTO``, ``WIR``).
        page_num: 1-based page index (matches ``extract_page`` / v1 conventions).
        field_id: Optional field id for diagnostics (set by caller).
    """

    doc_type: str = ""
    page_num: int = 1
    field_id: str | None = None
