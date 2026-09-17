"""PDF document / page properties for v2 (independent of v1 stamp extraction)."""

from __future__ import annotations

from pdf_parsing_v2_engine.document_properties.extractors import (
    build_document_properties_metadata,
    extract_page_layers_flag,
    extract_text_annotations_flag,
)
from pdf_parsing_v2_engine.document_properties.registry import (
    KEY_FILE_NAME,
    KEY_PAGE_ANNOTATIONS,
    KEY_PAGE_HEIGHT_MM,
    KEY_PAGE_LAYERS,
    KEY_PAGE_REAL_FORMAT,
    KEY_PAGE_WIDTH_MM,
    MANDATORY_PROPERTY_KEYS,
    PROPERTY_LABELS,
    RESERVED_FIELD_IDS,
    SCALE_POINTS_PER_MM,
)
from pdf_parsing_v2_engine.document_properties.template_defaults import (
    default_document_property_field_def,
    ensure_document_property_field_defs,
)

__all__ = [
    "KEY_FILE_NAME",
    "KEY_PAGE_ANNOTATIONS",
    "KEY_PAGE_HEIGHT_MM",
    "KEY_PAGE_LAYERS",
    "KEY_PAGE_REAL_FORMAT",
    "KEY_PAGE_WIDTH_MM",
    "MANDATORY_PROPERTY_KEYS",
    "PROPERTY_LABELS",
    "RESERVED_FIELD_IDS",
    "SCALE_POINTS_PER_MM",
    "build_document_properties_metadata",
    "default_document_property_field_def",
    "ensure_document_property_field_defs",
    "extract_page_layers_flag",
    "extract_text_annotations_flag",
]
