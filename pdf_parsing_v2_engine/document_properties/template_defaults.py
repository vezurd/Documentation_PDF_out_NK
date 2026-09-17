"""Default ``FieldDef`` rows for document-property fields in templates."""

from __future__ import annotations

from dataclasses import replace

from pdf_parsing_v2_engine.document_properties.registry import (
    MANDATORY_PROPERTY_KEYS,
    PROPERTY_LABELS,
)
from pdf_parsing_v2_engine.models import FieldDef, StampTemplate


def default_document_property_field_def(property_key: str) -> FieldDef:
    """Hidden template field: no stamp text extraction; value from document properties."""
    return FieldDef(
        id=property_key,
        label=PROPERTY_LABELS.get(property_key, property_key),
        bbox_mm=(0.0, 0.0, 0.0, 0.0),
        clean=None,
        validate_regex=None,
        expected="optional",
        padding_mm=None,
        field_type="data",
        expected_text=None,
        is_anchor=False,
        outside_stamp=False,
        origin="frame_bottom_right",
        stretch_to_page=(),
        bound_top=None,
        bound_bottom=None,
        bound_left=None,
        bound_right=None,
        document_property=property_key,
    )


def ensure_document_property_field_defs(template: StampTemplate) -> StampTemplate:
    """Append missing mandatory document-property fields (in registry order)."""
    have = {f.id for f in template.fields}
    extra: list[FieldDef] = []
    for key in MANDATORY_PROPERTY_KEYS:
        if key not in have:
            extra.append(default_document_property_field_def(key))
    if not extra:
        return template
    return replace(template, fields=list(template.fields) + extra)
