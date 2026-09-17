"""Document type classification lists (v2, detached from utils.string_parsing)."""

from __future__ import annotations

from collections.abc import Iterable
from typing import Any

LIST_OF_BBB: list[str] = ["BOM", "BOE", "BOQ", "MTO"]
LIST_OF_TEXT_DOC_TYPES: list[str] = ["OD", "CJ", "VO"]
LIST_OF_DWG_DOC_TYPES: list[str] = ["WIR", "LAY", "CAE", "GA", "PL", "DW", "NI"]


def template_covered_doc_types(templates: Iterable[Any]) -> set[str]:
    """Return the union of ``doc_types`` declared on loaded stamp templates."""
    covered: set[str] = set()
    for tmpl in templates:
        for doc_type in getattr(tmpl, "doc_types", None) or []:
            text = str(doc_type).strip()
            if text:
                covered.add(text)
    return covered


def uncovered_doc_type_files(
    file_doc_types: Iterable[tuple[str, str]],
    templates: Iterable[Any],
) -> list[tuple[str, str]]:
    """Return ``(file_name, doc_type)`` pairs not covered by any template.

    Empty / unparsed document types are reported as ``"?"``.
    """
    covered = template_covered_doc_types(templates)
    uncovered: list[tuple[str, str]] = []
    for file_name, doc_type in file_doc_types:
        parsed = str(doc_type or "").strip()
        if parsed not in covered:
            uncovered.append((str(file_name), parsed or "?"))
    return uncovered
