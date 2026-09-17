"""Tests for document-type coverage against stamp templates."""

from __future__ import annotations

from types import SimpleNamespace

from pdf_parsing_v2.v2_pipeline import format_uncovered_doc_types_warning
from pdf_parsing_v2_engine.doc_types import (
    LIST_OF_DWG_DOC_TYPES,
    uncovered_doc_type_files,
)


def test_ni_is_a_dwg_doc_type() -> None:
    assert "NI" in LIST_OF_DWG_DOC_TYPES


def test_uncovered_doc_type_files_skips_covered_and_flags_unknown() -> None:
    templates = [SimpleNamespace(doc_types=["WIR", "LAY", "NI"])]
    rows = [
        ("a.pdf", "WIR"),
        ("b.pdf", "XYZ"),
        ("c.pdf", "NI"),
        ("d.pdf", ""),
    ]
    out = uncovered_doc_type_files(rows, templates)
    assert out == [("b.pdf", "XYZ"), ("d.pdf", "?")]


def test_agcc_287_dwg_templates_include_ni() -> None:
    from pathlib import Path

    from pdf_parsing_v2_engine.models import StampTemplate

    root = Path(__file__).resolve().parents[1] / "pdf_parsing_v2_engine" / "templates" / "agcc_287"
    for name in ("dwg_page1.json", "dwg_page_rest.json"):
        tmpl = StampTemplate.from_json(str(root / name))
        assert "NI" in tmpl.doc_types


def test_format_uncovered_doc_types_warning_lists_types() -> None:
    text = format_uncovered_doc_types_warning(
        {
            "uncovered_doc_types": ["XYZ"],
            "uncovered_n_files": 1,
            "uncovered_files": ["b.pdf (XYZ)"],
        }
    )
    assert "XYZ" in text
    assert "b.pdf (XYZ)" in text
    assert "Нет шаблона" in text
