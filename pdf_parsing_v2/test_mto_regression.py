"""MTO page-1 batch probe: smoke test when sample PDFs exist; strict helper unit test."""

from __future__ import annotations

from pathlib import Path

import pytest

_PKG = Path(__file__).resolve().parent
_DEFAULT_MTO = _PKG / "templates" / "test_pdf" / "MTO"
_DEFAULT_TPL = _PKG / "templates" / "agcc_287" / "mto_page1.json"

_have_mto_pdf = _DEFAULT_MTO.is_dir() and any(_DEFAULT_MTO.glob("*.pdf"))
_have_template = _DEFAULT_TPL.is_file()


@pytest.mark.skipif(not _have_mto_pdf, reason="No PDFs in templates/test_pdf/MTO")
@pytest.mark.skipif(not _have_template, reason="mto_page1.json not available")
def test_mto_page1_batch_probe_smoke() -> None:
    from pdf_parsing_v2_engine.tools.mto_page1_batch_probe import run_mto_page1_batch_probe

    report = run_mto_page1_batch_probe(
        _DEFAULT_MTO,
        str(_DEFAULT_TPL),
        v2_cfg=None,
        print_report=False,
    )
    assert "error" not in report, report.get("error")
    assert report["rows"]
    for row in report["rows"]:
        assert "error" not in row, row


def test_check_mto_regression_strict_flags_excess_newlines() -> None:
    from pdf_parsing_v2_engine.tools.mto_page1_batch_probe import check_mto_regression_strict

    report = {
        "rows": [
            {
                "file": "fake.pdf",
                "rotation": 90,
                "newline_count_char_center": 500,
                "score_char_center": 1.0,
                "score_get_textbox": 1.0,
            },
        ],
    }
    reasons = check_mto_regression_strict(report, max_newlines_rot90=280)
    assert reasons and "newline_count_char_center=500" in reasons[0]


def test_check_mto_regression_strict_passes_clean_row() -> None:
    from pdf_parsing_v2_engine.tools.mto_page1_batch_probe import check_mto_regression_strict

    report = {
        "rows": [
            {
                "file": "ok.pdf",
                "rotation": 90,
                "newline_count_char_center": 50,
                "score_char_center": 1.0,
                "score_get_textbox": 1.0,
            },
        ],
    }
    assert check_mto_regression_strict(report, max_newlines_rot90=280) == []
