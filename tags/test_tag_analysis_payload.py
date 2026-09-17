"""XT and other hide_tag equipment must still appear in tag analysis reports."""

from __future__ import annotations

from tags.tag_analysis_payload import (
    build_tag_analysis_payload,
    render_tag_analysis_text_from_payload,
)
from tags.tag_parser import diff_lists_MTO_B


_MTO = "MTO-0001"
_FV_MATCH = "2612-PB-01-S-FV-1511"
_XT_ON_SHEET = "2612-PB-01-S-XT-9999"
_XT_IN_MTO = "2612-PB-01-S-XT-1512"


def _source_with_xt() -> dict:
    return {
        ("mto.pdf", 1, _MTO): [_FV_MATCH, _XT_IN_MTO],
        ("sheet.pdf", 1, "CS-0001"): [_FV_MATCH, _XT_ON_SHEET],
    }


def test_diff_lists_includes_xt_missing_from_sheets():
    missing = diff_lists_MTO_B([_FV_MATCH, _XT_ON_SHEET], [_FV_MATCH, _XT_IN_MTO])
    assert _XT_IN_MTO in missing
    assert _FV_MATCH not in missing


def test_payload_extra_and_missing_include_xt():
    payload = build_tag_analysis_payload(_source_with_xt(), main_doc_signature=_MTO)
    extra = next(c for c in payload["checks"] if c["id"] == "extra_in_sheets")
    extra_tags = [r["tag"] for r in extra["rows"]]
    assert _XT_ON_SHEET in extra_tags
    xt_row = next(r for r in extra["rows"] if r["tag"] == _XT_ON_SHEET)
    assert xt_row["reason"] == "hidden_by_rule"

    missing = next(c for c in payload["checks"] if c["id"] == "missing_in_mto")
    missing_tags = [r["tag"] for r in missing["rows"]]
    assert _XT_IN_MTO in missing_tags
    assert _FV_MATCH not in extra_tags
    assert _FV_MATCH not in missing_tags


def test_txt_report_includes_xt_in_both_tables():
    payload = build_tag_analysis_payload(_source_with_xt(), main_doc_signature=_MTO)
    text = render_tag_analysis_text_from_payload(
        payload, main_doc_title="МТО", main_doc_signature=_MTO
    )
    assert _XT_ON_SHEET in text
    assert _XT_IN_MTO in text
