from __future__ import annotations

from types import SimpleNamespace

from pdf_parsing_v2_engine.stamp_fields import (
    c_18_1,
    c_18_1_1,
    c_18_1_2,
    c_18_1_3,
    c_18_2,
    c_18_2_1,
    c_18_2_2,
    c_18_2_3,
    c_18_3,
    c_18_3_1,
    c_18_3_2,
    c_18_3_3,
)
from pdf_parsing_v2_rules.checks import prepare_revision_selection
from pdf_parsing_v2_rules.profile import get_profile


def _attrs(
    rev_1: str = "",
    rev_2: str = "",
    rev_3: str = "",
) -> dict[str, list[str]]:
    return {
        c_18_1: [],
        c_18_2: [],
        c_18_3: [],
        c_18_1_1: [rev_1] if rev_1 else [],
        c_18_1_2: [rev_2] if rev_2 else [],
        c_18_1_3: [rev_3] if rev_3 else [],
        c_18_2_1: ["date-1"],
        c_18_2_2: ["date-2"],
        c_18_2_3: ["date-3"],
        c_18_3_1: ["purpose-1"],
        c_18_3_2: ["purpose-2"],
        c_18_3_3: ["purpose-3"],
    }


def _doc(attrs: dict[str, list[str]]) -> SimpleNamespace:
    page = SimpleNamespace(page_num=1, dict_attributes=attrs)
    return SimpleNamespace(doc_OD_style_file_name="doc.pdf", pages=[page])


def test_prepare_revision_selection_uses_latest_numeric_revision() -> None:
    doc = _doc(_attrs("01", "02", "03"))

    rows = prepare_revision_selection([doc], get_profile("AGCC.287"))

    assert rows[0].result is True
    assert doc.pages[0].dict_attributes[c_18_1] == ["03"]
    assert doc.pages[0].dict_attributes[c_18_2] == ["date-3"]
    assert doc.pages[0].dict_attributes[c_18_3] == ["purpose-3"]


def test_prepare_revision_selection_combines_numeric_and_an_revision() -> None:
    doc = _doc(_attrs("01", "03", "AN02"))

    rows = prepare_revision_selection([doc], get_profile("AGCC.287"))

    assert rows[0].result is True
    assert doc.pages[0].dict_attributes[c_18_1] == ["03-AN02"]
    assert doc.pages[0].dict_attributes[c_18_2] == ["date-2"]
    assert doc.pages[0].dict_attributes[c_18_3] == ["purpose-2"]


def test_prepare_revision_selection_rejects_decreasing_numeric_order() -> None:
    doc = _doc(_attrs("03", "01", ""))

    rows = prepare_revision_selection([doc], get_profile("AGCC.287"))

    assert rows[0].result is False
    assert "18_1_1_Current_Revision: 03" in rows[0].text
    assert doc.pages[0].dict_attributes[c_18_1] == []
