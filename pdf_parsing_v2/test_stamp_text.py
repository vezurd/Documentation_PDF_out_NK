"""Unit tests for ``pdf_parsing_v2.stamp_text`` (primitives, junk, domain, pipelines)."""

from __future__ import annotations

import pytest

from pdf_parsing_v2_engine.stamp_text.context import StampTextContext
from pdf_parsing_v2_engine.stamp_text.junk import JunkMode, filter_junk_lines, junk_clean_lines
from pdf_parsing_v2_engine.stamp_text.primitives import remove_dates, remove_newlines, split_lines
from pdf_parsing_v2_engine.stamp_text.pipelines import (
    UNKNOWN_CLEANER_CODE,
    run_field_clean,
    unknown_cleaner_outcome,
)


def test_remove_dates_strips_dd_mm_pattern() -> None:
    assert remove_dates("foo 01.02.2023 bar") == "foo  bar"


def test_remove_newlines() -> None:
    assert remove_newlines("a\nb\n") == "ab"


def test_junk_drops_date_label_line() -> None:
    lines = split_lines("Дата: 1.1\nReal line\n")
    kept = filter_junk_lines(lines, JunkMode.DEFAULT)
    assert kept == ["Real line"]


def test_junk_mode_with_sheet_labels() -> None:
    lines = split_lines("Лист 1\nContent\n")
    kept_default = filter_junk_lines(lines, JunkMode.DEFAULT)
    assert kept_default == ["Лист 1", "Content"]
    kept_sheet = filter_junk_lines(lines, JunkMode.WITH_SHEET_LABELS)
    assert kept_sheet == ["Content"]


def test_junk_clean_lines_join_doc_title_style() -> None:
    text = "AGCC-1\nДата\nTitle bit\n"
    out = junk_clean_lines(text, JunkMode.DEFAULT)
    assert "Title bit" in out
    assert not any("Дата" in x for x in out)


def test_sheet_6_1_first_page_regex() -> None:
    ctx = StampTextContext(doc_type="WIR", page_num=1)
    o = run_field_clean("sheet_number_6_1", "лист 3.12 из 40", ctx)
    v = o.value
    assert "3" in v or "12" in v


def test_sheet_6_1_6_2_bbb_page2_ly_iz_listov() -> None:
    sample = "AGCC.287-2869-SOS.MTO-0001 л. 3 из 11 листов"
    ctx = StampTextContext(doc_type="MTO", page_num=2)
    assert run_field_clean("sheet_number_6_1", sample, ctx).value == "3"
    assert run_field_clean("sheet_number_6_2", sample, ctx).value == "11"


def test_sheet_6_1_6_2_bbb_page2_ly_iz_listov_case_insensitive() -> None:
    sample = "AGCC-X Л. 5 из 9 листов"
    ctx = StampTextContext(doc_type="MTO", page_num=3)
    assert run_field_clean("sheet_number_6_1", sample, ctx).value == "5"
    assert run_field_clean("sheet_number_6_2", sample, ctx).value == "9"


def test_pipeline_doc_title_strict_no_warning() -> None:
    ctx = StampTextContext()
    o = run_field_clean("doc_title", "AGCC-1\nДата\nTitle bit\n", ctx)
    assert o.clean_tier == "strict"
    assert not o.warnings
    assert "AGCC-1" in o.value


def test_pipeline_doc_title_relaxed_on_long_line() -> None:
    ctx = StampTextContext()
    long_body = " ".join(f"w{i}" for i in range(25))
    text = f"AGCC-X {long_body}\nnext\n"
    o = run_field_clean("doc_title", text, ctx)
    assert o.clean_tier == "relaxed"
    assert len(o.warnings) == 1
    assert o.warnings[0].code == "DOC_TITLE_RELAXED"


def test_unknown_cleaner_outcome() -> None:
    o = unknown_cleaner_outcome("nosuch", "rawx")
    assert o.value == "rawx"
    assert len(o.warnings) == 1
    assert o.warnings[0].code == UNKNOWN_CLEANER_CODE


def test_run_field_clean_unit_title() -> None:
    ctx = StampTextContext()
    o = run_field_clean("unit_title", "a\nb", ctx)
    assert o.value == "a b"
    assert not o.warnings
