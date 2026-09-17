"""Unit tests for TagClass(strict=False) and related helpers."""

from __future__ import annotations

import pytest

from tags.tag_classes import (
    TagClass,
    check_system_code_94s_4_1,
    check_tag_94s_4_2,
    check_tag_94s_4_4_additional_code,
    hide_tag,
)


def test_hide_tag_empty_equipment_not_hidden():
    assert hide_tag("") is False
    assert hide_tag("   ") is False
    assert hide_tag(None) is False


def test_hide_tag_xt_still_hidden():
    assert hide_tag("XT") is True


def test_strict_true_invalid_raises():
    with pytest.raises(ValueError):
        TagClass("bad", strict=True)
    with pytest.raises(ValueError):
        TagClass("a-b", strict=True)


def test_strict_false_invalid_sets_flags_and_defaults():
    tc = TagClass("a-b", strict=False)
    assert tc.is_valid is False
    assert tc.parse_error
    assert tc.wbs == ""
    assert tc.equipment == ""
    assert tc.out_row == 3
    assert tc.category_name == "Не классифицировано"
    assert tc.is_hidden() is False


def test_strict_false_valid_same_as_strict_true():
    s = "2612-PB-01-S-FV-1511"
    a = TagClass(s, strict=True)
    b = TagClass(s, strict=False)
    assert a.is_valid and b.is_valid
    assert a.wbs == b.wbs == "2612"
    assert a.equipment == b.equipment == "FV"
    assert isinstance(a.out_row, int)
    assert a.out_row == b.out_row


def test_category_name_mapping():
    tc = TagClass("2612-PB-01-S-FV-1511", strict=False)
    assert tc.is_valid
    assert tc.category_name in (
        "КСБ (НЕ в шкафах)",
        "КСБ (оборудование в шкафах)",
        "КСБ (шкафы SOT SX)",
        "Не КСБ / иностранные",
    )


def test_get_shield_number_invalid_returns_none():
    tc = TagClass("x", strict=False)
    assert tc.get_shield_number() is None


def test_check_94s_accepts_tagclass_invalid_returns_false():
    tc = TagClass("a-b", strict=False)
    assert not tc.is_valid
    assert check_system_code_94s_4_1(tc) is False
    assert check_tag_94s_4_2(tc) is False
    assert check_tag_94s_4_4_additional_code(tc) is False


def test_coerce_string_strict_false_no_raise_in_checks():
    # Too short: TagClass(..., strict=False) inside checks when we pass strict=False
    assert check_system_code_94s_4_1("ab", strict=False) is False
