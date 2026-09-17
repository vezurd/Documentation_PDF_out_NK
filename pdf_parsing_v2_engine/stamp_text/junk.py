"""Semantic line filtering (legacy ``string_Junk_Cleaner`` split into a dedicated layer)."""

from __future__ import annotations

from enum import IntEnum

from pdf_parsing_v2_engine.stamp_text.primitives import split_lines


class JunkMode(IntEnum):
    """Which label set to treat as junk triggers (legacy ``flag``)."""

    DEFAULT = 0
    """Default junk labels (dates, rev, stage, …)."""

    WITH_SHEET_LABELS = 1
    """Also treat lines containing ``Лист`` / ``Листов`` as junk."""


# Lines containing any of these substrings are dropped (legacy ``names`` list).
_JUNK_LABELS_BASE: tuple[str, ...] = (
    "Дата",
    "Date",
    "Рев.",
    "Rev.",
    "Назначение выпуска",
    "Ревизия",
    "Purpose of Issue",
    'ООО "Би.Си.Си."',
    "Примечание",
    "Стадия",
    'АО "НИПИГАЗ"',
    "В связи со значительными изменениями весь документ переделан",
)


def _labels_for_mode(mode: JunkMode) -> tuple[str, ...]:
    if mode == JunkMode.WITH_SHEET_LABELS:
        return _JUNK_LABELS_BASE + ("Лист", "Листов")
    return _JUNK_LABELS_BASE


def filter_junk_lines(lines: list[str], mode: JunkMode = JunkMode.DEFAULT) -> list[str]:
    """Drop lines that match legacy junk rules; keep order of survivors."""
    names = _labels_for_mode(mode)
    indexes: list[int] = []
    for i, line in enumerate(lines):
        for n in names:
            if n in line or line == "" or line == " ":
                indexes.append(i)
                break

    res: list[str] = []
    for i, line in enumerate(lines):
        if i not in indexes:
            res.append(line.strip())
    return res


def junk_clean_lines(text: str, mode: JunkMode = JunkMode.DEFAULT) -> list[str]:
    """Split, normalize, then apply junk filter (legacy ``string_Junk_Cleaner`` for str input)."""
    lines = split_lines(text)
    return filter_junk_lines(lines, mode)
