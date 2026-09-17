"""Units normalization and post-compare highlighting for DS vs MTO vs RFQ."""

from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Iterable

from base.base_classes import CheckElement, RowStd, RowType
from base.tables_columns import (
    RFQ_COMPARE_STATUS,
    RFQ_UNITS,
    UL_COMPARE_STATUS,
    UL_UNITS,
    UNITS,
    UNITS_2,
    UNITS_CHECK_STATUS,
    UNITS_CONVERSION_TRACE,
    UNITS_MTO,
)
from utils.colors import Color

_RFQ_ONLY_STATUS = "Новая"
_PACKING_ONLY_STATUS = "Только в УЛ"

_UNITS_TRAILING_DOTS_RE = re.compile(r"\.+$")


def normalize_units_text(value: object) -> str:
    """Strip whitespace and trailing dots (``шт.`` → ``шт``)."""
    text = str(value or "").strip()
    if not text:
        return ""
    return _UNITS_TRAILING_DOTS_RE.sub("", text).strip()


def join_unique_units_text(
    values: Iterable[object],
    *,
    separator: str = "; ",
) -> str:
    """Join unique non-empty unit status/trace fragments in deterministic order."""
    parts: list[str] = []
    seen: set[str] = set()
    for value in values:
        text = str(value or "").strip()
        if not text or text in seen:
            continue
        seen.add(text)
        parts.append(text)
    return separator.join(parts)


def merge_units_status_trace(target: RowStd, *sources: RowStd) -> None:
    """Merge UNITS_CHECK_STATUS / UNITS_CONVERSION_TRACE without losing existing values."""
    for column in (UNITS_CHECK_STATUS, UNITS_CONVERSION_TRACE):
        if column not in target.el:
            target.el[column] = CheckElement(None)
        source_values = [
            source.el[column].value
            for source in sources
            if column in source.el
        ]
        merged = join_unique_units_text(
            [target.get_value(column), *source_values]
        )
        target.el[column].value = merged or None


def copy_mto_units_fields(target: RowStd, mto_row: RowStd) -> None:
    """Copy normalized MTO units and merge status/trace from the MTO source row."""
    if UNITS_MTO not in target.el:
        target.el[UNITS_MTO] = CheckElement(None)
    normalized_units = normalize_units_text(mto_row.get_value(UNITS))
    target.el[UNITS_MTO].value = normalized_units or None
    merge_units_status_trace(target, mto_row)


def normalize_row_units(row: RowStd, *column_names: str) -> None:
    """Normalize units columns in place on one row."""
    for col_name in column_names:
        if col_name not in row.el:
            continue
        normalized = normalize_units_text(row.el[col_name].value)
        row.el[col_name].value = normalized or None


def normalize_rows_units(
    rows: list[RowStd],
    column_names: tuple[str, ...] = (UNITS,),
    *,
    row_types: tuple[str, ...] = (RowType.position_row,),
) -> int:
    """Normalize units on selected rows; returns count of touched rows."""
    touched = 0
    for row in rows:
        if row_types and row.row_type not in row_types:
            continue
        normalize_row_units(row, *column_names)
        touched += 1
    return touched


def normalize_mto_positions_units(rows: list[RowStd] | None) -> None:
    """Normalize ``UNITS`` on MTO position rows after load."""
    if not rows:
        return
    normalize_rows_units(rows, (UNITS,))


def normalize_mto_dict_units(mto_dict: dict[str, list[RowStd]]) -> None:
    """Normalize ``UNITS`` in every MTO list loaded for DS compare."""
    for rows in mto_dict.values():
        normalize_mto_positions_units(rows)


@dataclass
class UnitsCompareStats:
    """Counters for units consistency highlighting."""

    position_rows: int = 0
    all_match: int = 0
    mismatch: int = 0


def _units_pair_match(units: str, units_2: str) -> bool:
    if not units or not units_2:
        return False
    return units == units_2


def _units_triple_match(units: str, units_2: str, rfq_units: str) -> bool:
    if not units or not units_2 or not rfq_units:
        return False
    return units == units_2 == rfq_units


def highlight_compared_units_columns(
    rows: list[RowStd],
) -> UnitsCompareStats:
    """Color DS, MTO, RFQ, and packing units after grouped enrichment.

    Normal rows compare DS+MTO and append RFQ/UL when present. RFQ-only rows
    compare RFQ with UL when both are present. UL-only rows keep their cyan
    source highlighting.

    Args:
        rows: Final compared rows (with RFQ columns filled where matched).

    Returns:
        Summary counters.
    """
    stats = UnitsCompareStats()
    for row in rows:
        if row.row_type != RowType.position_row:
            continue
        for column in (
            UNITS,
            UNITS_2,
            RFQ_UNITS,
            UL_UNITS,
            RFQ_COMPARE_STATUS,
            UL_COMPARE_STATUS,
        ):
            if column not in row.el:
                row.el[column] = CheckElement(None)

        rfq_status = str(row.get_value(RFQ_COMPARE_STATUS) or "").strip()
        ul_status = str(row.get_value(UL_COMPARE_STATUS) or "").strip()
        if ul_status == _PACKING_ONLY_STATUS:
            continue

        u = normalize_units_text(row.get_value(UNITS))
        u2 = normalize_units_text(row.get_value(UNITS_2))
        ru = normalize_units_text(row.get_value(RFQ_UNITS))
        ulu = normalize_units_text(row.get_value(UL_UNITS))
        if not (u or u2 or ru):
            # UL-only and malformed UL-only rows keep cyan/red source colors.
            continue

        if rfq_status == _RFQ_ONLY_STATUS:
            if not (ru and ulu):
                continue
            compared = ((RFQ_UNITS, ru), (UL_UNITS, ulu))
        else:
            compared_list = [(UNITS, u), (UNITS_2, u2)]
            if ru:
                compared_list.append((RFQ_UNITS, ru))
            else:
                row.el[RFQ_UNITS].color = Color.no
            if ulu:
                compared_list.append((UL_UNITS, ulu))
            else:
                row.el[UL_UNITS].color = Color.no
            compared = tuple(compared_list)

        stats.position_rows += 1
        normalized_values = [value for _column, value in compared]
        matched = bool(normalized_values) and all(normalized_values)
        if matched:
            matched = len(set(normalized_values)) == 1
        color = Color.green if matched else Color.yellow
        for column, _value in compared:
            row.el[column].color = color

        if matched:
            stats.all_match += 1
        else:
            stats.mismatch += 1

    return stats
