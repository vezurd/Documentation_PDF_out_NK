"""Grouped DS vs MTO enrichment with RFQ TPK quantities.

Collapses RFQ rows by the same keys as grouped DS, then matches each
compared DS/MTO row by title + mark + code and fills balance columns.
"""

from __future__ import annotations

from dataclasses import dataclass

from base.base_classes import CheckElement, RowStd, RowType, TableComments
from base.tables_columns import (
    CODE,
    CODE_2,
    DS_SYSTEM,
    DS_TITLE,
    MTO_DS_DIFF,
    DELIVERY_OVERALL_STATUS,
    MTO_DS_RFQ_DIFF,
    NAME,
    RFQ_CODE,
    RFQ_COMPARE_STATUS,
    RFQ_NAME,
    RFQ_TYPE_MARK,
    RFQ_UNITS,
    RFQ_VALUES,
    ROW_TYPE,
    TYPE_MARK,
    UNITS,
    VALUES,
    VALUES_2,
)
from RFQ.ds_compare.ds_grouped_compare import group_ds_rows
from RFQ.ds_compare.ds_quantity_parse import quantity_as_float
from utils.colors import Color

_FLOAT_EPS = 1e-9

# RFQ match is always title (DS_TITLE) + mark (DS_SYSTEM) + code — not ds_name.
RFQ_MATCH_KEYS: tuple[str, ...] = ("title", "system", "code")
_DEBUG_MISS_SAMPLES = 5

STATUS_MATCH = "Совпадение"
STATUS_INCREASE = "Увеличение"
STATUS_SHORTFALL = "Недопоставка"
STATUS_NOT_IN_RFQ = "Нет в RFQ"
STATUS_RFQ_ALREADY = "RFQ учтён выше"
STATUS_NEW = "Новая"

_COLOR_BY_STATUS: dict[str, str] = {
    STATUS_MATCH: Color.green,
    STATUS_INCREASE: Color.match_added_mto,
    STATUS_SHORTFALL: Color.yellow,
}


@dataclass
class RfqCompareStats:
    """Counters for DS/MTO vs RFQ enrichment."""

    total_rows: int = 0
    position_rows: int = 0
    grouped_rfq_rows: int = 0
    rfq_matched: int = 0
    rfq_not_found: int = 0
    status_match: int = 0
    status_increase: int = 0
    status_shortfall: int = 0
    rfq_only_added: int = 0
    rfq_duplicate_skipped: int = 0


def _norm_match_text(value: object) -> str:
    """Normalize title/mark/code for lookup (strip, int-like floats)."""
    if value is None:
        return ""
    if isinstance(value, float) and value == int(value):
        return str(int(value))
    return str(value).strip()


def _codes_for_match(row: RowStd) -> list[str]:
    """Return non-empty DS ``CODE`` and MTO ``CODE_2`` (each once, in order)."""
    codes: list[str] = []
    seen: set[str] = set()
    for col_name in (CODE, CODE_2):
        text = _norm_match_text(row.get_value(col_name))
        if not text or text in seen:
            continue
        seen.add(text)
        codes.append(text)
    return codes


def _match_key(row: RowStd, code: str) -> tuple[str, str, str]:
    """Lookup key: (DS_TITLE, DS_SYSTEM, code)."""
    return (
        _norm_match_text(row.get_value(DS_TITLE)),
        _norm_match_text(row.get_value(DS_SYSTEM)),
        _norm_match_text(code),
    )


def _rfq_group_key(row: RowStd) -> tuple[str, str, str]:
    return _match_key(row, _norm_match_text(row.get_value(CODE)))


def _find_rfq_row(
    rfq_lookup: dict[tuple[str, str, str], RowStd],
    row: RowStd,
) -> tuple[RowStd | None, tuple[str, str, str] | None, str | None]:
    """Try RFQ lookup for each of ``CODE`` and ``CODE_2`` on the compared row."""
    for code in _codes_for_match(row):
        key = _match_key(row, code)
        hit = rfq_lookup.get(key)
        if hit is not None:
            return hit, key, code
    last_code = _codes_for_match(row)
    if last_code:
        return None, _match_key(row, last_code[-1]), last_code[-1]
    return None, _match_key(row, ""), None


def group_rfq_rows(rfq_rows: list[RowStd], group_keys: list[str] | None = None) -> list[RowStd]:
    """Collapse RFQ TPK position rows by title + mark + code."""
    del group_keys  # RFQ grouping ignores ds_name and other DS-only keys.
    return group_ds_rows(rfq_rows, list(RFQ_MATCH_KEYS))


def build_rfq_lookup(
    grouped_rfq_rows: list[RowStd],
    group_keys: list[str] | None = None,
) -> dict[tuple[str, str, str], RowStd]:
    """Map grouped RFQ rows by ``(DS_TITLE, DS_SYSTEM, CODE)``."""
    del group_keys
    lookup: dict[tuple[str, str, str], RowStd] = {}
    for row in grouped_rfq_rows:
        if row.row_type != RowType.position_row:
            continue
        code = _norm_match_text(row.get_value(CODE))
        if not code:
            continue
        lookup[_rfq_group_key(row)] = row
    return lookup


def _ensure_rfq_columns(row: RowStd) -> None:
    for col_name in (
        MTO_DS_DIFF,
        RFQ_VALUES,
        RFQ_UNITS,
        MTO_DS_RFQ_DIFF,
        RFQ_COMPARE_STATUS,
        DELIVERY_OVERALL_STATUS,
        RFQ_NAME,
        RFQ_CODE,
        RFQ_TYPE_MARK,
    ):
        if col_name not in row.el:
            row.el[col_name] = CheckElement(None)


def _clear_rfq_enrichment(row: RowStd) -> None:
    _ensure_rfq_columns(row)
    row.el[MTO_DS_DIFF].value = None
    row.el[MTO_DS_DIFF].color = Color.no
    row.el[RFQ_VALUES].value = None
    row.el[RFQ_VALUES].color = Color.no
    row.el[RFQ_UNITS].value = None
    row.el[RFQ_UNITS].color = Color.no
    row.el[MTO_DS_RFQ_DIFF].value = None
    row.el[MTO_DS_RFQ_DIFF].color = Color.no
    row.el[RFQ_COMPARE_STATUS].value = None
    row.el[RFQ_COMPARE_STATUS].color = Color.no
    row.el[DELIVERY_OVERALL_STATUS].value = None
    row.el[DELIVERY_OVERALL_STATUS].color = Color.no
    row.el[RFQ_NAME].value = None
    row.el[RFQ_NAME].color = Color.no
    row.el[RFQ_CODE].value = None
    row.el[RFQ_CODE].color = Color.no
    row.el[RFQ_TYPE_MARK].value = None
    row.el[RFQ_TYPE_MARK].color = Color.no


def _status_for_diff(diff: float) -> tuple[str, str]:
    if abs(diff) < _FLOAT_EPS:
        return STATUS_MATCH, _COLOR_BY_STATUS[STATUS_MATCH]
    if diff < 0:
        return STATUS_INCREASE, _COLOR_BY_STATUS[STATUS_INCREASE]
    return STATUS_SHORTFALL, _COLOR_BY_STATUS[STATUS_SHORTFALL]


def _set_delivery_overall_status(row: RowStd, diff: float) -> None:
    """Overall delivery status from a balance diff (MTO−DS or MTO−DS−RFQ)."""
    _ensure_rfq_columns(row)
    status_text, status_color = _status_for_diff(diff)
    row.el[DELIVERY_OVERALL_STATUS].value = status_text
    row.el[DELIVERY_OVERALL_STATUS].color = status_color


def _set_delivery_overall_status_new(row: RowStd) -> None:
    """RFQ-only row: overall status is ``Новая`` (same highlight as RFQ-only block)."""
    _ensure_rfq_columns(row)
    row.el[DELIVERY_OVERALL_STATUS].value = STATUS_NEW
    Color.set_el_color(row.el[DELIVERY_OVERALL_STATUS], Color.soft_cyan)


def _fill_rfq_detail_columns(
    out_row: RowStd,
    rfq_row: RowStd,
    *,
    with_cyan: bool = False,
) -> None:
    """Copy RFQ name / code / type_mark into export columns."""
    _ensure_rfq_columns(out_row)
    for src_col, dest_col in (
        (NAME, RFQ_NAME),
        (CODE, RFQ_CODE),
        (TYPE_MARK, RFQ_TYPE_MARK),
    ):
        _copy_rfq_field(
            out_row, rfq_row, src_col, dest_col, with_cyan=with_cyan
        )


def _copy_rfq_field(
    out_row: RowStd,
    rfq_row: RowStd,
    src_col: str,
    dest_col: str,
    *,
    with_cyan: bool = True,
) -> None:
    """Copy one RFQ source field into a destination column."""
    if src_col not in rfq_row.el:
        return
    value = rfq_row.el[src_col].value
    if value is None or str(value).strip() == "":
        return
    out_row.el[dest_col].value = value
    if with_cyan:
        Color.set_el_color(out_row.el[dest_col], Color.soft_cyan)


def _create_rfq_only_row(rfq_row: RowStd) -> RowStd:
    """Build a compared row for grouped RFQ lines absent from DS vs MTO."""
    out_row = RowStd.get_std_check_row({}, TableComments())
    out_row.row_type = RowType.position_row
    out_row.el[ROW_TYPE].value = RowType.position_row

    for col_name in (DS_TITLE, DS_SYSTEM):
        _copy_rfq_field(out_row, rfq_row, col_name, col_name)

    _fill_rfq_detail_columns(out_row, rfq_row, with_cyan=True)

    rfq_value = quantity_as_float(rfq_row.get_value(VALUES))
    rfq_units = rfq_row.get_value(UNITS)

    _ensure_rfq_columns(out_row)
    out_row.el[MTO_DS_DIFF].value = 0.0
    out_row.el[MTO_DS_DIFF].color = Color.no
    out_row.el[RFQ_VALUES].value = rfq_value
    Color.set_el_color(out_row.el[RFQ_VALUES], Color.soft_cyan)
    out_row.el[RFQ_UNITS].value = rfq_units
    Color.set_el_color(out_row.el[RFQ_UNITS], Color.soft_cyan)
    out_row.el[MTO_DS_RFQ_DIFF].value = -rfq_value
    Color.set_el_color(out_row.el[MTO_DS_RFQ_DIFF], Color.soft_cyan)
    out_row.el[RFQ_COMPARE_STATUS].value = STATUS_NEW
    Color.set_el_color(out_row.el[RFQ_COMPARE_STATUS], Color.soft_cyan)
    _set_delivery_overall_status_new(out_row)

    return out_row


def _append_unmatched_rfq_rows(
    compared_rows: list[RowStd],
    grouped_rfq: list[RowStd],
    matched_rfq_keys: set[tuple[str, str, str]],
) -> int:
    """Append RFQ-only rows that were not consumed by DS/MTO lookup."""
    added = 0
    for rfq_row in grouped_rfq:
        if rfq_row.row_type != RowType.position_row:
            continue
        key = _rfq_group_key(rfq_row)
        if key in matched_rfq_keys:
            continue
        if not key[2]:
            # No BCC code: cannot enter lookup, but quantity must appear in export.
            compared_rows.append(_create_rfq_only_row(rfq_row))
            added += 1
            continue
        compared_rows.append(_create_rfq_only_row(rfq_row))
        matched_rfq_keys.add(key)
        added += 1
    return added


def compare_ds_mto_with_rfq(
    compared_rows: list[RowStd],
    rfq_rows: list[RowStd],
    group_keys: list[str] | None = None,
) -> tuple[list[RowStd], RfqCompareStats]:
    """Enrich DS vs MTO rows with RFQ balance columns and status colors.

    For each ``position_row``:
      - ``MTO_DS_DIFF`` = ``VALUES_2`` − ``VALUES``
      - lookup grouped RFQ by ``DS_TITLE`` + ``DS_SYSTEM`` + code
      - for each non-empty ``CODE`` and ``CODE_2`` try RFQ lookup (DS code first)
      - ``MTO_DS_RFQ_DIFF`` = ``MTO_DS_DIFF`` − ``RFQ_VALUES``
      - ``RFQ_COMPARE_STATUS``: Совпадение / Увеличение / Недопоставка
      - unmatched grouped RFQ rows are appended at the end with status ``Новая``

    Args:
        compared_rows: Rows after grouped DS vs MTO compare (and optional MTO-only grouping).
        rfq_rows: Raw RFQ TPK rows from ``load_rfq_tpk_data``.
        group_keys: Ignored for RFQ match (always title + system + code).

    Returns:
        The same ``compared_rows`` list (mutated in place) and summary counters.
    """
    del group_keys
    grouped_rfq = group_rfq_rows(rfq_rows)
    rfq_lookup = build_rfq_lookup(grouped_rfq)

    stats = RfqCompareStats(
        total_rows=len(compared_rows),
        grouped_rfq_rows=len(grouped_rfq),
    )
    miss_samples: list[str] = []
    matched_rfq_keys: set[tuple[str, str, str]] = set()

    for row in compared_rows:
        if row.row_type != RowType.position_row:
            _clear_rfq_enrichment(row)
            continue

        stats.position_rows += 1
        _ensure_rfq_columns(row)

        ds_value = quantity_as_float(row.get_value(VALUES))
        mto_value = quantity_as_float(row.get_value(VALUES_2))
        mto_ds_diff = mto_value - ds_value
        row.el[MTO_DS_DIFF].value = mto_ds_diff
        row.el[MTO_DS_DIFF].color = Color.no

        rfq_row, tried_key, matched_code = _find_rfq_row(rfq_lookup, row)
        if rfq_row is None:
            stats.rfq_not_found += 1
            row.el[RFQ_VALUES].value = None
            row.el[RFQ_UNITS].value = None
            row.el[MTO_DS_RFQ_DIFF].value = None
            row.el[RFQ_COMPARE_STATUS].value = STATUS_NOT_IN_RFQ
            row.el[RFQ_COMPARE_STATUS].color = Color.no
            _set_delivery_overall_status(row, mto_ds_diff)
            if tried_key and len(miss_samples) < _DEBUG_MISS_SAMPLES:
                codes = _codes_for_match(row)
                miss_samples.append(
                    f"title={tried_key[0]!r}, mark={tried_key[1]!r}, "
                    f"tried_codes={codes!r}, lookup_size={len(rfq_lookup)}"
                )
            continue

        match_key = _match_key(row, matched_code) if matched_code else None
        if match_key and match_key in matched_rfq_keys:
            stats.rfq_duplicate_skipped += 1
            row.el[RFQ_VALUES].value = None
            row.el[RFQ_UNITS].value = None
            row.el[MTO_DS_RFQ_DIFF].value = None
            row.el[RFQ_COMPARE_STATUS].value = STATUS_RFQ_ALREADY
            row.el[RFQ_COMPARE_STATUS].color = Color.no
            row.el[RFQ_COMPARE_STATUS].comment = (
                f"Кол-во RFQ уже на другой строке сопоставления {match_key!r}"
            )
            _set_delivery_overall_status(row, mto_ds_diff)
            continue

        stats.rfq_matched += 1
        if match_key:
            matched_rfq_keys.add(match_key)
        if matched_code and matched_code != _norm_match_text(row.get_value(CODE)):
            row.el[RFQ_COMPARE_STATUS].comment = (
                f"RFQ matched by MTO code {matched_code!r}"
            )
        rfq_value = quantity_as_float(rfq_row.get_value(VALUES))
        rfq_diff = mto_ds_diff - rfq_value
        status_text, status_color = _status_for_diff(rfq_diff)

        row.el[RFQ_VALUES].value = rfq_value
        row.el[RFQ_VALUES].color = Color.no
        row.el[RFQ_UNITS].value = rfq_row.get_value(UNITS)
        row.el[RFQ_UNITS].color = Color.no
        row.el[MTO_DS_RFQ_DIFF].value = rfq_diff
        row.el[MTO_DS_RFQ_DIFF].color = status_color
        row.el[RFQ_COMPARE_STATUS].value = status_text
        row.el[RFQ_COMPARE_STATUS].color = status_color
        _fill_rfq_detail_columns(row, rfq_row, with_cyan=False)
        _set_delivery_overall_status(row, rfq_diff)

        if status_text == STATUS_MATCH:
            stats.status_match += 1
        elif status_text == STATUS_INCREASE:
            stats.status_increase += 1
        else:
            stats.status_shortfall += 1

    stats.rfq_only_added = _append_unmatched_rfq_rows(
        compared_rows,
        grouped_rfq,
        matched_rfq_keys,
    )
    if stats.rfq_duplicate_skipped:
        print(
            "RFQ match: duplicate key on extra DS/MTO rows (qty not repeated): "
            f"{stats.rfq_duplicate_skipped}"
        )
    if stats.rfq_only_added:
        stats.total_rows = len(compared_rows)
        print(
            "RFQ match: appended "
            f"{stats.rfq_only_added} RFQ-only row(s) with status "
            f"{STATUS_NEW!r}"
        )

    if miss_samples:
        print("RFQ match: sample (DS_TITLE, DS_SYSTEM, code) not found in RFQ lookup:")
        for line in miss_samples:
            print(f"  - {line}")
        sample_keys = list(rfq_lookup.keys())[:3]
        if sample_keys:
            print(
                "RFQ lookup sample keys (DS_TITLE, DS_SYSTEM, CODE) — first 3: "
                f"{sample_keys!r}"
            )

    return compared_rows, stats
