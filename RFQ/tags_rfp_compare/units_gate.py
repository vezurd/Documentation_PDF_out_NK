"""Common units conversion gate for RFP, MTO, and packing-list rows."""

from __future__ import annotations

from dataclasses import dataclass
from decimal import Decimal
from pathlib import Path
from typing import Iterable

from base.base_classes import RowStd, RowType
from base.tables_columns import (
    CODE,
    NAME,
    UNITS,
    UNITS_CHECK_STATUS,
    UNITS_CONVERSION_TRACE,
    VALUES,
    VALUES_2,
)
from RFQ.ds_compare.ds_quantity_parse import read_raw_quantity
from RFQ.packing_list_provider import PackingDataset, packing_row_key, packing_row_source
from RFQ.units_convert.models import (
    ConversionInvariant,
    ConversionPlan,
    ConversionRequest,
    GoogleUnitsIndex,
    RowStdBinding,
    STATUS_CONVERTED,
    STATUS_IDENTITY,
    STATUS_NO_GOOGLE,
    UnitsConversionError,
    parse_decimal_quantity,
)
from RFQ.units_convert import apply_conversion_plan, build_conversion_plan, build_google_units_index
from RFQ.units_convert.fractional_log import write_fractional_conversion_logs


_PROGRESS_CONTOUR_ORDER: tuple[tuple[str, str], ...] = (
    ("rfp", "RFP"),
    ("mto", "MTO"),
    ("ul", "УЛ"),
)


@dataclass(frozen=True)
class UnitsGateResult:
    """In-memory outcome of a successful units gate run."""

    plan: ConversionPlan
    google_index: GoogleUnitsIndex
    warnings: tuple[str, ...]
    summary: str


def format_units_gate_progress_summary(plan: ConversionPlan) -> str:
    """Build a compact Russian progress summary from plan actions.

    Contours are inferred from ``request_id`` prefixes ``rfp:``, ``mto:``, ``ul:``.
    Unknown prefixes are ignored safely.

    Args:
        plan: Validated conversion plan after ``build_conversion_plan``.

    Returns:
        Summary like ``RFP: проверено N, преобразовано C, без Google G; ...``.
    """
    checked: dict[str, int] = {key: 0 for key, _label in _PROGRESS_CONTOUR_ORDER}
    converted: dict[str, int] = {key: 0 for key, _label in _PROGRESS_CONTOUR_ORDER}
    no_google: dict[str, int] = {key: 0 for key, _label in _PROGRESS_CONTOUR_ORDER}
    identity: dict[str, int] = {key: 0 for key, _label in _PROGRESS_CONTOUR_ORDER}
    fractional: dict[str, int] = {key: 0 for key, _label in _PROGRESS_CONTOUR_ORDER}

    for action in plan.actions:
        prefix = action.request_id.partition(":")[0]
        if prefix not in checked:
            continue
        checked[prefix] += 1
        if action.fractional_result:
            fractional[prefix] += 1
        if action.status == STATUS_CONVERTED:
            converted[prefix] += 1
        elif action.status == STATUS_NO_GOOGLE:
            no_google[prefix] += 1
        elif action.status == STATUS_IDENTITY:
            identity[prefix] += 1

    parts: list[str] = []
    for key, label in _PROGRESS_CONTOUR_ORDER:
        segment = (
            f"{label}: проверено {checked[key]}, "
            f"преобразовано {converted[key]}, без Google {no_google[key]}"
        )
        if identity[key]:
            segment += f", без изменений {identity[key]}"
        if fractional[key]:
            segment += f", дробных {fractional[key]}"
        parts.append(segment)
    return "; ".join(parts)


def _rfp_location(row: RowStd, index: int) -> str:
    path = ""
    if getattr(row, "t_com", None) is not None:
        path = str(getattr(row.t_com, "file_full_path", "") or "").strip()
    suffix = f" row={index + 1}"
    return f"{path}{suffix}".strip() or f"rfp row={index + 1}"


def _mto_location(title_system: str, index: int) -> str:
    title = str(title_system or "").strip() or "<no title>"
    return f"mto title={title} row={index + 1}"


def _ul_location(row: RowStd) -> str:
    source_file, sheet, excel_row = packing_row_source(row)
    parts = [part for part in (source_file, sheet) if part]
    if excel_row is not None:
        parts.append(f"row={excel_row}")
    return " · ".join(parts) or "ul row"


def _values_2_active(row: RowStd) -> bool:
    element = row.el.get(VALUES_2)
    if element is None:
        return False
    value = element.value
    if value is None:
        return False
    if isinstance(value, str) and not value.strip():
        return False
    return True


def _tags_count(row: RowStd) -> int:
    tags = row.get_tags_list()
    return len(tags) if tags else 0


def _row_invariant(row: RowStd) -> ConversionInvariant:
    return ConversionInvariant.TAGS_EQUAL if _tags_count(row) else ConversionInvariant.NONE


def _ul_invariant(row: RowStd, quantity: Decimal) -> ConversionInvariant:
    tags_count = _tags_count(row)
    if tags_count:
        return ConversionInvariant.TAGS_NOT_EXCEED_QUANTITY
    return ConversionInvariant.NONE


def _ul_gate_quantity(row: RowStd) -> Decimal | None:
    try:
        quantity = parse_decimal_quantity(
            read_raw_quantity(row, VALUES),
            location=_ul_location(row),
        )
    except UnitsConversionError:
        return None
    if quantity < 0:
        return None
    if quantity != quantity.to_integral_value():
        return None
    return quantity


def _iter_ul_gate_rows(dataset: PackingDataset | None) -> Iterable[tuple[RowStd, Decimal]]:
    if dataset is None or not dataset.available:
        return
    for row in dataset.rows:
        if row.row_type != RowType.position_row:
            continue
        key = packing_row_key(row)
        if not all(key):
            continue
        quantity = _ul_gate_quantity(row)
        if quantity is None:
            continue
        yield row, quantity


def _position_name(row: RowStd) -> str:
    raw = row.get_value(NAME)
    if isinstance(raw, list):
        parts = [str(item).strip() for item in raw if str(item).strip()]
        return "; ".join(parts)
    return str(raw or "").strip()


def _append_position_request(
    *,
    requests: list[ConversionRequest],
    bindings: list[RowStdBinding],
    row: RowStd,
    request_id: str,
    contour: str,
    location: str,
    quantity_column: str,
    quantity: object,
    invariant: ConversionInvariant,
) -> None:
    code = str(row.get_value(CODE) or "").strip()
    units = str(row.get_value(UNITS) or "").strip()
    tags_count = _tags_count(row)
    item_name = _position_name(row)
    requests.append(
        ConversionRequest(
            request_id=request_id,
            contour=contour,
            code=code,
            source_unit=units,
            quantity=quantity,
            tags_count=tags_count,
            invariant=invariant,
            location=location,
            item_name=item_name,
        )
    )
    bindings.append(
        RowStdBinding(
            request_id=request_id,
            row=row,
            quantity_column=quantity_column,
            units_column=UNITS,
            status_column=UNITS_CHECK_STATUS,
            trace_column=UNITS_CONVERSION_TRACE,
        )
    )


def build_units_gate_plan(
    *,
    rfp_rows: list[RowStd],
    mto_data: dict[str, list[RowStd]],
    packing_dataset: PackingDataset | None,
    google_rows: list[RowStd],
    matrix_path: str | Path,
) -> tuple[ConversionPlan, GoogleUnitsIndex, list[RowStdBinding]]:
    """Build a validated conversion plan and bindings without applying it."""
    google_index = build_google_units_index(google_rows)
    requests: list[ConversionRequest] = []
    bindings: list[RowStdBinding] = []

    for index, row in enumerate(rfp_rows):
        if row.row_type != RowType.position_row:
            continue
        location = _rfp_location(row, index)
        invariant = _row_invariant(row)
        quantity = read_raw_quantity(row, VALUES)
        request_id = f"rfp:{index}:values"
        _append_position_request(
            requests=requests,
            bindings=bindings,
            row=row,
            request_id=request_id,
            contour="rfp",
            location=location,
            quantity_column=VALUES,
            quantity=quantity,
            invariant=invariant,
        )
        if _values_2_active(row):
            request_id_v2 = f"rfp:{index}:values_2"
            _append_position_request(
                requests=requests,
                bindings=bindings,
                row=row,
                request_id=request_id_v2,
                contour="rfp",
                location=f"{location} VALUES_2",
                quantity_column=VALUES_2,
                quantity=read_raw_quantity(row, VALUES_2),
                invariant=invariant,
            )

    mto_request_seq = 0
    for title_system, rows in (mto_data or {}).items():
        for index, row in enumerate(rows):
            if row.row_type != RowType.position_row:
                continue
            mto_request_seq += 1
            location = _mto_location(title_system, index)
            request_id = f"mto:{mto_request_seq}:values"
            _append_position_request(
                requests=requests,
                bindings=bindings,
                row=row,
                request_id=request_id,
                contour="mto",
                location=location,
                quantity_column=VALUES,
                quantity=read_raw_quantity(row, VALUES),
                invariant=_row_invariant(row),
            )

    for ul_index, (row, quantity) in enumerate(_iter_ul_gate_rows(packing_dataset)):
        location = _ul_location(row)
        request_id = f"ul:{ul_index}:values"
        _append_position_request(
            requests=requests,
            bindings=bindings,
            row=row,
            request_id=request_id,
            contour="ul",
            location=location,
            quantity_column=VALUES,
            quantity=quantity,
            invariant=_ul_invariant(row, quantity),
        )

    plan = build_conversion_plan(requests, google_index, matrix_path)
    return plan, google_index, bindings


def run_units_gate(
    *,
    rfp_rows: list[RowStd],
    mto_data: dict[str, list[RowStd]],
    packing_dataset: PackingDataset | None,
    google_rows: list[RowStd],
    matrix_path: str | Path,
    logs_dir: str | Path | None = None,
) -> UnitsGateResult:
    """Build and apply one common units conversion plan across active contours."""
    plan, google_index, bindings = build_units_gate_plan(
        rfp_rows=rfp_rows,
        mto_data=mto_data,
        packing_dataset=packing_dataset,
        google_rows=google_rows,
        matrix_path=matrix_path,
    )
    apply_conversion_plan(plan, bindings)
    for warning in plan.warnings:
        print(f"Предупреждение конвертации ед. изм.: {warning}")
    if logs_dir is not None:
        for path in write_fractional_conversion_logs(plan, Path(logs_dir)):
            print(f"written fractional conversion xlsx: {path} ({path.name})")
    return UnitsGateResult(
        plan=plan,
        google_index=google_index,
        warnings=plan.warnings,
        summary=format_units_gate_progress_summary(plan),
    )
