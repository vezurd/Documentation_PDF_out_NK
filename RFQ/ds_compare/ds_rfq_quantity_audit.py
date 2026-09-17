"""RFQ quantity totals: source file vs grouped pipeline vs export column."""

from __future__ import annotations

from dataclasses import dataclass

import openpyxl

from base.base_classes import RowStd, RowType
from base.tables_columns import CODE, RFQ_VALUES, VALUES
from RFQ.ds_compare.ds_quantity_parse import quantity_as_float, read_raw_quantity, try_parse_quantity
from RFQ.ds_compare.ds_rfq_grouped_compare import RFQ_MATCH_KEYS, group_rfq_rows

_FLOAT_EPS = 1e-6
# Excel column K = 11 (1-based); VALUES index 10 in RFQTSpecification.column_dict.
_EXCEL_VALUES_COL = 11


@dataclass
class RfqQuantityAudit:
    """Totals for RFQ quantity reconciliation after grouped DS vs MTO vs RFQ."""

    excel_k_all_numeric_sum: float
    excel_k_numeric_row_count: int
    loaded_position_sum: float
    loaded_position_count: int
    loaded_non_position_values_sum: float
    loaded_non_position_values_count: int
    grouped_sum: float
    grouped_position_count: int
    grouped_empty_code_sum: float
    grouped_empty_code_count: int
    export_rfq_values_sum: float
    export_rfq_values_row_count: int
    export_not_in_rfq_rows: int

    @property
    def grouped_vs_export_delta(self) -> float:
        """Positive => export column sums less than grouped RFQ (loss in output)."""
        return self.grouped_sum - self.export_rfq_values_sum

    @property
    def excel_vs_loaded_position_delta(self) -> float:
        """Positive => Excel K > loaded position; negative => loader has more qty than Excel K sum."""
        return self.excel_k_all_numeric_sum - self.loaded_position_sum

    @property
    def export_inflation(self) -> float:
        """Positive => export column sums more than grouped RFQ (duplicate rows before fix)."""
        return self.export_rfq_values_sum - self.grouped_sum

    @property
    def quantities_consistent(self) -> bool:
        """Grouped RFQ total matches naive sum of ``rfq_values`` in compared rows."""
        return abs(self.grouped_vs_export_delta) <= _FLOAT_EPS

    def format_short(self) -> str:
        """One-line summary for GUI status / ActionResult."""
        flag = "OK" if self.quantities_consistent else "РАСХОЖДЕНИЕ"
        extra = ""
        if self.export_inflation > _FLOAT_EPS:
            extra = f" | дубль в отчёте +{self.export_inflation:g}"
        elif self.grouped_vs_export_delta > _FLOAT_EPS:
            extra = f" | не в отчёте −{self.grouped_vs_export_delta:g}"
        if self.excel_vs_loaded_position_delta < -_FLOAT_EPS:
            extra += f" | Excel K −{abs(self.excel_vs_loaded_position_delta):g} vs загрузка"
        return (
            f"RFQ кол-во [{flag}]: файл K={self.excel_k_all_numeric_sum:g} | "
            f"загрузка position={self.loaded_position_sum:g} | "
            f"после группировки={self.grouped_sum:g} | "
            f"колонка «Кол-во RFQ»={self.export_rfq_values_sum:g}{extra}"
        )

    def format_detail(self) -> str:
        """Multi-line report for job log."""
        lines = [
            "RFQ quantity audit:",
            f"  Excel col K (all numeric cells): {self.excel_k_all_numeric_sum:g} "
            f"({self.excel_k_numeric_row_count} rows)",
            f"  Loaded position_row VALUES:     {self.loaded_position_sum:g} "
            f"({self.loaded_position_count} rows)",
        ]
        if self.loaded_non_position_values_count:
            lines.append(
                f"  Loaded other/head with VALUES:  {self.loaded_non_position_values_sum:g} "
                f"({self.loaded_non_position_values_count} rows) — often explain Excel K > pipeline"
            )
        lines.extend(
            [
                f"  After group_rfq_rows:           {self.grouped_sum:g} "
                f"({self.grouped_position_count} rows)",
            ]
        )
        if self.grouped_empty_code_count:
            lines.append(
                f"  Grouped without BCC code:       {self.grouped_empty_code_sum:g} "
                f"({self.grouped_empty_code_count} rows) — not exported to «Кол-во RFQ»"
            )
        lines.extend(
            [
                f"  Export «Кол-во RFQ» sum:        {self.export_rfq_values_sum:g} "
                f"({self.export_rfq_values_row_count} filled rows)",
                f"  Rows «Нет в RFQ» (empty col):   {self.export_not_in_rfq_rows}",
                f"  Delta grouped − export:         {self.grouped_vs_export_delta:g}",
            ]
        )
        if not self.quantities_consistent:
            lines.append(
                "  Hint: export < grouped — missing RFQ-only rows, empty CODE, or duplicate "
                "DS lines consuming keys without showing qty; export > grouped — same RFQ "
                "qty repeated on multiple compared rows."
            )
        if self.export_inflation > _FLOAT_EPS:
            lines.append(
                f"  Export inflation (dup RFQ on several DS/MTO rows): +{self.export_inflation:g} "
                "— after fix, re-run; duplicates show status «RFQ учтён выше»"
            )
        if self.excel_vs_loaded_position_delta > _FLOAT_EPS:
            lines.append(
                f"  Delta Excel K − loaded position: {self.excel_vs_loaded_position_delta:g} "
                "(extra numeric cells in K: subtotals, see RFQ_TPK_not_position_rows.txt)"
            )
        elif self.excel_vs_loaded_position_delta < -_FLOAT_EPS:
            lines.append(
                f"  Loaded position − Excel K: {abs(self.excel_vs_loaded_position_delta):g} "
                "(qty in position_row but cell K not numeric in xlsx, e.g. text/formula)"
            )
        return "\n".join(lines)


def _parse_excel_k_cell(cell: object) -> float | None:
    """Parse quantity from Excel column K using the same rules as RFQ load validation."""
    ok, parsed = try_parse_quantity(cell)
    if not ok:
        return None
    if cell is None or cell == "":
        return None
    return parsed


def sum_excel_column_k_numeric(rfq_path: str) -> tuple[float, int]:
    """Sum numeric cells in Excel column K (like a full-column SUM in Excel).

    Args:
        rfq_path: Path to RFQ TPK xlsx.

    Returns:
        (total, row_count) for cells with numeric values in column K.
    """
    path = str(rfq_path or "").strip()
    if not path:
        return 0.0, 0
    total = 0.0
    count = 0
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        ws = wb.active
        for row in ws.iter_rows(min_col=_EXCEL_VALUES_COL, max_col=_EXCEL_VALUES_COL, values_only=True):
            cell = row[0] if row else None
            parsed = _parse_excel_k_cell(cell)
            if parsed is None:
                continue
            total += parsed
            count += 1
    finally:
        wb.close()
    return total, count


def sum_position_values(rows: list[RowStd]) -> tuple[float, int]:
    """Sum ``VALUES`` on ``position_row`` only."""
    total = 0.0
    count = 0
    for row in rows:
        if row.row_type != RowType.position_row:
            continue
        total += quantity_as_float(read_raw_quantity(row, VALUES))
        count += 1
    return total, count


def sum_non_position_values(rows: list[RowStd]) -> tuple[float, int]:
    """Sum ``VALUES`` on non-position rows where the cell is a real number.

    Header/subtotal text (e.g. ``Кол-во``, ``ИТОГО``) is skipped — not validated as position_row.
    """
    total = 0.0
    count = 0
    for row in rows:
        if row.row_type == RowType.position_row:
            continue
        raw = read_raw_quantity(row, VALUES)
        ok, val = try_parse_quantity(raw)
        if not ok or abs(val) <= _FLOAT_EPS:
            continue
        total += val
        count += 1
    return total, count


def sum_export_rfq_values(compared_rows: list[RowStd]) -> tuple[float, int, int]:
    """Sum ``RFQ_VALUES`` in compared/export rows.

    Returns:
        (total, rows_with_value, rows_with_status_not_in_rfq_empty_col)
    """
    total = 0.0
    filled = 0
    not_in_rfq = 0
    from base.tables_columns import RFQ_COMPARE_STATUS

    for row in compared_rows:
        if row.row_type != RowType.position_row:
            continue
        status = row.get_value(RFQ_COMPARE_STATUS) if RFQ_COMPARE_STATUS in row.el else None
        el = row.el.get(RFQ_VALUES)
        if el is None or el.value is None:
            if status == "Нет в RFQ":
                not_in_rfq += 1
            continue
        total += quantity_as_float(el.value)
        filled += 1
    return total, filled, not_in_rfq


def compute_rfq_quantity_audit(
    rfq_path: str,
    rfq_rows: list[RowStd],
    compared_rows: list[RowStd],
) -> RfqQuantityAudit:
    """Build reconciliation totals for RFQ quantities through the pipeline."""
    excel_sum, excel_count = sum_excel_column_k_numeric(rfq_path)
    loaded_pos_sum, loaded_pos_count = sum_position_values(rfq_rows)
    other_sum, other_count = sum_non_position_values(rfq_rows)

    grouped = group_rfq_rows(rfq_rows, list(RFQ_MATCH_KEYS))
    grouped_sum, grouped_count = sum_position_values(grouped)

    empty_code_sum = 0.0
    empty_code_count = 0
    for row in grouped:
        if row.row_type != RowType.position_row:
            continue
        if not str(row.get_value(CODE) or "").strip():
            empty_code_sum += quantity_as_float(read_raw_quantity(row, VALUES))
            empty_code_count += 1

    export_sum, export_filled, not_in_rfq = sum_export_rfq_values(compared_rows)

    return RfqQuantityAudit(
        excel_k_all_numeric_sum=excel_sum,
        excel_k_numeric_row_count=excel_count,
        loaded_position_sum=loaded_pos_sum,
        loaded_position_count=loaded_pos_count,
        loaded_non_position_values_sum=other_sum,
        loaded_non_position_values_count=other_count,
        grouped_sum=grouped_sum,
        grouped_position_count=grouped_count,
        grouped_empty_code_sum=empty_code_sum,
        grouped_empty_code_count=empty_code_count,
        export_rfq_values_sum=export_sum,
        export_rfq_values_row_count=export_filled,
        export_not_in_rfq_rows=not_in_rfq,
    )
