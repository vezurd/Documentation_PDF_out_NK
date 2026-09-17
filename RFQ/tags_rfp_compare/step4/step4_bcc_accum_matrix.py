"""BCC accumulation matrix: one row per purchase code, no match/replacements.

Built from post-units_gate MTO / RFP(DS) / packing rows already in memory.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime
import math
import os
import shutil
import tempfile
import time

import xlsxwriter

from base.base_classes import RowStd, RowType
from base.tables_columns import (
    ANNOTATION,
    CODE,
    DS_NAME,
    NAME,
    TYPE_MARK,
    UNITS,
    VALUES,
    VENDOR,
)
from RFQ.ds_compare.ds_units_normalize import normalize_units_text
from RFQ.packing_list_provider import PackingDataset, packing_row_source
from RFQ.tags_rfp_compare.rfp_supply_status import is_excluded_from_supply
from RFQ.tags_rfp_compare.rfp_tags_utils import (
    append_timing_log,
    ensure_result_dir_exists,
)
from utils.colors import Color

GoogleIdentity = tuple[str, str, str, str]

_EMPTY_TITLE = "(без титула)"
_EMPTY_DS = "(без имени ДС)"
_EMPTY_UL = "(без файла УЛ)"

# Layout matches «Накопительная ведомость АГХК.xlsx»: identity, totals/diffs, wrap sources.
_FIXED_HEADERS = (
    "Код",
    "Наименование",
    "Тип / марка",
    "Вендор",
    "Ед. изм.",
    "Σ МТО",
    "Σ ДС",
    "Σ УЛ",
    "МТО − ДС",
    "ДС − УЛ",
    "МТО − УЛ",
    "Титулы МТО",
    "Источники ДС",
    "Источники УЛ",
)
_COL_UNITS = 4
_COL_QTY_FIRST = 5
_COL_QTY_LAST = 7
_COL_DIFF_FIRST = 8
_COL_DIFF_LAST = 10
_COL_WRAP_FIRST = 11
_COL_WRAP_LAST = 13

_HEADER_IDENTITY = "#D9D9D9"
_HEADER_TOTALS = "#F4B183"
_HEADER_MTO = "#BDD7EE"
_HEADER_DS = "#C6EFCE"
_HEADER_UL = "#FFE699"


@dataclass
class BccAccumRow:
    """One matrix row: coded BCC or empty-code position identity."""

    code: str
    name: str
    type_mark: str
    units: str
    vendor: str
    mto_titles_info: str
    ds_sources_info: str
    ul_sources_info: str
    qty_mto: float
    qty_ds: float
    qty_ul: float
    diff_mto_ds: float | None
    diff_ds_ul: float | None
    diff_mto_ul: float | None
    units_mismatch: bool
    mto_by_title: dict[str, float]
    ds_by_name: dict[str, float]
    ul_by_file: dict[str, float]


@dataclass
class BccAccumMatrix:
    """Pivot ready for Excel: rows plus dynamic column keys."""

    rows: list[BccAccumRow]
    mto_title_columns: list[str]
    ds_name_columns: list[str]
    ul_file_columns: list[str]


@dataclass
class _Bucket:
    code: str
    name: str = ""
    type_mark: str = ""
    units: str = ""
    vendor: str = ""
    mto_by_title: dict[str, float] = field(default_factory=dict)
    ds_by_name: dict[str, float] = field(default_factory=dict)
    ul_by_file: dict[str, float] = field(default_factory=dict)
    mto_units: set[str] = field(default_factory=set)
    ds_units: set[str] = field(default_factory=set)
    ul_units: set[str] = field(default_factory=set)


def build_google_identity_index(
    google_rows: list[RowStd] | None,
) -> dict[str, GoogleIdentity]:
    """Build first-match CODE → (name, type_mark, units, vendor).

    Args:
        google_rows: Rows from ``load_base()``. Empty CODE is skipped.

    Returns:
        Lookup keyed by stripped purchase code.
    """
    out: dict[str, GoogleIdentity] = {}
    for row in google_rows or []:
        code = _field(row, CODE)
        if not code or code in out:
            continue
        out[code] = (
            _field(row, NAME),
            _field(row, TYPE_MARK),
            _field(row, UNITS),
            _field(row, VENDOR),
        )
    return out


def build_bcc_accum_matrix(
    *,
    rfp_rows: list[RowStd] | None,
    mto_data: dict[str, list[RowStd]] | None,
    packing_dataset: PackingDataset | None,
    google_rows: list[RowStd] | None = None,
    google_identity: dict[str, GoogleIdentity] | None = None,
) -> BccAccumMatrix:
    """Aggregate MTO, RFP(DS) and UL quantities by purchase code.

    Args:
        rfp_rows: Post-gate RFP position rows (DS quantities).
        mto_data: Post-gate MTO rows keyed by title_system.
        packing_dataset: Shared packing cache; ``None`` skips UL columns.
        google_rows: Google base rows used when ``google_identity`` is omitted.
        google_identity: Prebuilt CODE lookup; wins over source-row text.

    Returns:
        Matrix with outer-joined codes and empty-code position rows.

    Notes:
        Does not mutate input ``RowStd`` (qty is parsed without ``get_value``).
    """
    identity = google_identity
    if identity is None:
        identity = build_google_identity_index(google_rows)

    buckets: dict[tuple, _Bucket] = {}
    anon_seq = 0

    for title_system, rows in (mto_data or {}).items():
        title = str(title_system).strip() or _EMPTY_TITLE
        for row in rows or []:
            qty = _qty(row)
            if qty is None:
                continue
            key, bucket = _bucket_for_row(row, buckets, anon_seq)
            if key[0] == "anon":
                anon_seq += 1
            _add_qty(bucket.mto_by_title, title, qty)
            _add_units(bucket.mto_units, row)
            _fill_identity(bucket, row)

    for row in rfp_rows or []:
        if is_excluded_from_supply(row):
            continue
        qty = _qty(row)
        if qty is None:
            continue
        key, bucket = _bucket_for_row(row, buckets, anon_seq)
        if key[0] == "anon":
            anon_seq += 1
        ds_name = _field(row, DS_NAME) or _EMPTY_DS
        _add_qty(bucket.ds_by_name, ds_name, qty)
        _add_units(bucket.ds_units, row)
        _fill_identity(bucket, row)

    packing_rows = packing_dataset.rows if packing_dataset is not None else []
    for row in packing_rows:
        qty = _qty(row)
        if qty is None:
            continue
        key, bucket = _bucket_for_row(row, buckets, anon_seq)
        if key[0] == "anon":
            anon_seq += 1
        source_file, _sheet, _excel_row = packing_row_source(row)
        if not source_file:
            source_file = _field(row, ANNOTATION)
        ul_file = os.path.basename(str(source_file).strip()) or _EMPTY_UL
        _add_qty(bucket.ul_by_file, ul_file, qty)
        _add_units(bucket.ul_units, row)
        _fill_identity(bucket, row)

    rows_out: list[BccAccumRow] = []
    mto_cols: set[str] = set()
    ds_cols: set[str] = set()
    ul_cols: set[str] = set()

    for bucket in buckets.values():
        _apply_google(bucket, identity)
        row = _finalize_row(bucket)
        rows_out.append(row)
        mto_cols.update(row.mto_by_title)
        ds_cols.update(row.ds_by_name)
        ul_cols.update(row.ul_by_file)

    rows_out.sort(
        key=lambda item: (
            0 if item.code else 1,
            item.code.casefold(),
            item.name.casefold(),
            item.type_mark.casefold(),
        )
    )
    return BccAccumMatrix(
        rows=rows_out,
        mto_title_columns=sorted(mto_cols),
        ds_name_columns=sorted(ds_cols),
        ul_file_columns=sorted(ul_cols),
    )


def save_bcc_accum_matrix_to_excel(
    matrix: BccAccumMatrix,
    result_dir: str,
) -> str:
    """Write ``Шаг4_Матрица_BCC_{timestamp}.xlsx`` next to Step4 output.

    Args:
        matrix: Aggregated matrix from ``build_bcc_accum_matrix``.
        result_dir: Destination folder (may be UNC).

    Returns:
        Absolute path of the published workbook.

    Raises:
        OSError: If the destination cannot be created or copied.
    """
    tmp_path = None
    publish_ok = False
    try:
        ensure_result_dir_exists(result_dir)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        file_name = f"Шаг4_Матрица_BCC_{timestamp}.xlsx"
        file_path = os.path.join(result_dir, file_name)
        tmp_fd, tmp_path = tempfile.mkstemp(prefix="step4_bcc_", suffix=".xlsx")
        os.close(tmp_fd)

        wb = xlsxwriter.Workbook(
            tmp_path,
            {"constant_memory": False, "strings_to_urls": False},
        )
        ws = wb.add_worksheet("Матрица BCC")
        num_cols = len(_FIXED_HEADERS)

        header_formats: dict[str, xlsxwriter.workbook.Format] = {}
        for key, color in (
            ("identity", _HEADER_IDENTITY),
            ("totals", _HEADER_TOTALS),
        ):
            header_formats[key] = wb.add_format(
                {
                    "bold": True,
                    "font_color": "#000000",
                    "bg_color": color,
                    "align": "center",
                    "valign": "vcenter",
                    "border": 1,
                    "text_wrap": True,
                }
            )
        cell_fmt = wb.add_format(
            {"align": "left", "valign": "vcenter", "border": 1}
        )
        wrap_fmt = wb.add_format(
            {
                "align": "left",
                "valign": "top",
                "border": 1,
                "text_wrap": True,
            }
        )
        num_fmt = wb.add_format(
            {"align": "right", "valign": "vcenter", "border": 1}
        )
        yellow_units = wb.add_format(
            {
                "align": "left",
                "valign": "vcenter",
                "border": 1,
                "bg_color": f"#{Color.yellow}",
            }
        )
        diff_zero = wb.add_format(
            {
                "align": "right",
                "valign": "vcenter",
                "border": 1,
                "bg_color": f"#{Color.match_matched}",
            }
        )
        diff_nonzero = wb.add_format(
            {
                "align": "right",
                "valign": "vcenter",
                "border": 1,
                "bg_color": f"#{Color.yellow}",
            }
        )

        for col_idx, label in enumerate(_FIXED_HEADERS):
            is_totals = _COL_QTY_FIRST <= col_idx <= _COL_DIFF_LAST
            kind = "totals" if is_totals else "identity"
            ws.write(0, col_idx, label, header_formats[kind])
        ws.set_row(0, 30)

        t_write_start = time.perf_counter()
        for row_i, item in enumerate(matrix.rows, start=1):
            values: list[object] = [
                item.code,
                item.name,
                item.type_mark,
                item.vendor,
                item.units,
                item.qty_mto,
                item.qty_ds,
                item.qty_ul,
                item.diff_mto_ds,
                item.diff_ds_ul,
                item.diff_mto_ul,
                item.mto_titles_info,
                item.ds_sources_info,
                item.ul_sources_info,
            ]

            info_line_count = max(
                item.mto_titles_info.count("\n") + 1 if item.mto_titles_info else 1,
                item.ds_sources_info.count("\n") + 1 if item.ds_sources_info else 1,
                item.ul_sources_info.count("\n") + 1 if item.ul_sources_info else 1,
            )
            for col_idx, value in enumerate(values):
                if col_idx == _COL_UNITS:
                    ws.write(
                        row_i,
                        col_idx,
                        value,
                        yellow_units if item.units_mismatch else cell_fmt,
                    )
                    continue
                if _COL_WRAP_FIRST <= col_idx <= _COL_WRAP_LAST:
                    ws.write(row_i, col_idx, value or "", wrap_fmt)
                    continue
                if _COL_QTY_FIRST <= col_idx <= _COL_QTY_LAST:
                    ws.write_number(row_i, col_idx, float(value or 0), num_fmt)
                    continue
                if _COL_DIFF_FIRST <= col_idx <= _COL_DIFF_LAST:
                    if value is None:
                        ws.write_blank(row_i, col_idx, None, cell_fmt)
                    elif float(value) == 0:
                        ws.write_number(row_i, col_idx, 0, diff_zero)
                    else:
                        ws.write_number(row_i, col_idx, float(value), diff_nonzero)
                    continue
                ws.write(row_i, col_idx, value if value is not None else "", cell_fmt)
            if info_line_count > 1:
                ws.set_row(row_i, min(15.0 * info_line_count + 2.0, 90.0))

        widths = {
            0: 14,
            1: 28,
            2: 18,
            3: 16,
            4: 10,
            11: 40,
            12: 40,
            13: 40,
        }
        for col_idx in range(num_cols):
            ws.set_column(col_idx, col_idx, widths.get(col_idx, 12))

        last_row = max(len(matrix.rows), 0)
        if last_row > 0 and num_cols > 0:
            ws.autofilter(0, 0, last_row, num_cols - 1)
        ws.freeze_panes(1, 5)

        write_bcc_accum_summary_sheet(wb, matrix)

        t_write = time.perf_counter() - t_write_start
        append_timing_log(
            result_dir, f"save_bcc_accum_matrix_to_excel::write: {t_write:.3f}s"
        )
        t_close_start = time.perf_counter()
        wb.close()
        append_timing_log(
            result_dir,
            f"save_bcc_accum_matrix_to_excel::wb_close: "
            f"{time.perf_counter() - t_close_start:.3f}s",
        )

        t_copy_start = time.perf_counter()
        if _path_is_unc(file_path):
            print("Накопительная матрица BCC: копирование на сетевой диск...", flush=True)
        shutil.copyfile(tmp_path, file_path)
        publish_ok = True
        append_timing_log(
            result_dir,
            f"save_bcc_accum_matrix_to_excel::copy_to_dest: "
            f"{time.perf_counter() - t_copy_start:.3f}s",
        )
        print(
            f"Накопительная матрица BCC: {len(matrix.rows)} строк -> {file_name}",
            flush=True,
        )
        return file_path
    except Exception:
        if tmp_path and os.path.isfile(tmp_path) and not publish_ok:
            print(f"Локальная копия матрицы BCC оставлена: {tmp_path}")
            tmp_path = None
        raise
    finally:
        if tmp_path:
            try:
                os.remove(tmp_path)
            except OSError:
                pass


def _field(row: RowStd, column: str) -> str:
    element = row.el.get(column)
    if element is None:
        return ""
    return str(element.value or "").strip()


def _qty(row: RowStd) -> float | None:
    if row.row_type != RowType.position_row:
        return None
    element = row.el.get(VALUES)
    if element is None:
        return None
    raw = element.value
    if raw is None or isinstance(raw, bool):
        return None
    if isinstance(raw, (int, float)):
        quantity = float(raw)
        return quantity if math.isfinite(quantity) else None
    text = str(raw).strip().replace(",", ".")
    if not text:
        return None
    try:
        quantity = float(text)
    except ValueError:
        return None
    return quantity if math.isfinite(quantity) else None


def _empty_identity_key(row: RowStd) -> tuple | None:
    name = _field(row, NAME)
    type_mark = _field(row, TYPE_MARK)
    units = normalize_units_text(_field(row, UNITS))
    vendor = _field(row, VENDOR)
    if name or type_mark or units or vendor:
        return ("empty", name, type_mark, units, vendor)
    return None


def _bucket_for_row(
    row: RowStd,
    buckets: dict[tuple, _Bucket],
    anon_seq: int,
) -> tuple[tuple, _Bucket]:
    code = _field(row, CODE)
    if code:
        key: tuple = ("code", code)
    else:
        empty_key = _empty_identity_key(row)
        if empty_key is not None:
            key = empty_key
        else:
            key = ("anon", anon_seq)
    bucket = buckets.get(key)
    if bucket is None:
        bucket = _Bucket(code=code)
        buckets[key] = bucket
    return key, bucket


def _add_qty(target: dict[str, float], key: str, qty: float) -> None:
    target[key] = target.get(key, 0.0) + qty


def _add_units(target: set[str], row: RowStd) -> None:
    units = normalize_units_text(_field(row, UNITS))
    if units:
        target.add(units)


def _fill_identity(bucket: _Bucket, row: RowStd) -> None:
    if not bucket.name:
        bucket.name = _field(row, NAME)
    if not bucket.type_mark:
        bucket.type_mark = _field(row, TYPE_MARK)
    if not bucket.units:
        bucket.units = _field(row, UNITS)
    if not bucket.vendor:
        bucket.vendor = _field(row, VENDOR)


def _apply_google(bucket: _Bucket, identity: dict[str, GoogleIdentity]) -> None:
    if not bucket.code:
        return
    entry = identity.get(bucket.code)
    if entry is None or not any(str(part or "").strip() for part in entry):
        return
    g_name, g_type, g_units, g_vendor = (_norm_part(part) for part in entry)
    bucket.name = g_name or bucket.name
    bucket.type_mark = g_type or bucket.type_mark
    bucket.units = g_units or bucket.units
    bucket.vendor = g_vendor or bucket.vendor


def _norm_part(value: object) -> str:
    return str(value or "").strip()


def _single_units(values: set[str]) -> tuple[str, bool]:
    if len(values) > 1:
        return "", True
    if len(values) == 1:
        return next(iter(values)), False
    return "", False


def _diff(
    qty_a: float,
    units_a: str,
    conflict_a: bool,
    qty_b: float,
    units_b: str,
    conflict_b: bool,
) -> float | None:
    a_has = qty_a != 0
    b_has = qty_b != 0
    if a_has and b_has:
        if conflict_a or conflict_b:
            return None
        if units_a and units_b and units_a != units_b:
            return None
        return qty_a - qty_b
    if a_has:
        return None if conflict_a else qty_a - qty_b
    if b_has:
        return None if conflict_b else qty_a - qty_b
    return 0.0


def write_bcc_accum_summary_sheet(
    wb: xlsxwriter.Workbook,
    matrix: BccAccumMatrix,
) -> None:
    """Write the «Сводка» worksheet: KPI, coverage tables, discrepancies.

    Args:
        wb: Open xlsxwriter workbook; a new worksheet is added.
        matrix: Aggregated matrix from ``build_bcc_accum_matrix``.
    """
    ws = wb.add_worksheet("Сводка")

    kpi_header_fmt = wb.add_format(
        {
            "bold": True,
            "bg_color": _HEADER_TOTALS,
            "border": 1,
            "align": "center",
            "valign": "vcenter",
        }
    )
    coverage_header_fmts = {
        "mto": wb.add_format(
            {
                "bold": True,
                "bg_color": _HEADER_MTO,
                "border": 1,
                "align": "center",
                "valign": "vcenter",
            }
        ),
        "ds": wb.add_format(
            {
                "bold": True,
                "bg_color": _HEADER_DS,
                "border": 1,
                "align": "center",
                "valign": "vcenter",
            }
        ),
        "ul": wb.add_format(
            {
                "bold": True,
                "bg_color": _HEADER_UL,
                "border": 1,
                "align": "center",
                "valign": "vcenter",
            }
        ),
    }
    cell_fmt = wb.add_format({"border": 1, "valign": "vcenter"})
    num_fmt = wb.add_format({"border": 1, "align": "right", "valign": "vcenter"})
    diff_zero = wb.add_format(
        {
            "border": 1,
            "align": "right",
            "valign": "vcenter",
            "bg_color": f"#{Color.match_matched}",
        }
    )
    diff_nonzero = wb.add_format(
        {
            "border": 1,
            "align": "right",
            "valign": "vcenter",
            "bg_color": f"#{Color.yellow}",
        }
    )
    title_fmt = wb.add_format({"bold": True})

    ws.write(0, 0, "Показатель", kpi_header_fmt)
    ws.write(0, 1, "Значение", kpi_header_fmt)

    kpi_rows = _build_bcc_summary_kpi(matrix)
    for row_idx, (label, value) in enumerate(kpi_rows, start=1):
        ws.write(row_idx, 0, label, cell_fmt)
        ws.write(row_idx, 1, value, num_fmt)

    coverage_sections = (
        ("mto", "Титулы МТО", 3, matrix.mto_title_columns, "mto_by_title"),
        ("ds", "ДС", 6, matrix.ds_name_columns, "ds_by_name"),
        ("ul", "УЛ", 9, matrix.ul_file_columns, "ul_by_file"),
    )
    for kind, title, start_col, names, mapping_attr in coverage_sections:
        header_fmt = coverage_header_fmts[kind]
        ws.write(0, start_col, title, header_fmt)
        ws.write(1, start_col, "Имя", header_fmt)
        ws.write(1, start_col + 1, "Кодов", header_fmt)
        ws.write(1, start_col + 2, "Σ кол-во", header_fmt)
        stats = _build_bcc_coverage_stats(matrix.rows, names, mapping_attr)
        for offset, (name, code_count, qty_sum) in enumerate(stats, start=2):
            ws.write(offset, start_col, name, cell_fmt)
            ws.write_number(offset, start_col + 1, code_count, num_fmt)
            ws.write_number(offset, start_col + 2, qty_sum, num_fmt)

    coverage_depth = max(
        (
            len(matrix.mto_title_columns),
            len(matrix.ds_name_columns),
            len(matrix.ul_file_columns),
        ),
        default=0,
    )
    disc_title_row = max(14, coverage_depth + 4)
    disc_header_row = disc_title_row + 1
    disc_first_data_row = disc_header_row + 1

    ws.write(disc_title_row, 0, "Расхождения", title_fmt)
    disc_headers = (
        "Код",
        "Наименование",
        "Σ МТО",
        "Σ ДС",
        "Σ УЛ",
        "МТО − ДС",
        "ДС − УЛ",
        "МТО − УЛ",
        "Конфликт ед. изм.",
    )
    for col_idx, label in enumerate(disc_headers):
        ws.write(disc_header_row, col_idx, label, kpi_header_fmt)

    disc_rows = _build_bcc_discrepancy_rows(matrix.rows)
    for offset, item in enumerate(disc_rows):
        excel_row = disc_first_data_row + offset
        ws.write(excel_row, 0, item.code, cell_fmt)
        ws.write(excel_row, 1, item.name, cell_fmt)
        ws.write_number(excel_row, 2, item.qty_mto, num_fmt)
        ws.write_number(excel_row, 3, item.qty_ds, num_fmt)
        ws.write_number(excel_row, 4, item.qty_ul, num_fmt)
        for col_idx, diff_value in (
            (5, item.diff_mto_ds),
            (6, item.diff_ds_ul),
            (7, item.diff_mto_ul),
        ):
            if diff_value is None:
                ws.write_blank(excel_row, col_idx, None, cell_fmt)
            elif diff_value == 0:
                ws.write_number(excel_row, col_idx, 0, diff_zero)
            else:
                ws.write_number(excel_row, col_idx, diff_value, diff_nonzero)
        ws.write(
            excel_row,
            8,
            "Да" if item.units_mismatch else "Нет",
            cell_fmt,
        )

    if disc_rows:
        ws.autofilter(
            disc_header_row,
            0,
            disc_header_row + len(disc_rows),
            len(disc_headers) - 1,
        )

    ws.set_column(0, 0, 28)
    ws.set_column(1, 1, 14)
    ws.set_column(3, 11, 14)


@dataclass(frozen=True)
class _BccDiscrepancyRow:
    code: str
    name: str
    qty_mto: float
    qty_ds: float
    qty_ul: float
    diff_mto_ds: float | None
    diff_ds_ul: float | None
    diff_mto_ul: float | None
    units_mismatch: bool


def _build_bcc_summary_kpi(matrix: BccAccumMatrix) -> list[tuple[str, int]]:
    rows = matrix.rows
    return [
        ("Строк всего", len(rows)),
        ("С кодом BCC", sum(1 for row in rows if row.code.strip())),
        ("Без кода", sum(1 for row in rows if not row.code.strip())),
        (
            "Только МТО",
            sum(
                1
                for row in rows
                if row.qty_mto != 0 and row.qty_ds == 0 and row.qty_ul == 0
            ),
        ),
        (
            "Только ДС",
            sum(
                1
                for row in rows
                if row.qty_ds != 0 and row.qty_mto == 0 and row.qty_ul == 0
            ),
        ),
        (
            "Только УЛ",
            sum(
                1
                for row in rows
                if row.qty_ul != 0 and row.qty_mto == 0 and row.qty_ds == 0
            ),
        ),
        (
            "Во всех трёх контурах",
            sum(
                1
                for row in rows
                if row.qty_mto != 0 and row.qty_ds != 0 and row.qty_ul != 0
            ),
        ),
        (
            "Ненулевой МТО − ДС",
            sum(
                1
                for row in rows
                if row.diff_mto_ds is not None and row.diff_mto_ds != 0
            ),
        ),
        (
            "Ненулевой ДС − УЛ",
            sum(
                1
                for row in rows
                if row.diff_ds_ul is not None and row.diff_ds_ul != 0
            ),
        ),
        (
            "Ненулевой МТО − УЛ",
            sum(
                1
                for row in rows
                if row.diff_mto_ul is not None and row.diff_mto_ul != 0
            ),
        ),
        ("Конфликт ед. изм.", sum(1 for row in rows if row.units_mismatch)),
    ]


def _build_bcc_coverage_stats(
    rows: list[BccAccumRow],
    names: list[str],
    mapping_attr: str,
) -> list[tuple[str, int, float]]:
    stats: list[tuple[str, int, float]] = []
    for name in sorted(names):
        code_count = 0
        qty_sum = 0.0
        for row in rows:
            mapping: dict[str, float] = getattr(row, mapping_attr)
            qty = mapping.get(name, 0)
            if qty not in (None, 0):
                code_count += 1
            qty_sum += qty or 0.0
        stats.append((name, code_count, qty_sum))
    return stats


def _build_bcc_discrepancy_rows(rows: list[BccAccumRow]) -> list[_BccDiscrepancyRow]:
    selected: list[_BccDiscrepancyRow] = []
    for row in rows:
        has_diff = any(
            diff is not None and diff != 0
            for diff in (row.diff_mto_ds, row.diff_ds_ul, row.diff_mto_ul)
        )
        if not row.units_mismatch and not has_diff:
            continue
        selected.append(
            _BccDiscrepancyRow(
                code=row.code,
                name=row.name,
                qty_mto=row.qty_mto,
                qty_ds=row.qty_ds,
                qty_ul=row.qty_ul,
                diff_mto_ds=row.diff_mto_ds,
                diff_ds_ul=row.diff_ds_ul,
                diff_mto_ul=row.diff_mto_ul,
                units_mismatch=row.units_mismatch,
            )
        )
    selected.sort(
        key=lambda item: abs(item.diff_mto_ds or 0),
        reverse=True,
    )
    return selected


def _format_qty_value(qty: float) -> str:
    if math.isfinite(qty) and qty == int(qty):
        return str(int(qty))
    text = f"{qty:.9g}"
    return text


def _format_qty_list(mapping: dict[str, float]) -> str:
    """Join ``{name} - {qty}`` pairs, one per line, names sorted."""
    parts: list[str] = []
    for name in sorted(mapping):
        qty = mapping[name]
        if qty == 0:
            continue
        parts.append(f"{name} - {_format_qty_value(qty)}")
    return "\n".join(parts)


def _finalize_row(bucket: _Bucket) -> BccAccumRow:
    qty_mto = sum(bucket.mto_by_title.values())
    qty_ds = sum(bucket.ds_by_name.values())
    qty_ul = sum(bucket.ul_by_file.values())
    mto_units, mto_conflict = _single_units(bucket.mto_units)
    ds_units, ds_conflict = _single_units(bucket.ds_units)
    ul_units, ul_conflict = _single_units(bucket.ul_units)
    contour_units = [u for u in (mto_units, ds_units, ul_units) if u]
    units_mismatch = bool(
        mto_conflict
        or ds_conflict
        or ul_conflict
        or (len(set(contour_units)) > 1)
    )
    return BccAccumRow(
        code=bucket.code,
        name=bucket.name,
        type_mark=bucket.type_mark,
        units=bucket.units,
        vendor=bucket.vendor,
        mto_titles_info=_format_qty_list(bucket.mto_by_title),
        ds_sources_info=_format_qty_list(bucket.ds_by_name),
        ul_sources_info=_format_qty_list(bucket.ul_by_file),
        qty_mto=qty_mto,
        qty_ds=qty_ds,
        qty_ul=qty_ul,
        diff_mto_ds=_diff(
            qty_mto, mto_units, mto_conflict, qty_ds, ds_units, ds_conflict
        ),
        diff_ds_ul=_diff(
            qty_ds, ds_units, ds_conflict, qty_ul, ul_units, ul_conflict
        ),
        diff_mto_ul=_diff(
            qty_mto, mto_units, mto_conflict, qty_ul, ul_units, ul_conflict
        ),
        units_mismatch=units_mismatch,
        mto_by_title=dict(bucket.mto_by_title),
        ds_by_name=dict(bucket.ds_by_name),
        ul_by_file=dict(bucket.ul_by_file),
    )


def _path_is_unc(path: str) -> bool:
    normalized = os.path.normpath(path)
    return normalized.startswith("\\\\") or normalized.startswith("//")
