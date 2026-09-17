"""
Fast Excel export for DS vs MTO (xlsxwriter, write-only), same data as
STDTable.to_excel_ds_vs_mto_spec / check_color_out, smaller and faster than openpyxl.

Column visibility, headers, widths, and outline: ``ds_vs_mto_excel_columns`` (like step4_6).
"""

from __future__ import annotations

import re
from typing import List, Sequence

import xlsxwriter

from base.base_classes import RowStd, RowType
from base.base_utils import tags_to_str
from base.tables_columns import (
    CODE_2,
    DS_PACKAGING_PRICE_EXCL_VAT,
    DS_SHIPPING_PRICE_EXCL_VAT,
    DS_TOTAL_PRICE_EXCL_VAT,
    DS_TOTAL_PRICE_INCL_VAT,
    DS_UNIT_PRICE_EXCL_VAT,
    DS_UNIT_PRICE_WITH_EXTRAS_EXCL_VAT,
    DS_VAT_AMOUNT,
    MTO_DS_DIFF,
    MTO_DS_RFQ_DIFF,
    RFQ_VALUES,
    UL_ORDERED_VALUES,
    UL_REMAINING_VALUES,
    UL_VALUES,
    VALUES,
    VALUES_2,
)
from RFQ.ds_compare.ds_vs_mto_excel_columns import (
    DEFAULT_COLUMN_WIDTH,
    DS_VS_MTO_OUTPUT_COLUMNS_CONFIG,
    ColumnDef,
    HEADER_FILL_COLORS,
    visible_columns,
)
from RFQ.ds_compare.ds_sources_preflight import SourceFileRecord, SourceManifest
from utils.colors import Color
import utils.path

# Same float coercion as check_color_out in base_excel_out.py (+ common purchase columns)
_FLOAT_ATTS = {
    DS_UNIT_PRICE_EXCL_VAT,
    VALUES,
    DS_SHIPPING_PRICE_EXCL_VAT,
    DS_PACKAGING_PRICE_EXCL_VAT,
    DS_UNIT_PRICE_WITH_EXTRAS_EXCL_VAT,
    DS_TOTAL_PRICE_EXCL_VAT,
    DS_VAT_AMOUNT,
    DS_TOTAL_PRICE_INCL_VAT,
    VALUES_2,
    MTO_DS_DIFF,
    RFQ_VALUES,
    MTO_DS_RFQ_DIFF,
    UL_ORDERED_VALUES,
    UL_VALUES,
    UL_REMAINING_VALUES,
}


_AGG_REPLACED_FROM_RE = re.compile(r"\[AGG_REPLACED_FROM:\s*([^\]]+)\]")


def _split_agg_replaced_codes(comment: object) -> list[str]:
    """Return source codes from ``[AGG_REPLACED_FROM: ...]`` comments."""
    text = str(comment or "")
    if not text:
        return []
    match = _AGG_REPLACED_FROM_RE.search(text)
    if not match:
        return []
    return [part.strip() for part in match.group(1).split(",") if part.strip()]


def _join_unique_codes_for_filter(base_code: object, extra_codes: list[str]) -> str | None:
    """Build a human-filterable code string without changing compare data."""
    codes: list[str] = []
    for raw in [base_code, *extra_codes]:
        text = str(raw or "").strip()
        if text and text not in codes:
            codes.append(text)
    return " / ".join(codes) if codes else None


def _cell_value_for_excel(
    row: RowStd,
    col_name: str,
    *,
    expand_aggregated_replacement_codes: bool = True,
) -> str | int | float | None:
    el = row.el.get(col_name)
    if el is None:
        return None
    value = el.value
    if col_name == CODE_2 and expand_aggregated_replacement_codes:
        extra_codes = _split_agg_replaced_codes(el.comment)
        if extra_codes:
            return _join_unique_codes_for_filter(value, extra_codes)
    if isinstance(value, list):
        return tags_to_str(value)
    if col_name in _FLOAT_ATTS:
        try:
            return float(value)
        except (ValueError, TypeError):
            return value
    return value


def _build_col_group_options(
    visible: list[ColumnDef], num_cols: int
) -> dict[int, dict]:
    """Excel outline: same rules as step4_6 save_match_result_to_excel."""
    col_group_opts: dict[int, dict] = {}
    for ci, d in enumerate(visible):
        if d.group_level is not None:
            col_group_opts[ci] = {
                "level": d.group_level,
                "hidden": d.group_collapsed,
            }
        elif d.group_collapsed:
            col_group_opts[ci] = {"level": 0, "hidden": True}

    def _is_outline_group_hidden(ci: int) -> bool:
        o = col_group_opts.get(ci)
        if not o or not o.get("hidden"):
            return False
        return int(o.get("level", 0)) >= 1

    prev_collapsed = False
    for ci in range(num_cols):
        is_collapsed = _is_outline_group_hidden(ci)
        if prev_collapsed and not is_collapsed:
            existing = col_group_opts.get(ci, {})
            merged = {**existing, "collapsed": True}
            if "level" not in merged:
                merged["level"] = 0
            col_group_opts[ci] = merged
        prev_collapsed = is_collapsed
    return col_group_opts


def _write_source_table(
    workbook,
    sheet_name: str,
    records: list[SourceFileRecord],
    *,
    metadata: list[tuple[str, str]] | None = None,
) -> None:
    """Write one source-file worksheet with optional cache metadata."""
    worksheet = workbook.add_worksheet(sheet_name)
    header_format = workbook.add_format(
        {
            "bold": True,
            "font_color": "#000000",
            "bg_color": "#D9EAF7",
            "border": 1,
            "align": "center",
            "valign": "vcenter",
        }
    )
    cell_format = workbook.add_format(
        {"border": 1, "align": "left", "valign": "top"}
    )
    metadata_format = workbook.add_format({"bold": True})

    row = 0
    for label, value in metadata or []:
        worksheet.write(row, 0, label, metadata_format)
        worksheet.write(row, 1, value)
        row += 1
    if metadata:
        row += 1

    header_row = row
    headers = ("Ключ", "Роль", "Имя файла", "Путь", "Изменён", "Статус", "Примечание")
    for column, header in enumerate(headers):
        worksheet.write(header_row, column, header, header_format)

    for record in records:
        row += 1
        values = (
            record.key,
            record.role,
            str(record.path).replace("\\", "/").rsplit("/", 1)[-1],
            record.path,
            record.modified_at,
            record.status,
            record.note,
        )
        for column, value in enumerate(values):
            worksheet.write(row, column, value, cell_format)

    worksheet.set_column(0, 0, 22)
    worksheet.set_column(1, 2, 20)
    worksheet.set_column(3, 3, 100)
    worksheet.set_column(4, 5, 22)
    worksheet.set_column(6, 6, 70)
    if records:
        worksheet.autofilter(header_row, 0, row, len(headers) - 1)
    worksheet.freeze_panes(header_row + 1, 0)


def _write_source_sheets(workbook, manifest: SourceManifest) -> None:
    """Append DS, MTO, and packing source worksheets."""
    _write_source_table(workbook, "Источники ДС", manifest.ds)
    _write_source_table(workbook, "Источники МТО", manifest.mto)
    _write_source_table(
        workbook,
        "Источники УЛ",
        manifest.packing,
        metadata=[
            ("Корневая папка", manifest.packing_root),
            ("Fingerprint", manifest.packing_fingerprint),
            ("Кэш", manifest.packing_cache_path),
            ("Свод", manifest.packing_summary_path),
        ],
    )


def save_ds_vs_mto_excel_xlsxwriter(
    base: List[RowStd],
    out_dir: str,
    file_prefix: str,
    suffix: str = "_xw",
    row_not_print_list: Sequence[RowType] = (
        RowType.empty_row,
        RowType.other_row,
        RowType.head_row,
    ),
    column_config: list[ColumnDef] | None = None,
    show_internal_column_names: bool = False,
    output_file_name: str | None = None,
    include_comments: bool = True,
    expand_aggregated_replacement_codes: bool = True,
    source_manifest: SourceManifest | None = None,
) -> str | None:
    """
    Write DS vs MTO rows with xlsxwriter.

    Args:
        base: Rows to export.
        out_dir: Output directory.
        file_prefix: File name prefix (e.g. ``РОБОТ_СРАВНЕНИЕ_``).
        suffix: Inserted before ``.xlsx`` (e.g. ``_xw``).
        row_not_print_list: Row types to skip.
        column_config: If None, uses ``DS_VS_MTO_OUTPUT_COLUMNS_CONFIG`` from
            ``ds_vs_mto_excel_columns.py`` (edit ``output`` / ``width`` / ``group_*`` there
            or pass a custom list of ``ColumnDef`` in the same order as DsVsMto).
        show_internal_column_names: If True, appends the program column name to
            each header, e.g. ``Кол-во\\n(VALUES)``. Intended for debugging.
        output_file_name: If set, used as the full output file name (``.xlsx`` optional).
        include_comments: If False, cell comments are not written (faster export).
        expand_aggregated_replacement_codes: If True, shows aggregated replacement
            source codes in ``CODE_2`` for human Excel filtering.
        source_manifest: Files used to build the report; written to three sheets.
    """
    if not base:
        return None

    cfg = column_config if column_config is not None else DS_VS_MTO_OUTPUT_COLUMNS_CONFIG
    visible = visible_columns(cfg)
    if not visible:
        print("DS vs MTO xlsxwriter: no columns (all output=False)")
        return None

    col_width: dict[int, float] = {}
    col_section: dict[int, str] = {}
    for ci, d in enumerate(visible):
        col_width[ci] = float(d.width or DEFAULT_COLUMN_WIDTH)
        col_section[ci] = d.section
    ncols = len(visible)
    col_names = [d.col_name for d in visible]

    data_rows: List[RowStd] = [
        r for r in base if r.row_type not in row_not_print_list
    ]

    base_name = str(base[0].t_com.file_name or "out").removesuffix(".xlsx").removesuffix(
        ".XLSX"
    )
    if output_file_name:
        out_name = output_file_name
        if not out_name.lower().endswith(".xlsx"):
            out_name = f"{out_name}.xlsx"
    else:
        out_name = f"{file_prefix}_{base_name}{suffix}.xlsx"
    file_name = utils.path.file_path_plus_file_name(out_dir, out_name)
    try:
        utils.path.make_dir(out_dir)
    except OSError as e:
        print(f"DS vs MTO xlsxwriter: cannot create {out_dir}: {e}")
        return None

    try:
        wb = xlsxwriter.Workbook(
            file_name, {"constant_memory": False, "remove_timezone": True}
        )
    except OSError as e:
        print(f"DS vs MTO xlsxwriter: cannot open workbook {file_name}: {e}")
        return None

    ws = wb.add_worksheet("DS vs MTO")

    custom_header_formats: dict[str, object] = {}

    def header_format_for_column(col_def: ColumnDef) -> object:
        color_key = (
            str(col_def.header_color).strip()
            if col_def.header_color
            else HEADER_FILL_COLORS.get(col_def.section, HEADER_FILL_COLORS["ds"])
        )
        if not color_key.startswith("#"):
            color_key = f"#{color_key}"
        cached = custom_header_formats.get(color_key)
        if cached is not None:
            return cached
        fmt = wb.add_format(
            {
                "bold": True,
                "font_color": "#000000",
                "bg_color": color_key,
                "align": "center",
                "valign": "vcenter",
                "border": 1,
                "text_wrap": True,
            }
        )
        custom_header_formats[color_key] = fmt
        return fmt

    def header_label_with_column_name(label: str, col_name: str) -> str:
        """Show user header and internal column name for MVP report tuning."""
        label_text = str(label or col_name).strip()
        return f"{label_text}\n({col_name})"

    base_cell_fmt = wb.add_format(
        {
            "align": "left",
            "valign": "vcenter",
            "border": 1,
        }
    )
    fill_format_cache: dict[str, object] = {}

    def get_fill_format(color_hex: str) -> object:
        if color_hex not in fill_format_cache:
            ch = str(color_hex).strip()
            if not ch.startswith("#"):
                ch = f"#{ch}"
            fill_format_cache[color_hex] = wb.add_format(
                {
                    "align": "left",
                    "valign": "vcenter",
                    "border": 1,
                    "bg_color": ch,
                }
            )
        return fill_format_cache[color_hex]

    if show_internal_column_names:
        ws.set_row(0, 72)
    else:
        ws.set_row(0, 54)
    for ci, d in enumerate(visible):
        label = d.header_label or d.col_name
        col_width[ci] = float(d.width or DEFAULT_COLUMN_WIDTH)
        hdr_fmt = header_format_for_column(d)
        if show_internal_column_names:
            label = header_label_with_column_name(label, d.col_name)
        ws.write(0, ci, label, hdr_fmt)

    n_data = len(data_rows)
    comment_write_errors = 0
    for r_i, row in enumerate(data_rows):
        excel_row = r_i + 1
        for ci, col_name in enumerate(col_names):
            value = _cell_value_for_excel(
                row,
                col_name,
                expand_aggregated_replacement_codes=expand_aggregated_replacement_codes,
            )
            el = row.el.get(col_name)
            if el is None:
                ws.write(excel_row, ci, value, base_cell_fmt)
                continue
            if el.color and el.color != Color.no:
                try:
                    fmt = get_fill_format(str(el.color))
                except (TypeError, ValueError):
                    fmt = base_cell_fmt
            else:
                fmt = base_cell_fmt
            ws.write(excel_row, ci, value, fmt)
            cmt = el.comment if include_comments else None
            if cmt and str(cmt).strip():
                try:
                    ws.write_comment(
                        excel_row, ci, str(cmt), {"width": 200, "height": 100}
                    )
                except Exception:
                    comment_write_errors += 1

    col_group_opts = _build_col_group_options(visible, ncols)
    for ci in range(ncols):
        w = col_width.get(ci, float(DEFAULT_COLUMN_WIDTH))
        opts = col_group_opts.get(ci)
        if opts:
            ws.set_column(ci, ci, w, None, opts)
        else:
            ws.set_column(ci, ci, w, None, {"level": 0})

    if n_data > 0:
        ws.autofilter(0, 0, n_data, ncols - 1)
    ws.freeze_panes(1, 0)
    if source_manifest is not None:
        _write_source_sheets(wb, source_manifest)
    try:
        wb.close()
    except OSError as e:
        print(f"DS vs MTO xlsxwriter: save failed {e}")
        return None
    if comment_write_errors:
        print(
            "DS vs MTO xlsxwriter WARNING: "
            f"не записано комментариев={comment_write_errors}; файл={file_name}"
        )
    try:
        print(
            f"\tDS vs MTO (xlsxwriter): {file_name}\n"
        )
    except (UnicodeEncodeError, OSError):
        print(
            f"\tDS vs MTO (xlsxwriter): {ascii(file_name)}\n"
        )
    return file_name


def _output_name_without_comments(output_file_name: str) -> str:
    """Insert ``_без_комм`` before ``.xlsx`` in a full output file name."""
    name = output_file_name.strip()
    if name.lower().endswith(".xlsx"):
        return f"{name[:-5]}_без_комм.xlsx"
    return f"{name}_без_комм.xlsx"


def _suffix_without_comments(suffix: str) -> str:
    base = suffix or ""
    if base.endswith("_без_комм"):
        return base
    return f"{base}_без_комм"


def save_ds_vs_mto_excel_per_export_mode(
    base: List[RowStd],
    out_dir: str,
    file_prefix: str,
    *,
    excel_export_mode: str = "both",
    suffix: str = "_xw",
    row_not_print_list: Sequence[RowType] = (
        RowType.empty_row,
        RowType.other_row,
        RowType.head_row,
    ),
    column_config: list[ColumnDef] | None = None,
    show_internal_column_names: bool = False,
    output_file_name: str | None = None,
    expand_aggregated_replacement_codes: bool = True,
    source_manifest: SourceManifest | None = None,
) -> str | None:
    """Write one or two xlsx files depending on ``excel_export_mode``.

    Args:
        excel_export_mode: ``both``, ``with_comments``, or ``without_comments``.

    Returns:
        Path of the primary file (with comments if ``both``), or the only file written.
    """
    from RFQ.ds_compare.ds_compare_config import normalize_excel_export_mode

    mode = normalize_excel_export_mode(excel_export_mode)
    common = {
        "base": base,
        "out_dir": out_dir,
        "file_prefix": file_prefix,
        "row_not_print_list": row_not_print_list,
        "column_config": column_config,
        "show_internal_column_names": show_internal_column_names,
        "expand_aggregated_replacement_codes": expand_aggregated_replacement_codes,
        "source_manifest": source_manifest,
    }
    primary: str | None = None

    if mode in ("both", "with_comments"):
        primary = save_ds_vs_mto_excel_xlsxwriter(
            **common,
            suffix=suffix,
            output_file_name=output_file_name,
            include_comments=True,
        )

    if mode in ("both", "without_comments"):
        if output_file_name:
            no_cmt_name = _output_name_without_comments(output_file_name)
            no_cmt_suffix = suffix
        else:
            no_cmt_name = None
            no_cmt_suffix = (
                suffix if mode == "without_comments" else _suffix_without_comments(suffix)
            )
        path_no_cmt = save_ds_vs_mto_excel_xlsxwriter(
            **common,
            suffix=no_cmt_suffix,
            output_file_name=no_cmt_name,
            include_comments=False,
        )
        if mode == "without_comments":
            primary = path_no_cmt

    return primary
