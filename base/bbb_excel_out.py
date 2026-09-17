"""
Вывод результатов анализа BBB в Excel (xlsxwriter).
Использует конфигурацию столбцов из bbb_output_config.py.
"""

import os
from datetime import datetime
from typing import List

import xlsxwriter

from base.base_classes import RowStd, RowType
from base.base_utils import tags_to_str
from base.bbb_output_config import (
    ColumnDef, get_bbb_output_config, DEFAULT_COLUMN_WIDTH, HEADER_FILL_COLOR,
)
from utils.colors import Color


def save_bbb_to_excel(
    data: List[RowStd],
    doc_type: str,
    out_dir: str,
    file_prefix: str,
    row_not_print_list=(RowType.empty_row, RowType.head_row),
) -> str | None:
    """Сохраняет результат проверки BBB в Excel с заголовками и ширинами из конфигурации."""
    if not data:
        print(f"  Нет данных для вывода {doc_type}")
        return None

    config = get_bbb_output_config(doc_type)
    active_cols: List[ColumnDef] = [c for c in config if c.output]

    os.makedirs(out_dir, exist_ok=True)

    base_file_name = data[0].t_com.file_name.replace(".xlsx", "")
    timestamp = datetime.now().strftime("%Y_%m_%d_%H_%M")
    file_name = f"{file_prefix}_{base_file_name}_{timestamp}.xlsx"
    file_path = os.path.join(out_dir, file_name)

    wb = xlsxwriter.Workbook(file_path, {"constant_memory": False})
    ws = wb.add_worksheet(doc_type)

    header_fmt = wb.add_format({
        "bold": True,
        "font_color": "#000000",
        "bg_color": HEADER_FILL_COLOR,
        "align": "center",
        "valign": "vcenter",
        "border": 1,
        "text_wrap": True,
    })

    cell_base_fmt = wb.add_format({
        "align": "left",
        "valign": "vcenter",
        "border": 1,
        "text_wrap": True,
    })

    cell_center_fmt = wb.add_format({
        "align": "center",
        "valign": "vcenter",
        "border": 1,
        "text_wrap": True,
    })

    fill_format_cache: dict[str, xlsxwriter.workbook.Format] = {}

    def _get_fill_format(color_hex: str, centered: bool = False):
        key = f"{color_hex}_{centered}"
        if key not in fill_format_cache:
            fill_format_cache[key] = wb.add_format({
                "align": "center" if centered else "left",
                "valign": "vcenter",
                "border": 1,
                "text_wrap": True,
                "bg_color": f"#{color_hex}" if not color_hex.startswith("#") else color_hex,
            })
        return fill_format_cache[key]

    _CENTER_COLS = {
        "numbers", "tags", "code", "units", "values", "mass", "vendor",
        "type_mark", "bbb_marka", "bbb_revision", "bbb_work_code",
        "bbb_position", "bbb_consumption_rate",
        "bbb_unit_price", "bbb_total_price", "bbb_assembly",
        "bbb_labor_unit", "bbb_labor_total", "bbb_equipment_unit",
        "bbb_equipment_total", "bbb_labor_rate", "bbb_equipment_rate",
        "bbb_direct_cost_rate", "bbb_direct_cost_total",
    }

    _NUMERIC_COLS = {
        "values", "mass",
        "bbb_consumption_rate", "bbb_unit_price", "bbb_total_price",
        "bbb_labor_unit", "bbb_labor_total", "bbb_equipment_unit",
        "bbb_equipment_total", "bbb_labor_rate", "bbb_equipment_rate",
        "bbb_direct_cost_rate", "bbb_direct_cost_total",
    }

    ws.set_row(0, 45)
    for ci, col_def in enumerate(active_cols):
        ws.write(0, ci, col_def.header_label, header_fmt)
        w = col_def.width or DEFAULT_COLUMN_WIDTH
        ws.set_column(ci, ci, w)

    data_rows = [r for r in data if r.row_type not in row_not_print_list]

    def _to_numeric(val, col_name):
        if col_name not in _NUMERIC_COLS or not isinstance(val, str):
            return val
        if not val or '\n' in val or val.startswith('='):
            return val
        try:
            return float(val)
        except (ValueError, TypeError):
            return val

    for row_idx, row in enumerate(data_rows):
        excel_row = row_idx + 1
        for ci, col_def in enumerate(active_cols):
            col_name = col_def.col_name
            value = row.el[col_name].value
            if isinstance(value, list):
                value = tags_to_str(value)

            value = _to_numeric(value, col_name)

            el_color = row.el[col_name].color
            is_centered = col_name in _CENTER_COLS

            if el_color != Color.no:
                fmt = _get_fill_format(el_color, centered=is_centered)
            else:
                fmt = cell_center_fmt if is_centered else cell_base_fmt

            ws.write(excel_row, ci, value, fmt)

            cell_comment = row.el[col_name].comment
            if cell_comment:
                ws.write_comment(excel_row, ci, cell_comment,
                                 {"width": 400, "height": 150})

    total_rows = len(data_rows)
    if total_rows > 0:
        ws.autofilter(0, 0, total_rows, len(active_cols) - 1)
    ws.freeze_panes(1, 0)

    wb.close()
    print(f"\tЗапись в Excel: {total_rows} строк -> {file_name}")
    return file_path
