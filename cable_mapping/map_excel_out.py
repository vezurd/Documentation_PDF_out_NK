from __future__ import annotations

import openpyxl
from openpyxl.styles import Border, Side, Alignment
from openpyxl.comments import Comment

import utils.path
from utils.colors import Color
from cable_mapping.cm_clss.cm_CableTable import CableTable
from cable_mapping.cm_clss.cm_CableTableRow import CableTableRow
from cable_mapping.cm_clss.cm_ColumnsCableTable import ColumnsCableTable
from cable_mapping.constants import *

excel_template_check_color = r"templates\out_template_check_cj_color.xlsx"


def tpk_normalized_out(tpk_std: list[CableTableRow],
                       source_file_name: str,
                       out_dir: str,
                       col_list: tuple[str, ...] = ColumnsCableTable.tpk_normalized_list) -> str | None:
    """Writes normalized TPK rows to xlsx in ``out_dir`` (same folder as source)."""
    if not tpk_std:
        return None

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "ТПК normalized"
    ws.append(list(col_list))
    for row in tpk_std:
        ws.append(row.to_list(col_list))

    base_name = source_file_name.strip()
    if base_name.lower().endswith(".xlsx"):
        base_name = base_name[:-5]
    file_name_out = f"tpk_normalized_{base_name}.xlsx"
    file_name = utils.path.file_path_plus_file_name(out_dir, file_name_out)
    try:
        utils.path.make_dir(out_dir)
        wb.save(file_name)
        wb.close()
        print(f"\tDONE tpk_normalized Excel\n\t{file_name}\n")
        return file_name
    except PermissionError as err:
        print(err)
        print(f"\tFAIL tpk_normalized Excel\n\t{file_name}\n")
        wb.close()
        return None


def check_color_out_map(cable_cls: CableTable,
                        col_for_out_list: dict,
                        out_dir,
                        file_prefix,
                        excel_template_check_color=excel_template_check_color,
                        comments_dbg=0):

    # Загружаем шаблон Excel
    wb = openpyxl.load_workbook(excel_template_check_color)
    # GRAB THE ACTIVE WORKSHEET
    ws = wb.get_sheet_by_name("Кабельный журнал")
    base = cable_cls.table
    #
    ws_summ = wb.get_sheet_by_name("Сумма кабелей")
    summ_base = cable_cls.get_sum_table_list()
    #
    ws_freq = wb.get_sheet_by_name("Частотный анализ")
    freq_dict = cable_cls.get_frequency_tags_list()
    """
    Определяем форматирование в ячейках:
        horizontal (
        vertical
        wrap_text (True, False)

    HorizontalAlignmentsType
        "general", "left", "center", "right", "fill", "justify", "centerContinuous", "distributed"

    VerticalAlignmentsType
        "top", "center", "bottom", "justify", "distributed"

    """
    formating_dict = {
        CJ_CABLE_LABEL: ["center", "center", True],
        CJ_CABLE_FULL_TAG: ["center", "center", True],
        CJ_FROM: ["center", "center", True],
        CJ_FROM_TERMINAL: ["center", "center", True],
        CJ_WHERE: ["center", "center", True],
        CJ_WHERE_TERMINAL: ["center", "center", True],
        CJ_CABLE_MARK: ["center", "center", True],
        CABLE_LAYING_IN_CABIN: ["center", "center", True],
        CABLE_LAYING_IN_TRAY: ["center", "center", True],
        CABLE_LAYING_IN_CABLE_CHANEL: ["center", "center", True],
        CABLE_LAYING_IN_OPEN: ["center", "center", True],
        CABLE_LAYING_IN_FLUTED: ["center", "center", True],
        CABLE_LAYING_IN_METAL_HOUSE: ["center", "center", True],
        CABLE_LAYING_IN_TUBE: ["center", "center", True],
        CABLE_LAYING_IN_TRENCH: ["center", "center", True],
        CJ_TOTAL_LENGTH: ["center", "center", True],
        CJ_ANNOTAION: ["center", "center", True]
    }
    thin_border = Border(left=Side(style='thin'),
                         right=Side(style='thin'),
                         top=Side(style='thin'),
                         bottom=Side(style='thin'))
    """
        Вывод таблицы ТАБЛИЦА СОЕДИНЕНИЙ И ПОДКЛЮЧЕНИЙ
    """
    row_index = 3
    for ind, row in base.items():
        row_index += 1
        column_index = 0
        for att in col_for_out_list:
            column_index += 1
            cell = ws.cell(column=column_index, row=row_index)

            value = row.el[att].value
            cell.value = value

            # Добавляем комментарии в ячейку если они есть.
            cell_comment = row.el[att].comment

            if len(cell_comment) > 0:
                cell.comment = Comment(cell_comment, "Python_NK", width=400, height=150)
                cell.value = str(cell.value)
                if comments_dbg:
                    cell.value += "\n" + cell_comment

            if att in formating_dict:
                horizontal = formating_dict[att][0]
                vertical = formating_dict[att][1]
                wrap_text = formating_dict[att][2]
                cell.alignment = Alignment(horizontal=horizontal, vertical=vertical, wrap_text=wrap_text)
                cell.border = thin_border

            color_index = row.el[att].color
            if color_index != Color.no:
                try:
                    cell.fill = openpyxl.styles.PatternFill(start_color=color_index,
                                                            end_color=color_index,
                                                            fill_type="solid")
                except ValueError:
                    print(f"color_index: <{column_index}>; value: <{value}>")

    """
    Вывод таблицы СВОДКА КАБЕЛЕЙ И ПРОВОДОВ
    """
    horizontal = "center"
    vertical = "center"
    wrap_text = True
    row_index = 1
    for row in summ_base:
        row_index += 1
        column_index = 0
        for cell_value in row:
            column_index += 1

            cell = ws_summ.cell(column=column_index, row=row_index)
            cell.value = cell_value

            cell.alignment = Alignment(horizontal=horizontal, vertical=vertical, wrap_text=wrap_text)
            cell.border = thin_border

    """
    Вывод в таблицу Частоту тегов
    """
    row_index = 1
    for k, v in freq_dict:
        row_index += 1

        key_cell = ws_freq.cell(column=1, row=row_index)
        value_cell = ws_freq.cell(column=2, row=row_index)

        key_cell.value = k
        value_cell.value = v

        key_cell.alignment = Alignment(horizontal=horizontal, vertical=vertical, wrap_text=wrap_text)
        value_cell.alignment = Alignment(horizontal=horizontal, vertical=vertical, wrap_text=wrap_text)

        key_cell.border = thin_border
        value_cell.border = thin_border

    # Save the file
    base_file_name = base[1].t_com.file_full_name.strip(".xlsx")
    file_name = out_dir + file_prefix + "_" + base_file_name + ".xlsx"
    try:
        utils.path.make_dir(out_dir)
        wb.save(file_name)
        print(f"\tDONE check_cj_color base Excel\n\t{file_name}\n")
        return file_name
    except PermissionError as err:
        print(err)
        print(f"\tFAIL check_cj_color base Excel (excel_base_out)\n\t{file_name}\n")
        return None
