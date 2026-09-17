"""
Функция проверки кабельного журнала существующего.
Проверяем:
    1. Код кабеля (CJ_CABLE_CODE)
    2. Дополнительный код (CJ_ADDITIONAL_CODE)

"""
import tags.tag_classes
from utils.colors import Color

from cable_mapping.cm_clss.cm_CableTable import CableTable
from cable_mapping.cm_clss.cm_CableTableRow import CableTableRow
from cable_mapping.cm_functions import cm_from_where_terminal
from cable_mapping.constants import *
from cable_mapping.mapping.get_cj import get_map_table_temp_row
from cable_mapping.mapping.map_calsses import MapTable


def cj_check(cable_table_cls: CableTable, map_table: MapTable, dbg=0):
    i = 0
    if dbg:
        print("\nCJ_CHECK - Функция проверки кабельного журнала существующего\n")
    for index, row in cable_table_cls.table.items():
        i += 1
        if isinstance(row, CableTableRow):
            # Создаем временную строчку CableTableRow
            temp_row = CableTableRow(row.row_type, row.t_com, row.system)
            # Копируем значения полей из row
            att_list = [CJ_FROM, CJ_WHERE, CJ_CABLE_MARK]
            for att in att_list:
                temp_row.el[att].value = row.el[att].value

            # Ищем в MapTable сочетание для row и записываем:
            #   CJ_CABLE_CODE
            #   CJ_ADDITIONAL_CODE
            #   CJ_FROM_TERMINAL
            #   CJ_WHERE_TERMINAL
            get_map_table_temp_row(temp_row, map_table,False)
            temp_row.update()

            # Список сравниваемых атрибутов при проверке строки из
            compare_list = [CJ_CABLE_CODE, CJ_ADDITIONAL_CODE, CJ_FROM_TERMINAL, CJ_WHERE_TERMINAL]
            compare_list_shift = [CJ_CABLE_CODE, CJ_ADDITIONAL_CODE]
            for att in compare_list:
                att_l = att
                #
                reg_from_map_table = temp_row.el[att].value
                result = cm_from_where_terminal.regular_compare(row.el[att].value, reg_from_map_table)
                #
                # if row.el[att].value != temp_row.el[att].value:
                if not result:
                    if dbg:
                        print(f"{i} "
                              f"|{result}| "
                              f"|{row.el[att].value}| "
                              f"|{reg_from_map_table}| "
                              f"({row.get_mapping_signature()})")
                    if att in compare_list_shift:
                        att = CJ_CABLE_LABEL
                    row.el[att].color = Color.red
                    row.el[att].comment += f"|{temp_row.el[att].value}| \n(по MapTable ({att_l})\n"
                else:
                    if att in compare_list_shift:
                        att = CJ_CABLE_LABEL
                    #
                    if att == CJ_FROM_TERMINAL or att == CJ_WHERE_TERMINAL:
                        if "XT" in row.el[att].value:
                            row.el[att].color = Color.soft_yellow
                            row.el[att].comment += f"Проверить правильность клемм\n"
                    #
                    if row.el[att].color == Color.no:
                        row.el[att].color = Color.green


if __name__ in {"__main__"}:

    reg = r"X2 \(разъем RJ-45\)"
    text = "X2 (разъем RJ-45)"
    print(cm_from_where_terminal.regular_compare(text, reg), text, reg, "\n")

    text = "3252-PB-01-S-PI-1911"
    reg = "[TAG]"
    print(cm_from_where_terminal.regular_compare(text, reg), text, reg, "\n")

    text = "3252-S-PI-1911"
    reg = "[TAG]"
    print(cm_from_where_terminal.regular_compare(text, reg), text, reg, "\n")
