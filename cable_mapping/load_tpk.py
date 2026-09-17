import openpyxl

from base.base_class_std_table import STDTable
from cable_mapping.cm_clss.cm_CableLayingTypes import CableLayingTypes
from cable_mapping.cm_clss.cm_CableTable import CableTable
from cable_mapping.cm_clss.cm_ElementCT import ElementCT
from cable_mapping.cm_clss.cm_CableTableRow import CableTableRow
from cable_mapping.cm_clss.cm_ColumnsCableTable import ColumnsCableTable
from cable_mapping.cm_clss.cm_TableCommentsCabMap import TableCommentsCabMap, DocumentCabMap
from cable_mapping.constants import *
from cable_mapping.utils_cm import print_list_std_row_func
from utils.string_parsing import print_att_list_table

tpk_dict = [0, 1, 2, 3, 4, 5, 6, 7, 8, 10, 11]


def load_from_xlsx_file(doc: DocumentCabMap, dbg=0, start_index=2, max_len=25):
    # Открываем файл...

    print(f"\tОткрываем |{doc.file_full_path}| <load_from_xlsx_file>")
    wb = openpyxl.load_workbook(doc.file_full_path.strip(), True, rich_text=False)
    # grab the <sheet_name> worksheet or any active
    if doc.sheet_name == -1:
        ws = wb.active
    else:
        try:
            ws = wb.get_sheet_by_name(doc.sheet_name)
        except KeyError:
            print(f"Ошибка при открытии листа <{doc.sheet_name}> в файле <{doc.file_full_path}>")
            return []
    table_raw = []
    # Чтение данных из листа
    row_index = 0
    for row in ws.iter_rows():
        row_index += 1
        if row_index < start_index:
            continue
        raw_row = []  # Создаем переменную для внесения значений из строк
        cell_index = 0  # Считаем кол-во ячеек в строке
        for cell in row:
            if cell_index in tpk_dict:  # Проверяем что нужный столбец есть
                raw_row.append(cell.value)
            cell_index += 1
        table_raw.append(raw_row)  # Заносим ряд RowStd в список
    wb.close()
    # for i in table_raw_obj:
    #     for k in i.el:
    #         if k == "tags":
    #             if i.el[k].value != "":
    #                 print(k, i.el[k].value, i.el[k].struck_value)
    if dbg:
        print_att_list_table(table_raw, max_len=max_len)
    return table_raw


def tpk_raw_to_std(table_raw: list[list], t_com=TableCommentsCabMap(), dbg=0, max_len=25):
    tpk_std = []
    # tpk_kv_dict = ColumnsCableTable.TPK.column_dict

    for t_row in table_raw:
        if t_row[1] != -1:
            std_row = CableTableRow(t_com=t_com)
            for k, att in ColumnsCableTable.TPK.column_dict.items():
                value = t_row[k]
                std_row.el[att] = ElementCT(value)
            tpk_std.append(std_row)

    if dbg:
        print_list_std_row_func(tpk_std, max_len=max_len)

    return tpk_std


def tpk_normalize_laying_types(tpk_raw_std: list[CableTableRow],
                               cable_laying_class: CableLayingTypes,
                               dbg=0,
                               max_len=25):
    def is_material_row(c_row: CableTableRow):
        value = c_row.el[TPK_R9_CABLE_LAYING_TYPE].value
        if value is not None and value != "":
            return True
        else:
            return False

    def is_blank_row(c_row: CableTableRow):
        from_value = c_row.el[CJ_FROM].value
        laying_type_value = c_row.el[TPK_R9_CABLE_LAYING_TYPE].value

        if (from_value is None or from_value == "") and (laying_type_value is None or laying_type_value == ""):
            return True
        else:
            return False

    def is_full_row(c_row: CableTableRow):
        from_value = c_row.el[CJ_FROM].value
        laying_type_value = c_row.el[TPK_R9_CABLE_LAYING_TYPE].value

        if (from_value is not None and from_value != "") and (
                laying_type_value is not None and laying_type_value != ""):
            return True
        else:
            return False

    exit_flag = False
    tpk_std = []
    flag_new_row = False  # Следит за добавлением полных строчек
    row_index = 0
    for row in tpk_raw_std:
        row_index += 1
        # Если это строка с не пустым FROM и типом прокладки
        if is_full_row(row):
            # flag_new_row = True
            norm_row = row
            tpk_std.append(norm_row)

        # Если в строке указан способ прокладки - добавляем его значение в активную строку norm_row
        if is_material_row(row):
            laying_material = row.el[TPK_R9_CABLE_LAYING_TYPE].value
            if isinstance(laying_material, str):
                laying_material = laying_material.strip()
            laying_type = cable_laying_class.get_type_by_material(laying_material)
            if laying_type == -1:
                print(f"load_tpk.tpk_normalize_laying_types: ОШИБКА!"
                      f"   <{laying_material}> - нет в базе прокладки, номер строки {row_index}"
                      f"    ({row.t_com.file_full_name}), "
                      f"VALUE=[{row.el[CJ_FROM].value}], CABLE_LAYING_TYPE=[{row.el[TPK_R9_CABLE_LAYING_TYPE].value}]")
                exit_flag = True
            if not exit_flag:
                laying_value = row.el[TPK_R10_CABLE_LAYING_LENGTH].value
                # print(row_index, laying_material,laying_type, laying_value)
                norm_row.add_laying_type_value(laying_type, laying_value)
        # Просто информируем что есть пустая строка в ТПК
        if is_blank_row(row):
            if dbg:
                print("load_tpk.tpk_normalize_laying_types: Пустая строка", row_index)
    if exit_flag:
        exit(0)
    if dbg:
        print_list_std_row_func(tpk_std, max_len=max_len, title=" AFTER tpk_normalize_laying_types")

    return tpk_std


def tpk_normalize_column_from(tpk_std: list[CableTableRow]):
    row_index = 0
    for row in tpk_std:
        row_index += 1
        val = str(row.el[CJ_FROM].value).replace("*", "").strip()
        val2 = str(row.el[CJ_WHERE].value).replace("*", "").strip()
        val_list = val.split(" ", 1)
        val2_list = val2.split(" ", 1)

        if len(val_list) == 1:
            pass
        elif len(val_list) == 2:
            row.el[CJ_FROM].value = val_list[0]
            row.el[CJ_FROM_TERMINAL].value = val_list[1]
        else:
            raise Exception (f"ОШИБКА: tpk_normalize_column_from:\n"
                  f"   строка:{row_index}; <{val}> - нет обработки разбития на составные части")
        if len(val2_list) == 1:
            pass
        elif len(val2_list) == 2:
            row.el[CJ_WHERE].value = val2_list[0]
            row.el[CJ_WHERE_TERMINAL].value = val2_list[1]
        else:
            raise Exception(f"ОШИБКА: tpk_normalize_column_from:\n"
                  f"   строка:{row_index}; <{val2}> - нет обработки разбития на составные части")


def tpk_normalize(tpk_raw_std: list[CableTableRow], cable_laying_class: CableLayingTypes, dbg=0, max_len=25):

    tpk_std = tpk_normalize_laying_types(tpk_raw_std, cable_laying_class, dbg=dbg, max_len=max_len)
    tpk_normalize_column_from(tpk_std)

    # for row in tpk_std:
    #     row.el[CJ_CABLE_MARK].value = row.el[CJ_CABLE_MARK].value

    if dbg:
        print_list_std_row_func(tpk_std, max_len=max_len, title="tpk_normalize")
    return tpk_std
