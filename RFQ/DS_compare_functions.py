from typing import List

import openpyxl

from RFQ.func_get_unique_row_list import get_unique_row_list
from base.base_cheks import *
from base.base_class_std_table import STDTable
from base.base_excel_out import check_color_out
from base.tables_columns import CODE, TAGS, VALUES, NUMBERS, TYPE_MARK, NAME, TAGS_2, VALUES_2, NUMBERS_2, TYPE_MARK_2, \
    NAME_2, DIFF_FLAG_TAG, DIFF_FLAG_VALUE, DIFF_FLAG_NAME, DIFF_FLAG_TYPE
from utils.colors import Color
from utils.path import open_dir
from collections import Counter
from base.tables_columns import ColNames
from utils.string_parsing import print_att_list_table


def file_name_remove_junk(in_file_name: str):
    """
    Для сокращения имени файла в MTO_vs_MTO и имени файла
    убираем лишнее:
        Номер договора
        Расширение файл
        Язык документа
        МТО-0001
    :param in_file_name:
    :return:
    """
    in_file_name = in_file_name.replace("AGCC.287-", "").strip(".xlsx")
    in_file_name = in_file_name.replace("_RU", "")
    in_file_name = in_file_name.replace(".MTO-0001", "")

    return in_file_name



def ds_mto_compare(all_base_unique: List[List[RowStd]], path_out_dir, open_dir_flag=1):
    base_compare = []  # Список в который будем добавлять строки для вывода
    #
    # Строка заголовка таблицы
    title_row = RowStd()
    mto1_name = file_name_remove_junk(all_base_unique[0][0].t_com.file_name)
    mto2_name = file_name_remove_junk(all_base_unique[1][0].t_com.file_name)
    title_row.el[TAGS].value = f"Теги из \n{mto1_name}"
    title_row.el[TAGS_2].value = f"Теги из \n{mto2_name}"
    title_row.el[NAME].value = (f"Наименование по\n"
                                f"{mto1_name}")
    title_row.el[NAME_2].value = (f"Наименование по\n"
                                  f"{mto2_name}")
    title_row.el[TYPE_MARK].value = (f"Тип, марка, обозначение по\n"
                                     f"{mto1_name}")
    title_row.el[TYPE_MARK_2].value = (f"Тип, марка, обозначение по\n"
                                       f"{mto2_name}")
    title_row.el[CODE].value = "Закупочный код"
    title_row.el[VALUES].value = f"Кол-во по\n {mto1_name}"
    title_row.el[VALUES_2].value = f"Кол-во по\n {mto2_name}"
    title_row.el[NUMBERS].value = f"Номер(а) строк в\n {mto1_name}"
    title_row.el[DS_NUMBER].value = f"Номер(а) ДС"
    title_row.el[NUMBERS_2].value = f"Номер(а) строк в\n {mto2_name}"
    title_row.el[DIFF_FLAG_TAG].value = f"Различия Тегов"
    title_row.el[DIFF_FLAG_VALUE].value = f"Различия Кол-ва"
    title_row.el[DIFF_FLAG_NAME].value = f"Различия Наименований"
    title_row.el[DIFF_FLAG_TYPE].value = f"Различия Тип, Марка, Обозначение"
    title_row.el[DIFF_FLAG_MATCH].value = f"Строки Совпадают"
    base_compare.append(title_row)
    #
    #
    mto_att_dict_list = [
        # Столбцы для МТО1:
        {
            TAGS: TAGS,
            VALUES: VALUES,
            NUMBERS: NUMBERS,
            NAME: NAME,
            TYPE_MARK: TYPE_MARK,
            CODE: CODE,
            DS_NUMBER: DS_NUMBER},
        # Столбцы для МТО2:
        {
            TAGS: TAGS_2,
            VALUES: VALUES_2,
            NUMBERS: NUMBERS_2,
            NAME: NAME_2,
            TYPE_MARK: TYPE_MARK_2,
            CODE: CODE
        }
    ]

    # Перебираем все базы (предполагаем что 2) и заносим значения в строки по коду
    # all_base_unique - список List[List[RowStd]] c базами МТО1, МТО2 в которых уже "схлопнуты" строчки по кодам БСС
    for i in range(len(all_base_unique)):
        for row in all_base_unique[i]:
            code = row.el[CODE].value
            #  Проверяем на пустое поле CODE
            #  Если пустое - создаем новую строку и заносим в нужные столбцы по mto_att_dict_list[i] значения
            if code == "":
                finding_row = RowStd(row.t_com)
                for att, column in mto_att_dict_list[i].items():
                    finding_row.el[column].value = row.el[att].value
                base_compare.append(finding_row)
                continue
            # Ищем в base_compare сроку по коду code
            # Если не нашли - то добавляем пустую строку в base_compare
            # Если нашли - то заносим значения в нужные столбцы по mto_att_dict_list[i]
            finding_row = RowStd.get_row_by_code(code, base_compare)
            if finding_row is None:
                finding_row = RowStd(row.t_com)
                base_compare.append(finding_row)
            #
            # Заносим значения в нужные столбцы по mto_att_dict_list[i]
            if isinstance(finding_row, RowStd):
                for att, column in mto_att_dict_list[i].items():
                    finding_row.el[column].value = row.el[att].value
            #
            else:
                print(f"\tERROR in TWO_MTO_COMPARE: <{code}> - code, <{finding_row}> code_check")
                print("Программа завершена с ошибкой")
                exit(0)

    # ПРОВЕРКИ И ПОДСВЕЧИВАНИЕ В BASE_COMPARE
    for row in base_compare:
        # Пропускаем строку заголовка
        if row.el[DIFF_FLAG_TAG].value == "Различия Тегов":
            continue
        #
        result_tags_diff_1 = tags_diff(row.el[TAGS].value, row.el[TAGS_2].value)
        result_tags_diff_2 = tags_diff(row.el[TAGS_2].value, row.el[TAGS].value)
        doubles_check_1 = doubles_check(row.el[TAGS].value)
        doubles_check_2 = doubles_check(row.el[TAGS_2].value)
        if not list_vs_list(row.el[TAGS].value, row.el[TAGS_2].value):
            row.el[TAGS].color = Color.green
            row.el[TAGS_2].color = Color.green
        else:
            row.el[TAGS].color = Color.yellow
            row.el[TAGS_2].color = Color.yellow
            row.el[DIFF_FLAG_TAG].value = 1
            err_text = (f"Теги не совпадают в сравниваемых МТО\n"
                        f"В {mto1_name} нет тегов ({len(result_tags_diff_1)} шт.):\n"
                        f"{result_tags_diff_1}\n\n"
                        f"В {mto2_name} нет тегов ({len(result_tags_diff_2)} шт.):\n"
                        f"{result_tags_diff_2}\n")

            row.el[TAGS].comment += err_text
            row.el[TAGS_2].comment += err_text
            row.el[DIFF_FLAG_TAG].comment += err_text
            if doubles_check_1:
                row.el[TAGS].color = Color.red
                err_text = f"\nДУБЛИРОВАНИЕ ТЕГОВ в {mto1_name}\n"
                for t in doubles_check_1:
                    err_text += f"  {t}\n"
                row.el[TAGS].comment += err_text
            if doubles_check_2:
                row.el[TAGS_2].color = Color.red
                err_text = f"\nДУБЛИРОВАНИЕ ТЕГОВ в {mto2_name}\n"
                for t in doubles_check_2:
                    err_text += f"  {t}\n"
                row.el[TAGS_2].comment += err_text
        #
        if row.el[NAME].value == row.el[NAME_2].value:
            row.el[NAME].color = Color.green
            row.el[NAME_2].color = Color.green
        else:
            row.el[NAME].color = Color.yellow
            row.el[NAME_2].color = Color.yellow
            row.el[DIFF_FLAG_NAME].value = 1
            err_text = f"Наименования не совпадают в сравниваемых МТО\n"
            row.el[NAME].comment += err_text
            row.el[NAME_2].comment += err_text
            row.el[DIFF_FLAG_NAME].comment += err_text
        #
        if row.el[TYPE_MARK].value == row.el[TYPE_MARK_2].value:
            row.el[TYPE_MARK].color = Color.green
            row.el[TYPE_MARK_2].color = Color.green
        else:
            row.el[TYPE_MARK].color = Color.yellow
            row.el[TYPE_MARK_2].color = Color.yellow
            row.el[DIFF_FLAG_TYPE].value = 1
            err_text = f"Наименования не совпадают в сравниваемых МТО\n"
            row.el[TYPE_MARK].comment += err_text
            row.el[TYPE_MARK_2].comment += err_text
            row.el[DIFF_FLAG_TYPE].comment += err_text
        #
        if str(row.el[VALUES].value) == str(row.el[VALUES_2].value):
            row.el[VALUES].color = Color.green
            row.el[VALUES_2].color = Color.green
        else:
            row.el[VALUES].color = Color.yellow
            row.el[VALUES_2].color = Color.yellow
            row.el[DIFF_FLAG_VALUE].value = 1
            err_text = f"Кол-во не совпадают в сравниваемых МТО\n"
            row.el[VALUES].comment += err_text
            row.el[VALUES_2].comment += err_text
            row.el[DIFF_FLAG_VALUE].comment += err_text
        #
        if row.el[CODE].value == "":
            row.el[CODE].color = Color.red
        #

    # Словарь для вывода столбцов в EXCEL
    column_out_dict = {
        0: TAGS,
        1: TAGS_2,
        2: NAME,
        3: NAME_2,
        4: TYPE_MARK,
        5: TYPE_MARK_2,
        6: CODE,
        7: VALUES,
        8: VALUES_2,
        9: NUMBERS,
        10: DS_NUMBER,
        11: NUMBERS_2,
        12: DIFF_FLAG_TAG,
        13: DIFF_FLAG_VALUE,
        14: DIFF_FLAG_NAME,
        15: DIFF_FLAG_TYPE,
        16: DIFF_FLAG_MATCH,
    }
    return_file_path = check_color_out(base_compare,
                                       column_out_dict,
                                       path_out_dir,
                                       f"{all_base_unique[1][0].t_com.file_name}_сравнение_с_ДС",
                                       excel_template_check_color=r"templates\out_template_DS_compare.xlsx",
                                       row_index=0,
                                       comment_height=300)
    if open_dir_flag:
        open_dir(return_file_path)
    return base_compare


def start(all_base: List[List[RowStd]], check_position_row=0):
    all_base_unique = []
    # print_base_pt(all_base)
    for base in all_base:
        """
        "Схлопываем" позиции из base(RowStd) в список уникальных по коду (CODE)
        """
        base_unique = get_unique_row_list(base, check_position_flag=check_position_row)
        #
        all_base_unique.append(base_unique)
        # print_base_pt(base_unique)

    #  Нахождение разницы между базами
    diff_base = []
    empty_row = ["", "", "", "", ""]
    longest_base = -1
    longest_base_index = -1
    for i in range(len(all_base_unique)):
        if len(all_base_unique[i]) > longest_base:
            longest_base_index = i
            longest_base = len(all_base_unique[i])

    for row in all_base_unique[longest_base_index]:
        code = row.el[CODE].value
        #  Проверяем на пустое поле CODE
        if code == "":
            diff_base.append(RowStd.get_row_copy(row, row.t_com))
            continue

        for i in range(len(all_base_unique)):
            flag = 0
            if i != longest_base_index:
                # c_base = all_base_unique[i]
                c_row = RowStd.get_row_by_code(code, all_base_unique[i])
                if isinstance(c_row, RowStd):
                    att_list = [VALUES, TAGS]
                    for att in att_list:
                        if att == TAGS:
                            flag = list_vs_list(row.el[att].value, c_row.el[att].value)
                        elif row.el[att].value == c_row.el[att].value:
                            flag = 0
                        else:
                            flag = 1
                elif c_row is None:
                    flag = 1
                    c_row = empty_row
                if flag == 1:
                    diff_base.append(row)
                    diff_base.append(c_row)
                    diff_base.append(empty_row)

    # print_base_pt(diff_base)
    return diff_base, all_base_unique


def list_vs_list(a: List, b: List):
    # if a is None and b is None:
    #     flag = 0
    if not a and not b:
        flag = 0
    else:
        try:
            if isinstance(a, List):
                a.sort()
            if isinstance(b, List):
                b.sort()
            if a == b:
                flag = 0
            else:
                flag = 1
        except TypeError:
            print("a", a)
            print("b", b)
            flag = 1
    return flag


def tags_diff(a: List, b: List):
    diff_list = []
    if not a and not b:
        return diff_list
    else:
        if not a:
            a = []
        if not b:
            b = []
        diff_list = list(set(a) - set(b))
    return diff_list


def doubles_check(checked_list: list[str]):
    doub_list = []
    counter = Counter(checked_list)
    doubles = {element: count for element, count in counter.items() if count > 1}
    if not doubles:
        pass
    else:
        for k, v in doubles.items():
            doub_list.append(f"{k}: {v} шт.")
    return doub_list


def print_base_pt(base: List[RowStd], dbg=1):
    if dbg:
        print()
    if len(base) == 0:
        if dbg:
            print(["0", "Не найдено различий", "0", "0", "0"])
        return [["0", "Не найдено различий", "0", "0", "0"]]
    if dbg:
        print(base[0].t_com.file_name)
    out_list = []
    att_list = [NUMBERS, CODE, VALUES, TAGS]
    from prettytable import PrettyTable
    table = PrettyTable()
    table.field_names = [NUMBERS, CODE, VALUES, TAGS, "File_name"]
    # out_list.append([NUMBERS, CODE, VALUES, TAGS, "File_name"])
    table.border = 1
    table.align = "l"
    for row in base:
        if isinstance(row, RowStd):
            table_row = []
            for att in att_list:
                table_row.append(str(row.el[att].value))
            table_row.append(row.t_com.file_name)
            table.add_row(table_row)
            out_list.append(table_row)
        else:
            table.add_row(row)
            out_list.append(row)
    if dbg:
        print(f"\n\tPRINT_BASE_PT:")
        print(table)
    return out_list


excel_template_RFQ_compare = r"templates\out_template_DS_compare.xlsx"


def excel_base_out(in_list, dir_base="base_check/"):
    in_list = print_base_pt(in_list, 0)
    # Загружаем шаблон Excel
    wb = openpyxl.load_workbook(excel_template_RFQ_compare)

    # grab the active worksheet
    ws = wb.active

    # Добавляем к шаблону строчки с данными
    for i in range(len(in_list)):
        try:
            ws.append(in_list[i])
        except TypeError:
            print(in_list[i])

    # Save the file
    try:
        wb.save(dir_base + "RFQ_diff_list" + ".xlsx")
    except PermissionError as err:
        print(err)
        print("\tFAIL RFQ DIFF LISR (excel_base_out)\n")
