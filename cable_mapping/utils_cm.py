from os import listdir
from os.path import isfile, join

from cable_mapping.cm_clss.cm_CableTableRow import CableTableRow
from cable_mapping.cm_clss.cm_ColumnsCableTable import ColumnsCableTable
from cable_mapping.cm_clss.cm_TableCommentsCabMap import DocumentCabMap
from utils.string_parsing import print_att_list_table


def get_files_list(pdf_path, endswith=(".pdf", ".PDF")):
    curr_proj = []
    # Читаем все файлы из папки (без директорий)
    try:
        only_files = [f for f in listdir(pdf_path) if isfile(join(pdf_path, f))]
    except FileNotFoundError as e:
        print(e)
        return -1
    # Заносим все пути файлов и создаем массив документов - уникальное свойство - file_name (без пути).

    for next_file in only_files:
        for e in endswith:
            if next_file.endswith(e) and next_file[0] != "~":
                curr_proj.append(pdf_path + "\\" + next_file)

    # curr_proj = list(sorted(curr_proj, key=lambda x: x.doc_Number_for_sort))
    return curr_proj


def print_list_std_row_func(tpk_std: list[CableTableRow], dbg=1, max_len=25, title="",
                            row_dict=ColumnsCableTable.main_list):
    print_list = []
    for row in tpk_std:
        print_list.append(row.to_list(row_dict))
    print_att_list_table(print_list, max_len=max_len, title_row=row_dict, title=title)


def print_file_list_to_console_cab_map(cm_proj):
    # Вывод таблицы в консоль
    from prettytable import PrettyTable
    table = PrettyTable()
    table.field_names = ["Полный путь файла", "Имя файла", "Система"]
    table.border = 0
    table.align = "l"
    for document in cm_proj:
        if isinstance(document, DocumentCabMap):
            table.add_row([document.file_full_path, document.file_name, document.t_com.system])
        else:
            print(f"{document} not class DocumentCabMap")
    print(table)
