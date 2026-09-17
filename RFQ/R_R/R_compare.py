import openpyxl

import utils
from base.base_classes import *
from base.base_excel_out import check_color_out, excel_base_out
from base.base_mto import get_std_from_excel_file
from base.tables_columns import ColNames, TAGS, ROW_TYPE
from tags import tag_parser
from utils.error_log import ErrorLog
from utils.path import open_dir
from utils.prints_to_console import print_file_list_to_console
import base
import base.t_comm_initial_classes as t_com_init_cls

debug_print_files_list = 1
debug_print_progress = 1

"""
2025.04
По запросу от Рыжкова

"""


def r_start(dir_path):
    ErrorLog.add_log("Сканирование фалов")
    ErrorLog.exit_path = dir_path

    # Массив для хранения всех Документов (файлов) текущего проекта
    curr_proj = utils.path.get_files_single(dir_path,
                                            endswith=(".xlsx", ".XLSX"),
                                            forbidden_endswith=(".xls", ".XLS"),
                                            sub_folders=False)
    if curr_proj == -1:
        ErrorLog.add_error(f"По пути <{dir_path}> файлы не найдены.")
        ErrorLog.exit()

    ############################
    if debug_print_files_list:
        print_file_list_to_console(curr_proj)

    # 1 Открываем по списку curr_proj все файлы и добавляем в pre_zip_base
    pre_zip_base = []
    # Сброс для tag_analyzer
    tag_parser.reset()
    for document in curr_proj:
        if "RFP" in document.file_name:
            # print(f"    Open MTO - {document.file_name}")
            t_com = TableComments(file_full_path=document.file_full_path,
                                  dir_path="-1",
                                  tabel_class=t_com_init_cls.RFP,
                                  )
            mto_base = get_std_from_excel_file(t_com=t_com)
            RowStd.set_color_by_row_type_in_list(mto_base)
            for row in mto_base:
                tags = str(row.el[TAGS].value)
                tag_parser.add_context(tags, document.file_name, ColNames.RFP.sheet_name)

        else:
            out_text = ""
            wb = openpyxl.load_workbook(document.file_full_path, True, rich_text=False)
            try:
                for sheet in wb:
                    if sheet.sheet_state != "hidden":
                        ws = wb[sheet.title]
                        for row in ws.values:
                            for value in row:
                                out_text += str(value) + "; "
                        tag_parser.add_context(out_text, document.file_name, sheet.title)
            finally:
                wb.close()

    # Путь для вывода результата
    path_out_dir = utils.path.get_path_out_dir(dir_path)
    tag_parser.analyze(path_out_dir, main_doc_title="RFP", main_doc_signature="RFP")

    # Вывод в файл

    column_dict = ColNames.DS_zin.column_dict
    column_dict[99] = ROW_TYPE

    return_file_path = base.base_excel_out.check_color_out(mto_base,
                                                           column_dict,
                                                           path_out_dir,
                                                           "RFP",
                                                           row_index=0)
    excel_base_out(mto_base, dir_base=path_out_dir)

    if return_file_path:
        open_dir(return_file_path)
    else:
        ErrorLog.add_error(f"ERROR: OPEN_DIR: Путь для вывода файла не найден")

    ErrorLog.add_log("Сканирование фалов ЗАВЕРШЕНО")


if __name__ in {"__main__"}:
    mto_path = r"C:\YandexDisk\темп\RFQ_2"
    # mto_path = r"C:\YandexDisk\темп\ЗИП"
    r_start(mto_path)
