
# Overall Assessment:
# The code handles file scanning, error logging, and data processing for ZIP codes derived from Excel files
# in a specific directory. It combines utility functions and data extraction to create a consolidated output while
# logging significant events and errors during execution.

# Key Improvement Opportunities:
# There are potential issues with error handling, particularly regarding unhandled exceptions.
# The code could also benefit from clearer separation of concerns and more modular design
# to enhance readability and maintainability.

import utils
from RFQ.ZIP.zip_code_rules import zip_code_rule
from RFQ.func_get_unique_row_list import get_unique_row_list
from base.base_classes import *
from base.base_excel_out import check_color_out
from base.base_google import load_base
from base.base_mto import get_mto_std_from_file
from base.tables_columns import ColNames, CODE, ZIP_CODE
from utils.error_log import ErrorLog
from utils.path import open_dir
from utils.prints_to_console import print_file_list_to_console

debug_print_files_list = 1
debug_print_progress = 1


def zip_start(dir_path):
    """
        Scans the specified directory for Excel files relevant to the project,
        processes them to extract unique data rows, and checks them against a Google database for ZIP code assignments.

        Input:
            dir_path (str): The path to the directory where the files are located.

        Output:
            Generates an Excel output file with processed data.

        Key Behaviors:
            - Collects MTO files, processes their data, and logs errors.
            - Outputs results to a specific directory.
    """
    ErrorLog.add_error("ЗИП сканирование фалов")  # Log the starting of the ZIP scan
    ErrorLog.exit_path = dir_path  # Set the exit path for errors
    ErrorLog.error_log = []  # Initialize the error log

    # Массив для хранения всех Документов (файлов) текущего проекта
    curr_proj = utils.path.get_files_single(dir_path,
                                            endswith=(".xlsx", ".XLSX"),
                                            forbidden_endswith=(".xls", ".XLS"),
                                            sub_folders=1)
    if curr_proj == -1:
        ErrorLog.add_error(f"По пути <{dir_path}> файлы не найдены.")
        ErrorLog.exit()

    ############################
    if debug_print_files_list:
        print_file_list_to_console(curr_proj)

    # 1 Открываем по списку curr_proj все файлы и добавляем в pre_zip_base
    pre_zip_base = []  # Initialize a list for holding unique data rows
    for document in curr_proj:
        if "MTO" in document.file_name:
            # print(f"    Open MTO - {document.file_name}")
            mto_base = get_mto_std_from_file(t_com=TableComments(file_full_path=document.file_full_path,
                                                                 dir_path="-1",
                                                                 tabel_class=t_com_init_cls.MTO
                                                                 )
                                             )
            # Основная функция по добавлению строк в общий список
            get_unique_row_list(mto_base, pre_zip_base)  # Собираем в уникальные строки по закупочному коду

    # Путь для вывода результата
    path_out_dir = utils.path.get_path_out_dir(dir_path)

    # 3 Загружаем гугл-базу
    code_base_std = load_base()

    # 4 Обработка базы pre_zip_base
    for row in pre_zip_base:  # Iterate over unique rows
        if isinstance(row, RowStd):  # Ensure the row is of the correct type
            code = row.el[CODE].value  # Extract the code value from the row
            row_code_base = RowStd.get_row_by_code(code, code_base_std)  # Retrieve corresponding row from code base

            if isinstance(row_code_base, RowStd):  # Check if a valid row was found in the code base
                zip_code = row_code_base.el[ZIP_CODE].value  # Get ZIP code from the found row
                row.el[ZIP_CODE].value = zip_code  # Assign the ZIP code to the current row
                zip_code_rule(zip_code, row)  # Validate the ZIP code against pre-defined rules
            else:
                ErrorLog.add_error(f"ERROR: {row.get_text_for_debug()}")  # Log error for missing code
        else:
            raise Exception  # Raise exception if the row is not valid

    # Вывод в файл
    column_dict = ColNames.ZIP.column_dict

    return_file_path = check_color_out(pre_zip_base,
                                       column_dict,
                                       path_out_dir,
                                       f"пре_ЗИП_из_{len(curr_proj)}_файлов",
                                       excel_template_check_color=r"templates\out_template_pre_ZIP.xlsx",
                                       row_index=1)

    open_dir(return_file_path)

    ErrorLog.add_log("ЗИП сканирование фалов ЗАВЕРШЕНО")


if __name__ in {"__main__"}:
    mto_path = r"\\bcc\eng\PrDoc\377_НИПИГАЗ\АГХК\КСБ\РД\8950\99_Вспомогательные\МТО ред. формат"
    # mto_path = r"C:\YandexDisk\темп\ЗИП"
    zip_start(mto_path)
