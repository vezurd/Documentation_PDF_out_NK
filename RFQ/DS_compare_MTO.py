import datetime

import RFQ.Value_Compare

import utils
import utils.path
import RFQ.DS_compare_functions
from RFQ.func_get_unique_row_list import get_unique_row_list
from base.base_classes import *
from base.base_mto import get_mto_std_from_file, get_std_from_excel_file
from pdf_parsing_v2_engine.document import V2Document
from utils import path
from utils.prints_to_console import print_file_list_to_console
from utils.save_to_file import save_string_to_file

debug_print_files_list = 1
debug_print_progress = 1

def ds_compare_folder_start(mto_folder, ds_path):
    # Массив для хранения всех Документов (файлов) текущего проекта
    curr_proj = utils.path.get_files_single(mto_folder, endswith=(".xlsx", ".XLSX"))
    if curr_proj == -1:
        print(f"По пути <{mto_folder}> файлы не найдены.")
        exit(0)

    ############################
    if debug_print_files_list:
        from prettytable import PrettyTable
        table = PrettyTable()
        table.field_names = ["Полный путь файла", "Имя файла"]
        table.border = 0
        table.align = "l"
        for document in curr_proj:
            table.add_row([document.file_full_path, document.file_name])
        print(table)

    # Обработка файла с ДС
    ds_t_com = TableComments(file_full_path=ds_path, dir_path="-1", tabel_class=t_com_init_cls.DSZin)
    print(f"    Open DS - {ds_path}")
    # Загружаем все строки
    ds_base_full = get_std_from_excel_file(ds_t_com)
    error_log = [datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")]
    print(error_log)
    for document in curr_proj:
        mto_path = document.file_full_path
        ds_compare_start(mto_path,ds_path,
                         ds_base_full=ds_base_full,info_flag=0, open_dir_flag=False,error_log=error_log)
    base_path = r"\\bcc\eng\PrDoc\377_НИПИГАЗ\АГХК\КСБ\РД\Таблички_графики\14_сравнение с ДС\\"
    base_path_err = r"\\bcc\eng\PrDoc\377_НИПИГАЗ\АГХК\КСБ\РД\Таблички_графики\14_сравнение с ДС\error_log.txt"
    error_string = ""
    for x in error_log:
        error_string = error_string + "\n"+ x

    save_string_to_file(error_string,base_path_err)
    print(error_log)
    path.open_dir(base_path)

def ds_compare_start(mto_path, ds_path,
                     info_flag=1,
                     ds_base_full=None,
                     open_dir_flag=1,
                     error_log=None):
    """

    :param ds_base_unique:
    :param mto_path:
    :param ds_path:
    :return:
    """

    # Что делать со строчками которые не проходят по проверке кода закупочного
    # 0 - добавляем все в сравнение, в т.ч. пустые строки
    # 1 - добавляем только прошедшие проверку на коды, только валидные строки
    # 2 - убираем пустые строки, не выалидные но не пустые - не убираем
    check_position_flag = 1

    if info_flag == 1:
        print("ДС vs MTO сравнение")
        print(f"Файл ДС:\t{ds_path}")
        print(f"МТО фал:\t{mto_path}")

    mto_obj = V2Document.from_file_path(mto_path)
    if info_flag:
        mto_obj.print_debug()


    if "MTO" == mto_obj.doc_Type:
        print(f"    Open MTO - {mto_obj.file_name}")
        mto_base = get_mto_std_from_file(t_com=TableComments(file_full_path=mto_obj.file_full_path,
                                                             dir_path="-1",
                                                             tabel_class=t_com_init_cls.MTO,
                                                             )
                                         , dbg=0)
        mto_base_unique = get_unique_row_list(mto_base, check_position_flag=check_position_flag)
    else:
        print(f"    Файл - {mto_obj.file_name} не содержит символов MTO в названии. Программа завершена")
        exit(0)

    if ds_base_full is None:
        # Обработка файла с ДС
        ds_t_com = TableComments(file_full_path=ds_path, dir_path="-1", tabel_class=t_com_init_cls.DSZin)
        print(f"    Open DS - {ds_path}")
        # Загружаем все строки
        ds_base_full = get_std_from_excel_file(ds_t_com)

    # Отсортировываем строки только относящиеся к нужной МТО
    ds_base = []
    for row in ds_base_full:
        ds_title = row.el[DS_TITLE].value
        ds_system = row.el[DS_SYSTEM].value
        if ds_title == mto_obj.doc_Title_4d and ds_system == mto_obj.doc_Marka:
            ds_base.append(row)

    # Схлопываем по уникальным кодам
    if ds_base:
        ds_base_unique = get_unique_row_list(ds_base, check_position_flag=check_position_flag)
        base_path = r"\\bcc\eng\PrDoc\377_НИПИГАЗ\АГХК\КСБ\РД\Таблички_графики\14_сравнение с ДС\\"
        path_out_dir = base_path  # +"\\"+ mto_obj.doc_OD_style_file_name

        # 4
        all_bases_unique = [ds_base_unique, mto_base_unique]
        RFQ.DS_compare_functions.ds_mto_compare(all_bases_unique, path_out_dir, open_dir_flag=open_dir_flag)
        if info_flag:
            print("DS_COMPARE FINISH")
    else:
        err_msg = f"В ДС не найден титул {mto_obj.file_name}"
        if error_log:
            error_log.append(err_msg)
        print(err_msg)

    #



