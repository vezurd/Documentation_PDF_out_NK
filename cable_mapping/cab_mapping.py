import os

from cable_mapping import load_tpk
import base
from cable_mapping.cm_clss.cm_CableLayingTypes import CableLayingTypes
from cable_mapping.cm_clss.cm_ColumnsCableTable import ColumnsCableTable
from cable_mapping.cm_clss.cm_TableCommentsCabMap import DocumentCabMap, TableCommentsCabMap
from cable_mapping.cm_functions.cm_get_cable_numbers import get_cable_numbers_from_cj
from cable_mapping.lolad_docx_cj import load_cj_from_docx
from cable_mapping.mapping import load_map_from_google
from cable_mapping.mapping.get_cj import get_cj
from cable_mapping.mapping.get_summary import get_summary_from_cj_table
from cable_mapping.mapping.load_map_from_google import load_no_out_cables_google_base, open_laying_google_base
from cable_mapping.utils_cm import get_files_list, print_file_list_to_console_cab_map
from utils.error_log import ErrorLog
import cable_mapping.map_excel_out
from utils.path import open_dir
from cable_mapping.cj_check.cj_check_main import *


def cab_mapping_start(dir_path,
                      dbg=1):
    dir_name = os.path.basename(dir_path)
    open_path = ""
    print(dir_name)
    ErrorLog.exit_path = dir_path  # Задаем начальный путь для сохранения логов ошибок
    ErrorLog.error_log = []
    # Ищем excel файлы в папке
    cm_proj_file_list = get_files_list(dir_path, endswith=(".xlsx", ".XLSX", ".docx", ".DOCX"))
    if cm_proj_file_list == -1:
        print(f"По пути <{dir_path}> файлы не найдены.")
        exit(0)
    else:
        # Заполняем cable_mapping_project
        cm_proj = []
        for file_path in cm_proj_file_list:
            file_name = os.path.basename(file_path)
            if DocumentCabMap.is_tpk_normalized_output(file_name):
                print(f"Пропуск (результат нормализации ТПК, не вход): {file_name}")
                continue
            cm_proj.append(DocumentCabMap(file_path, dir_path=dir_path))
        print_file_list_to_console_cab_map(cm_proj)

    # Если список файлов в папке не пустой...
    if cm_proj:
        if dbg:
            print("###################################################")
            print("# Загружаем сырой массив из базы оборудования гул #")
            print("####################################################")
        code_base_data_std = base.base_google.load_base()
        map_raw = load_map_from_google.load_map_google_base(dbg=0, max_len=45)
        no_out_cables_raw = load_no_out_cables_google_base(dbg=0, max_len=45)
        open_laying_raw = open_laying_google_base(dbg=0, max_len=45)

        if dbg:
            print("Создаем класс MapTable")
        map_cls = MapTable(map_raw, code_base_data_std, no_out_cables_raw, t_com=cm_proj[0].t_com, dbg=0)
        # Вывод всей mapping table в консоль (для отладки — почему комбинация не найдена / дублируется):
        print_map_table_to_console = False  # True — вывести всю таблицу
        if print_map_table_to_console:
            map_cls.print_to_console()
        # map_cls.SummaryCables.print_summary_cables()
        # map_cls.SummaryCables.print_no_out_cables()

        if dbg:
            print("Создаем класс со словарем МАТЕРИАЛ:СПОСОБ ПРОКЛАДКИ")
        cable_laying_class = CableLayingTypes(code_base_data_std, open_laying_raw)
        # temp_list = cable_laying_class.get_list()
        # for i in range(len(temp_list)):
        #     print(temp_list[i])
        #
        """
        
        """
        tpk_cable_cls = CableTable(t_com=TableCommentsCabMap(dir_path=cm_proj[0].t_com.dir_path,
                                                             file_full_path=cm_proj[0].t_com.dir_path,
                                                             file_full_name=cm_proj[0].t_com.file_full_name,
                                                             tabel_type="tpk_cable_cls"),
                                   dbg=0)
        # tpk_cable_cls.print_table_to_console()
        # cj_std = []
        #
        flag_old_cj = False  # Флаг загрузки старого КЖ для сравнения кабелей

        for doc in cm_proj:
            if dbg:
                print("Загрузка КЖ из ТПК")
            if doc.doc_type == FILE_NAME_TPK[1] and 1 == 1:
                if dbg:
                    print("Загружаем сырой массив из ТПК")
                tpk_raw = load_tpk.load_from_xlsx_file(doc, dbg=0, start_index=2, max_len=400)
                tpk_std = load_tpk.tpk_raw_to_std(tpk_raw, t_com=doc.t_com, dbg=0, max_len=30)

                if dbg:
                    print("Нормализуем ТПК таблицу (способы прокладки заносим в столбцы)")
                tpk_std = load_tpk.tpk_normalize(tpk_std, cable_laying_class, dbg=0, max_len=30)
                cable_mapping.map_excel_out.tpk_normalized_out(
                    tpk_std,
                    source_file_name=doc.file_name,
                    out_dir=doc.dir_path,
                )
                for row in tpk_std:
                    tpk_cable_cls.add_std_row(row)
            # """

            elif doc.doc_type == FILE_NAME_CJ[1] and 1 == 1:
                if dbg:
                    print("Загрузка КЖ из старого КЖ")
                """
                Работа с данными
                """
                # Создаем объект класса CableTable
                cj_exist_cable_cls: CableTable = CableTable(t_com=doc.t_com, dbg=0)
                if dbg:
                    print("Загружаем данные из docx")
                load_cj_from_docx(doc, cj_exist_cable_cls, dbg=0, start_index=0, max_len=400)
                flag_old_cj = True
                # Обрабатываем сырые данные:
                # разбиваем CJ_CABLE_LABEL на CJ_CABLE_CODE | CJ_CABLE_NUMBER | CJ_ADDITIONAL_CODE
                cj_exist_cable_cls.cable_label_to_code_number_add_code()

                """
                Вывод результата работы для отладки
                """
                # Вывод маленькой таблицы сумм кабелей
                # cj_exist_cable_cls.print_sum_table_to_console(mode="full")
                # Вывод основной таблицы кабельных соединений
                # cj_exist_cable_cls.print_table_to_console(file_name="cj_exist_cable.txt")

                if dbg:
                    print("Проверка кабельного журнала")
                cj_check(cj_exist_cable_cls, map_cls)

                if dbg:
                    print("Вывод в EXCEL результата проверки старого КЖ")
                open_path = cable_mapping.map_excel_out.check_color_out_map(cj_exist_cable_cls,
                                                                            ColumnsCableTable.cj_list,
                                                                            cm_proj[0].t_com.out_dir,
                                                                            "cj_",
                                                                            comments_dbg=1)

        if tpk_cable_cls.index > 0:
            # Ищем сочетания кодов в MapTable и вносим его в таблицу CJ
            get_cj(map_cls, tpk_cable_cls.table)
            get_summary_from_cj_table(map_cls, tpk_cable_cls)

            tpk_cable_cls.print_table_to_console(column_list=ColumnsCableTable.cj_list_extend)

            # Если старый КЖ найден в папке, то ищем кабели в нем
            if flag_old_cj:
                """
                get_cable_numbers_from_cj
                    Присваивает номера кабелей, места расключения из старого КЖ
                    Проверяет нумерацию кабелей после обновления номеров
                """
                get_cable_numbers_from_cj(tpk_cable_cls, cj_exist_cable_cls, map_cls, dbg=0)

                """
                Вывод результата работы для отладки
                """
                tpk_cable_cls.print_table_to_console(column_list=ColumnsCableTable.cj_list_extend)

            # tpk_cable_cls.print_sum_table_to_console()

            open_path = cable_mapping.map_excel_out.check_color_out_map(tpk_cable_cls,
                                                                        ColumnsCableTable.cj_list,
                                                                        cm_proj[0].t_com.out_dir,
                                                                        "cj_",
                                                                        comments_dbg=0)
        if open_path != "":
            # open_dir(cm_proj[0].t_com.out_dir)
            open_dir(open_path)
            ErrorLog.print_err_log(open_path, open_file_flag=0)
    print("END CAB_MAPPING")


if __name__ in {"__main__"}:
    dir_path = r"C:\YandexDisk\темп\_cable_mapping\3252"
    open_path = ""
    if 1 == 0:
        print("Запуск проверки ТПК папки")
        # Окно выбора ПАПКИ
        dir_path = folder_select.getDirectory()
        print(dir_path, "<open_pdf_folder>")
        # Обработка исключения
        if is_file_path(dir_path) is False:
            exit(0)
    cab_mapping_start(dir_path)
