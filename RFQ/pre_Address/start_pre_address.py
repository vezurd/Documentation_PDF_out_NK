import utils
from base.base_cheks import check_position_row
from base.base_classes import *
from base.base_excel_out import check_color_out
from base.base_google import load_base
from base.base_mto import get_mto_std_from_file
from base.tables_columns import ColNames, CODE, MTO_NAME, ZIP_CODE, ZIP_BUILDING
from utils.error_log import ErrorLog
from utils.path import open_dir
from utils.prints_to_console import print_file_list_to_console
from utils.string_parsing import getDocTitle
import base.t_comm_initial_classes as t_com_init_cls


debug_print_files_list = 1
debug_print_progress = 1

"""
START_PRE_ADDRESS
По запросу Куропаткина (2025.04.02) для подготовки таблицы адресов.

1. Сканируем папку и собираем все МТО, в т.ч. в подкаталогах
2. Собираем все МТО в один список строк
3. Не учитываем материалы в метрах
4. Разделяем на отдельные строчки позиции с тегами, если кол-во тегов и кол-во совпадают. 
    Если не совпадают - не раскрываем.
5. Добавляем для каждой строки ЗИП код для удобства фильтрации
"""


def start_pre_address(dir_path):
    ErrorLog.add_log("PRE_ADDRESS: сканирование фалов")
    ErrorLog.exit_path = dir_path

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
    pre_zip_base = []
    for document in curr_proj:
        if "MTO" in document.file_name:
            # print(f"    Open MTO - {document.file_name}")
            mto_base = get_mto_std_from_file(t_com=TableComments(file_full_path=document.file_full_path,
                                                                 dir_path="-1",
                                                                 tabel_class=t_com_init_cls.MTO,
                                                                 )
                                             )
            # Основная функция по добавлению строк в общий список
            t_com = TableComments(dir_path=mto_base[0].t_com.dir_path,
                                  file_full_path=mto_base[0].t_com.file_full_path,
                                  file_full_name=mto_base[0].t_com.file_name,
                                  tabel_class=mto_base[0].t_com.tabel_class,
                                  )
            for row in mto_base:
                """
                Проверка относится ли тип строки к позиции
                """
                if not check_position_row(row):
                    continue
                elif isinstance(row, RowStd):
                    code = row.el[CODE].value
                    row.el[MTO_NAME].value.append(t_com.file_name)

                    # Для примера AGCC.287-7417-SOS.MTO-0001_01_RU.pdf - ищем часть "7417"
                    row.el[ZIP_BUILDING].value = getDocTitle(t_com.file_name)
                    #
                    #  Проверяем на пустое поле CODE
                    if code == "":
                        pre_zip_base.append(RowStd.get_row_copy(row, t_com))
                        continue

                    units_list_no = ["м", "м.", "М", "М."]
                    list_for_out = ("BCC0000348",
                                    "BCC0001989",
                                    "BCC0000657",
                                    "BCC0000658",
                                    "BCC0000659",
                                    "BCC0000660",
                                    "BCC0000664",
                                    "BCC0000666",
                                    "BCC0000667",
                                    "BCC0000668",
                                    "BCC0000669",
                                    "BCC0000670",
                                    "BCC0000671",
                                    "BCC0000672",
                                    "BCC0002852",
                                    "BCC0000419")
                    if row.el[CODE].value in list_for_out:
                        pre_zip_base.append(RowStd.get_row_copy(row, t_com))

                    # if row.el[UNITS].value not in units_list_no:
                    #     try:
                    #         row.el[VALUES].value = integer_from_any(row.el[VALUES].value)
                    #     except ValueError as e:
                    #         print(f"Error:\n"
                    #               f"   {t_com.file_name}\n"
                    #               f"   {t_com.file_full_path}\n"
                    #               f"   {code}")
                    #         raise e
                    #     if row.el[VALUES].value == 1:
                    #         pre_zip_base.append(RowStd.get_row_copy(row, t_com))
                    #     else:
                    #         count_value = int(row.el[VALUES].value)
                    #         count_tags = len(row.el[TAGS].value)
                    #         if count_value == count_tags:
                    #             for x in range(count_value):
                    #                 temp_row = (RowStd.get_row_copy(row, t_com))
                    #                 temp_row.el[VALUES].value = 1
                    #                 temp_row.el[TAGS].value = [row.el[TAGS].value[x]]
                    #                 pre_zip_base.append(temp_row)
                    #         else:
                    #             temp_row = (RowStd.get_row_copy(row, t_com))
                    #             # Если количество тегов не нулевое и не совпадает с кол-вом в позиции - отмечаем желтым
                    #             if row.el[TAGS].value:
                    #                 temp_row.el[VALUES].color = Color.yellow
                    #                 temp_row.el[TAGS].color = Color.yellow
                    #                 row.el[VALUES].comment += (f"Кол-во {count_value} "
                    #                                            f"не равно кол-ву тегов ({count_tags})\n")
                    #             pre_zip_base.append(temp_row)

    # Путь для вывода результата
    path_out_dir = utils.path.get_path_out_dir(dir_path)

    # 3 Загружаем гугл-базу
    code_base_std = load_base()

    # 4 Обработка базы pre_zip_base
    for row in pre_zip_base:
        if isinstance(row, RowStd):
            code = row.el[CODE].value
            row_code_base = RowStd.get_row_by_code(code, code_base_std)

            if isinstance(row_code_base, RowStd):
                zip_code = row_code_base.el[ZIP_CODE].value
                row.el[ZIP_CODE].value = zip_code
                # zip_code_rule(zip_code, row)
            else:
                ErrorLog.add_error(f"ERROR: {row.get_text_for_debug()}")
                # raise Exception(f"\t pass : {code}  (4 Обработка базы pre_zip_base)")

        else:
            raise Exception

    # Вывод в файл

    column_dict = ColNames.ZIP.column_dict_adress
    return_file_path = check_color_out(pre_zip_base,
                                       column_dict,
                                       path_out_dir,
                                       f"теги_построчно_из_{len(curr_proj)}_файлов",
                                       excel_template_check_color=r"templates\out_template_pre_ADRESS.xlsx",
                                       row_index=1)

    open_dir(return_file_path)

    ErrorLog.add_log("PRE_ADDRESS сканирование фалов ЗАВЕРШЕНО")


if __name__ in {"__main__"}:
    mto_path = r"\\bcc\eng\PrDoc\377_НИПИГАЗ\АГХК\КСБ\РД\ЗИП"
    # mto_path = r"C:\YandexDisk\темп\ЗИП"
    start_pre_address(mto_path)
