import utils.path
from base import t_comm_initial_classes
from base.base_cheks import base_vs_base_correction
from base.base_class_std_table import STDTable
from base.base_classes import TableComments
from base.base_google import load_base
from base.base_xlsx_load import list_raw_to_std
from base.tables_columns import CODE, TYPE_MARK, NAME, VENDOR, MASS, UNITS, BD_TABLE_NAME, BD_ID
from nano_cad.import_db import import_db
from nano_cad.write_db import write_db
from utils.error_log import ErrorLog
from utils.path import get_path_from_file_path
from utils.string_parsing import print_att_list_table


def load_db(file_path, sheet_name=0, column_dict=0):
    #
    out_dir = utils.path.get_path_out_dir(get_path_from_file_path(file_path))
    ErrorLog.start(exit_path=out_dir)
    # Открываем файл...
    ErrorLog.add_log(f"\tОткрываем |{file_path}| <load_db>")
    # Инициализация
    #
    column_list_view = [CODE, TYPE_MARK, NAME, VENDOR, MASS, UNITS, BD_TABLE_NAME, BD_ID]

    #
    t_com = TableComments(file_full_path=file_path,
                          dir_path="-1",
                          tabel_class=t_comm_initial_classes.EQUIPMENT)

    # 01 Читаем базу Нанокада
    #
    table_raw_obj = import_db(file_path, dbg=True)
    #
    # Нормализуем таблицу
    input_base_std = list_raw_to_std(table_raw_obj, t_com)

    # Печать в консоль (БЕЗ изменений)
    # print_att_list_table(STDTable.to_list(input_base_std, column_list_view),
    #                      max_len=45,
    #                      title="LOAD_DB input_base_std out",
    #                      title_row=column_list_view)
    # Печать в файл (БЕЗ изменений)
    # STDTable.to_excel_nanocad_db(input_base_std, out_dir, "01_not_changed")
    #
    # 02 Сравниваем с гугл базой
    code_base_data_std = load_base()
    #
    columns_for_check = [CODE, TYPE_MARK, NAME, VENDOR, MASS, UNITS]
    #
    input_base_std = base_vs_base_correction(input_base_std, code_base_data_std, columns_for_check)
    #
    # Печать в файл (БЕЗ изменений)
    STDTable.to_excel_nanocad_db(input_base_std, out_dir, "02_compare")
    # 03 Запись в базу
    # Копируем старую базу перед корректировкой
    out_dir = utils.path.normalize_path(out_dir)
    file_name = "bac_" + utils.path.get_file_name_from_full_file_path(file_path)
    dest_file_path = utils.path.file_path_plus_file_name(out_dir, file_name)
    ErrorLog.add_log(f"Копируем БД:\n"
                     f"\t{file_path}\n"
                     f"\t{dest_file_path}\n")
    utils.path.copy_file(file_path, dest_file_path)
    #
    # Записываем изменения в БД
    #
    write_db(file_path, input_base_std)

    # 99 Завершение
    utils.path.open_dir(out_dir)


    return table_raw_obj


if __name__ in {"__main__"}:
    # dir_path = r"C:\YandexDisk\темп\AGCC.287-2879-SOS_nC_v.24.0"
    # dir_path = r"C:/YandexDisk/темп/Новая папка/AGCC.287-6400.SKUD.db"
    dir_path = r"C:\YandexDisk\темп\DB\9110-PB-21.db"
    # file_path = r"C:\YandexDisk\темп\XML\Здание 1 Этаж 0.xml"

    open_path = ""
    # if 1 == 0:
    #     print("Запуск проверки ТПК папки")
    #     # Окно выбора ПАПКИ
    #     dir_path = folder_select.getDirectory()
    #     print(dir_path, "<open_pdf_folder>")
    #     # Обработка исключения
    #     if is_file_path(dir_path) is False:
    #         exit(0)
    load_db(dir_path)
