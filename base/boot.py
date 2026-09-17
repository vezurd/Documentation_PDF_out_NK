import utils.path
from base.base_classes import *
from base.base_xlsx_load import load_from_xlsx_file, list_raw_to_std


def base_co_get_std_from_file(t_com: TableComments):
    # 01.1 Ищем "СО.xlsx"
    if t_com.file_full_path != "-1":
        if t_com.dir_path != "-1":
            curr_mto = utils.path.get_files_single(t_com.dir_path, [".xlsx"])
            if curr_mto == -1:
                print(f"По пути <{t_com.dir_path}> файлы .xlsx не найдены.")
                exit(0)
            for doc in curr_mto:
                print(doc.file_full_path)
                if doc.file_name == "СО.xlsx":
                    t_com.file_full_path = doc.file_full_path
    # 01.2 Читаем лист
    boot_co_raw_obj = load_from_xlsx_file(t_com)
    # 01.3 Нормализуем данные (по столбцам, формату, лишним символам и т.п.)
    boot_co_std = list_raw_to_std(boot_co_raw_obj, t_com)
    return boot_co_std


