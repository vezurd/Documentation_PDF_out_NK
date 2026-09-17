from base.base_classes import *
from base.base_xlsx_load import load_from_xlsx_file, list_raw_to_std



def aveva_get_std_from_file(t_com: TableComments):
    # 01.2 Читаем лист
    boot_co_raw_obj = load_from_xlsx_file(t_com)
    # 01.3 Нормализуем данные (по столбцам, формату, лишним символам и т.п.)
    boot_co_std = list_raw_to_std(boot_co_raw_obj, t_com)
    return boot_co_std


def equipment_get_std_from_file(t_com: TableComments):
    # 01.2 Читаем лист
    input_raw_obj = load_from_xlsx_file(t_com)
    # 01.3 Нормализуем данные (по столбцам, формату, лишним символам и т.п.)
    input_base_std = list_raw_to_std(input_raw_obj, t_com)
    return input_base_std



