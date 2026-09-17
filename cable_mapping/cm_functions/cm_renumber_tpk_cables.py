"""
    Перебираем номера строк которые не нашлись в старом КЖ и обновляем номера кабелей
"""
from utils.colors import Color
from cable_mapping.cm_clss.cm_CableTable import CableTable
from cable_mapping.cm_clss.cm_CableTableRow import CableTableRow
from cable_mapping.constants import *
from tags.decimal_format import decimal_format
from tags.tag_classes import TagClass


def renumber_tpk_cables(new_cab_cls: CableTable, dbg=1, mode="auto"):
    numbers_dict = {
        "SOT": [],
        "SKUD": [],
        "POS": [],
        "SOS": [],
        "SPP;SOO": [],
    }

    non_numbered_cables = []
    """
    Вариант сравнения без автоматического определения системы, по столбцу row.el[CJ_SYSTEM].value
    CJ_SYSTEM - берется из названия файла ТПК - ТПК_SOS.xlsx например.
    """
    if mode == "by_access":
        for row_index, row in new_cab_cls.table.items():
            if isinstance(row, CableTableRow):
                row_system = row.el[CJ_SYSTEM].value
                row_num_flag = row.el[CJ_CABLE_CODE_FLAG].value
                row_cable_number = row.el[CJ_CABLE_NUMBER].value
                if row_num_flag == TEXT_CABLE_CODE_FLAG_1:  # "Номер заменен из предыдущего КЖ"
                    numbers_dict[row_system].append(row_cable_number)
                else:
                    non_numbered_cables.append(row_index)

    """
    Вариант сравнения без автоматического определения системы, по столбцу row.el[CJ_SYSTEM].value
    CJ_SYSTEM - берется из названия файла ТПК - ТПК_SOS.xlsx например.
    """
    if mode == "auto":
        for row_index, row in new_cab_cls.table.items():
            if isinstance(row, CableTableRow):
                # Берем код системы из тега первого устройства
                tag_from = TagClass(row.el[CJ_FROM].value)
                row_system = tag_from.sub_system_cj  # SKUD, SOS, POS, SOS, SPP;SOO
                #
                row_num_flag = row.el[CJ_CABLE_CODE_FLAG].value
                if row_num_flag == TEXT_CABLE_CODE_FLAG_1:  # "Номер заменен из предыдущего КЖ"
                    numbers_dict[row_system].append(row.el[CJ_CABLE_NUMBER].value)
                else:
                    non_numbered_cables.append(row_index)

    #
    if dbg:
        print("Номер заменен из предыдущего КЖ")
        for key, value in numbers_dict.items():
            print(key, value)
        print("\nСписок НЕ замененных кабелей")
        for i in non_numbered_cables:
            row = new_cab_cls.table[i]
            if isinstance(row, CableTableRow):
                print(row.get_cable_number(), row.get_cable_full_tag())
    #
    end_number = 9999
    if non_numbered_cables:
        for i in non_numbered_cables:
            row = new_cab_cls.table[i]
            if isinstance(row, CableTableRow):
                # Берем код системы из тега первого устройства и определяем начальную нумерацию для кабеля этой системы
                tag_from = TagClass(row.el[CJ_FROM].value)

                row_system = tag_from.sub_system_cj  # SKUD, SOS, POS, SOS, SPP;SOO
                #
                # Собираем все номера кабелей в один список
                all_exist_numbers = []
                for k, v in numbers_dict.items():
                    for n in v:
                        all_exist_numbers.append(int(n))
                all_exist_numbers = set(all_exist_numbers)
                # Находим стартовый номер для кабеля по номеру системы
                start_number = int(f"0{tag_from.sum_system_number}01")  # "100" для SKUD, например
                # Создаем список со всеми возможными номерами от стартового до 9999
                all_numbers = set(range(start_number, end_number))
                # Находим разницу между списком без пропусков и существующими кабелями
                sorted_diff = [i for i in sorted(all_numbers.difference(all_exist_numbers))]
                # берем
                next_number = decimal_format(sorted_diff[0])
                numbers_dict[row_system].append(next_number)
                #
                row.set_cable_number(next_number)
                row.el[CJ_CABLE_LABEL].color = Color.soft_cyan
                row.el[CJ_CABLE_LABEL].comment += f"{TEXT_CABLE_CODE_FLAG_2}\n"
                row.el[CJ_CABLE_CODE_FLAG].value = TEXT_CABLE_CODE_FLAG_2
