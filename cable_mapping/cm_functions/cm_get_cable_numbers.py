from utils.colors import Color
from cable_mapping.cm_clss.cm_CableTable import CableTable
from cable_mapping.cm_clss.cm_CableTableRow import CableTableRow
from cable_mapping.cm_functions import cm_from_where_terminal
from cable_mapping.cm_functions.cm_renumber_tpk_cables import renumber_tpk_cables
from cable_mapping.constants import *
from cable_mapping.mapping.map_calsses import MapTable, PROHIBITION_CODE_LIST
from cable_mapping.mapping.summary_cables import SummaryRow
from utils.error_log import ErrorLog
from utils.string_parsing import print_att_list_table

"""
get_cable_numbers_from_cj - сравнивает две таблицы кабелей на номера кабелей.
"""


def get_cable_numbers_from_cj(new_cab_cls: CableTable, old_cab_cls: CableTable, map_table: MapTable, dbg=1):
    # Формируем словарь из старого кабельного журнала
    #   Ключ - уникальное сочетание Откуда-Куда-На_какие_клеммы_Какой_кабель и пр.
    #   Значение - номер кабеля (0001, 1101 и т.п.)
    if dbg:
        ErrorLog.add_log("\nGET_CABLE_NUMBERS_FROM_CJ - Функция проверки кабельного журнала существующего\n")

    #
    old_dict = {}
    for index, row in old_cab_cls.table.items():
        if isinstance(row, CableTableRow):
            key = row.get_cable_no_number_tag()
            value = row.get_cable_number()
            if key in old_dict:
                ErrorLog.add_error(f"ERROR get_cable_numbers_from_cj: "
                                   f"Ключ {key} [{value}] (ind:{index}) уже есть в словаре -> {old_dict[key]}\n"
                                   f"Завершение программы")
                ErrorLog.exit()
            else:
                old_dict.setdefault(key, [value, index])
        else:
            ErrorLog.add_error(f"ERROR get_cable_numbers_from_cj: Not isinstance SummaryRow index:{index}")

    # for k, v in old_dict.items():
    #     print(v, k)
    # print("\n")

    # Матрица проверки полей
    #   0 - сравнение без преобразование
    #   1 - сравнение с регулярным выражением
    matrix = {
        CJ_CABLE_CODE: 0,
        CJ_ADDITIONAL_CODE: 0,
        CJ_FROM: 0,
        CJ_FROM_TERMINAL: 1,
        CJ_WHERE: 0,
        CJ_WHERE_TERMINAL: 1,
        CJ_CABLE_MARK: 0,  # ??
        # CJ_ANNOTAION: 0,
    }
    compare_rank_max = 100 * len(matrix)
    i_ind = 0
    for index, row in new_cab_cls.table.items():
        if isinstance(row, CableTableRow):
            # Формируем строку с ключами для новой строки из cj_ТПК
            n_key = row.get_cable_no_number_tag()
            #
            # Только для вывода в консоль и комментарии
            i_ind += 1
            l_value = [row.get_cable_number(), index]
            #
            # result_flag = False  # Флаг сравнения двух строк в КЖ и ТПК
            match_old_key = []  # список всех подходящих строк из старого КЖ
            match_old_key_compare_rank = {}  # список степени сходства
            #  Ищем сочетание в MapTable для строки из ТПК
            id_index = map_table.get_code_by_row(row)
            #  Если нашли в MapTable, то идем по старому КЖ и ищем подобное, как в n_key, сочетание в нем
            if id_index not in PROHIBITION_CODE_LIST:
                for old_key, old_value in old_dict.items():
                    result_flag = True  # Флаг сравнения двух строк в КЖ и ТПК
                    #
                    compare_rank = 0
                    """
                    Далее перебираем весь словарь с ключами строк из старого КЖ
                    Если при сравнении всех столбцов текущей строки (row, n_key) и строки из старого КЖ
                    флаг result_flag останется True - то добавляем old_key в список match_old_key
                    match_old_key - по итогу проверки должен содержать или 0 (не найдено) или 1 совпадение.
                    
                        compare_rank - это степень схожести строки для случаев когда и по маске подходит 
                    и при буквальном сравнении.
                    Например, при маске "SX(TAG) - > SK(TAG)" - все четыре строки удовлетворяют, но при сравнении по 
                    конкретным тегам можно понять что искомая строка в старом КЖ только 
                    одна (например 1513-SS-01-S-ARK-1811). В таком случае для нее compare_rank будет выше. 
                    И compare_rank для искомой строки будет максимальным и в одном экземпляре.
                    
                    compare_rank_max - это максимальное число которое можно набрать при сравнении строк. 
                    Величина известна это количество сравниваемых параметров на максимальную оценку - 100 * len(matrix)
    --------------------+---------------------+---------------------+---------------------+----------------------+---
    | CJ_FROM             | CJ_FROM_TERMINAL    | CJ_WHERE            | CJ_WHERE_TERMINAL    | CJ_CABLE_MARK         
    +---------------------+---------------------+---------------------+----------------------+-----------------------
    | 1520-SS-01-N-SX-1002| 1520-SS-01-S-PP-1001| 1513-SS-01-S-SK-1201| 1513-SS-01-S-ARK-1811| U/UTP Cat6 ZH нг(А)-HF
    | 1520-SS-01-N-SX-1002| 1520-SS-01-S-PP-1001| 1513-SS-01-S-SK-1201| 1513-SS-01-S-ARK-1812| U/UTP Cat6 ZH нг(А)-HF
    | 1520-SS-01-N-SX-1002| 1520-SS-01-S-PP-1001| 1513-SS-01-S-SK-1201| 1513-SS-01-S-MR-1811 | U/UTP Cat6 ZH нг(А)-HF
    | 1520-SS-01-N-SX-1002| 1520-SS-01-S-PP-1001| 1513-SS-01-S-SK-1201| 1513-SS-01-S-PI-1811 | U/UTP Cat6 ZH нг(А)-HF
    +---------------------+---------------------+---------------------+----------------------+-----------------------
                    """
                    for i in range(len(old_key)):
                        matrix_key = old_key[i].split("@")[0]  # f.e. CJ_CABLE_CODE
                        value_old_key = str(old_key[i].split("@")[1]).strip()  # 3252-PB-01-S-SK-1001
                        value_new_key = str(n_key[i].split("@")[1]).strip()  # 3252-PB-01-S-SK-1001
                        check_mode = matrix[matrix_key]  # f.e. CJ_CABLE_CODE -> 0, CJ_FROM_TERMINAL -> 1

                        if check_mode == 0:
                            if value_old_key != value_new_key:
                                # print(f"<{value_old_key}> <{value_new_key}>")
                                result_flag = False
                                compare_rank = 0
                            else:
                                compare_rank += 100
                        elif check_mode == 1:
                            if value_old_key == value_new_key:
                                compare_rank += 100
                            else:
                                # Преобразуем столбцы CJ в столбцы MapClass
                                map_column = map_table.get_by_cj_id_column(matrix_key)
                                # Находим регулярное выражение в списке MapClass по нужной колонке
                                reg_exp = map_table.table[id_index][map_column]
                                # Сравниваем старое значение с регулярным выражением из MapClass
                                # !!! При этом не учитываются порты, клеммы обозначенные в рег выражении как #, а так
                                # же теги оборудования - ПРОВЕРЯЕТСЯ ТОЛЬКО СОВПАДЕНИЕ МАСКИ
                                result = cm_from_where_terminal.regular_compare(value_old_key, reg_exp)
                                compare_rank += 50
                                # if dbg:
                                #     print(f"{i_ind} {result} -> {value_old_key} | {reg_exp}")
                                if not result:
                                    result_flag = False
                                    compare_rank = 0
                    if result_flag:
                        match_old_key.append([old_key, compare_rank])
                        match_old_key_compare_rank.setdefault(compare_rank, 0)
                        match_old_key_compare_rank[compare_rank] += 1
                        continue
                    """
                    Конец кода по сравнению строк ТПК и старого КЖ
                    """
            elif id_index in PROHIBITION_CODE_LIST:
                row.el[CJ_CABLE_CODE].value = id_index
                row.el[CJ_CABLE_LABEL].color = Color.red
                row.el[CJ_CABLE_LABEL].comment += f"Сочетание не найдено в MapTable.({id_index})\n"

            if dbg:
                ErrorLog.add_log(f"{i_ind} {len(match_old_key)} -> {match_old_key}")
            #
            # Находим номер строки в списке с наибольшим compare_rank
            match_index = -1
            _ = 0
            for x in range(len(match_old_key)):
                if match_old_key[x][1] > _:
                    match_index = x
                    _ = match_old_key[x][1]

            # Если найдена подходящая строка из старого КЖ то берем из него значения для:
            #       CJ_CABLE_NUMBER
            #       CJ_FROM_TERMINAL
            #       CJ_WHERE_TERMINAL
            if len(match_old_key) == 1 or match_old_key_compare_rank.get(compare_rank_max) == 1:
                cj_cable_number = old_dict[match_old_key[match_index][0]][0]
                cj_from_terminal = row.get_from_key_value_by_column(match_old_key[match_index][0], CJ_FROM_TERMINAL)
                cj_where_terminal = row.get_from_key_value_by_column(match_old_key[match_index][0], CJ_WHERE_TERMINAL)
                row.set_cable_number(cj_cable_number)
                row.set_from_terminal(cj_from_terminal)
                row.set_where_terminal(cj_where_terminal)

                # Помечаем CJ_CABLE_CODE_FLAG что бы отличать от нумерованных автоматически
                row.el[CJ_CABLE_CODE_FLAG].value = TEXT_CABLE_CODE_FLAG_1
                #

                if dbg:
                    ErrorLog.add_log(
                        f"\tTrue\t{l_value} {n_key} -> ({row.get_mapping_signature()})\t "
                        f"номер заменен из предыдущего КЖ")
                row.el[CJ_CABLE_LABEL].color = Color.green
                row.el[CJ_CABLE_LABEL].comment += (f"Подключение найдено. "
                                                   f"Номер заменен из предыдущего КЖ "
                                                   f"({old_dict[match_old_key[match_index][0]]})\n")
            elif len(match_old_key) == 0:
                row.el[CJ_CABLE_LABEL].color = Color.yellow
                row.el[CJ_CABLE_LABEL].comment += f"Подключение не найдено в предыдущем КЖ\n"
                if dbg:
                    ErrorLog.add_log(f"\tFalse\t{l_value} {n_key} -> ({row.get_mapping_signature()})\t"
                                     f" отсутствует в предыдущем КЖ")
            else:
                ErrorLog.add_error(f"ERROR GET_CABLE_NUMBERS_FROM_CJ: Найдено больше одного сочетания строк:")
                err_out_list = []
                title_row = CableTableRow.get_cable_no_number_header_row()
                for x in match_old_key:
                    err_out_list.append(CableTableRow.get_row_of_value_from_cable_no_number(x[0]))
                for k, v in match_old_key_compare_rank.items():
                    ErrorLog.add_error(f"{k} : {v} (максимально возможное значение {compare_rank_max})")
                ErrorLog.add_error(print_att_list_table(err_out_list,
                                                        max_len=100,
                                                        title_row=title_row,
                                                        titles_print_flag=False),
                                   print_flag=0)
                ErrorLog.add_error(f"Программа завершена с ошибкой!")
                ErrorLog.exit()
        else:
            ErrorLog.add_error(f"ERROR GET_CABLE_NUMBERS_FROM_CJ: Not isinstance SummaryRow index:{index}")
    # Окончание цикла присвоения номеров из старого КЖ
    """
    Перебираем номера строк которые не нашлись в старом КЖ и обновляем номера кабелей
    """
    renumber_tpk_cables(new_cab_cls, dbg=1)
