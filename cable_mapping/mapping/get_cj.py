from cable_mapping.cm_functions.cm_from_where_terminal import get_clean_terminal_from_where_no_reg
from cable_mapping.mapping.map_calsses import *
from utils.colors import Color
from utils.error_log import ErrorLog


def get_map_table_temp_row(row: CableTableRow, map_table: MapTable, clean_reg=True):
    """
    Ищем в MapTable сочетание для row и записываем:
       CJ_CABLE_CODE
       CJ_ADDITIONAL_CODE
       CJ_FROM_TERMINAL
       CJ_WHERE_TERMINAL
    """

    #  Ищем сочетание в MapTable - если есть - заполняем поля
    id_index = map_table.get_code_by_row(row)
    if id_index not in PROHIBITION_CODE_LIST:
        # Заполняем код кабеля - I, A, C, P и т.п.
        row.el[CJ_CABLE_CODE].value = map_table.table[id_index][MAP_CODE]
        #
        map_from_terminal = map_table.table[id_index][MAP_FROM_TERMINAL]
        map_where_terminal = map_table.table[id_index][MAP_WHERE_TERMINAL]
        #
        if clean_reg:
            map_from_terminal = get_clean_terminal_from_where_no_reg(map_from_terminal)
            map_where_terminal = get_clean_terminal_from_where_no_reg(map_where_terminal)
        #
        row.el[CJ_FROM_TERMINAL].concatenate_str_values(map_from_terminal)
        row.el[CJ_WHERE_TERMINAL].concatenate_str_values(map_where_terminal)

        tag_from_cls = TagClass(row.el[CJ_FROM].value)
        tag_where_cls = TagClass(row.el[CJ_WHERE].value)
        if tag_from_cls.additional_code is not None and tag_where_cls.additional_code is not None:
            if tag_from_cls.additional_code != tag_where_cls.additional_code:
                err_text = (f"ОШИБКА: РАЗНОЕ СОЧЕТАНИЕЕ ДОПОЛНИТЕЛЬНЫХ КОДОВ: map_classes.get_cj:\n"
                            f"    {tag_from_cls.tag}:{tag_where_cls.tag}; id_index:{id_index}\n")
                print(err_text)
                row.el[CJ_CABLE_LABEL].color = Color.red
                row.el[CJ_CABLE_LABEL].comment += err_text
            else:
                row.el[CJ_ADDITIONAL_CODE].value = tag_where_cls.additional_code
        elif tag_where_cls.additional_code is not None:
            row.el[CJ_ADDITIONAL_CODE].value = tag_where_cls.additional_code
    elif id_index in PROHIBITION_CODE_LIST:
        row.el[CJ_CABLE_CODE].value = id_index


def get_cj(map_table: MapTable, cj_std: dict[int, CableTableRow]):
    for ind, row in cj_std.items():
        # Нумеруем кабель
        row.el[CJ_CABLE_NUMBER].value = map_table.get_next_cable_number(row)

        # Ищем в MapTable сочетание для row и записываем:
        #   CJ_CABLE_CODE
        #   CJ_ADDITIONAL_CODE
        #   CJ_FROM_TERMINAL
        #   CJ_WHERE_TERMINAL
        get_map_table_temp_row(row, map_table)

        # Собираем название кабеля
        # row.el[CJ_CABLE_LABEL].value = row.get_cable_label()
        # row.el[CJ_CABLE_FULL_TAG].value = row.get_cable_full_tag()
        row.update()

        # Собираем общую длину по способам прокладки
        row.el[CJ_TOTAL_LENGTH].value = 0
        for att in CABLE_LAYING_TYPES_LIST:
            if row.el[att].value != "":
                row.el[CJ_TOTAL_LENGTH].value += int(row.el[att].value)
            # Проверяем на нулевую длину сумм прокладки
        if row.el[CJ_TOTAL_LENGTH].value == 0:
            err_text = f"ОШИБКА! СУММАРНАЯ ДЛИНА РАВНА 0\n"
            print(err_text)
            row.el[CJ_TOTAL_LENGTH].color = Color.red
            row.el[CJ_TOTAL_LENGTH].comment += err_text
            # Проверяем на общую длину относительно ТПК
        if row.el[TPK_R8_TOTAL_LENGTH].value != "":
            row.el[TPK_R8_TOTAL_LENGTH].value = math.ceil(float(row.el[TPK_R8_TOTAL_LENGTH].value.replace(",", ".")))
            if row.el[CJ_TOTAL_LENGTH].value < row.el[TPK_R8_TOTAL_LENGTH].value:
                err_text = (f"ОШИБКА! ДЛИНА КАБЕЛЯ ({row.el[CJ_TOTAL_LENGTH].value}) МЕНЬШЕ ЧЕМ В ТПК "
                            f"({row.el[TPK_R8_TOTAL_LENGTH].value})\n")
                print(err_text)
                row.el[CJ_TOTAL_LENGTH].color = Color.red
                row.el[CJ_TOTAL_LENGTH].comment += err_text

    print_err_no_found_combination(map_table)
    print_err_duplicates(map_table)


def print_err_duplicates(map_table: MapTable):
    print_list = []
    for k, v in map_table.err_duplicated_dict.items():
        row = [k, v]
        print_list.append(row)
    if print_list:
        text = print_att_list_table(print_list,
                                    max_len=400,
                                    title_row=["Комбинация:", "Кол-во:"],
                                    title="СПИСОК КОМБИНАЦИЙ ДУБЛИРОВАННЫХ В MAPPING TABLE:",
                                    titles_print_flag=1)
        ErrorLog.add_error(text)
        if map_table.dbg and map_table.t_com.out_dir is not None:
            save_text_to_file(text, file_name="err_duplicates.txt", out_dir=map_table.t_com.out_dir)
    else:
        ErrorLog.add_error("СПИСОК ДУБЛИРОВАНИЙ В MAPPING TABLE ПУСТ")


def print_err_no_found_combination(map_table: MapTable):
    # print(f"СПИСОК КОМБИНАЦИЙ НЕ НАЙДЕНЫХ В MAPPING TABLE:")
    print_list = []
    for k, v in map_table.err_no_found_dict.items():
        row = [k, v]
        print_list.append(row)
    if print_list:
        text = print_att_list_table(print_list,
                                    max_len=400,
                                    title_row=["Комбинация:", "Кол-во:"],
                                    title="СПИСОК КОМБИНАЦИЙ НЕ НАЙДЕНЫХ В MAPPING TABLE:",
                                    titles_print_flag=1)
        ErrorLog.add_error(text, print_flag=0)
        if map_table.dbg and map_table.t_com.out_dir is not None:
            save_text_to_file(text, file_name="err_no_found_combination.txt", out_dir=map_table.t_com.out_dir)
    else:
        ErrorLog.add_error("СПИСОК КОМБИНАЦИЙ НЕ НАЙДЕНЫХ В MAPPING TABLE ПУСТ")
