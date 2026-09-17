import math

from base.base_classes import RowStd
from utils.colors import Color
from cable_mapping.cm_clss.cm_TableCommentsCabMap import TableCommentsCabMap
from cable_mapping.cm_clss.cm_CableTableRow import CableTableRow
from cable_mapping.constants import *
from cable_mapping.mapping.summary_cables import CableSummary
from tags.decimal_format import decimal_format
from tags.tag_classes import TagClass
from utils.save_to_file import save_text_to_file
from utils.string_parsing import print_att_list_table, cable_x_fixer

# MAP_IDM = "idm"  #
MAP_CODE = "type"  # A, I, C, etc
MAP_FROM = "from"
MAP_FROM_TERMINAL = "fromDesc"
MAP_WHERE = "where"
MAP_WHERE_TERMINAL = "whereDesc"
MAP_SYSTEM_FROM = "system_from"
MAP_SYSTEM_WHERE = "system_where"
MAP_CABLE = "cable"

NOT_FOUND_CODE = "НЕТ_КОДА"
NOT_UNIQUE_CODE = "НЕ_УНИКАЛЬНОЕ_СОЧЕТАНИЕ"
PROHIBITION_CODE_LIST = [NOT_UNIQUE_CODE, NOT_FOUND_CODE]


class MapTable:
    main_list = (
        MAP_CODE,
        MAP_SYSTEM_FROM,
        MAP_FROM,
        MAP_FROM_TERMINAL,
        MAP_SYSTEM_WHERE,
        MAP_WHERE,
        MAP_WHERE_TERMINAL,
        MAP_CABLE,
    )

    def __init__(self, map_raw: list[list],
                 code_base_data_std: list[RowStd],
                 no_out_cables_raw: list[list],
                 t_com=TableCommentsCabMap(), dbg=0):

        self.err_no_found_dict = {}
        self.err_duplicated_dict = {}
        self.t_com = t_com
        self.dbg = dbg

        # Проверяем первую строчку, создаем словарь заголовков ИМЯ_СТОЛБЦА:НОМЕР_СТОЛБЦА
        self.map_table_column_dict = {}
        for x in range(len(map_raw[0])):
            self.map_table_column_dict[x] = map_raw[0][x]

        self.table = {}
        for x in range(len(map_raw)):
            # Пропускаем строчку с заголовком
            if x == 0:
                continue
            row = {}
            for y in range(len(map_raw[x])):
                # Перебираем столбцы, если номер столбца есть в словаре - добавляем значение в row
                value = map_raw[x][y]
                if y in self.map_table_column_dict:
                    att = self.map_table_column_dict[y]
                    row[att] = value
            self.table[x] = row
        # Преобразуем названия кабелей в список
        for k, v in self.table.items():
            v[MAP_CABLE] = v[MAP_CABLE].strip()
            v[MAP_CABLE] = v[MAP_CABLE].strip(";")
            map_cable_list = v[MAP_CABLE].split(";")
            for a in range(len(map_cable_list)):
                map_cable_list[a] = cable_x_fixer(map_cable_list[a])
            v[MAP_CABLE] = map_cable_list
        # Переменная счетчика сквозной нумерации
        self.cable_number = 0

        # Проверка
        self.check_duplicates_in_map_table()
        #
        self.SummaryCables = CableSummary(code_base_data_std, no_out_cables_raw)

    """
        check_duplicates_in_map_table:
        вспомогательная проверка после создания класса MapTable
        Проверяет на дубликаты все сочетания параметров
    """

    def check_duplicates_in_map_table(self):

        for k, v in self.table.items():
            f1 = v[MAP_FROM]
            w1 = v[MAP_WHERE]
            s1 = v[MAP_SYSTEM_FROM]
            s2 = v[MAP_SYSTEM_WHERE]
            flag1 = True
            flag2 = -1
            for c1 in v[MAP_CABLE]:
                ind_list = self.find_map_row(f1, w1, c1, s1, s2)
                if len(ind_list) > 1:
                    self.add_err_duplicates(f1, w1, c1, s1, s2)

    def add_err_duplicates(self, cj_from1, cj_where1, cj_cable, cj_system1, cj_system2):
        key = f"[{cj_system1}-{cj_from1}] [{cj_system2}-{cj_where1}] [{cj_cable}]"
        if key in self.err_duplicated_dict:
            self.err_duplicated_dict[key] += 1
        else:
            self.err_duplicated_dict[key] = 1

    """
        find_map_row - Алгоритм сравнения и поиска строки в MapTable.table
        Если совпадают все параметры то ряд найден.
        Ны выходе получается список индексов строк их MapTable.table
        Если одно значение - комбинация найдена и уникальна
        Если ни одного - нет сочетания
        Если больше одного - комбинация не уникальна
    """

    def find_map_row(self, f1, w1, c1, s1, s2):
        index_list = []
        for k, v in self.table.items():
            if (f1 == v[MAP_FROM]
                    and w1 == v[MAP_WHERE]
                    and s1 == v[MAP_SYSTEM_FROM]
                    and s2 == v[MAP_SYSTEM_WHERE]
                    and cable_x_fixer(c1) in v[MAP_CABLE]):
                index_list.append(k)
        return index_list

    def error_text(self, id_index):
        err_from = self.table[id_index][MAP_FROM]
        err_where = self.table[id_index][MAP_WHERE]
        err_cable = self.table[id_index][MAP_CABLE]
        err_system1 = self.table[id_index][MAP_SYSTEM_FROM]
        err_system2 = self.table[id_index][MAP_SYSTEM_WHERE]
        err_code = self.table[id_index][MAP_CODE]
        return (f"<{id_index + 1}> | Index; <{err_code}> | Код кабеля\n"
                f"[{err_system1}-{err_from}] [{err_system2}-{err_where}] [{err_cable}]\n")

    def print_to_console(self, max_len=400):
        out_list = []
        for index, row in self.table.items():
            out_row_list = []
            for att in self.main_list:
                out_row_list.append(row[att])
            out_list.append(out_row_list)
        text = print_att_list_table(out_list, max_len=max_len, title_row=self.main_list, title="MapTable.table")
        if self.dbg and self.t_com.out_dir is not None:
            save_text_to_file(text, file_name="map_table.txt", out_dir=self.t_com.out_dir)

    def get_code_by_row(self, row: CableTableRow):
        # unique_flag = True
        code_row_id = NOT_FOUND_CODE
        try:
            tag_from_cls = TagClass(row.el[CJ_FROM].value)
        except (ValueError,Exception):
            print("ERROR in GET_CODE_BY_ROW")
            # print(row.debug_info())
            print(row.print_debug())
            print("END GET_CODE_BY_ROW")

            exit(0)

        tag_where_cls = TagClass(row.el[CJ_WHERE].value)
        cj_from1 = tag_from_cls.equipment
        cj_where1 = tag_where_cls.equipment
        cj_system1 = tag_from_cls.sub_system_cj
        cj_system2 = tag_where_cls.sub_system_cj
        cj_cable = row.el[CJ_CABLE_MARK].value

        # Ищем список подходящих комбинаций
        ind_find_list = self.find_map_row(cj_from1, cj_where1, cj_cable, cj_system1, cj_system2)

        # Если нашлось уникальное совпадение - возвращаем код строки
        if len(ind_find_list) == 1:
            code_row_id = ind_find_list[0]
        elif len(ind_find_list) > 1:
            code_row_id = NOT_UNIQUE_CODE
            print(f"ОШИБКА ДУБЛИРОВАНИЯ: map_classes.get_code_by_row:")
            err_combination = f"[{cj_system1}-{cj_from1}] [{cj_system2}-{cj_where1}] [{cj_cable}]\n"
            print(err_combination)
            self.add_err_duplicates(cj_from1, cj_where1, cj_cable, cj_system1, cj_system2)
        elif len(ind_find_list) < 1:
            code_row_id = NOT_FOUND_CODE
            err_combination = f"[{cj_system1}-{cj_from1}] [{cj_system2}-{cj_where1}] [{cj_cable}]\n"
            self.add_err_no_found_combination(cj_from1, cj_where1, cj_cable, cj_system1, cj_system2)
            # err_text = (f"ОШИБКА: НЕ НАЙДЕНО СОЧЕТАНИЕ: map_classes.get_code_by_row:\n" + err_combination)
            # print(err_text)
            row.el[CJ_CABLE_LABEL].color = Color.red
            row.el[CJ_CABLE_LABEL].comment += (f"\nНЕ НАЙДЕНО СОЧЕТАНИЕ:\n" + err_combination)

        return code_row_id

    def get_next_cable_number(self, row: CableTableRow):
        if row.el[CJ_CABLE_NUMBER].value == "":
            self.cable_number += 1
        else:
            return row.el[CJ_CABLE_NUMBER].value
        return decimal_format(self.cable_number)

    def add_err_no_found_combination(self, cj_from1, cj_where1, cj_cable, cj_system1, cj_system2):
        key = f"[{cj_system1}-{cj_from1}] [{cj_system2}-{cj_where1}] [{cj_cable}]"
        if key in self.err_no_found_dict:
            self.err_no_found_dict[key] += 1
        else:
            self.err_no_found_dict[key] = 1

    def get_from_terminal (self, id_index):
        return self.table[id_index][MAP_FROM_TERMINAL]

    def get_by_cj_id_column (self, cj_column):
        main_list_to_cj = {
            CJ_CABLE_CODE: MAP_CODE,
            CJ_FROM: MAP_FROM,
            CJ_FROM_TERMINAL: MAP_FROM_TERMINAL,
            CJ_WHERE: MAP_WHERE,
            CJ_WHERE_TERMINAL: MAP_WHERE_TERMINAL,
            CJ_CABLE_MARK: MAP_CABLE,
        }
        return main_list_to_cj[cj_column]




