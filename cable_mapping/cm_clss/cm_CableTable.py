from cable_mapping.cm_clss.cm_ColumnsCableTable import ColumnsCableTable
from cable_mapping.cm_clss.cls_cable_lable import CableLabelCls
from cable_mapping.cm_clss.cm_CableTableRow import CableTableRow
from cable_mapping.cm_clss.cm_TableCommentsCabMap import TableCommentsCabMap
from cable_mapping.constants import *
from cable_mapping.mapping.summary_cables import SummaryRow
from tags.tag_parser import get_tag
from utils.error_log import ErrorLog
from utils.save_to_file import save_text_to_file
from utils.string_parsing import cable_x_fixer, print_att_list_table

"""
CableTable:
Класс для хранения и работы с таблицей кабелей и подключений
    .__next_index - возвращает индекс следующей строки в self.table
    #
    .add_std_row - Функция добавления стандартной строки, проверка на тип строки
    .add_std_table - Добавление стандартных строк CableTableRow
    #
    .add_raw_table - Добавление сырых таблиц по словарю. В словаре указывается номер столбца и его тип
    #
    .print_table_to_console - печать в консоль таблицы self.table
"""


class CableTable:
    def __init__(self, t_com=TableCommentsCabMap(), dbg=0):
        self.table = {}  # [index]CableTableRow
        self.index = 0

        self.summary_table = {0: SummaryRow("none", "none", "none")}
        self.summary_table.pop(0)
        self.summary_index = 0

        self.frequency_tags = {"8350-CCB-01-S-BZ-3322": 10}
        self.frequency_tags.pop("8350-CCB-01-S-BZ-3322")

        self.t_com = t_com
        self.dbg = dbg

    """
    Функции для добавления строк в table
    """

    # Инкрементор индекса словаря table
    def __next_index(self):
        self.index += 1
        return self.index

    # Функция добавления стандартной строки, проверка на тип строки
    def add_std_row(self, std_row: CableTableRow):
        if isinstance(std_row, CableTableRow):
            self.table[self.__next_index()] = std_row
        else:
            ErrorLog.add_error(f"CableTable.add_std_row: std_row not CableTableRow class!\n"
                               f"   table[index]:{self.index}")
            ErrorLog.exit()

    """
    Добавление сырых таблиц по словарю
    В словаре указывается номер столбца и его тип
    """

    def add_raw_table(self, raw_table: list[list], column_dict: dict):
        if raw_table and column_dict:
            for raw_row in raw_table:
                std_row = CableTableRow(self.t_com)
                for k, v in column_dict.items():
                    value = raw_row[v]
                    if isinstance(value, str):
                        value.strip()
                    std_row.el[k].value = value
                self.add_std_row(std_row)

    """
       Добавление стандартных строк CableTableRow
    """

    def add_std_table(self, std_table: list[CableTableRow]):
        for std_row in std_table:
            self.add_std_row(std_row)

    """
    SUMMARY
    """

    def __find_sum_row(self, row: SummaryRow):
        cable_mark_fix = cable_x_fixer(row.cable_full_name)
        for ind, t_row in self.summary_table.items():
            if cable_x_fixer(t_row.cable_full_name) == cable_mark_fix:
                return ind
        return -1

    def __sum_next_index(self):
        self.summary_index += 1
        return self.summary_index

    def sum_add_summary_row(self, row: SummaryRow):
        if isinstance(row, SummaryRow):
            self.summary_table[self.__sum_next_index()] = row
        else:
            err_text = (f"CableTable.sum_add_summary_row\n"
                        f"  <{row}> не является объектом класса SummaryRow")
            ErrorLog.add_error(err_text)

    def sum_add_cable_table_row(self, row: CableTableRow):
        if isinstance(row, CableTableRow):
            self.summary_table[self.__sum_next_index()] = row
        else:
            err_text = (f"CableTable.sum_add_cable_table_row\n"
                        f"  <{row}> не является объектом класса CableTableRow")
            ErrorLog.add_error(err_text)

    def sum_concatenate_summary_row(self, row: SummaryRow):
        if isinstance(row, SummaryRow):
            ind = self.__find_sum_row(row)
            # print(row.cable_full_name, ind)
            if ind == -1:
                self.sum_add_summary_row(row)
            else:
                self.summary_table[ind].add_length_value(row)
        else:
            err_text = (f"CableTable.sum_concatenate_summary_row\n"
                        f"  <{row}> не является объектом класса SummaryRow")
            ErrorLog.add_error(err_text)

    """
    Вспомогательные функции
    """

    # разбиваем CJ_CABLE_LABEL на CJ_CABLE_CODE | CJ_CABLE_NUMBER | CJ_ADDITIONAL_CODE
    def cable_label_to_code_number_add_code(self):
        for index, row in self.table.items():
            if isinstance(row, CableTableRow):
                cable_label = row.el[CJ_CABLE_LABEL].value
                cable_cls = CableLabelCls(cable_label)
                row.el[CJ_CABLE_CODE].value = cable_cls.cable_code
                row.el[CJ_CABLE_NUMBER].value = cable_cls.cable_number
                row.el[CJ_ADDITIONAL_CODE].value = cable_cls.additional_code
            else:
                ErrorLog.add_error("ERROR CABLE_LABEL_TO_CODE_NUMBER_ADD_CODE")
                ErrorLog.exit()

    # Вывод основной таблицы кабельных соединений
    def print_table_to_console(self,
                               column_list=ColumnsCableTable.main_list,
                               max_len=400,
                               file_name="table_cable.txt"):
        out_list = []
        for index, row in self.table.items():
            out_list.append(row.to_list(column_list))
        text = print_att_list_table(out_list,
                                    max_len=max_len,
                                    title_row=column_list,
                                    title="CableTable.table")
        if self.dbg and self.t_com.out_dir is not None:
            save_text_to_file(text, file_name=file_name, out_dir=self.t_com.out_dir)

    # Вывод маленькой таблицы сумм кабелей
    def print_sum_table_to_console(self, max_len=400, mode="std"):
        out_list = self.get_sum_table_list(mode=mode)
        text = print_att_list_table(out_list,
                                    max_len=max_len,
                                    title=f"CableTable.summary_table (mode: {mode})")
        if self.dbg and self.t_com.out_dir is not None:
            save_text_to_file(text, file_name="table_summary.txt", out_dir=self.t_com.out_dir)

    def get_sum_table_list(self, mode="std", sort=True):
        out_list = []
        for index, row in self.summary_table.items():
            if isinstance(row, SummaryRow):
                if mode == "std":
                    out_list.append(row.get_cj_sum_table_list())
                elif mode == "full":
                    out_list.append(row.get_all_arr_list())
                else:
                    ErrorLog.add_error(f"No Such Mode : {mode} (CableTable.print_sum_table_to_console)",
                                       print_flag=1)
                    ErrorLog.exit()
        if sort:
            out_list.sort(key=lambda x: x[1])
            out_list.sort(key=lambda x: x[0])
        return out_list

    def get_frequency_tags_list(self):
        looking_list = [
            CJ_FROM,
            CJ_FROM_TERMINAL,
            CJ_WHERE,
            CJ_WHERE_TERMINAL
        ]
        for index, row in self.table.items():
            if isinstance(row, CableTableRow):
                for att in looking_list:
                    tag_list = get_tag(row.el[att].value)
                    if tag_list:
                        for tag in tag_list:
                            if tag in self.frequency_tags:
                                self.frequency_tags[tag] +=1
                            else:
                                self.frequency_tags[tag] = 1
        if self.frequency_tags:
            self.frequency_tags = sorted(self.frequency_tags.items(), key=lambda item: item[1], reverse=True)

        return self.frequency_tags


