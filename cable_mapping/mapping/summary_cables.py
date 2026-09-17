from base.base_classes import *
from base.tables_columns import CODE, TYPE_MARK

from cable_mapping.constants import *
from utils.error_log import ErrorLog
from utils.string_parsing import print_att_list_table, cable_x_fixer

NO_OUT_CABLES = "no_out_cables"


class SummaryRow:
    def __init__(self,
                 cable_full_name,
                 mark,
                 number_of_cores,
                 value=0,
                 code="",
                 variants=[],
                 r_system="NO_SYS"):
        """

        :type variants: object
        """
        self.variants = variants
        self.cable_full_name = cable_full_name
        self.mark = mark
        self.number_of_cores = number_of_cores
        self.code = code
        self.value = value
        self.system = r_system

    def get_all_arr_list(self):
        our_row = [
            self.cable_full_name,
            self.mark,
            self.number_of_cores,
            self.value,
            self.code,
            self.variants,
            self.system
        ]
        return our_row

    def get_cj_sum_table_list(self):
        out_row = [
            self.mark,
            self.number_of_cores,
            self.value]
        return out_row

    def add_length_value(self, inp_val):
        if inp_val == "":
            value = 0
        elif isinstance(inp_val, SummaryRow):
            value = int(inp_val.value)
        else:
            value = int(inp_val)
        self.value = int(self.value) + value


"""
CableSummary:
Класс для работы с таблицей суммирования типов и длин кабелей в КЖ.
Является составной частью класса MapTable.
Для работы собирает из гугл базы данные:
    1. По месту разбития кабеля. 
        Пример КПСВВнг(А)-LS 1x2x0,5 записываем в базе как "КПСВВнг(А)-LS; 1x2x0,5", разбиваем [КПСВВнг(А)-LS, 1x2x0,5]
    2. По типам кабелей которые не надо выводить в сумму и МТО. Берем из таблицы "no_out_cables"
        Пример: Штатный кабель ALM-RD-S06 - 1 м
        https://docs.google.com/spreadsheets/d/1VyHiGjcbc9yow3D5ratq3tSq-A2-S-mmSrwRfA4oNo8/edit?gid=1337835792#gid=1337835792&range=A2        
"""


class CableSummary:
    def __init__(self, base_data_std: list[RowStd],
                 no_out_cables_raw: list[list]):
        self.summary_dict = {}  # Словарь кабелей с разбитием по столбцам
        self.no_out_dict = {}  # Словарь кабелей которые не надо выводить в сумму
        # Добавление типов кабелей для суммы кабелей в КЖ
        for row in base_data_std:
            if row.el[G_BASE_CABLE_LAYING_FLAG].value == G_BASE_CABLE_SUMMARY_FLAG_VALUE:
                cable_full_name = str(row.el[TYPE_MARK].value).strip()
                value = str(row.el[G_BASE_CABLE_LAYING_VARIANTS].value)
                code = str(row.el[CODE].value).strip()
                value_list = value.split(";")
                if len(value_list) == 2:
                    mark = str(value_list[0]).strip()
                    number_of_cores = str(value_list[1]).strip()
                else:
                    print(f"CableLayingTypes -> CABLE_SUMMARY:"
                          f"   Для позиции {cable_full_name}({code}) не правильно задано разбитие на марку и число жил\n"
                          f"    {value}")
                    exit(0)
                # Раскладываем варианты в список
                variants = str(row.el[G_BASE_CABLE_SUMMARY_VARIANTS].value).strip()
                if variants == "":
                    variants = []
                else:
                    variants = variants.split(";")
                    for _ in range(len(variants)):
                        variants[_] = cable_x_fixer(variants[_])

                if cable_full_name not in self.summary_dict:
                    self.summary_dict[cable_full_name] = SummaryRow(cable_full_name, mark, number_of_cores,
                                                                    code=code,
                                                                    variants=variants)
                else:
                    print(f"CableLayingTypes.CABLE_SUMMARY:"
                          f"   ПОВТОР! {cable_full_name}({code}) - уже есть в списке кабелей прокладки\n")
        ##
        # Заполняем лист кабелей no_out
        ##
        ind_no_out = -1
        for x in range(len(no_out_cables_raw)):
            for y in range(len(no_out_cables_raw[x])):
                if no_out_cables_raw[x][y] == NO_OUT_CABLES:
                    ind_no_out = y
                elif y == ind_no_out:
                    self.no_out_dict[x] = str(no_out_cables_raw[x][y]).strip()

    """
    ОСНОВНЫЕ ФУНКЦИИ
    """

    def return_new_row_by_id(self, index):
        if index in self.summary_dict:
            row = self.summary_dict[index]
            return SummaryRow(row.cable_full_name, row.mark, row.number_of_cores, row.value, row.code)

    def get_summary_row_by_cable_cj_mark(self, cable_mark):
        cable_mark_fix = cable_x_fixer(cable_mark)
        for ind, row in self.summary_dict.items():
            if cable_x_fixer(row.cable_full_name) == cable_mark_fix:
                return self.return_new_row_by_id(ind)
            if row.variants:
                for _ in row.variants:
                    if cable_x_fixer(_) == cable_mark_fix:
                        return self.return_new_row_by_id(ind)
        for ind, row in self.no_out_dict.items():
            if cable_x_fixer(row) == cable_mark_fix:
                return NO_OUT_CABLES
        err_text = (f"CAB_MAPPING.GET_SUMMARY_FROM_CJ_TABLE.GET_SUMMARY_FROM_CJ_TABLE.GET_SUMMARY_ROW_BY_CABLE_CJ_MARK\n"
                    f"  <{cable_mark}> не найден в summary_dict и no_out_dict")
        ErrorLog.add_error(err_text)
        return "Кабель не найден"

    """
    Вспомогательные функции
    """

    def print_summary_cables(self):
        out_list = []
        for k, v in self.summary_dict.items():
            if isinstance(v, SummaryRow):
                out_list.append(v.get_all_arr_list())
        if out_list:
            print_att_list_table(out_list, max_len=400, title="CableLayingTypes.print_summary_cables")
        else:
            print("CableLayingTypes.print_summary_cables пустой")

    def print_no_out_cables(self):
        out_list = []
        for k, v in self.no_out_dict.items():
            out_list.append([k, v])
        if out_list:
            print_att_list_table(out_list, max_len=400, title="CableLayingTypes.print_no_out_cables")
        else:
            print("CableLayingTypes.print_no_out_cables пустой")
