import json
import math

from prettytable import PrettyTable

from utils.colors import Color
from cable_mapping.cm_clss.cm_ElementCT import ElementCT
from cable_mapping.cm_clss.cm_ColumnsCableTable import ColumnsCableTable
from cable_mapping.cm_clss.cm_TableCommentsCabMap import TableCommentsCabMap
from cable_mapping.constants import *
from tags.tag_classes import TagClass


class CableTableRow:
    def __init__(self,
                 row_type=None,
                 t_com=TableCommentsCabMap(),
                 system="NO_SYS"):

        self.row_type = row_type
        self.t_com = t_com
        self.system = system
        if self.system == "NO_SYS" and t_com.system != "NO_SYS":
            self.system = t_com.system

        self.el = {}

        for att in ColumnsCableTable.main_list:
            if att == CJ_SYSTEM:
                self.el[CJ_SYSTEM] = ElementCT(self.system)
            else:
                self.el[att] = ElementCT("")

    @staticmethod
    def get_std_row(dic: dict, t_com: TableCommentsCabMap()):
        """
        Запрос на создание стандартной строки CableTableRow
        :param t_com:
        :type dic: dict
        """
        row = CableTableRow(t_com=t_com)

        for k, v in dic.items():
            if k in row.el:
                if isinstance(v, str):
                    v = v.strip()
                row.el[k] = ElementCT(v, Color.no)
        # Заполняем остальные поля стандартного ряда пустыми значениями
        for att in ColumnsCableTable.main_list:
            if att == CJ_SYSTEM:
                row.el[CJ_SYSTEM].value = t_com.system
            elif row.el[att].value is None:
                v = ""
                row.el[att] = ElementCT(v, Color.no)

        return row

    def to_list(self, att_list):
        out_list = []
        for att in att_list:
            out_list.append(str(self.el[att].value))
        return out_list

    def add_laying_type_value(self, laying_type, value: str):
        cur_value = self.el[laying_type].value
        if cur_value == "":
            cur_value = 0
        value = value.replace(",", ".")
        num = math.ceil(float(value))
        self.el[laying_type].value = int(cur_value) + num

    """
    GET сегмент
    """
    matrix = {
        CJ_CABLE_CODE: 1,
        CJ_ADDITIONAL_CODE: 1,
        CJ_FROM: 1,
        CJ_FROM_TERMINAL: 1,
        CJ_WHERE: 1,
        CJ_WHERE_TERMINAL: 1,
        CJ_CABLE_MARK: 1,  # ??
        CJ_ANNOTAION: 0,
    }

    @staticmethod
    def get_cable_no_number_header_row():
        value = []
        for key, flag in CableTableRow.matrix.items():
            if flag:
                value.append(str(key))
        return tuple(value)

    # Получаем код без номера в конце (для функции сравнения номеров старого и нового КЖ)
    def get_cable_no_number_tag(self):
        value = []
        for key, flag in self.matrix.items():
            if flag:
                value.append(str(key + "@" + self.el[key].value))
        return tuple(value)

    @staticmethod
    def get_from_key_value_by_column(input_key: tuple, column: str):
        for i in range(len(input_key)):
            input_key_column = input_key[i].split("@")[0]  # f.e. CJ_CABLE_CODE
            input_key_value = input_key[i].split("@")[1]  # 3252-PB-01-S-SK-1001
            if input_key_column == column:
                return input_key_value
        return "NO_FOUND_KEY_BY_COLUMN"

    @staticmethod
    def get_row_of_value_from_cable_no_number(input_key: tuple):
        value = []
        for i in range(len(input_key)):
            value.append(input_key[i].split("@")[1])  # 3252-PB-01-S-SK-1001
        return value

    # Получаем код без номера в конце (для функции сравнения номеров старого и нового КЖ)
    def get_cable_full_tag(self):
        full_tag = (self.el[CJ_FROM].value + "/"
                    + self.el[CJ_WHERE].value
                    + "-" + self.el[CJ_CABLE_LABEL].value)
        return full_tag

    def get_cable_label(self):
        cab_label = self.el[CJ_CABLE_CODE].value
        add_code = self.el[CJ_ADDITIONAL_CODE].value
        if add_code != "":
            cab_label += f".{add_code}"
        cab_label += f"-{self.el[CJ_CABLE_NUMBER].value}"
        return cab_label

    def get_cable_number(self):
        return self.el[CJ_CABLE_NUMBER].value

    def get_mapping_signature(self):
        t1 = TagClass(self.el[CJ_FROM].value)
        t2 = TagClass(self.el[CJ_WHERE].value)
        sys1 = t1.sub_system_cj
        eq1 = t1.equipment
        sys2 = t2.sub_system_cj
        eq2 = t2.equipment
        return f"{sys1}-{eq1}->{sys2}-{eq2}"

    """
    SETTERS
    """

    def set_cable_number(self, cable_number):
        self.el[CJ_CABLE_NUMBER].value = cable_number
        self.update()

    def set_from_terminal(self, from_terminal):
        self.el[CJ_FROM_TERMINAL].value = from_terminal

    def set_where_terminal(self, where_terminal):
        self.el[CJ_WHERE_TERMINAL].value = where_terminal

    def update(self):
        self.el[CJ_CABLE_LABEL].value = self.get_cable_label()
        self.el[CJ_CABLE_FULL_TAG].value = self.get_cable_full_tag()

    """
    DEBUG
    """

    def debug_info(self):
        """
        Краткая отладочная информация о текущем элементе.
        """
        info = {
            'row_type': self.row_type,
            'system': self.system,
            't_com.system': self.t_com.system,
            'cable_label': self.get_cable_label(),
            'cable_full_tag': self.get_cable_full_tag(),
            'elements': {k: v.value for k, v in self.el.items()}
        }

        return f"CableTableRow Debug Info:\n{json.dumps(info, indent=2, ensure_ascii=False)}"

    from prettytable import PrettyTable

    def debug_info(self):
        """
        Выводит отладочную информацию о текущем элементе класса в виде таблицы.
        Показывает все основные атрибуты и значения элементов.
        """
        debug_output = []
        debug_output.append("=" * 60)
        debug_output.append("DEBUG INFO - CableTableRow")
        debug_output.append("=" * 60)

        # Таблица с основными атрибутами
        main_table = PrettyTable()
        main_table.field_names = ["Attribute", "Value"]
        main_table.align = "l"
        main_table.add_row(["row_type", self.row_type])
        main_table.add_row(["system", self.system])
        main_table.add_row(["t_com.system", self.t_com.system])

        debug_output.append("MAIN ATTRIBUTES:")
        debug_output.append(main_table.get_string())

        # Таблица с элементами
        elements_table = PrettyTable()
        elements_table.field_names = ["Key", "Value"]
        elements_table.align = "l"

        for key, element in self.el.items():
            value = str(element.value) if element.value is not None else "None"
            color = getattr(element, "color", None) if element.color is not None else "None"
            # Обрезаем длинные значения для лучшего отображения
            if len(value) > 50:
                value = value[:47] + "..."
            elements_table.add_row([key, value])

        debug_output.append("\nELEMENTS:")
        debug_output.append(elements_table.get_string())

        # Таблица с вычисляемыми значениями (опционально)
        computed_table = PrettyTable()
        computed_table.field_names = ["Computed Value", "Result"]
        computed_table.align = "l"

        try:
            computed_table.add_row(["Cable label", self.get_cable_label()])
        except:
            computed_table.add_row(["Cable label", "ERROR"])

        try:
            computed_table.add_row(["Cable full tag", self.get_cable_full_tag()])
        except:
            computed_table.add_row(["Cable full tag", "ERROR"])

        try:
            computed_table.add_row(["Mapping signature", self.get_mapping_signature()])
        except:
            computed_table.add_row(["Mapping signature", "ERROR"])

        debug_output.append("\nCOMPUTED VALUES:")
        debug_output.append(computed_table.get_string())

        debug_output.append("=" * 60)

        return "\n".join(debug_output)

    def print_debug(self):
        """
        Просто печатает отладочную информацию в консоль.
        """
        print(self.debug_info())