from base.base_classes import RowStd
from base.tables_columns import CODE, TYPE_MARK
from cable_mapping.constants import G_BASE_CABLE_LAYING_FLAG, G_BASE_CABLE_LAYING_FLAG_VALUE, G_BASE_CABLE_LAYING_TYPE, \
    G_BASE_CABLE_LAYING_VARIANTS
from utils.error_log import ErrorLog

"""
CableLayingTypes:
Считывает сочетания из гугл базы
https://docs.google.com/spreadsheets/d/1P_9LcZ2LGqWRUcjvun5G1r3cd78vGe6fhxPo7tasqDc/edit?usp=sharing
    Если для позиции указано значение "КНС_ДЛЯ_КЖ" то в следующих двух столбцах ищем: 
        - тип прокладки (например "В металлорукаве")
        - альтернативные названия материала (например TA-GN 100×60; TA-GN 100x60)
Функции:
    get_type_by_material    ВХОД:   тип материала (например TA-GN 100×60)
                            ВЫХОД:  способ прокладки (например "В металлорукаве") или ошибка (-1)

Вспомогательные функции для отладки:   
    get_list                ВЫХОД: types_dict-> list[list]                        
"""

OPEN_LAYING = "open_laying"

class CableLayingTypes:
    def __init__(self, base: list[RowStd], open_laying_raw: list[list]):
        self.types_dict = {}  # Словарь типов проклади (короб, лоток и пр)

        def add_to_dict_k_v(key, value):
            if key not in self.types_dict:
                self.types_dict[key] = value
            else:
                ErrorLog.add_error(f"CableLayingTypes:"
                                   f"   ПОВТОР! {key} - уже есть в типах прокладки\n"
                                   f"    {row.el[CODE].value} -> {value}")

        def variants_to_list(in_text: str):
            if in_text != "":
                val_list = in_text.split(";")
                for x in range(len(val_list)):
                    val_list[x] = str(val_list[x]).strip()
            else:
                val_list = []
            return val_list

        for row in base:
            # Добавление типов прокладки для материала
            if row.el[G_BASE_CABLE_LAYING_FLAG].value == G_BASE_CABLE_LAYING_FLAG_VALUE:
                value = row.el[G_BASE_CABLE_LAYING_TYPE].value
                key = row.el[TYPE_MARK].value
                add_to_dict_k_v(key, value)

                types_list_variants = variants_to_list(row.el[G_BASE_CABLE_LAYING_VARIANTS].value)

                if types_list_variants:
                    for var_type in types_list_variants:
                        add_to_dict_k_v(var_type, value)

        ##
        # Заполняем способы открытой прокладки
        ##
        ind_no_out = -1
        for x in range(len(open_laying_raw)):
            for y in range(len(open_laying_raw[x])):
                if open_laying_raw[x][y] == OPEN_LAYING:
                    ind_no_out = y
                elif y == ind_no_out:
                    value = "Открыто по конструкциям"
                    key = str(open_laying_raw[x][y]).strip()
                    add_to_dict_k_v(key, value)

                    # types_list_variants = variants_to_list(row.el[G_BASE_CABLE_LAYING_VARIANTS].value)
                    # self.no_out_dict[x] = str(open_laying_raw[x][y]).strip()

    def get_list(self):
        out_list = []
        for k, v in self.types_dict.items():
            out_list.append([k, v])
        return out_list

    def get_type_by_material(self, material):
        if material in self.types_dict:
            out_text = self.types_dict[material]
        else:
            out_text = -1
        return out_text
