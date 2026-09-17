"""
Конфигурация столбцов вывода для BBB (BOE, BOM, BOQ).
По аналогии с OUTPUT_COLUMNS_CONFIG из step4_6_save_match_result_to_excel.py.
"""

from dataclasses import dataclass
from typing import List

from base.tables_columns import (
    NUMBERS, TITLE, CODE, NAME, TYPE_MARK, VENDOR, UNITS, VALUES, MASS,
    ANNOTATION, TAGS, ROW_TYPE, SECTION_TYPE,
    BBB_MARKA, BBB_SPEC_NUMBER, BBB_REVISION, BBB_OBJECT_NAME,
    BBB_MTR_TYPE, BBB_MTR_GROUP, BBB_SUPPLY_ZONE, BBB_WORK_CODE,
    BBB_CONSUMPTION_RATE, BBB_UNIT_PRICE, BBB_TOTAL_PRICE,
    BBB_ASSEMBLY, BBB_WORK_NAME, BBB_LABOR_UNIT, BBB_LABOR_TOTAL,
    BBB_EQUIPMENT_UNIT, BBB_EQUIPMENT_TOTAL, BBB_LABOR_RATE,
    BBB_EQUIPMENT_RATE, BBB_DIRECT_COST_RATE, BBB_DIRECT_COST_TOTAL,
)


@dataclass
class ColumnDef:
    col_name: str
    header_label: str
    output: bool = True
    width: float = 15


DEFAULT_COLUMN_WIDTH = 15

HEADER_FILL_COLOR = "#dbbcdb"


BOE_COLUMNS_CONFIG: List[ColumnDef] = [
    ColumnDef(NUMBERS,          "№ п/п",                                            width=5),
    ColumnDef(TITLE,            "Титул",                                            width=7),
    ColumnDef(BBB_MARKA,        "Марка",                                            width=8),
    ColumnDef(BBB_SPEC_NUMBER,  "Номер тех. Спецификации",                          width=29,   output=False),
    ColumnDef(BBB_REVISION,     "Ревизия",                                          width=10),
    ColumnDef(BBB_OBJECT_NAME,  "Наименование объекта / системы",                   width=17),
    ColumnDef(BBB_MTR_TYPE,     "Тип МТР (оборудование/материалы)",                 width=17,   output=False),
    ColumnDef(BBB_MTR_GROUP,    "Группа МТР",                                       width=40),
    ColumnDef(BBB_SUPPLY_ZONE,  "Зона ответственности Поставки",                    width=22,   output=False),
    ColumnDef(BBB_WORK_CODE,    "Код работ (единичных расценок)",                    width=16),
    ColumnDef(TAGS,             "Позиция",                                          width=26),
    ColumnDef(NAME,             "Наименование и техническая характеристика",         width=75),
    ColumnDef(TYPE_MARK,        "Тип, марка, обозначение документа, опросного листа", width=28),
    ColumnDef(CODE,             "Код продукции",                                    width=17),
    ColumnDef(VENDOR,           "Поставщик",                                        width=12),
    ColumnDef(UNITS,            "Ед. измерения",                                    width=12),
    ColumnDef(VALUES,           "Количество",                                       width=13),
    ColumnDef(MASS,             "Масса 1 ед., кг",                                  width=8),
    ColumnDef(ANNOTATION,       "Примечание",                                       width=14),
    ColumnDef(BBB_UNIT_PRICE,   "Стоимость за ед., без НДС 20%, руб.",              width=12,   output=False),
    ColumnDef(BBB_TOTAL_PRICE,  "Стоимость всего, без НДС, руб.",                   width=12,   output=False),
    ColumnDef(BBB_ASSEMBLY,     "Сборка",                                           width=9,    output=False),
    ColumnDef(ROW_TYPE,         "Тип строки",                                       width=12,   output=True),
    ColumnDef(SECTION_TYPE,     "Секция MTO",                                       width=35,   output=True),
]


BOM_COLUMNS_CONFIG: List[ColumnDef] = [
    ColumnDef(NUMBERS,              "№ п/п",                                            width=5),
    ColumnDef(TITLE,                "Титул",                                            width=7),
    ColumnDef(BBB_MARKA,            "Марка",                                            width=8),
    ColumnDef(BBB_SPEC_NUMBER,      "Номер тех. Спецификации",                          width=29,   output=False),
    ColumnDef(BBB_REVISION,         "Ревизия",                                          width=10),
    ColumnDef(BBB_OBJECT_NAME,      "Наименование объекта / системы",                   width=17),
    ColumnDef(BBB_MTR_TYPE,         "Тип МТР (оборудование/материалы)",                 width=17,   output=False),
    ColumnDef(BBB_MTR_GROUP,        "Группа МТР",                                       width=40),
    ColumnDef(BBB_SUPPLY_ZONE,      "Зона ответственности Поставки",                    width=22,   output=False),
    ColumnDef(BBB_WORK_CODE,        "Код работ (единичных расценок)",                    width=16),
    ColumnDef(TAGS,                 "Позиция",                                          width=23),
    ColumnDef(NAME,                 "Наименование и техническая характеристика",         width=75),
    ColumnDef(TYPE_MARK,            "Тип, марка, обозначение документа, опросного листа", width=28),
    ColumnDef(CODE,                 "Код продукции",                                    width=17),
    ColumnDef(VENDOR,               "Поставщик",                                        width=13),
    ColumnDef(UNITS,                "Единица измерения",                                width=12),
    ColumnDef(VALUES,               "Количество с учетом нормы расхода",                width=14),
    ColumnDef(BBB_CONSUMPTION_RATE, "Норма расхода (справочно)",                        width=12),
    ColumnDef(MASS,                 "Масса 1 ед., кг",                                  width=8),
    ColumnDef(ANNOTATION,           "Примечание",                                       width=20),
    ColumnDef(BBB_UNIT_PRICE,       "Стоимость за ед., без НДС 20%, руб.",              width=12,   output=False),
    ColumnDef(BBB_TOTAL_PRICE,      "Стоимость всего, без НДС, руб.",                   width=12,   output=False),
    ColumnDef(BBB_ASSEMBLY,         "Сборка",                                           width=9,   output=False),
    ColumnDef(ROW_TYPE,             "Тип строки",                                       width=12,   output=True),
    ColumnDef(SECTION_TYPE,         "Секция MTO",                                       width=35,   output=True),
]


BOQ_COLUMNS_CONFIG: List[ColumnDef] = [
    ColumnDef(NUMBERS,              "№ п/п",                                            width=5),
    ColumnDef(TITLE,                "Титул",                                            width=7),
    ColumnDef(BBB_MARKA,            "Марка комплекта",                                  width=12),
    ColumnDef(BBB_REVISION,         "Ревизия",                                          width=10),
    ColumnDef(BBB_OBJECT_NAME,      "Наименование объекта / системы",                   width=17),
    ColumnDef(BBB_WORK_CODE,        "Код работ (единичных расценок)",                    width=16),
    ColumnDef(BBB_WORK_NAME,        "Наименование работ (единичных расценок)",           width=75),
    ColumnDef(UNITS,                "Единица измерения",                                width=13),
    ColumnDef(VALUES,               "Количество (всего)",                               width=13),

    ColumnDef(BBB_LABOR_UNIT,       "Трудо-затраты за ед., чел./час.",                  width=11,   output=False),
    ColumnDef(BBB_LABOR_TOTAL,      "Трудо-затраты всего, чел./час.",                   width=11,   output=False),
    ColumnDef(BBB_EQUIPMENT_UNIT,   "Строит. оборудование за ед., маш./час.",           width=16,   output=False),
    ColumnDef(BBB_EQUIPMENT_TOTAL,  "Строит. оборудование всего, маш./час.",            width=14,   output=False),
    ColumnDef(BBB_LABOR_RATE,       "Расценка на оплату труда, руб.",                   width=11,   output=False),
    ColumnDef(BBB_EQUIPMENT_RATE,   "Расценка на строит. оборудование (ЭММ), руб.",     width=16,   output=False),
    ColumnDef(BBB_DIRECT_COST_RATE, "Расценка на прямые затраты, руб.",                 width=11,   output=False),
    ColumnDef(BBB_DIRECT_COST_TOTAL, "Прямые затраты всего, руб.",                      width=12,   output=False),
    
    ColumnDef(ANNOTATION,           "Примечание",                                       width=14),
    ColumnDef(ROW_TYPE,             "Тип строки",                                       width=12,   output=True),
]


BBB_COLUMNS_CONFIGS = {
    "BOE": BOE_COLUMNS_CONFIG,
    "BOM": BOM_COLUMNS_CONFIG,
    "BOQ": BOQ_COLUMNS_CONFIG,
}


def get_bbb_output_config(doc_type: str) -> List[ColumnDef]:
    return BBB_COLUMNS_CONFIGS.get(doc_type, BOE_COLUMNS_CONFIG)
