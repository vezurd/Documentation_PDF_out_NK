from base.base_classes import *
from base.tables_columns import (
    UNITS, CODE, TAGS, VALUES, VENDOR, MASS, NUMBERS, TYPE_MARK, NAME,
    DS_NAME, DS_NUMBER, DS_TITLE, DS_SPECIFICATION, DS_CODE_1C, NAME_2, VALUES_2,
    RFP_SUPPLY_STATUS,
)


def get_row_type(row: RowStd):
    # Проверяем на пустые ячейки
    column_dict = {
        DS_NAME: 1,  # Имя ДС, например ДС15, ДС16 или просто 15 , 16
        DS_NUMBER: 1,  # Порядковый нормер позиции в ДС 1, 2, 3 и т.д
        DS_TITLE: 1,  # 2265-KSB в данном случае
        DS_SPECIFICATION: 1,  # AGCC.287-2265-KSB.MTO-0001
        TAGS: 3,  # 2265-SH-02-S-UZ-0711
        DS_CODE_1C:3,  # 002054119
        CODE:1,  # BCC0001016
        NAME:1,  # Плата источника питания для установки в корпус Elsys-MB
        TYPE_MARK:3,  # Elsys-SWPS-2И
        VALUES:1,  # 1
         UNITS:1,  # шт
         NAME_2:3,
         RFP_SUPPLY_STATUS: 3,  # Optional; old nets without the column still load.
         # Optional: get_std_check_row turns missing cells into "".
         # rfp_parts_net.xlsx has no separate lot column unless written at index 17.
         VALUES_2:3,
    }
    #
    if ("Наименование МТР" in row.el[NAME].value
            or "Код 1С СОУ" in row.el[DS_CODE_1C].value
            or "Технические характеристики" in row.el[TYPE_MARK].value
            or "Ед. изм" in row.el[UNITS].value):
        return RowType.other_row
    # POSITION DICT
    if RowType.check_row(row, column_dict):
        return RowType.position_row

    return RowType.other_row
