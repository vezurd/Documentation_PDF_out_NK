from base.base_classes import *
from base.tables_columns import UNITS, CODE, TAGS, VALUES, VENDOR, MASS, NUMBERS, TYPE_MARK, NAME


def get_row_type(row: RowStd):
    if ("Наименование и техническая" in row.el[NAME].value
            or "Тип, марка," in row.el[TYPE_MARK].value
            or "Код продукции" in row.el[CODE].value
            or "Поставщик" in row.el[VENDOR].value):
        return RowType.head_row
    # POSITION DICT
    column_dict = {
        TAGS: 3,
        NUMBERS: 3,
        NAME: 1,
        TYPE_MARK: 3,
        CODE: 3,
        VENDOR: 3,
        UNITS: 1,
        VALUES: 1,
        MASS: 3,
    }
    if RowType.check_row(row, column_dict):
        return RowType.position_row
    # SECTION DICT
    column_dict = {
        TAGS: 0,
        NUMBERS: 1,
        NAME: 1,
        TYPE_MARK: 0,
        CODE: 0,
        VENDOR: 0,
        UNITS: 0,
        VALUES: 0,
        MASS: 0,
    }
    if RowType.check_row(row, column_dict):
        return RowType.section_row
    # CABINET DICT
    column_dict = {
        TAGS: 3,
        NUMBERS: 1,
        NAME: 1,
        TYPE_MARK: 1,
        CODE: 0,
        VENDOR: 0,
        UNITS: 0,
        VALUES: 0,
        MASS: 3,
    }
    if RowType.check_row(row, column_dict):
        return RowType.cabinet_title_row

    # SYSTEM DICT
    column_dict = {
        TAGS: 1,
        NUMBERS: 0,
        NAME: 0,
        TYPE_MARK: 0,
        CODE: 0,
        VENDOR: 0,
        UNITS: 0,
        VALUES: 0,
        MASS: 0,
    }
    if RowType.check_row(row, column_dict):
        row.el[NAME].value = row.el[TAGS].value
        row.el[TAGS].value = ""
        return RowType.system_row

    # EMPTY DICT
    column_dict = {
        TAGS: 0,
        NUMBERS: 0,
        NAME: 0,
        TYPE_MARK: 0,
        CODE: 0,
        VENDOR: 0,
        UNITS: 0,
        VALUES: 0,
        MASS: 0,
    }
    if RowType.check_row(row, column_dict):
        return RowType.empty_row

    return RowType.other_row
