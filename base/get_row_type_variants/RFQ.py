from base.base_classes import *
from base.tables_columns import UNITS, CODE, TAGS, VALUES, NUMBERS, TYPE_MARK, NAME, TITLE, SPECIFICATION_NAME

column_dict = {
    NUMBERS: 1,
    TITLE: 3,
    SPECIFICATION_NAME: 3,
    TAGS: 3,
    CODE: 1,
    NAME: 1,
    TYPE_MARK: 3,
    UNITS: 1,
    VALUES: 1,
        }


def get_row_type(row: RowStd):
    if ("Наименование, описание" in row.el[NAME].value
            or "Код продукта" in row.el[CODE].value
            or "Кол-во" in row.el[VALUES].value):
        return RowType.other_row
    # POSITION DICT
    column_dict = {
        NUMBERS: 3,
        TITLE: 3,
        SPECIFICATION_NAME: 3,
        TAGS: 3,
        CODE: 1,
        NAME: 1,
        TYPE_MARK: 3,
        UNITS: 1,
        VALUES: 1,
    }
    if RowType.check_row(row, column_dict):
        return RowType.position_row
    # EMPTY DICT
    column_dict = {
        NUMBERS: 0,
        TITLE: 0,
        SPECIFICATION_NAME: 0,
        TAGS: 0,
        CODE: 0,
        NAME: 0,
        TYPE_MARK: 0,
        UNITS: 0,
        VALUES: 0,
    }
    if RowType.check_row(row, column_dict):
        return RowType.empty_row

    return RowType.other_row
