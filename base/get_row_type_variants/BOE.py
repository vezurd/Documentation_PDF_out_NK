from base.base_classes import *
from base.tables_columns import UNITS, CODE, TAGS, VALUES, VENDOR, MASS, NUMBERS, TYPE_MARK, NAME


def _is_column_numbers_row(row: RowStd) -> bool:
    """Строка с порядковыми номерами столбцов (1, 2, 3, ... 22) под заголовком."""
    nums = str(row.el[NUMBERS].value or "").strip()
    name = str(row.el[NAME].value or "").strip()
    code = str(row.el[CODE].value or "").strip()
    if nums == "1" and name.isdigit() and code.isdigit():
        return True
    return False


def get_row_type(row: RowStd):
    name_val = str(row.el[NAME].value or "")
    type_mark_val = str(row.el[TYPE_MARK].value or "")
    code_val = str(row.el[CODE].value or "")
    vendor_val = str(row.el[VENDOR].value or "")
    numbers_val = str(row.el[NUMBERS].value or "").strip()

    if ("Наименование и техническая" in name_val
            or "Наименование и тех" in name_val
            or "Тип, марка" in type_mark_val
            or "Код продукции" in code_val
            or "Поставщик" in vendor_val):
        return RowType.head_row

    if _is_column_numbers_row(row):
        return RowType.head_row

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

    column_dict = {
        TAGS: 0,
        NUMBERS: 3,
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

    return RowType.other_row
