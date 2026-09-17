from base.base_classes import *
from base.tables_columns import UNITS, VALUES, NUMBERS, BBB_WORK_CODE, BBB_WORK_NAME


def _is_column_numbers_row(row: RowStd) -> bool:
    """Строка с порядковыми номерами столбцов (1, 2, 3, ... 18) под заголовком."""
    nums = str(row.el[NUMBERS].value or "").strip()
    units = str(row.el[UNITS].value or "").strip()
    vals = str(row.el[VALUES].value or "").strip()
    if nums == "1" and units.isdigit() and vals.isdigit():
        return True
    return False


def get_row_type(row: RowStd):
    numbers_val = str(row.el[NUMBERS].value or "").strip()
    work_code_val = str(row.el[BBB_WORK_CODE].value or "").strip()
    work_name_val = str(row.el[BBB_WORK_NAME].value or "").strip()
    units_val = str(row.el[UNITS].value or "").strip()
    values_val = str(row.el[VALUES].value or "").strip()

    if ("Единица измерения" in units_val
            or "Количество" in values_val
            or "Код работ" in work_code_val
            or "Наименование работ" in work_name_val
            or "п/п" in numbers_val):
        return RowType.head_row

    if _is_column_numbers_row(row):
        return RowType.other_row

    column_dict = {
        NUMBERS: 3,
        BBB_WORK_CODE: 1,
        UNITS: 1,
        VALUES: 1,
    }
    if RowType.check_row(row, column_dict):
        return RowType.position_row

    column_dict = {
        NUMBERS: 0,
        BBB_WORK_CODE: 0,
        BBB_WORK_NAME: 0,
        UNITS: 0,
        VALUES: 0,
    }
    if RowType.check_row(row, column_dict):
        return RowType.empty_row

    if work_name_val and not units_val and not values_val:
        return RowType.section_row

    return RowType.other_row
