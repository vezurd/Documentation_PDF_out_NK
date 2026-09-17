from base.base_classes import *
from base.tables_columns import UNITS, CODE, TAGS, VALUES, VENDOR, MASS, NUMBERS, TYPE_MARK, NAME, ANNOTATION


def get_row_type(row: RowStd):
    # Проверяем на пустые ячейки   
        # """
        # 0 - должно быть пустым
        # 1 - должно быть НЕ пустым
        # 3 - не важно значение
        # """ 
    column_dict = {
        NUMBERS: 3,  
        TAGS: 3,  
        NAME: 3, 
        VALUES: 3,  
        ANNOTATION: 3, 
        CODE:1, 
        
    }
    #
    if ("Наименование" in row.el[NAME].value
            or "Примечание" in row.el[ANNOTATION].value
            or "Код BCC" in row.el[CODE].value
            or "Ед. изм" in row.el[UNITS].value):
        return RowType.other_row
    # POSITION DICT
    if RowType.check_row(row, column_dict):
        return RowType.position_row

    return RowType.empty_row
