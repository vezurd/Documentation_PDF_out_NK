
# Этот код определяет функцию `get_unique_row_list`, которая обрабатывает список объектов `RowStd`
# с целью создания уникального списка строк на основе определенных атрибутов. Он использует ведение журнала для
# обработки ошибок и включает проверки на допустимые типы данных, что делает его необходимым для обеспечения
# целостности данных при обработке.

# Ключевые возможности для улучшения включают оценку использования `eval` на предмет потенциальных рисков для
# безопасности и производительности,
# улучшение обработки ошибок и, возможно, разбиение сложных разделов на более мелкие функции для лучшей читаемости.


from typing import List

from base.base_cheks import *
from base.tables_columns import CODE, TAGS, VALUES, NUMBERS, MTO_NAME
from utils.error_log import ErrorLog


def get_unique_row_list(base: List[RowStd], base_unique=None, check_position_flag=1):
    """
    Retrieve a unique list of RowStd objects based on specific attributes.

    Parameters:
    base (List[RowStd]): A list of RowStd objects to process.
    base_unique (List[RowStd], optional): An optional list to accumulate unique rows.
    check_position_flag (int, optional): Flag indicating whether to check the position of rows.

    Returns:
    List[RowStd]: A unique list of RowStd objects.
    """
    # doc_name = base[0].t_com.doc_od_style_file_name
    if base_unique is None:
        base_unique = []

    t_com = TableComments(dir_path=base[0].t_com.dir_path,
                          file_full_path=base[0].t_com.file_full_path,
                          file_full_name=base[0].t_com.file_name,
                          tabel_class=base[0].t_com.tabel_class,
                          )

    for row in base:
        """
        Проверка относится ли тип строки к позиции
        """
        if check_position_flag:
            # Skip rows that do not pass the position check.
            if not check_position_row(row):
                continue

        code = row.el[CODE].value
        # Append the file name to the MTO_NAME field for tracking purposes.
        row.el[MTO_NAME].value.append(t_com.file_name)

        # Check for empty CODE field to add a copy of the row directly.
        if code == "":
            base_unique.append(RowStd.get_row_copy(row, t_com))
            continue

        # Attempt to find a corresponding unique row based on the CODE.
        code_check = RowStd.get_row_by_code(code, base_unique)
        if isinstance(code_check, RowStd):
            row_unique = code_check
            att_list = [VALUES, NUMBERS, TAGS, MTO_NAME, DS_NUMBER]
            for att in att_list:
                if att == VALUES:
                    # Handle value updates, accounting for the potential presence of an equation.
                    inc_val = str(row.el[att].value).replace(",", ".")
                    base_val = str(row_unique.el[att].value).replace(",", ".")
                    if "=" in inc_val:
                        inc_val = inc_val.replace("=", "")
                        # IMPROVEMENT: Using eval can introduce security risks; consider alternative parsing.
                        try:
                            inc_val = eval(inc_val)
                        except NameError as e:
                            print(row.get_text_for_debug())
                            raise e
                    try:
                        # If base_val is empty, default it to zero.
                        if base_val == "":
                            base_val = 0
                        if inc_val == "":
                            inc_val = 0
                        # Update the unique row value by summing the existing and new values.
                        row_unique.el[att].value = float(base_val) + float(inc_val)
                    except ValueError as e:
                        e.add_note(row.get_text_for_debug())
                        raise e
                elif att == TAGS or att == MTO_NAME:
                    # Ensure the unique row's TAGS or MTO_NAME is initialized before concatenation.
                    if row_unique.el[att].value is None:
                        row_unique.el[att].value = []
                    row_unique.el[att].value = row_unique.el[att].value + row.el[att].value
                elif att == NUMBERS:
                    row_unique.el[att].value = str(row_unique.el[att].value) + ", " + str(row.el[att].value)
                elif att == DS_NUMBER:
                    # Ensure the unique row's is initialized before concatenation.
                    if row_unique.el[att].value is None:
                        row_unique.el[att].value = []
                    if row.el[att].value not in row_unique.el[att].value:
                        try:
                            row_unique.el[att].value.append(row.el[att].value)
                        except AttributeError as e:
                            e.add_note(row.get_text_for_debug())
        elif code_check is None:
            base_unique.append(RowStd.get_row_copy(row, t_com))
        else:
            ErrorLog.add_error(f"    ERROR Value_Compare: <{code}> - code, <{code_check}> code_check")
    return base_unique


