import sqlite3

from base.base_classes import RowStd
from base.tables_columns import *
from nano_cad.fields_tables.tables_fields import table_fields
from nano_cad.fields_in_tables import TableFields


def get_list_from_db(cursor):
    out_list = []
    while True:
        next_row = cursor.fetchone()
        if next_row:
            out_list.append(next_row[0])
        else:
            break
    return out_list


def import_db(db, dbg=True):
    conn = sqlite3.connect(db)
    cursor = conn.cursor()

    #  Собираем названия таблиц из базы
    cursor.execute("SELECT name FROM sqlite_master WHERE type='table';")
    tabel_list = get_list_from_db(cursor)
    #
    table_raw_obj = []

    for tab_name in tabel_list:
        # print("\t", tab_name)
        # Собираем имена столбцов из таблицы БД
        # cursor.execute(f"SELECT name FROM PRAGMA_TABLE_INFO('{tab_name}');")
        # fields_str, row_list = sql_get_fields(get_list_from_db(cursor))

        if tab_name in table_fields:
            tab_class = TableFields(tab_name)
            # Если у таблицы есть нужные поля
            if tab_class.get_cursor_fields_str():
                selected_columns = tab_class.get_cursor_fields_str()
                query = f"SELECT {selected_columns} FROM '{tab_name}'"
                try:
                    cursor.execute(query)
                except sqlite3.OperationalError as e:
                    print("ERROR QUERY: ", query)
                    raise e

                while True:
                    next_row = cursor.fetchone()
                    if next_row:
                        mto_row = RowStd()  # Создаем переменную для внесения значений из строк МТО
                        mto_row.el[BD_NAME].value = db  # Имя файла базы данных
                        mto_row.el[BD_TABLE_NAME].value = tab_name  # Имя таблице в БД
                        i = 0
                        for std_row_column, db_column_name in tab_class.get_dict().items():
                            mto_row.el[std_row_column].value = next_row[i]
                            i += 1

                        table_raw_obj.append(mto_row)
                    else:
                        break

    column_dict = ColNames.NanocadDB.column_dict

    # for row in db_list:
    #     mto_row = RowStd()  # Создаем переменную для внесения значений из строк МТО
    #     for k, v in column_dict.items():
    #         mto_row.el[v] = CheckElement(row[k])  # заносим результат
    #     table_raw_obj.append(mto_row)  # Заносим ряд RowStd в список
    conn.close()
    return table_raw_obj
