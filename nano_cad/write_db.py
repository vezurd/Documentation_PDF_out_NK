import sqlite3
from base.base_classes import RowStd
from base.tables_columns import *
from nano_cad.fields_in_tables import TableFields
from utils.colors import *


def write_db(db, std_table: list[RowStd]):
    # Устанавливаем соединение с базой данных
    print(f"Внесение изменений в БД: {db}")
    counter = 0
    connection = sqlite3.connect(db)
    cursor = connection.cursor()
    #
    for row in std_table:
        db_table = row.el[BD_TABLE_NAME].value
        db_id = row.el[BD_ID].value
        for att in row.el.keys():
            el = row.el[att]
            db_column = TableFields.get_db_column_by_table_std_name(db_table, att)
            if db_column and el.color == Color.correction_db:
                query = f"UPDATE {db_table} SET {db_column} = \'{el.value}\' WHERE id = {db_id}"
                print(query)
                cursor.execute(query)
                counter += 1

    # Сохраняем изменения и закрываем соединение
    connection.commit()
    connection.close()
    print(f"\tЗавершено внесение изменений:\n"
          f"\t\t{counter} изменено записей\n")
