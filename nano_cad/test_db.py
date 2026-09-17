import sqlite3


def search_in_all_tables(db_path, search_string):
    """Поиск строки во всех таблицах SQLite базы"""

    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()

    # Получаем список всех таблиц
    cursor.execute("SELECT name FROM sqlite_master WHERE type='table';")
    tables = cursor.fetchall()

    results = []

    for table in tables:
        table_name = table[0]

        # Получаем информацию о колонках таблицы
        cursor.execute(f"PRAGMA table_info({table_name})")
        columns = cursor.fetchall()

        # Проверяем только текстовые колонки
        text_columns = []
        for col in columns:
            col_name = col[1]
            col_type = col[2].upper()
            if 'TEXT' in col_type or 'CHAR' in col_type or 'VARCHAR' in col_type:
                text_columns.append(col_name)

        # Ищем в каждой текстовой колонке
        for column in text_columns:
            try:
                query = f"""
                SELECT '{table_name}' as table_name, 
                       '{column}' as column_name,
                       rowid,
                       {column} as found_value
                FROM {table_name}
                WHERE {column} LIKE '%{search_string}%'
                """
                cursor.execute(query)
                rows = cursor.fetchall()

                if rows:
                    results.extend(rows)

            except Exception as e:
                # Пропускаем ошибки (например, BLOB колонки)
                pass

    conn.close()
    return results


# Использование
if __name__ == "__main__":
    db_path = r'C:\YandexDisk\темп\DB\AGCC.287-5601-SOT.db'
    search_text = '5601-PB-01-S-PP-9601'

    results = search_in_all_tables(db_path, search_text)

    if results:
        print(f"Найдено {len(results)} совпадений:")
        for table, column, rowid, value in results:
            print(f"Таблица: {table}, Колонка: {column}, RowID: {rowid}")
            print(f"Значение: {value}")
            print("-" * 50)
    else:
        print("Совпадений не найдено")