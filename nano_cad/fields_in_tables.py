from base.tables_columns import *
from nano_cad.fields_tables.tables_fields import table_fields
from utils.error_log import ErrorLog


class TableFields:
    def __init__(self, tab_name):
        self.tab_name = tab_name

    def get_cursor_fields_str(self):
        out_str = ""
        if self.tab_name in table_fields:
            tab_dict = table_fields[self.tab_name]
            if tab_dict:
                for k, v in tab_dict.items():
                    out_str = out_str + v + ", "
        out_str = out_str.strip()
        out_str = out_str.strip(",")
        return out_str

    def get_dict(self):
        return table_fields[self.tab_name]

    @staticmethod
    def get_db_column_by_table_std_name(table, std_name):
        if table in table_fields:
            if std_name in table_fields[table]:
                return table_fields[table][std_name]
            else:
                # ErrorLog.add_error(
                #     f"ERROR (GET_DB_COLUMN_BY_TABLE_STD_NAME): Столбец {std_name} не найден в таблице {table}")
                return None
        else:
            # ErrorLog.add_error(
            #     f"ERROR (GET_DB_COLUMN_BY_TABLE_STD_NAME): Таблицы {table} не найдена")
            return None
