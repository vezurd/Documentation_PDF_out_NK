
from base.base_classes import RowStd
from base.tables_columns import *
from prettytable import PrettyTable
import pandas as pd
import os
from datetime import datetime


class DsMtoTable:
    def __init__(self, ds_base: list[RowStd], debug: bool = False):
        self.ds_base = ds_base
        self.debug = debug
        
    def get_ds_mto_list(self) -> dict[str, list[str]]:
        """
        Возвращает словарь где ключи - это DS_NAME, значения - это список уникальных DS_TITLE + "-" + DS_SYSTEM
        """
        result_dict = {}
        
        for row in self.ds_base:
            ds_name = row.get_value(DS_NAME)
            ds_title = row.get_value(DS_TITLE)
            ds_system = row.get_value(DS_SYSTEM)
            
            if ds_name:  # Проверяем, что DS_NAME не пустой
                # Формируем значение как DS_TITLE + "-" + DS_SYSTEM
                value = f"{ds_title}-{ds_system}"
                
                # Если DS_NAME уже есть в словаре, добавляем к существующему множеству
                if ds_name in result_dict:
                    result_dict[ds_name].add(value)
                else:
                    result_dict[ds_name] = {value}
        
        # Преобразуем множества обратно в списки для возврата
        for ds_name in result_dict:
            result_dict[ds_name] = list(result_dict[ds_name])
        
        if self.debug:
            self._debug_print_ds_mto_list(result_dict)
            
        return result_dict
    
    def _debug_print_ds_mto_list(self, ds_dict: dict[str, list[str]]):
        """
        Функция отладки с выводом словаря в консоль через prettytable и сохранением в xlsx
        """
        if not ds_dict:
            print("Словарь DS_MTO пуст")
            return
            
        table = PrettyTable()
        table.field_names = ["DS_NAME", "Количество", "DS_TITLE-DS_SYSTEM"]
        
        # Подготавливаем данные для xlsx
        xlsx_data = []
        
        total_specifications = 0
        for ds_name, ds_values in ds_dict.items():
            count = len(ds_values)
            total_specifications += count
            
            # Каждое имя MTO выводим с новой строки
            values_str = "\n".join(ds_values)
            
            table.add_row([ds_name, count, values_str])
            
            # Для xlsx каждое значение MTO в отдельной строке
            for ds_value in ds_values:
                xlsx_data.append({
                    "DS_NAME": ds_name,
                    "Количество": count,
                    "DS_TITLE-DS_SYSTEM": ds_value
                })
        
        print("\n=== Отладочная информация: DS_MTO словарь ===")
        print(table)
        print(f"Уникальных DS_NAME: {len(ds_dict)}")
        print(f"Всего спецификаций: {total_specifications}")
        
        # Сохранение в xlsx файл
        # try:
        #     # Создаем папку для отладочных файлов, если её нет
        #     debug_dir = "debug_output"
        #     if not os.path.exists(debug_dir):
        #         os.makedirs(debug_dir)
        #
        #     # Генерируем имя файла с временной меткой
        #     timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        #     filename = f"ds_mto_debug_{timestamp}.xlsx"
        #     filepath = os.path.join(debug_dir, filename)
        #
        #     # Создаем DataFrame и сохраняем в xlsx
        #     df = pd.DataFrame(xlsx_data)
        #     df.to_excel(filepath, index=False, sheet_name="DS_MTO_Debug")
        #
        #     print(f"Отладочная таблица сохранена в файл: {filepath}")
        #
        # except Exception as e:
        #     print(f"Ошибка при сохранении в xlsx: {e}")
        
        print("=" * 50)
    
    def get_rows_by_ds_title_system(self, ds_title_system: str) -> list[RowStd]:
        """
        Возвращает список RowStd по заданному DS_TITLE-DS_SYSTEM
        
        Args:
            ds_title_system: Строка в формате "DS_TITLE-DS_SYSTEM" (например, "2210-KSB")
            
        Returns:
            list[RowStd]: Список строк, соответствующих заданному DS_TITLE-DS_SYSTEM
        """
        result_rows = []
        
        for row in self.ds_base:
            ds_title = row.get_value(DS_TITLE)
            ds_system = row.get_value(DS_SYSTEM)
            
            # Формируем строку для сравнения
            current_ds_title_system = f"{ds_title}-{ds_system}"
            
            if current_ds_title_system == ds_title_system:
                result_rows.append(row)
        
        if self.debug:
            print(f"\n=== Поиск по DS_TITLE-DS_SYSTEM: {ds_title_system} ===")
            print(f"Найдено строк: {len(result_rows)}")
            print("=" * 50)
        
        return result_rows
    
    def calculate_sum_by_column(self, rows: list[RowStd], column: str) -> float:
        """
        Вычисляет сумму значений в указанном столбце для списка строк
        
        Args:
            rows: Список RowStd для суммирования
            column: Название столбца для суммирования (VALUES или VALUES_2)
            
        Returns:
            float: Сумма всех значений в указанном столбце
        """
        if column not in [VALUES, VALUES_2]:
            raise ValueError(f"Недопустимый столбец: {column}. Используйте {VALUES} или {VALUES_2}")
        
        total_sum = 0.0
        
        for row in rows:
            value = row.get_value(column)
            if value is not None:
                try:
                    # Преобразуем в число и добавляем к сумме
                    numeric_value = float(value)
                    total_sum += numeric_value
                except (ValueError, TypeError):
                    # Пропускаем значения, которые нельзя преобразовать в число
                    if self.debug:
                        print(f"Предупреждение: не удалось преобразовать значение '{value}' в число для столбца {column}")
                    continue
        
        if self.debug:
            print(f"\n=== Суммирование по столбцу {column} ===")
            print(f"Обработано строк: {len(rows)}")
            print(f"Сумма: {total_sum}")
            print("=" * 50)
        
        return total_sum
    
    def calculate_sum_by_code(self, rows: list[RowStd], column: str, code: str) -> dict[str, float]:
        """
        Вычисляет сумму значений в указанном столбце, группируя по коду
        
        Args:
            rows: Список RowStd для суммирования
            column: Название столбца для суммирования (VALUES или VALUES_2)
            code: Название столбца с кодом (CODE или CODE_2)
            
        Returns:
            dict[str, float]: Словарь где ключи - это коды, значения - суммы по столбцу
        """
        if column not in [VALUES, VALUES_2]:
            raise ValueError(f"Недопустимый столбец для суммирования: {column}. Используйте {VALUES} или {VALUES_2}")
        
        if code not in [CODE, CODE_2]:
            raise ValueError(f"Недопустимый столбец для кода: {code}. Используйте {CODE} или {CODE_2}")
        
        result_dict = {}
        
        for row in rows:
            code_value = row.get_value(code)
            column_value = row.get_value(column)
            
            if code_value is not None:
                # Преобразуем код в строку для использования как ключ
                code_str = str(code_value)
                
                if column_value is not None:
                    try:
                        # Преобразуем значение в число
                        numeric_value = float(column_value)
                        
                        # Добавляем к сумме для данного кода
                        if code_str in result_dict:
                            result_dict[code_str] += numeric_value
                        else:
                            result_dict[code_str] = numeric_value
                            
                    except (ValueError, TypeError):
                        # Пропускаем значения, которые нельзя преобразовать в число
                        if self.debug:
                            print(f"Предупреждение: не удалось преобразовать значение '{column_value}' в число для столбца {column}")
                        continue
                else:
                    # Если значение столбца None, добавляем 0 для данного кода
                    if code_str not in result_dict:
                        result_dict[code_str] = 0.0
        
        if self.debug:
            print(f"\n=== Суммирование по коду {code} и столбцу {column} ===")
            print(f"Обработано строк: {len(rows)}")
            print(f"Найдено уникальных кодов: {len(result_dict)}")
            for code_key, sum_value in result_dict.items():
                print(f"  {code_key}: {sum_value}")
            print("=" * 50)
        
        return result_dict
    
    def compare_dictionaries(self, dict1: dict[str, float], dict2: dict[str, float], tolerance: float = 1e-6) -> dict[str, dict]:
        """
        Сравнивает два словаря с суммами по кодам и возвращает детальный отчет о различиях
        
        Args:
            dict1: Первый словарь для сравнения (результат calculate_sum_by_code)
            dict2: Второй словарь для сравнения (результат calculate_sum_by_code)
            tolerance: Допустимая погрешность для сравнения чисел с плавающей точкой
            
        Returns:
            dict[str, dict]: Словарь с результатами сравнения, содержащий:
                - 'identical': bool - одинаковы ли словари
                - 'only_in_first': dict - коды, которые есть только в первом словаре
                - 'only_in_second': dict - коды, которые есть только во втором словаре
                - 'different_values': dict - коды с разными значениями
                - 'summary': dict - общая статистика
        """
        result = {
            'identical': True,
            'only_in_first': {},
            'only_in_second': {},
            'different_values': {},
            'summary': {
                'total_codes_first': len(dict1),
                'total_codes_second': len(dict2),
                'common_codes': 0,
                'total_difference': 0.0
            }
        }
        
        # Получаем все уникальные ключи из обоих словарей
        all_keys = set(dict1.keys()) | set(dict2.keys())
        
        for key in all_keys:
            value1 = dict1.get(key, 0.0)
            value2 = dict2.get(key, 0.0)
            
            if key not in dict1:
                # Код есть только во втором словаре
                result['only_in_second'][key] = value2
                result['identical'] = False
            elif key not in dict2:
                # Код есть только в первом словаре
                result['only_in_first'][key] = value1
                result['identical'] = False
            else:
                # Код есть в обоих словарях - сравниваем значения
                result['summary']['common_codes'] += 1
                difference = abs(value1 - value2)
                
                if difference > tolerance:
                    result['different_values'][key] = {
                        'first_value': value1,
                        'second_value': value2,
                        'difference': difference,
                        'relative_difference': difference / max(abs(value1), abs(value2)) if max(abs(value1), abs(value2)) > 0 else 0
                    }
                    result['identical'] = False
                    result['summary']['total_difference'] += difference
        
        if self.debug:
            print(f"\n=== Сравнение словарей ===")
            print(f"Словари идентичны: {result['identical']}")
            print(f"Кодов в первом словаре: {result['summary']['total_codes_first']}")
            print(f"Кодов во втором словаре: {result['summary']['total_codes_second']}")
            print(f"Общих кодов: {result['summary']['common_codes']}")
            print(f"Кодов только в первом: {len(result['only_in_first'])}")
            print(f"Кодов только во втором: {len(result['only_in_second'])}")
            print(f"Кодов с разными значениями: {len(result['different_values'])}")
            print(f"Общая разность: {result['summary']['total_difference']}")
            
            if result['only_in_first']:
                print("\nКоды только в первом словаре:")
                for key, value in result['only_in_first'].items():
                    print(f"  {key}: {value}")
            
            if result['only_in_second']:
                print("\nКоды только во втором словаре:")
                for key, value in result['only_in_second'].items():
                    print(f"  {key}: {value}")
            
            if result['different_values']:
                print("\nКоды с разными значениями:")
                for key, diff_info in result['different_values'].items():
                    print(f"  {key}: {diff_info['first_value']} vs {diff_info['second_value']} (разность: {diff_info['difference']:.6f})")
            
            print("=" * 50)
        
        return result
        