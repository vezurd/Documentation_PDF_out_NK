"""
Модуль для фильтрации MTO файлов
Нахождение корректных МТО, фильтрация МТО из папок old и других временных папок
"""

import os
from typing import List
from prettytable import PrettyTable

from utils.path import get_files_single


def filter_mto_files(mto_path: str, debug: bool = False, flat_structure: bool = False) -> List:
    """
    Находит и фильтрует MTO файлы, исключая файлы из папок old и других временных папок.
    Учитывает структуру: корневая папка -> 4-символьная папка (например 2225) -> 
    (опционально подпапка из списка POS, SPP, SOS и т.д.) -> МТО файл
    
    Args:
        mto_path: Путь к директории с MTO файлами
        debug: Флаг отладки для вывода подробной информации
        flat_structure: Если True — файлы лежат прямо в корне mto_path (без подпапок),
                        проверка структуры папок пропускается.
        
    Returns:
        Список отфильтрованных MTO файлов
    """
    # Список папок для исключения (регистронезависимо)
    excluded_folders = ['old', 'temp', 'tmp', 'backup', 'archive', 'архив', 'старые', "test","__результат_", "Проверка"]
    
    # Список допустимых подпапок (регистронезависимо)
    valid_subfolders = [
        'POS', 'SPP', 'SOS', 'SKUD', 'SOT', 'SOO', 'KSB', 'KSB1', 'KSB2', 
        'SOS1', 'SKUD1', 'SOT1', 'KBI', 'SAGD', 'KSB3', 'KSB4', 'POS1', 
        'SOO1', 'POS2', 'SOT2', 'SOO2', 'POS3', 'SOT3', 'SOO3', 'POS4', 
        'SOT4', 'SOO4', 'SOT5', 'POS5', 'SOO5', 'SOT.1'
    ]
    
    all_files = get_files_single(mto_path, endswith=(".xlsx", ".XLSX"), sub_folders=True)
    if all_files == -1:
        print(f"Сканирование MTO: файлы не найдены в {mto_path}")
        return []
    
    # Этап 1: Фильтруем файлы содержащие "МТО" в имени (регистронезависимо)
    mto_files = [
        f for f in all_files 
        if "МТО" in f.file_name.upper() or "MTO" in f.file_name.upper()
    ]
    
    if debug:
        _print_files_table("Все найденные MTO файлы", mto_files, ["№", "Имя файла", "Полный путь"])
    
    if not mto_files:
        print(f"Сканирование MTO: файлы с 'МТО' в имени не найдены в {mto_path}")
        return []
    
    # Нормализуем корневой путь для вычисления относительных путей
    mto_path_normalized = os.path.normpath(mto_path)
    
    # Этап 2: Фильтруем файлы из исключенных папок и проверяем структуру
    filtered_files = []
    excluded_files = []
    
    for mto_file in mto_files:
        # Получаем путь к файлу
        file_path = mto_file.file_full_path
        # Нормализуем путь и разбиваем на части
        path_parts = os.path.normpath(file_path).split(os.sep)
        
        # Получаем относительный путь (без корневого адреса)
        relative_path = _get_relative_path(file_path, mto_path_normalized)
        
        # Проверяем, есть ли в пути исключенные папки
        is_excluded = False
        excluded_folder = None
        
        for part in path_parts:
            part_upper = part.upper()
            for excluded in excluded_folders:
                if excluded.upper() in part_upper:
                    is_excluded = True
                    excluded_folder = part
                    break
            if is_excluded:
                break
        
        # Проверяем структуру папок: должна быть 4-символьная папка и опционально подпапка из списка
        is_valid_structure = _check_folder_structure(path_parts, mto_path_normalized, valid_subfolders)
        
        if is_excluded:
            excluded_files.append((mto_file, excluded_folder, relative_path))
        elif not flat_structure and not is_valid_structure:
            excluded_files.append((mto_file, "Неверная структура папок", relative_path))
        else:
            filtered_files.append((mto_file, relative_path))
    
    print(f"Сканирование: Excel: {len(all_files)} | МТО: {len(mto_files)}"
          f" → Исключено: {len(excluded_files)}, корректных: {len(filtered_files)}")
    
    if debug:
        if excluded_files:
            _print_files_table(
                "Исключенные файлы", 
                [f[0] for f in excluded_files],
                ["№", "Имя файла", "Причина", "Относительный путь", "Полный путь"],
                additional_data=[(f[1], f[2]) for f in excluded_files]
            )
            # Вывод полных путей исключенных файлов
            print("\nПолные пути исключенных файлов:")
            for idx, (excluded_file, excluded_folder, rel_path) in enumerate(excluded_files, 1):
                print(f"  {idx}. {excluded_file.file_full_path} (причина: {excluded_folder}, путь: {rel_path})")
        
        if filtered_files:
            _print_files_table(
                "Корректные MTO файлы (для загрузки)", 
                [f[0] for f in filtered_files],
                ["№", "Имя файла", "Относительный путь", "Полный путь"],
                additional_data=[f[1] for f in filtered_files]
            )
            # Вывод полных путей прошедших фильтрацию файлов
            print("\nПолные пути файлов, прошедших фильтрацию:")
            for idx, (filtered_file, rel_path) in enumerate(filtered_files, 1):
                print(f"  {idx}. {filtered_file.file_full_path} (путь: {rel_path})")
    
    # Возвращаем только файлы без дополнительной информации
    return [f[0] for f in filtered_files]


def _get_relative_path(file_path: str, root_path: str) -> str:
    """
    Получает относительный путь файла от корневой папки
    Возвращает путь вида: 2225\\KSB или просто 2225 (если файл сразу в 4-символьной папке)
    
    Args:
        file_path: Полный путь к файлу
        root_path: Корневой путь
        
    Returns:
        Относительный путь (например: 2225\\KSB или 2225)
    """
    try:
        # Получаем директорию файла
        file_dir = os.path.dirname(file_path)
        
        # Нормализуем пути
        file_dir_norm = os.path.normpath(file_dir)
        root_path_norm = os.path.normpath(root_path)
        
        # Вычисляем относительный путь
        rel_path = os.path.relpath(file_dir_norm, root_path_norm)
        
        # Если путь начинается с .., значит файл вне корневой папки
        if rel_path.startswith('..'):
            # Пытаемся извлечь последние части пути
            path_parts = file_dir_norm.split(os.sep)
            root_parts = root_path_norm.split(os.sep)
            
            # Находим где начинается различие
            for i in range(min(len(path_parts), len(root_parts))):
                if path_parts[i] != root_parts[i]:
                    # Берем части пути после корня
                    if i < len(path_parts):
                        return os.sep.join(path_parts[i:])
                    break
            return os.path.basename(file_dir)
        
        # Если относительный путь - это текущая директория, возвращаем только имя папки
        if rel_path == '.':
            return os.path.basename(file_dir)
        
        return rel_path
    except Exception:
        # В случае ошибки возвращаем только имя папки файла
        return os.path.basename(os.path.dirname(file_path))


def _check_folder_structure(path_parts: List[str], root_path: str, valid_subfolders: List[str]) -> bool:
    """
    Проверяет структуру папок: должна быть 4-символьная папка и опционально подпапка из списка
    
    Структура должна быть:
    - root_path / 4-символьная_папка / (опционально подпапка из valid_subfolders) / файл
    
    Args:
        path_parts: Разбитый на части путь
        root_path: Корневой путь
        valid_subfolders: Список допустимых подпапок
        
    Returns:
        True если структура корректна, False иначе
    """
    # Нормализуем пути
    root_parts = os.path.normpath(root_path).split(os.sep)
    root_parts = [p for p in root_parts if p]  # Убираем пустые элементы
    
    # Находим индекс начала относительно корневой папки
    # Ищем где заканчивается корневой путь
    root_parts_clean = [p for p in root_parts if p]  # Убираем пустые элементы
    path_parts_clean = [p for p in path_parts if p]  # Убираем пустые элементы
    
    # Ищем начало пути относительно корня
    start_idx = 0
    for i in range(len(root_parts_clean)):
        if i < len(path_parts_clean) and root_parts_clean[i] == path_parts_clean[i]:
            start_idx = i + 1
        else:
            break
    
    # Проверяем, что путь содержит корневой путь
    if len(path_parts_clean) <= start_idx:
        return False
    
    # Первая папка после корня должна быть 4-символьной (например: 2225, 2250)
    first_folder = path_parts_clean[start_idx] if start_idx < len(path_parts_clean) else None
    if not first_folder:
        return False
    
    # Проверяем, что папка состоит из 4 символов (обычно цифры)
    if len(first_folder) != 4:
        return False
    
    # Если есть еще одна папка (подпапка), она должна быть из списка допустимых
    # Файл может быть сразу в 4-символьной папке или в подпапке из списка
    if len(path_parts_clean) > start_idx + 1:
        # Проверяем, не является ли следующая часть именем файла (с расширением)
        second_part = path_parts_clean[start_idx + 1]
        # Если это не файл (нет расширения .xlsx), значит это подпапка
        if not second_part.upper().endswith('.XLSX'):
            second_folder_upper = second_part.upper()
            
            # Проверяем, есть ли эта папка в списке допустимых (регистронезависимо)
            is_valid_subfolder = any(
                valid.upper() == second_folder_upper
                for valid in valid_subfolders
            )
            if not is_valid_subfolder:
                return False
            
            # Если это подпапка из списка, следующая часть ДОЛЖНА быть файлом
            # Не должно быть дополнительных папок после подпапки из списка
            if len(path_parts_clean) > start_idx + 2:
                third_part = path_parts_clean[start_idx + 2]
                # Если третья часть - это не файл (нет расширения .xlsx), значит это еще одна папка
                # Это недопустимо - после подпапки из списка должен быть только файл
                if not third_part.upper().endswith('.XLSX'):
                    return False
    
    # Если структура соответствует требованиям, возвращаем True
    return True


def _print_files_table(title: str, files: List, columns: List[str], additional_data: List = None):
    """
    Выводит таблицу с информацией о файлах
    
    Args:
        title: Заголовок таблицы
        files: Список файлов для вывода
        columns: Список названий колонок
        additional_data: Дополнительные данные для каждой строки (опционально)
                        Может быть списком значений или кортежей для нескольких столбцов
    """
    if not files:
        return
    
    table = PrettyTable()
    table.field_names = columns
    table.border = True
    table.align = "l"
    
    for idx, file_obj in enumerate(files):
        row = [idx + 1, file_obj.file_name]
        
        # Добавляем дополнительные данные если есть
        if additional_data and idx < len(additional_data):
            add_data = additional_data[idx]
            # Если это кортеж, добавляем каждый элемент отдельно
            if isinstance(add_data, tuple):
                for item in add_data:
                    row.append(item)
            else:
                row.append(add_data)
        
        # Добавляем путь (последний столбец, если есть в columns)
        if "Полный путь" in columns or len(columns) > len(row):
            full_path = file_obj.file_full_path
            if len(full_path) > 80:
                full_path = "..." + full_path[-77:]
            row.append(full_path)
        
        table.add_row(row)
    
    print(f"\n{title}:")
    print(table)


if __name__ in {"__main__"}:
    # Пример запуска фильтрации
    import sys
    
    if len(sys.argv) > 1:
        mto_path = sys.argv[1]
    else:
        mto_path = r"C:\Python\_ГОТОВЫЕ"
    
    debug = True
    
    filtered_files = filter_mto_files(mto_path, debug=debug)
    print(f"\nИтого отфильтровано корректных MTO файлов: {len(filtered_files)}")

