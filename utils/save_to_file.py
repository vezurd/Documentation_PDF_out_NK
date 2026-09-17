import os

import utils.path
from pathlib import Path


def save_text_to_file(text, file_name, out_dir, open_file_flag=0):
    try:
        out_dir = utils.path.get_path_from_file_path(out_dir)
        utils.path.make_dir(out_dir)
        full_path_to_file = Path(out_dir, file_name.strip("\\").strip("/"))
        with open(full_path_to_file, "w", encoding="utf-8") as f:
            f.write(text)
        if open_file_flag:
            utils.path.open_file(full_path_to_file)
    except PermissionError as err:
        print(err)
        print(f"ERROR: Не удалось сохранить текст в файл - {file_name}")


def save_string_to_file(content, file_path, encoding='utf-8'):
    """
    Сохраняет строку в файл.

    Args:
        content (str): Строка для сохранения
        file_path (str): Путь к файлу для сохранения
        encoding (str): Кодировка файла (по умолчанию utf-8)

    Returns:
        bool: True если успешно, False если ошибка
    """
    try:
        # Создаем директорию если она не существует
        directory = os.path.dirname(file_path)
        if directory and not os.path.exists(directory):
            os.makedirs(directory, exist_ok=True)

        # Сохраняем содержимое в файл
        with open(file_path, 'w', encoding=encoding) as file:
            file.write(content)

        print(f"Файл успешно сохранен: {file_path}")
        return True

    except Exception as e:
        print(f"Ошибка при сохранении файла {file_path}: {e}")
        return False