import datetime
import os
import sys
from os import listdir
from os.path import isfile, join
from pathlib import Path
from types import NoneType
import subprocess
import threading
import time

import utils


# def get_path_out_dir(base_path, dir_result_prefix="/__результат_проверки_"):
#     now = datetime.datetime.now()
#     date = now.strftime("%Y.%m.%d_%HH-%MM")
#     # print(date, "<get_path_out_dir>")
#     path_out_dir = base_path + dir_result_prefix + date + "/"  # Путь для выкладывания результата анализа файлов
#     return path_out_dir
def get_path_out_dir(base_path, dir_result_prefix="/__результат_проверки_"):
    """
    Создает путь для выходной директории результатов.
    Если base_path указывает на файл, использует директорию файла.
    """
    now = datetime.datetime.now()
    date = now.strftime("%Y.%m.%d_%HH-%MM")

    # Проверяем, является ли base_path файлом
    if os.path.isfile(base_path):
        # Если это файл, берем его директорию
        base_dir = os.path.dirname(base_path)
    else:
        # В противном случае используем как директорию
        base_dir = base_path

    # Создаем путь для результатов (используем os.path.join для кроссплатформенности)
    result_dir_name = dir_result_prefix.lstrip('/\\') + date  # Убираем начальные слеши
    path_out_dir = os.path.join(base_dir, result_dir_name)

    # Добавляем разделитель в конец для совместимости
    path_out_dir = os.path.normpath(path_out_dir) + os.sep

    return path_out_dir

def get_path_from_file_path(file_path):
    if os.path.isfile(file_path):
        p = Path(file_path)
        path = str(p.parent)
    elif os.path.isdir(file_path):
        path = file_path
    else:
        utils.error_log.ErrorLog.add_error(f"GET_PATH_FROM_FILE_PATH: \"{file_path}\" не является ни файлом ни директорией")
        utils.error_log.ErrorLog.exit()
        path = None
    return path


def get_file_name_from_full_file_path(full_file_path):
    p = Path(full_file_path)
    path = str(p.name)
    return path

def file_path_plus_file_name(path, name):
    p = Path(path, name)
    return str(p)

def make_dir(path_out_dir):
    path_out_dir = normalize_path(path_out_dir)
    os.makedirs(path_out_dir, exist_ok=True)


def normalize_path(path):
    return str(Path(path))


def get_directory_path(path):
    """
    Извлекает путь к папке из переданного пути.
    Если путь ведет к файлу - возвращает папку, содержащую файл.
    Если путь ведет к папке - возвращает саму папку.
    """
    path_obj = Path(path)

    if path_obj.is_file():
        # Это файл - возвращаем родительскую папку
        return str(path_obj.parent)
    elif path_obj.is_dir():
        # Это папка - возвращаем саму папку
        return str(path_obj)
    else:
        # Путь не существует, но мы все равно пытаемся определить намерение
        # Если есть расширение файла, считаем что это файл
        if path_obj.suffix:
            return str(path_obj.parent)
        else:
            # Нет расширения - считаем что это папка
            return str(path_obj)


def _open_dir_impl(dir_path: str) -> None:
    """
    Opens directory in file explorer. Runs in a background thread to avoid GIL
    crash when called after long-running tasks (e.g. with tqdm) under debugpy.
    """
    time.sleep(0.3)  # Let tqdm and other threads settle before spawning subprocess
    try:
        os.startfile(dir_path)
        print(f"\t\tПапка успешно открыта")
    except Exception:
        try:
            subprocess.run(['explorer', dir_path], check=True)
        except Exception as e2:
            print(f"\t\tОшибка при открытии папки: {e2}")
            try:
                escaped_path = f'"{dir_path}"'
                os.system(f'explorer {escaped_path}')
            except Exception as e3:
                print(f"\t\tКритическая ошибка: не удалось открыть папку: {e3}")


def open_dir(path):
    if path is None:
        print("Передан объект None - нет результата работы программы")
        return

    # Получаем корректный путь к папке (fast, no subprocess)
    dir_path = get_directory_path(path)
    dir_path = normalize_path(dir_path)

    print(f"\t\tОткрываем папку: {dir_path}")

    # Проверяем, существует ли папка
    if not Path(dir_path).exists():
        print(f"\t\tВнимание: папка не существует: {dir_path}")
        Path(dir_path).mkdir(parents=True, exist_ok=True)
        print(f"\t\tСоздана папка: {dir_path}")

    # Открываем в проводнике в фоновом потоке (избегаем GIL crash с debugpy/tqdm)
    threading.Thread(target=_open_dir_impl, args=(dir_path,), daemon=True).start()


def open_file(path):
    if not isinstance(path, NoneType):
        path = str(Path(path))
        print("\t\t" + path, "<open_file>")
        os.system(r"explorer.exe " + path)
    else:
        print("Передан объект None - нет результата работы программы")


def is_file_path(file_path):
    if file_path is None or file_path == "":
        print(f"\tФайл не выбран")
        return False
    else:
        return True


def get_files_single(dir_path,
                     endswith=(".pdf", ".PDF"),
                     forbidden_endswith=(),
                     sub_folders=False,
                     factory=None):
    """Collect files from *dir_path* filtered by extension.

    Args:
        factory: callable(file_path) → document object. When ``None`` (default),
            uses ``V2Document.from_file_path`` (lazy import; no v2 at module level).
    """
    if factory is None:
        from pdf_parsing_v2_engine.document import V2Document

        factory = V2Document.from_file_path

    curr_proj = []
    only_files = []
    for root, dirs, files in os.walk(dir_path):
        for name in files:
            if sub_folders:
                if "результат_проверки" in root:
                    continue
                only_files.append(os.path.join(root, name))
            else:
                if root == dir_path:
                    only_files.append(os.path.join(root, name))

    for next_file in only_files:
        if "~" in next_file:
            continue
        next_file_lower = next_file.lower()
        for e in endswith:
            if next_file_lower.endswith(e.lower()):
                curr_proj.append(factory(next_file))
                break
        else:
            for e in forbidden_endswith:
                if next_file_lower.endswith(e.lower()):
                    utils.error_log.ErrorLog.add_error(next_file + f"    \"{e}\" - необходимо изменить формат")
    curr_proj = list(sorted(curr_proj, key=lambda x: x.doc_Number_for_sort))
    return curr_proj


def copy_file(path_from, path_destination):
    import shutil
    shutil.copy2(path_from, path_destination)
