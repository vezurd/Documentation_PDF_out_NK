import os
import time
import utils.path
from base.base_classes import *
from base.base_xlsx_load import load_from_xlsx_file, list_raw_to_std
from base.cabinet_composition_info import add_cabinet_composition_info
from base.section_type_info import add_section_type_info


def trim_empty_rows_from_end(input_base_std: list[RowStd], max_empty_rows: int = 50, debug=False) -> list[RowStd]:
    """
    Обрезает пустые строки в конце массива для оптимизации последующей работы.
    
    Args:
        input_base_std: Список строк RowStd
        max_empty_rows: Максимальное количество пустых строк в конце, при превышении которых происходит обрезка
        
    Returns:
        Обрезанный список строк
    """
    if not input_base_std:
        return input_base_std
    
    # Подсчитываем количество пустых строк в конце
    empty_count = 0
    for i in range(len(input_base_std) - 1, -1, -1):
        if input_base_std[i].row_type == RowType.empty_row:
            empty_count += 1
        else:
            break
    
    # Если количество пустых строк превышает лимит, обрезаем массив
    if empty_count > max_empty_rows:
        if debug:
            print(f"Обнаружено {empty_count} пустых строк в конце таблицы. Обрезаем до {max_empty_rows} строк для оптимизации.")
        return input_base_std[:len(input_base_std) - empty_count + 2]
    
    return input_base_std


def get_mto_std_from_file(t_com: TableComments, dbg=0):
    file_full_path = str(getattr(t_com, "file_full_path", "") or "").strip()
    dir_path = str(getattr(t_com, "dir_path", "") or "").strip()

    # Режим папки: всегда ищем MTO в выбранной папке.
    if dir_path and dir_path != "-1":
        curr_mto = utils.path.get_files_single(t_com.dir_path, [".xlsx"])
        if curr_mto == -1:
            print(f"По пути <{t_com.dir_path}> файлы .xlsx не найдены.")
            return None
        found_mto = False
        for doc in curr_mto:
            # print("\t"+doc.file_full_path, "<get_mto_std_from_file>")
            if doc.doc_Type == "MTO":
                t_com.file_full_path = doc.file_full_path
                t_com.file_name = utils.path.get_file_name_from_full_file_path(doc.file_full_path)
                found_mto = True
                break
        if not found_mto:
            print(f"По пути <{t_com.dir_path}> файл MTO не найден.")
            return None

    file_full_path = str(getattr(t_com, "file_full_path", "") or "").strip()
    if not file_full_path or file_full_path == "-1":
        if dir_path and dir_path != "-1":
            print(f"MTO путь не задан или не найден для папки <{dir_path}>.")
        else:
            print("MTO путь не задан.")
        return None

    return get_std_from_excel_file(t_com, debug=bool(dbg))


def get_std_from_excel_file(t_com: TableComments, debug: bool = False):
    # debug = True
    if debug:
        start_total = time.time()
        start = time.time()
        print("01.2 Читаем лист")
    input_raw_obj = load_from_xlsx_file(t_com)
    if debug:
        print(f"    -> Время: {time.time() - start:.3f} сек")
        start = time.time()
        print("01.3 Нормализуем данные (по столбцам, формату, лишним символам и т.п.)")
    input_base_std = list_raw_to_std(input_raw_obj, t_com)
    if debug:
        print(f"    -> Время: {time.time() - start:.3f} сек")
        start = time.time()
        print("Обрезаем избыточные пустые строки в конце для оптимизации")
    input_base_std = trim_empty_rows_from_end(input_base_std)
    if debug:
        print(f"    -> Время: {time.time() - start:.3f} сек")
        start = time.time()
        print("Добавляем обработку состава шкафов")
    input_base_std = add_cabinet_composition_info(input_base_std)
    if debug:
        print(f"    -> Время: {time.time() - start:.3f} сек")
        start = time.time()
        print("Определяем типы секций (section_type)")
    source_file = os.path.basename(getattr(t_com, "file_full_path", "") or "")
    input_base_std = add_section_type_info(input_base_std, source_file=source_file)
    if debug:
        print(f"    -> Время: {time.time() - start:.3f} сек")
        print(f"=== Общее время выполнения: {time.time() - start_total:.3f} сек ===")
    return input_base_std


