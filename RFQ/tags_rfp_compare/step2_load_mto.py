"""
Этап 2: Загрузка данных из MTO файлов
Сканирование директории с поиском файлов Excel содержащих "МТО" в имени
"""

from dataclasses import dataclass
from typing import List, Set, Dict, Literal, Optional
from prettytable import PrettyTable
from collections import defaultdict
import os
from datetime import datetime
import time

import openpyxl
from tqdm import tqdm
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

from base.base_classes import RowStd, TableComments, RowType
import base.t_comm_initial_classes as t_com_init_cls
from base.base_mto import get_mto_std_from_file
from base.tables_columns import (
    DS_TITLE,
    TAGS, NUMBERS, NAME, TYPE_MARK, CODE, VENDOR, UNITS, VALUES, MASS, ANNOTATION,
)
from RFQ.tags_rfp_compare.mto_file_filter import filter_mto_files
from utils.file_name_converts import ProjectFileName
from utils.path import make_dir
from RFQ.tags_rfp_compare.rfp_tags_utils import (
    DEFAULT_UNITS_SPLIT_BAN,
    ensure_result_dir_exists,
    append_timing_log,
    save_input_fingerprint,
)
from utils.cache_utils import cache_manager


Section = Literal["meta", "mto"]


@dataclass
class ColumnDef:
    """Описание столбца для выгрузки Step2 в Excel."""
    col_name: str
    header_label: str
    section: Section
    output: bool = True
    width: int | None = None


TITLE_FILTER_COL = "__title_filter"
SYSTEM_FILTER_COL = "__system_filter"
DEFAULT_COLUMN_WIDTH = 15


MTO_OUTPUT_COLUMNS_CONFIG: List[ColumnDef] = [
    ColumnDef(TITLE_FILTER_COL, "Титул (TITLE)", "meta", width=14),
    ColumnDef(SYSTEM_FILTER_COL, "Система (SYSTEM)", "meta", width=16),
    ColumnDef(TAGS, "Теги", "mto", width=30),
    ColumnDef(NUMBERS, "Номера", "mto", width=12),
    ColumnDef(NAME, "Наименование", "mto", width=35),
    ColumnDef(TYPE_MARK, "Тип/марка", "mto", width=25),
    ColumnDef(CODE, "Код", "mto", width=20),
    ColumnDef(VENDOR, "Производитель", "mto", width=18),
    ColumnDef(UNITS, "Ед. изм.", "mto", width=10),
    ColumnDef(VALUES, "Количество", "mto", width=10),
    ColumnDef(MASS, "Масса", "mto", width=10),
    ColumnDef(ANNOTATION, "Аннотация", "mto", width=40),
]

MTO_COLLAPSE_KEY_COLUMNS = [
    TITLE_FILTER_COL,
    SYSTEM_FILTER_COL,
    TAGS,
    NUMBERS,
    NAME,
    TYPE_MARK,
    CODE,
    VENDOR,
    UNITS,
    MASS,
    ANNOTATION,
]


def step2_load_mto_data(
    mto_path: str,
    result_dir: str = None,
    rfp_data: List[RowStd] = None,
    debug: bool = False,
    column_optimization_hash: str = None,
    save_input_fingerprints: bool = True,
    export_load_results_excel: bool = True,
    export_positions_database_excel: bool = False,
    flat_mto_structure: bool = False,
    units_split_ban: Optional[List[str]] = None,
) -> Dict[str, List[RowStd]]:
    """
    Этап 2: Сканирование пути mto_path с поиском файлов Excel содержащих "МТО" в имени
    и загрузка данных через функцию похожую на get_mto_list_from_ds
    
    Args:
        mto_path: Путь к директории с MTO файлами
        result_dir: Путь к папке результатов (не используется, для совместимости)
        rfp_data: Данные из RFP для возможного использования при загрузке MTO
        debug: Флаг отладки для вывода подробной информации
        
    Returns:
        Словарь: title_system -> список строк RowStd из MTO файлов с этим title_system
    """
    print("=" * 80)
    print(f"ЭТАП 2: Загрузка MTO")
    print(f"Путь: {mto_path}")
    
    # Этап 2.1: Фильтрация MTO файлов
    filter_start = time.perf_counter()
    mto_files = filter_mto_files(mto_path, debug=debug, flat_structure=flat_mto_structure)
    append_timing_log(result_dir, f"step2_filter_mto_files: {time.perf_counter() - filter_start:.3f}s")
    
    if not mto_files:
        print("MTO файлы для загрузки не найдены.")
        return []

    _check_duplicate_mto_files(mto_files)

    if save_input_fingerprints:
        try:
            save_input_fingerprint(
                result_dir=result_dir,
                source_name="MTO",
                base_path=mto_path,
                file_paths=[f.file_full_path for f in mto_files],
            )
        except Exception:
            pass
    
    # # Этап 2.1.5: Проверка соответствия title_system между RFP и MTO файлами
    # if rfp_data:
    #     mto_files = _check_title_system_match(rfp_data, mto_files, result_dir, debug=debug)
    #     if not mto_files:
    #         print("После проверки соответствия title_system не осталось файлов для загрузки.")
    #         return []
    
    # Этап 2.2: Загрузка данных из каждого MTO файла (инкрементальный кэш)
    mto_data: Dict[str, List[RowStd]] = defaultdict(list)
    load_results = []
    errors: List[str] = []
    current_paths = {f.file_full_path for f in mto_files}
    file_meta_by_path = {
        f.file_full_path: {
            "title": str(getattr(f, "doc_Title_4d", "") or "").strip(),
            "system": str(getattr(f, "doc_Marka", "") or "").strip(),
        }
        for f in mto_files
    }

    # Загружаем mto_agg кэш: только contributions для текущих файлов с совпадающим mtime
    file_contributions = {}
    extra_key = column_optimization_hash if column_optimization_hash else None
    cached_contributions = cache_manager.load_mto_data_from_cache(mto_path, verbose=False, extra_key=extra_key)
    if cached_contributions:
        for fp, contrib in cached_contributions.items():
            if fp not in current_paths:
                continue
            try:
                if os.path.exists(fp) and abs(os.path.getmtime(fp) - contrib.get('mtime', 0)) < 1e-6:
                    file_contributions[fp] = contrib
            except Exception:
                pass
    load_loop_start = time.perf_counter()
    total_files = len(mto_files)
    cache_count = len(file_contributions)
    cache_info = f" (кэш: {cache_count})" if cache_count else ""
    print(f"Загрузка: {total_files} файлов{cache_info}")

    for mto_file in tqdm(mto_files, total=total_files, desc="MTO загрузка", unit="файл", ncols=80, dynamic_ncols=False):
        try:
            file_path = mto_file.file_full_path
            file_title_system = ProjectFileName.scan_title_system(mto_file.file_name)
            if file_title_system:
                file_title_system = file_title_system.strip()

            # Берём из mto_agg кэша или загружаем (per-file кэш / парсинг)
            contrib = file_contributions.get(file_path)
            if contrib is None:
                rows = None
                if cache_manager.is_cache_valid(file_path, extra_key=extra_key):
                    rows = cache_manager.load_from_cache(file_path, verbose=False, extra_key=extra_key)
                if rows is None:
                    t_com = TableComments(
                        file_full_path=file_path,
                        dir_path="-1",
                        tabel_class=t_com_init_cls.MTO
                    )
                    rows = get_mto_std_from_file(t_com, dbg=0)
                    cache_manager.save_to_cache(file_path, rows, verbose=False, extra_key=extra_key)
                file_mtime = os.path.getmtime(file_path) if os.path.exists(file_path) else 0.0
                contrib = {'mtime': file_mtime, 'title_system': file_title_system, 'rows': rows or []}
                file_contributions[file_path] = contrib

            rows = contrib.get('rows') or []
            if rows:
                if file_title_system:
                    mto_data[file_title_system].extend(rows)
                else:
                    mto_data[None].extend(rows)
                load_results.append({
                    'file_name': mto_file.file_name,
                    'status': 'Успешно',
                    'rows_count': len(rows),
                    'title_system': file_title_system or "не определен",
                    'error': None
                })
            else:
                load_results.append({
                    'file_name': mto_file.file_name,
                    'status': 'Пустой файл',
                    'rows_count': 0,
                    'title_system': file_title_system or "не определен",
                    'error': None
                })
        except Exception as e:
            error_msg = str(e)
            file_title_system = ProjectFileName.scan_title_system(mto_file.file_name) if hasattr(mto_file, 'file_name') else None
            if file_title_system:
                file_title_system = file_title_system.strip()
            load_results.append({
                'file_name': mto_file.file_name,
                'status': 'Ошибка',
                'rows_count': 0,
                'title_system': file_title_system or "не определен",
                'error': error_msg[:50] + "..." if len(error_msg) > 50 else error_msg
            })
            errors.append(f"{mto_file.file_name}: {e}")
            if file_path in file_contributions:
                del file_contributions[file_path]

    for err in errors:
        print(f"Ошибка: {err}")

    # Удаляем из кэша записи файлов, которых больше нет
    for fp in list(file_contributions.keys()):
        if fp not in current_paths:
            del file_contributions[fp]
    cache_manager.save_mto_data_to_cache(mto_path, file_contributions, verbose=False, extra_key=extra_key)

    append_timing_log(result_dir, f"step2_load_mto_loop_total: {time.perf_counter() - load_loop_start:.3f}s")

    if export_load_results_excel and load_results and result_dir:
        excel_path = _save_load_results_table_to_excel(load_results, result_dir)
        if excel_path:
            print(f"Результаты загрузки: {os.path.basename(excel_path)}")

    if export_positions_database_excel and result_dir:
        mto_positions_excel_path = _save_mto_positions_database_to_excel(
            mto_data=dict(mto_data),
            result_dir=result_dir,
            file_meta_by_path=file_meta_by_path,
            units_split_ban=units_split_ban,
        )
        if mto_positions_excel_path:
            print(f"База position_row MTO: {os.path.basename(mto_positions_excel_path)}")
    
    return dict(mto_data)


def _check_title_system_match(rfp_data: List[RowStd], mto_files: List, result_dir: str = None, debug: bool = False) -> List:
    """
    Проверяет соответствие title_system между RFP данными и MTO файлами.
    
    Args:
        rfp_data: Список строк RFP данных
        mto_files: Список MTO файлов для проверки
        debug: Флаг отладки
        
    Returns:
        Отфильтрованный список MTO файлов, которые совпадают по title_system
    """
    print("\n" + "=" * 80)
    print("ЭТАП 2.1.5: Проверка соответствия title_system между RFP и MTO")
    print("=" * 80)
    
    # Пункт 1: Собираем уникальные значения DS_TITLE из rfp_data
    rfp_title_systems: Set[str] = set()
    for row in rfp_data:
        if row.row_type != RowType.position_row:
            continue
        title_system = row.get_value(DS_TITLE)
        if title_system:
            rfp_title_systems.add(str(title_system).strip())
    
    print(f"\nНайдено уникальных title_system в RFP: {len(rfp_title_systems)}")
    if debug:
        print("Список title_system из RFP:")
        for ts in sorted(rfp_title_systems):
            print(f"  - {ts}")
    
    # Пункт 2: Собираем title_system из имен MTO файлов
    mto_title_systems_dict = defaultdict(list)  # title_system -> список файлов
    
    for mto_file in mto_files:
        file_name = mto_file.file_name
        title_system = ProjectFileName.scan_title_system(file_name)
        if title_system:
            title_system = title_system.strip()
        
        if title_system:
            mto_title_systems_dict[title_system].append(mto_file)
        else:
            if debug:
                print(f"  ⚠️  Не удалось определить title_system для файла: {file_name}")
    
    # Проверяем на дубликаты после завершения цикла
    duplicate_title_systems = {}
    mto_title_systems = {}  # title_system -> один файл (для дальнейшей работы)
    
    for title_system, files in mto_title_systems_dict.items():
        if len(files) > 1:
            # Дубликат найден
            duplicate_title_systems[title_system] = files
        else:
            # Уникальный title_system
            mto_title_systems[title_system] = files[0]
    
    # Выводим сообщения о дубликатах после завершения цикла
    if duplicate_title_systems:
        print(f"\n⚠️  ВНИМАНИЕ: Обнаружены дубликаты title_system в MTO файлах:")
        for title_system, files in duplicate_title_systems.items():
            print(f"  title_system '{title_system}' встречается в {len(files)} файлах:")
            for file in files:
                print(f"    - {file.file_name}")
        print("\n⚠️  Работа остановлена из-за дубликатов title_system в MTO файлах.")
        exit(0)
    
    print(f"Найдено уникальных title_system в MTO файлах: {len(mto_title_systems)}")
    if debug:
        print("Список title_system из MTO файлов:")
        for ts, file in sorted(mto_title_systems.items()):
            print(f"  - {ts} (файл: {file.file_name})")
    
    # Пункт 3: Сравнение списков и вывод таблицы
    all_title_systems = rfp_title_systems | set(mto_title_systems.keys())
    
    comparison_table = PrettyTable()
    comparison_table.field_names = ["title_system", "В RFP", "В MTO", "Результат сравнения"]
    comparison_table.border = True
    comparison_table.align = "l"
    
    matched_files = []  # Файлы для дальнейшей загрузки
    
    for title_system in sorted(all_title_systems):
        in_rfp = "✓" if title_system in rfp_title_systems else ""
        in_mto = "✓" if title_system in mto_title_systems else ""
        
        if title_system in rfp_title_systems and title_system in mto_title_systems:
            result = "Совпадают"
            matched_files.append(mto_title_systems[title_system])
        elif title_system in rfp_title_systems:
            result = "Нет в списке MTO файлов"
        else:
            result = "Нет в списке RFP данных"
        
        comparison_table.add_row([title_system, in_rfp, in_mto, result])
    
    print("\n" + "=" * 80)
    print("Таблица сравнения title_system:")
    print("=" * 80)
    print(comparison_table)
    
    # Статистика
    matched_count = len([ts for ts in all_title_systems 
                        if ts in rfp_title_systems and ts in mto_title_systems])
    only_rfp_count = len([ts for ts in all_title_systems 
                         if ts in rfp_title_systems and ts not in mto_title_systems])
    only_mto_count = len([ts for ts in all_title_systems 
                         if ts not in rfp_title_systems and ts in mto_title_systems])
    
    print(f"\nСтатистика:")
    print(f"  Совпадают: {matched_count}")
    print(f"  Только в RFP: {only_rfp_count}")
    print(f"  Только в MTO: {only_mto_count}")
    
    # Сохранение таблицы в Excel
    if result_dir:
        excel_path = _save_comparison_table_to_excel(
            all_title_systems,
            rfp_title_systems,
            mto_title_systems,
            result_dir
        )
        if excel_path:
            print(f"\nТаблица сравнения сохранена в: {excel_path}")
    
    # Пункт 4: Возвращаем только совпадающие файлы
    print(f"\nДля дальнейшей загрузки передано MTO файлов: {len(matched_files)}")    
    
    return matched_files


def _check_duplicate_mto_files(mto_files: List) -> None:
    """
    Проверяет наличие дубликатов MTO файлов по комбинации (title_system, doc_Number).
    Файлы с одинаковым title_system но разным doc_Number — валидный случай (разные спецификации).
    Дубликат: два или более файла с одинаковым title_system И одинаковым doc_Number.
    """
    groups = defaultdict(list)
    for mto_file in mto_files:
        ts = mto_file.doc_Short_Title or ProjectFileName.scan_title_system(mto_file.file_name)
        if ts:
            ts = ts.strip()
        doc_num = getattr(mto_file, 'doc_Number', None) or ''
        groups[(ts, doc_num)].append(mto_file)

    duplicates = {key: files for key, files in groups.items() if len(files) > 1}
    if not duplicates:
        return

    print("\n" + "=" * 80)
    print("ОШИБКА: Обнаружены дубликаты MTO файлов (одинаковые title_system + doc_Number)")
    print("=" * 80)

    table = PrettyTable()
    table.field_names = ["№", "title_system", "doc_Number", "Кол-во файлов", "Имена файлов", "Пути файлов"]
    table.align = "l"
    table.align["№"] = "r"
    table.align["Кол-во файлов"] = "r"

    for idx, ((ts, doc_num), files) in enumerate(sorted(duplicates.items()), 1):
        file_names = "\n".join(f.file_name for f in files)
        file_paths = "\n".join(getattr(f, "file_full_path", "") or f.file_name for f in files)
        table.add_row([idx, ts or "не определен", doc_num or "—", len(files), file_names, file_paths])

    print(table)
    print(f"\nВсего групп-дубликатов: {len(duplicates)}, "
          f"затронуто файлов: {sum(len(f) for f in duplicates.values())}")
    print("\nРабота остановлена. Устраните дубликаты и запустите повторно.")
    exit(0)


def _save_comparison_table_to_excel(
    all_title_systems: Set[str],
    rfp_title_systems: Set[str],
    mto_title_systems: dict,
    result_dir: str) -> str:
    """
    Сохраняет таблицу сравнения title_system в Excel файл
    
    Args:
        all_title_systems: Множество всех title_system
        rfp_title_systems: Множество title_system из RFP
        mto_title_systems: Словарь title_system -> файл из MTO
        result_dir: Путь к папке результатов
        
    Returns:
        Путь к сохраненному Excel файлу или None в случае ошибки
    """
    try:
        # Создаем папку результатов если её нет
        ensure_result_dir_exists(result_dir)
        
        # Создаем новую книгу Excel
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Сравнение title_system"
        
        # Заголовки
        headers = ["title_system", "В RFP", "В MTO", "Результат сравнения"]
        ws.append(headers)
        
        # Стили для заголовков
        header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
        header_font = Font(bold=True, color="FFFFFF")
        header_alignment = Alignment(horizontal="center", vertical="center")
        border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )
        
        for col_num, header in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col_num)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = header_alignment
            cell.border = border
        
        # Заполняем данные
        for title_system in sorted(all_title_systems):
            in_rfp = "✓" if title_system in rfp_title_systems else ""
            in_mto = "✓" if title_system in mto_title_systems else ""
            
            if title_system in rfp_title_systems and title_system in mto_title_systems:
                result = "Совпадают"
                fill_color = "C6EFCE"  # Светло-зеленый
            elif title_system in rfp_title_systems:
                result = "Нет в списке MTO файлов"
                fill_color = "FFC7CE"  # Светло-красный
            else:
                result = "Нет в списке RFP данных"
                fill_color = "FFEB9C"  # Светло-желтый
            
            row_data = [title_system, in_rfp, in_mto, result]
            ws.append(row_data)
            
            # Применяем стили к строке
            for col_num in range(1, len(headers) + 1):
                cell = ws.cell(row=ws.max_row, column=col_num)
                cell.border = border
                cell.alignment = Alignment(horizontal="left", vertical="center")
                
                # Применяем цвет фона к колонке "Результат сравнения"
                if col_num == 4:
                    cell.fill = PatternFill(start_color=fill_color, end_color=fill_color, fill_type="solid")
        
        # Настраиваем ширину колонок
        ws.column_dimensions['A'].width = 30
        ws.column_dimensions['B'].width = 12
        ws.column_dimensions['C'].width = 12
        ws.column_dimensions['D'].width = 35
        
        # Добавляем автофильтр для заголовков столбцов
        last_col_letter = get_column_letter(len(headers))
        ws.auto_filter.ref = f'A1:{last_col_letter}{ws.max_row}'
        
        # Замораживаем первую строку
        ws.freeze_panes = 'A2'
        
        # Сохраняем файл
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        file_name = f"Шаг2_MTO_сравнение_title_system_{timestamp}.xlsx"
        file_path = os.path.join(result_dir, file_name)
        
        wb.save(file_path)
        return file_path
        
    except Exception as e:
        print(f"Ошибка при сохранении таблицы сравнения в Excel: {e}")
        return None


def _save_load_results_table_to_excel(load_results: List[dict], result_dir: str) -> str:
    """
    Сохраняет таблицу результатов загрузки MTO файлов в Excel файл
    
    Args:
        load_results: Список словарей с результатами загрузки
        result_dir: Путь к папке результатов
        
    Returns:
        Путь к сохраненному Excel файлу или None в случае ошибки
    """
    try:
        # Создаем папку результатов если её нет
        ensure_result_dir_exists(result_dir)
        
        # Создаем новую книгу Excel
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Результаты загрузки MTO"
        
        # Проверяем, есть ли title_system в данных
        has_title_system = any('title_system' in result for result in load_results)
        
        # Заголовки
        if has_title_system:
            headers = ["№", "Имя файла", "title_system", "Статус", "Кол-во строк", "Ошибка"]
        else:
            headers = ["№", "Имя файла", "Статус", "Кол-во строк", "Ошибка"]
        
        ws.append(headers)
        
        # Стили для заголовков
        header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
        header_font = Font(bold=True, color="FFFFFF")
        header_alignment = Alignment(horizontal="center", vertical="center")
        border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )
        
        for col_num, header in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col_num)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = header_alignment
            cell.border = border
        
        # Заполняем данные
        for idx, result in enumerate(load_results, 1):
            if has_title_system:
                row_data = [
                    idx,
                    result['file_name'],
                    result.get('title_system', ''),
                    result['status'],
                    result['rows_count'],
                    result.get('error') or ""
                ]
            else:
                row_data = [
                    idx,
                    result['file_name'],
                    result['status'],
                    result['rows_count'],
                    result.get('error') or ""
                ]
            
            ws.append(row_data)
            
            # Применяем стили к строке
            for col_num in range(1, len(headers) + 1):
                cell = ws.cell(row=ws.max_row, column=col_num)
                cell.border = border
                cell.alignment = Alignment(horizontal="left", vertical="center")
                
                # Применяем цвет фона в зависимости от статуса
                status = result['status']
                if status == 'Успешно':
                    fill_color = "C6EFCE"  # Светло-зеленый
                elif status == 'Ошибка':
                    fill_color = "FFC7CE"  # Светло-красный
                elif status == 'Пустой файл':
                    fill_color = "FFEB9C"  # Светло-желтый
                else:
                    fill_color = None
                
                # Применяем цвет к колонке "Статус"
                if col_num == (4 if has_title_system else 3) and fill_color:
                    cell.fill = PatternFill(start_color=fill_color, end_color=fill_color, fill_type="solid")
                
                # Выравнивание для числовых колонок
                if col_num == 1 or (col_num == (5 if has_title_system else 4)):  # № или Кол-во строк
                    cell.alignment = Alignment(horizontal="right", vertical="center")
        
        # Настраиваем ширину колонок
        ws.column_dimensions['A'].width = 8  # №
        ws.column_dimensions['B'].width = 50  # Имя файла
        if has_title_system:
            ws.column_dimensions['C'].width = 20  # title_system
            ws.column_dimensions['D'].width = 15  # Статус
            ws.column_dimensions['E'].width = 12  # Кол-во строк
            ws.column_dimensions['F'].width = 50  # Ошибка
        else:
            ws.column_dimensions['C'].width = 15  # Статус
            ws.column_dimensions['D'].width = 12  # Кол-во строк
            ws.column_dimensions['E'].width = 50  # Ошибка
        
        # Добавляем автофильтр для заголовков столбцов
        last_col_letter = get_column_letter(len(headers))
        ws.auto_filter.ref = f'A1:{last_col_letter}{ws.max_row}'
        
        # Замораживаем первую строку
        ws.freeze_panes = 'A2'
        
        # Сохраняем файл
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        file_name = f"Шаг2_MTO_результаты_загрузки_{timestamp}.xlsx"
        file_path = os.path.join(result_dir, file_name)
        
        wb.save(file_path)
        return file_path
        
    except Exception as e:
        print(f"Ошибка при сохранении таблицы результатов загрузки в Excel: {e}")
        return None


def _print_load_results_table(load_results: List[dict]):
    """
    Выводит таблицу с результатами загрузки MTO файлов
    
    Args:
        load_results: Список словарей с результатами загрузки
    """
    table = PrettyTable()
    # Проверяем, есть ли title_system в данных
    has_title_system = any('title_system' in result for result in load_results)
    
    if has_title_system:
        table.field_names = ["№", "Имя файла", "title_system", "Статус", "Кол-во строк", "Ошибка"]
    else:
        table.field_names = ["№", "Имя файла", "Статус", "Кол-во строк", "Ошибка"]
    
    table.border = True
    table.align = "l"
    table.align["№"] = "r"
    table.align["Кол-во строк"] = "r"
    
    for idx, result in enumerate(load_results):
        if has_title_system:
            row = [
                idx + 1,
                result['file_name'],
                result.get('title_system', ''),
                result['status'],
                result['rows_count'],
                result.get('error') or ""
            ]
        else:
            row = [
                idx + 1,
                result['file_name'],
                result['status'],
                result['rows_count'],
                result.get('error') or ""
            ]
        table.add_row(row)
    
    print("\nРезультаты загрузки MTO файлов:")
    print(table)


def _get_mto_output_column_defs() -> List[ColumnDef]:
    return [d for d in MTO_OUTPUT_COLUMNS_CONFIG if d.output]


def _get_title_system_for_row(
    row: RowStd,
    file_meta_by_path: Dict[str, Dict[str, str]],
    per_path_cache: Dict[str, tuple[str, str]],
) -> tuple[str, str]:
    file_path = getattr(getattr(row, "t_com", None), "file_full_path", "") or ""
    if file_path in per_path_cache:
        return per_path_cache[file_path]

    meta = file_meta_by_path.get(file_path, {})
    title = str(meta.get("title") or "").strip()
    system = str(meta.get("system") or "").strip()
    per_path_cache[file_path] = (title, system)
    return title, system


def _save_mto_positions_database_to_excel(
    mto_data: Dict[str, List[RowStd]],
    result_dir: str,
    file_meta_by_path: Dict[str, Dict[str, str]],
    units_split_ban: Optional[List[str]] = None,
) -> str | None:
    """
    Сохраняет единую базу загруженных MTO position_row в Excel.
    Формат: TITLE, SYSTEM + колонки MTO (по конфигу MTO_OUTPUT_COLUMNS_CONFIG).
    """
    try:
        ensure_result_dir_exists(result_dir)
        t_total_start = time.perf_counter()

        column_defs = _get_mto_output_column_defs()
        headers = [d.header_label for d in column_defs]

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "MTO position_row"
        ws.append(headers)

        t_collapse_start = time.perf_counter()
        print("Шаг2 MTO export: подготовка данных и схлопывание...")
        usb = list(units_split_ban) if units_split_ban is not None else list(DEFAULT_UNITS_SPLIT_BAN)
        export_rows_before, export_rows_after_split, collapsed_export_rows = _build_collapsed_mto_export_rows(
            mto_data=mto_data,
            file_meta_by_path=file_meta_by_path,
            units_ban_list=usb,
        )
        t_collapse = time.perf_counter() - t_collapse_start
        append_timing_log(result_dir, f"step2_export_positions::prepare_and_collapse: {t_collapse:.3f}s")
        print(
            f"Шаг2 MTO export collapse: rows_before={export_rows_before}, "
            f"rows_after_split={export_rows_after_split}, rows_after_collapse={len(collapsed_export_rows)}"
        )

        t_write_start = time.perf_counter()
        total_rows = len(collapsed_export_rows)
        progress_step = 5000
        print(f"Шаг2 MTO export: запись в Excel {total_rows} строк...")
        for idx, export_row in enumerate(collapsed_export_rows, start=1):
            excel_row = [export_row.get(col_def.col_name, "") for col_def in column_defs]
            ws.append(excel_row)
            if idx % progress_step == 0:
                print(f"  Step2 export write progress: {idx}/{total_rows}")
        t_write = time.perf_counter() - t_write_start
        append_timing_log(result_dir, f"step2_export_positions::write_rows: {t_write:.3f}s")

        for col_idx, col_def in enumerate(column_defs, start=1):
            width = col_def.width or DEFAULT_COLUMN_WIDTH
            ws.column_dimensions[get_column_letter(col_idx)].width = width

        last_col_letter = get_column_letter(len(headers))
        ws.auto_filter.ref = f'A1:{last_col_letter}{ws.max_row}'
        ws.freeze_panes = 'A2'

        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        file_name = f"Шаг2_MTO_база_position_row_{timestamp}.xlsx"
        file_path = os.path.join(result_dir, file_name)
        print("Шаг2 MTO export: сохранение файла (wb.save)...")
        t_save_start = time.perf_counter()
        wb.save(file_path)
        t_save = time.perf_counter() - t_save_start
        append_timing_log(result_dir, f"step2_export_positions::wb_save: {t_save:.3f}s")
        append_timing_log(result_dir, f"step2_export_positions::total: {time.perf_counter() - t_total_start:.3f}s")
        return file_path
    except Exception as e:
        print(f"Ошибка при сохранении базы MTO position_row в Excel: {e}")
        return None


def _build_collapsed_mto_export_rows(
    mto_data: Dict[str, List[RowStd]],
    file_meta_by_path: Dict[str, Dict[str, str]],
    units_ban_list: List[str],
) -> tuple[int, int, List[Dict[str, object]]]:
    """
    Формирует отдельную (не мутирующую mto_data) экспортную проекцию и схлопывает её.
    Перед схлопыванием раскрывает MTO-строки по тегам по логике Step4:
      - только если tags_count > 1
      - для ед. изм. из units_ban_list раскрытие по тегам не делается
      - VALUES распределяется по 1 на тег, остаток/дефицит переносится как 0/предупреждение не пишем.

    Ключ схлопывания:
      TITLE + SYSTEM + все MTO-поля кроме VALUES.
    Агрегация:
      VALUES суммируется.
    """
    per_path_cache: Dict[str, tuple[str, str]] = {}
    collapsed: Dict[tuple, Dict[str, object]] = {}
    order: List[tuple] = []
    export_rows_before = 0
    export_rows_after_split = 0
    units_ban_set = {str(u).strip().lower() for u in (units_ban_list or [])}

    for rows in mto_data.values():
        for row in rows:
            if row.row_type != RowType.position_row:
                continue
            export_rows_before += 1
            if export_rows_before % 10000 == 0:
                print(f"  Step2 export prep progress: processed {export_rows_before} position_row")

            title, system = _get_title_system_for_row(row, file_meta_by_path, per_path_cache)
            base_row_dict: Dict[str, object] = {
                TITLE_FILTER_COL: title,
                SYSTEM_FILTER_COL: system,
            }
            for col in (NUMBERS, NAME, TYPE_MARK, CODE, VENDOR, UNITS, MASS, ANNOTATION):
                base_row_dict[col] = _normalize_export_value(row.el[col].value)

            tags_list = row.get_tags_list()
            values_int = _to_int_safe(row.el[VALUES].value)
            units_norm = str(row.get_value(UNITS) or "").strip().lower()
            should_split_tags = bool(tags_list and len(tags_list) > 1 and units_norm not in units_ban_set)

            split_rows: List[Dict[str, object]] = []
            if should_split_tags:
                remaining_values = values_int
                for tag in tags_list:
                    val = 1 if remaining_values > 0 else 0
                    if remaining_values > 0:
                        remaining_values -= 1
                    split_row = dict(base_row_dict)
                    split_row[TAGS] = _normalize_export_value([tag])
                    split_row[VALUES] = float(val)
                    split_rows.append(split_row)
            else:
                single_row = dict(base_row_dict)
                single_row[TAGS] = _normalize_export_value(tags_list)
                single_row[VALUES] = _to_float_safe(row.el[VALUES].value)
                split_rows.append(single_row)

            for row_dict in split_rows:
                export_rows_after_split += 1
                if export_rows_after_split % 20000 == 0:
                    print(f"  Step2 export prep progress: built {export_rows_after_split} rows after split")

                key = tuple(row_dict.get(col, "") for col in MTO_COLLAPSE_KEY_COLUMNS)
                if key not in collapsed:
                    new_row = dict(row_dict)
                    collapsed[key] = new_row
                    order.append(key)
                else:
                    collapsed[key][VALUES] = _to_float_safe(collapsed[key].get(VALUES)) + _to_float_safe(row_dict.get(VALUES))

    return export_rows_before, export_rows_after_split, [collapsed[k] for k in order]


def _normalize_export_value(value) -> str:
    if isinstance(value, list):
        return ", ".join(str(v) for v in value) if value else ""
    if value is None:
        return ""
    return str(value).strip()


def _to_float_safe(value) -> float:
    try:
        if value in (None, ""):
            return 0.0
        return float(value)
    except (ValueError, TypeError):
        try:
            return float(str(value).replace(",", "."))
        except (ValueError, TypeError):
            return 0.0


def _to_int_safe(value) -> int:
    try:
        if value in (None, ""):
            return 0
        return int(float(value))
    except (ValueError, TypeError):
        try:
            return int(float(str(value).replace(",", ".")))
        except (ValueError, TypeError):
            return 0


if __name__ in {"__main__"}:
    # Пример запуска этапа 2
    import sys
    
    if len(sys.argv) > 1:
        mto_path = sys.argv[1]
    else:
        mto_path = r"C:\Python\_ГОТОВЫЕ"
    
    debug = len(sys.argv) > 2 and sys.argv[2].lower() in ['1', 'true', 'debug']
    
    mto_data = step2_load_mto_data(mto_path, debug=debug)
    total_rows = sum(len(rows) for rows in mto_data.values())
    print(f"\nИтого загружено строк из MTO: {total_rows}")
    print(f"Количество групп по title_system: {len(mto_data)}")

