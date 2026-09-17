
"""
Этап 3: Загрузка данных из вендорской документации VO
Загрузка данных из Excel файла или файлов
"""

import os
from typing import Dict, List, Optional
from collections import defaultdict
import time

import openpyxl
from tqdm import tqdm
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from datetime import datetime
from prettytable import PrettyTable

from base.base_classes import RowStd, TableComments, RowType
from base.base_mto import get_std_from_excel_file
import base.t_comm_initial_classes as t_com_init_cls
from base.tables_columns import ANNOTATION, TAGS, IN_CABINET
from tags.tag_classes import TagClass
from tags.tag_parser import get_tag
from utils.file_name_converts import ProjectFileName
from utils.path import get_files_single
from RFQ.tags_rfp_compare.rfp_tags_utils import (
    ensure_result_dir_exists,
    get_result_dir_path,
    append_timing_log,
    save_input_fingerprint,
)
from utils.cache_utils import cache_manager


def step3_load_vo_data(
    vo_path: str,
    result_dir: str = None,
    debug: bool = False,
    export_to_excel: bool = True,
    save_input_fingerprints: bool = True,
) -> Dict[str, List[RowStd]]:
    """
    Этап 3: Загрузка данных из вендорской документации VO
    (Excel файл или файлы)
    
    Args:
        vo_path: Путь к директории с VO данными (Excel файлы)
        result_dir: Путь к папке результатов (не используется, для совместимости)
        debug: Флаг отладки
        export_to_excel: Флаг выгрузки отчетов в Excel
        
    Returns:
        Словарь: title_system -> список строк RowStd из VO файлов
    """
    print("=" * 80)
    print(f"ЭТАП 3: Загрузка VO")
    print(f"Путь: {vo_path}")

    vo_data: Dict[str, List[RowStd]] = defaultdict(list)
    empty_title_rows: Dict[str, List[RowStd]] = defaultdict(list)

    scan_start = time.perf_counter()
    all_files = get_files_single(vo_path, endswith=(".xlsx", ".XLSX"))
    append_timing_log(result_dir, f"step3_scan_vo_dir: {time.perf_counter() - scan_start:.3f}s")
    if not all_files:
        print(f"Ошибка: По пути <{vo_path}> файлы не найдены.")
        return []

    if save_input_fingerprints:
        try:
            save_input_fingerprint(
                result_dir=result_dir,
                source_name="VO",
                base_path=vo_path,
                file_paths=[f.file_full_path for f in all_files],
            )
        except Exception:
            pass

    file_paths = [f.file_full_path for f in all_files]
    cache_start = time.perf_counter()
    cached = cache_manager.load_vo_data_from_cache(
        vo_path, file_paths, verbose=False,
        result_dir=result_dir, log_timing=append_timing_log
    )
    append_timing_log(result_dir, f"step3_vo_cache_total: {time.perf_counter() - cache_start:.3f}s")
    if cached:
        vo_data, empty_title_rows = cached
        total_files = len(all_files)
        print(f"Файлов: {total_files} (из кэша)")
    else:
        total_files = len(all_files)
        print(f"Файлов: {total_files}")
        errors: List[str] = []
        load_files_start = time.perf_counter()
        for vo_file in tqdm(all_files, total=total_files, desc="VO загрузка", unit="файл", ncols=80):
            rows = None
            try:
                t_com = TableComments(
                    file_full_path=vo_file.file_full_path,
                    dir_path="-1",
                    tabel_class=t_com_init_cls.VO_MTO
                )
                file_path = vo_file.file_full_path
                if cache_manager.is_cache_valid(file_path):
                    rows = cache_manager.load_from_cache(file_path, verbose=False)
                if rows is None:
                    rows = get_std_from_excel_file(t_com)
                    cache_manager.save_to_cache(file_path, rows, verbose=False)
            except Exception as e:
                errors.append(f"{vo_file.file_name}: {e}")
            if rows:
                cabinet_tags = get_tag(vo_file.file_name) if vo_file.file_name else []
                in_cabinet_value = cabinet_tags[0] if cabinet_tags else ""
                if in_cabinet_value:
                    for row in rows:
                        if row.row_type == RowType.position_row:
                            row.el[IN_CABINET].value = in_cabinet_value
                _add_rows_to_vo_data(vo_data, rows, empty_title_rows)
        for err in errors:
            print(f"Ошибка: {err}")
        append_timing_log(result_dir, f"step3_vo_load_files: {time.perf_counter() - load_files_start:.3f}s")
        save_cache_start = time.perf_counter()
        cache_manager.save_vo_data_to_cache(vo_path, vo_data, empty_title_rows, file_paths, verbose=False)
        append_timing_log(result_dir, f"step3_vo_save_cache: {time.perf_counter() - save_cache_start:.3f}s")
    
    if empty_title_rows and export_to_excel:
        empty_rows_count = sum(len(rows) for rows in empty_title_rows.values())
        if not result_dir:
            result_dir = get_result_dir_path(vo_path)
        t0 = time.perf_counter()
        excel_path = _save_empty_title_systems_to_excel(empty_title_rows, result_dir)
        append_timing_log(result_dir, f"step3_vo_save_empty_title_excel: {time.perf_counter() - t0:.3f}s")
        if excel_path:
            print(f"Пустые title_system: {empty_rows_count} строк → {os.path.basename(excel_path)}")
    if vo_data and export_to_excel:
        if not result_dir:
            result_dir = get_result_dir_path(vo_path)
        t0 = time.perf_counter()
        excel_path = _save_vo_rows_to_excel(vo_data, result_dir)
        append_timing_log(result_dir, f"step3_vo_save_all_rows_excel: {time.perf_counter() - t0:.3f}s")
        if excel_path:
            print(f"Строки VO → {os.path.basename(excel_path)}")
    if debug:
        _print_vo_table_debug(vo_data)

    return vo_data


def _add_rows_to_vo_data(
    vo_data: Dict[str, List[RowStd]],
    rows: List[RowStd],
    empty_title_rows: Optional[Dict[str, List[RowStd]]] = None):
    for row in rows:
        if row.row_type != RowType.position_row:
            continue
        annotation_value = row.get_value(ANNOTATION)
        if annotation_value:
            annotation_str = str(annotation_value).strip()
        else:
            annotation_str = ""
        title_system = ProjectFileName.scan_title_system(annotation_str) if annotation_str else None
        if isinstance(title_system, str):
            title_system = title_system.strip() or None
        if not title_system and empty_title_rows is not None:
            file_key = row.t_com.file_full_path or row.t_com.file_name or "Неизвестный файл"
            empty_title_rows[file_key].append(row)
        vo_data[title_system].append(row)


def _print_vo_table_debug(vo_data: Dict[str, List[RowStd]], max_rows: int = 20):
    if not vo_data:
        print("VO таблица пуста, нечего выводить.")
        return

    summary_table = PrettyTable()
    summary_table.field_names = ["title_system", "Кол-во строк"]
    summary_table.border = True
    summary_table.align = "l"
    summary_table.align["Кол-во строк"] = "r"

    for title_system in sorted(vo_data.keys(), key=lambda x: (x is None, x or "")):
        rows_count = len(vo_data[title_system])
        title_display = title_system if title_system else "не определен"
        summary_table.add_row([title_display, rows_count])

    print("\nСводка загруженной VO (prettytable):")
    print(summary_table)

    sample_rows = []
    total_rows = 0
    for title_system in sorted(vo_data.keys(), key=lambda x: (x is None, x or "")):
        for row in vo_data[title_system]:
            if row.row_type in (RowType.empty_row, RowType.other_row):
                continue
            total_rows += 1
            if len(sample_rows) < max_rows:
                sample_rows.append(row)

    if not sample_rows:
        print("VO таблица не содержит строк для вывода (после фильтрации).")
        return

    column_dict = sample_rows[0].t_com.column_dict or t_com_init_cls.VO_MTO.column_dict
    columns = [column_dict[k] for k in sorted(column_dict.keys())]

    table = PrettyTable()
    table.field_names = ["№"] + columns
    table.border = True
    table.align = "l"
    table.align["№"] = "r"

    for idx, row in enumerate(sample_rows, 1):
        row_values = [row.el[col].value for col in columns]
        table.add_row([idx] + row_values)

    print("\nПример строк VO (prettytable):")
    print(f"Показаны строки: {len(sample_rows)} из {total_rows}")
    print(table)


def _save_empty_title_systems_to_excel(
    empty_title_rows: Dict[str, List[RowStd]],
    result_dir: str) -> Optional[str]:
    """
    Сохраняет информацию о строках с пустым title_system в Excel файл
    
    Args:
        empty_title_rows: Словарь file_path -> список строк RowStd
        result_dir: Путь к папке результатов
        
    Returns:
        Путь к сохраненному Excel файлу или None в случае ошибки
    """
    try:
        ensure_result_dir_exists(result_dir)
        
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Пустой title_system"
        
        headers = ["№", "Имя файла", "Полный путь", "Кол-во строк", "Примеры ANNOTATION"]
        ws.append(headers)
        
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
        
        for idx, file_path in enumerate(sorted(empty_title_rows.keys()), 1):
            rows = empty_title_rows[file_path]
            file_name = os.path.basename(file_path) if file_path else "Неизвестный файл"
            
            annotation_samples = []
            for row in rows:
                annotation_value = row.get_value(ANNOTATION)
                annotation_display = str(annotation_value).strip() if annotation_value else ""
                if not annotation_display:
                    annotation_display = "<пусто>"
                if annotation_display not in annotation_samples:
                    annotation_samples.append(annotation_display)
                if len(annotation_samples) >= 5:
                    break
            
            row_data = [
                idx,
                file_name,
                file_path,
                len(rows),
                "; ".join(annotation_samples)
            ]
            ws.append(row_data)
            
            for col_num in range(1, len(headers) + 1):
                cell = ws.cell(row=ws.max_row, column=col_num)
                cell.border = border
                cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        
        ws.column_dimensions['A'].width = 6
        ws.column_dimensions['B'].width = 40
        ws.column_dimensions['C'].width = 80
        ws.column_dimensions['D'].width = 14
        ws.column_dimensions['E'].width = 60
        
        last_col_letter = get_column_letter(len(headers))
        ws.auto_filter.ref = f'A1:{last_col_letter}{ws.max_row}'
        ws.freeze_panes = 'A2'
        
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        file_name = f"Шаг3_VO_пустой_title_system_{timestamp}.xlsx"
        file_path = os.path.join(result_dir, file_name)
        wb.save(file_path)
        return file_path
        
    except Exception as e:
        print(f"Ошибка при сохранении списка пустых title_system в Excel: {e}")
        return None


def _save_vo_rows_to_excel(
    vo_data: Dict[str, List[RowStd]],
    result_dir: str) -> Optional[str]:
    """
    Сохраняет все строки VO в Excel файл с отдельной колонкой title_system
    
    Args:
        vo_data: Словарь title_system -> список строк RowStd
        result_dir: Путь к папке результатов
        
    Returns:
        Путь к сохраненному Excel файлу или None в случае ошибки
    """
    try:
        ensure_result_dir_exists(result_dir)
        
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "VO"
        
        all_rows = []
        for title_system in sorted(vo_data.keys(), key=lambda x: (x is None, x or "")):
            for row in vo_data[title_system]:
                all_rows.append((title_system, row))
        
        if not all_rows:
            return None
        
        column_dict = all_rows[0][1].t_com.column_dict or t_com_init_cls.VO_MTO.column_dict
        columns = [column_dict[k] for k in sorted(column_dict.keys())]
        
        headers = [
            "title_system",
            "file_name",
            "status_title_file_name_vs_tags",
            "status_title_tags_vs_annotation"
        ] + columns
        ws.append(headers)
        
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
        
        def _excel_safe_value(value):
            if isinstance(value, list):
                return "; ".join(str(item) for item in value)
            return value

        def _get_tags_titles(tags_value):
            titles = set()
            if not tags_value:
                return titles
            tags_list = tags_value if isinstance(tags_value, list) else [tags_value]
            for tag in tags_list:
                if not tag:
                    continue
                try:
                    title = TagClass(str(tag)).wbs
                except Exception:
                    continue
                if title:
                    titles.add(str(title).strip())
            return titles

        def _compare_title_and_tags(title_titles, tags_titles):
            title_titles = {str(t).strip() for t in title_titles if str(t).strip()} if title_titles else set()
            if not title_titles and not tags_titles:
                return "Нет title и тегов"
            if not title_titles:
                return "Нет title"
            if not tags_titles:
                return "Нет тегов"
            if title_titles.intersection(tags_titles):
                return "Совпадает"
            return "Не совпадает"

        for title_system, row in all_rows:
            title_display = title_system if title_system else ""
            file_name = row.t_com.file_name or os.path.basename(row.t_com.file_full_path) or ""
            file_name_tags = get_tag(file_name) if file_name else []
            file_name_titles = _get_tags_titles(file_name_tags)
            annotation_value = row.get_value(ANNOTATION)
            annotation_str = str(annotation_value).strip() if annotation_value else ""
            annotation_title = ProjectFileName.scan_title(annotation_str) if annotation_str else None
            annotation_titles = {str(annotation_title).strip()} if annotation_title else set()
            tags_titles = _get_tags_titles(row.get_value(TAGS))
            status_file_vs_tags = _compare_title_and_tags(file_name_titles, tags_titles)
            status_tags_vs_annotation = _compare_title_and_tags(annotation_titles, tags_titles)
            row_values = [_excel_safe_value(row.el[col].value) for col in columns]
            ws.append([title_display, file_name, status_file_vs_tags, status_tags_vs_annotation] + row_values)
            
            for col_num in range(1, len(headers) + 1):
                cell = ws.cell(row=ws.max_row, column=col_num)
                cell.border = border
                cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        
        ws.column_dimensions['A'].width = 30
        ws.column_dimensions['B'].width = 40
        ws.column_dimensions['C'].width = 30
        ws.column_dimensions['D'].width = 30
        for idx in range(5, len(headers) + 1):
            col_letter = get_column_letter(idx)
            ws.column_dimensions[col_letter].width = 30
        
        last_col_letter = get_column_letter(len(headers))
        ws.auto_filter.ref = f'A1:{last_col_letter}{ws.max_row}'
        ws.freeze_panes = 'A2'
        
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        file_name = f"Шаг3_VO_все_строки_{timestamp}.xlsx"
        file_path = os.path.join(result_dir, file_name)
        wb.save(file_path)
        return file_path
        
    except Exception as e:
        print(f"Ошибка при сохранении строк VO в Excel: {e}")
        return None


if __name__ in {"__main__"}:
    # Пример запуска этапа 3
    import sys
    
    if len(sys.argv) > 1:
        vo_path = sys.argv[1]
    else:
        vo_path = r"\\bcc\eng\PrDoc\377_НИПИГАЗ\АГХК\КСБ\РКД\Материалы шкафов из РКД для робота"
    
    debug = True
    vo_data = step3_load_vo_data(vo_path, debug=debug)
    total_rows = sum(len(rows) for rows in vo_data.values())
    print(f"\nИтого загружено строк из VO: {total_rows}")

