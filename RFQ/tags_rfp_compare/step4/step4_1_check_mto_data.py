"""
Этап 4_1: Проверка данных из MTO
Проверка соответствия количества тегов и VALUES, разбиение строк с несколькими тегами
"""

import os
import time
from typing import List, Dict, Optional
from datetime import datetime
from collections import defaultdict
from prettytable import PrettyTable

import openpyxl
from tqdm import tqdm
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

from base.base_classes import RowStd, RowType, CheckElement
from base.tables_columns import TAGS, VALUES, CODE, MTO_NAME, DS_TITLE, NAME, UNITS
from RFQ.tags_rfp_compare.rfp_tags_utils import (
    DEFAULT_UNITS_SPLIT_BAN,
    ensure_result_dir_exists,
    append_timing_log,
    log_mto_tag_distribution,
)
from RFQ.tags_rfp_compare.step1_load_rfp import try_exact_int_quantity
from RFQ.tags_rfp_compare.tag_count_mismatch_excel import save_tag_count_mismatch_excel


def _should_debug(
    debug: bool,
    title_system: str,
    code: str,
    tags: List[str],
    debug_tag: List[str],
    debug_code: List[str],
    debug_title_system: List[str]) -> bool:
    """
    Проверяет, нужно ли выводить отладочную информацию.
    Все указанные (не None/пустые) фильтры должны совпадать (AND между фильтрами).
    Внутри каждого фильтра-списка работает логика OR (совпадение с любым элементом).
    """
    if not debug:
        return False
    if not (debug_tag or debug_code or debug_title_system):
        return True
    # debug_title_system: проверяем, что title_system входит в список
    if debug_title_system and title_system not in debug_title_system:
        return False
    # debug_code: проверяем, что code входит в список
    if debug_code and code not in debug_code:
        return False
    # debug_tag: проверяем, что хотя бы один тег из tags входит в debug_tag
    if debug_tag and (not tags or not any(t in debug_tag for t in tags)):
        return False
    return True


def _debug_print(debug_log: List[str], *args, should_log: bool = True, **kwargs) -> None:
    """
    Выводит отладочную информацию в консоль и опционально в debug_log.
    
    Args:
        debug_log: Список строк для накопления debug-лога
        *args: Аргументы для print
        should_log: Если True, записывает в debug_log. Если False, только в консоль.
        **kwargs: Kwargs для print (sep, end)
    """
    print(*args, **kwargs)
    if debug_log is None or not should_log:
        return
    sep = kwargs.get("sep", " ")
    end = kwargs.get("end", "\n")
    debug_log.append(sep.join(str(arg) for arg in args) + end)


def check_mto_data(mto_data: Dict[str, List[RowStd]], 
                result_dir: str, 
                debug: bool = False,
                debug_tag: List[str] = None,
                debug_code: List[str] = None,
                debug_title_system: List[str] = None,
                debug_log: List[str] = None,
                defer_reports: bool = False,
                units_ban: Optional[List[str]] = None,
                ) -> tuple[Dict[str, List[RowStd]], dict]:
    """
    Проверка данных из MTO:
    1. Проверка соответствия количества тегов и VALUES
    2. Разбиение строк с несколькими тегами на отдельные строки
    
    Args:
        mto_data: Словарь title_system -> список строк RowStd из MTO
        result_dir: Путь к папке результатов
        debug: Флаг отладки
        debug_tag: Тег для точечной отладки (AND с другими фильтрами)
        debug_code: Код для точечной отладки (AND с другими фильтрами)
        debug_title_system: title_system для точечной отладки (AND с другими фильтрами)
        defer_reports: Если True — не печатает таблицы и не пишет Excel,
            а возвращает диагностику в словаре (для агрегации в main-процессе).
        
    Returns:
        (mto_data, diagnostics) — обновлённый словарь + диагностика
        diagnostics: {"count_mismatches": [...], "duplicate_tags": [...]}
    """
    _empty_diag: dict = {"count_mismatches": [], "duplicate_tags": []}

    if not defer_reports:
        print("=" * 80)
        print("ЭТАП 4_1: Проверка данных из MTO")
        print("=" * 80)
    
    if not mto_data:
        if not defer_reports:
            print("MTO данные пусты, проверка не требуется")
        return {}, _empty_diag
    
    if isinstance(mto_data, list):
        if not defer_reports:
            print("⚠️  ВНИМАНИЕ: mto_data передан как список, преобразуем в словарь")
        mto_data = {"не определен": mto_data}

    if not defer_reports:
        log_mto_tag_distribution(result_dir, mto_data, label="MTO tag distribution (before split)")

    # Список для хранения несовпадений количества тегов и VALUES
    count_mismatches = []
    ub_list = list(DEFAULT_UNITS_SPLIT_BAN) if units_ban is None else list(units_ban)
    units_ban_set = {str(u).strip().lower() for u in ub_list}
    tag_count_check_time = 0.0
    split_no_tags_time = 0.0
    split_tags_time = 0.0
    loop_total_start = time.perf_counter()
    check_element_cls = CheckElement
    row_std_cls = RowStd
    
    items_list = list(mto_data.items())
    total_items = len(items_list)
    warnings: List[str] = []
    
    for title_system, mto_rows in tqdm(items_list, total=total_items, desc="ЭТАП 4_1 MTO", unit="система", ncols=80, disable=defer_reports):
        show_debug_info = _should_debug(
            debug,
            title_system,
            "",
            [],
            debug_tag,
            debug_code,
            debug_title_system
        )
        
        if show_debug_info:
            print(f"\nОбработка title_system: {title_system or 'не определен'}")
            print(f"  Всего строк: {len(mto_rows)}")
        
        # Сохраняем исходное количество строк для отчета
        original_count = len(mto_rows)
        
        new_rows: List[RowStd] = []
        new_rows_append = new_rows.append
        for row in mto_rows:
            # Проверяем только строки с типом position_row
            if row.row_type != RowType.position_row:
                new_rows_append(row)
                continue
            
            tags_list = row.get_tags_list()
            values_value = row.get_value(VALUES)
            values_int = try_exact_int_quantity(values_value)
            if values_int is None:
                new_rows_append(row)
                continue

            units_value = row.get_value(UNITS) or ""
            units_norm = str(units_value).strip().lower()
            is_unit_banned = units_norm in units_ban_set

            # Если тегов нет, разбиваем по VALUES (если возможно)
            if not tags_list:
                if values_int <= 1 or is_unit_banned:
                    new_rows_append(row)
                    continue
                if values_int > 1:
                    t_start = time.perf_counter()
                    new_rows.extend(row_std_cls.batch_copy_light(row, values_int, values_override=1))
                    split_no_tags_time += time.perf_counter() - t_start
                continue
            
            tags_count = len(tags_list)
            
            # Проверка 1: Соответствие количества тегов и VALUES
            t_start = time.perf_counter()
            if tags_count != values_int:
                code_value = row.get_value(CODE) or ""
                mto_name_value = row.get_value(MTO_NAME) or ""
                ds_title_value = row.get_value(DS_TITLE) or ""
                count_mismatches.append({
                    'title_system': title_system or "не определен",
                    'code': str(code_value),
                    'mto_name': str(mto_name_value),
                    'ds_title': str(ds_title_value),
                    'tags_count': tags_count,
                    'values': values_int,
                    'tags': ', '.join(tags_list) if tags_list else ""
                })
            tag_count_check_time += time.perf_counter() - t_start
            
            # Проверка 2: Разбиение строк с несколькими тегами
            if tags_count > 1 and not is_unit_banned:
                t_start = time.perf_counter()
                remaining_values = values_int
                code_value = row.get_value(CODE) or ""
                
                for tag in tags_list:
                    val = 1 if remaining_values > 0 else 0
                    if remaining_values > 0:
                        remaining_values -= 1
                    else:
                        warnings.append(f"Недостаточно VALUES для тега {tag} в строке {code_value} (title_system: {title_system or 'не определен'})")
                    new_row = row_std_cls.get_row_copy_light(row, tags_override=[tag], values_override=val)
                    new_rows_append(new_row)
                    
                    if _should_debug(
                        debug,
                        title_system,
                        code_value,
                        [tag],
                        debug_tag,
                        debug_code,
                        debug_title_system
                    ):
                        actual_value = new_row.el[VALUES].value
                        _debug_print(
                            debug_log,
                            f"    [DEBUG][check_mto_data][split_row] Разбита строка MTO: CODE={code_value}, TAG={tag}, VALUES={actual_value}"
                        )
                
                if remaining_values > 0:
                    warnings.append(f"Осталось {remaining_values} нераспределенных VALUES в строке {code_value} (title_system: {title_system or 'не определен'})")
                
                split_tags_time += time.perf_counter() - t_start
            else:
                new_rows_append(row)
        
        # Обновляем mto_data с обновленными строками (изменения уже внесены в исходный список)
        mto_data[title_system] = new_rows
        
        if show_debug_info:
            print(f"  После обработки строк: {len(new_rows)} (было: {original_count})")
            
            # Выводим уже раскрытые строки для дебага
            print(f"\n  Раскрытые строки (только position_row с тегами):")
            for row in new_rows:
                if row.row_type == RowType.position_row:
                    tags_list = row.get_tags_list()
                    if tags_list:
                        code_val = row.get_value(CODE) or ""
                        values_val = row.get_value(VALUES) or 0
                        tags_str = ', '.join(tags_list)
                        print(f"    CODE={code_val}, TAGS=[{tags_str}], VALUES={values_val}")
    
    loop_total = time.perf_counter() - loop_total_start
    append_timing_log(result_dir, f"check_mto_data::loop_total: {loop_total:.3f}s")
    append_timing_log(result_dir, f"check_mto_data::tag_count_check: {tag_count_check_time:.3f}s")
    append_timing_log(result_dir, f"check_mto_data::split_no_tags: {split_no_tags_time:.3f}s")
    append_timing_log(result_dir, f"check_mto_data::split_tags: {split_tags_time:.3f}s")
    
    if not defer_reports:
        for w in warnings:
            print(f"Ошибка: {w}")

    duplicate_tags = _check_duplicate_tags(mto_data)

    if defer_reports:
        return mto_data, {"count_mismatches": count_mismatches, "duplicate_tags": duplicate_tags}

    if count_mismatches:
        t_start = time.perf_counter()
        _print_count_mismatches_table(count_mismatches)
        append_timing_log(result_dir, f"check_mto_data::print_count_mismatches: {time.perf_counter() - t_start:.3f}s")
        if result_dir:
            t_start = time.perf_counter()
            excel_path = _save_count_mismatches_to_excel(count_mismatches, result_dir)
            append_timing_log(result_dir, f"check_mto_data::save_count_mismatches_excel: {time.perf_counter() - t_start:.3f}s")
            if excel_path:
                print(f"Несовпадения тегов/VALUES ({len(count_mismatches)}) -> {os.path.basename(excel_path)}")

    if duplicate_tags:
        _print_duplicate_tags_table(duplicate_tags)
        if result_dir:
            t_start = time.perf_counter()
            excel_path = _save_duplicate_tags_to_excel(duplicate_tags, result_dir)
            append_timing_log(result_dir, f"check_mto_data::save_duplicate_tags_excel: {time.perf_counter() - t_start:.3f}s")
            if excel_path:
                print(f"Дублирующиеся теги MTO ({len(duplicate_tags)}) -> {os.path.basename(excel_path)}")  

    return mto_data, {"count_mismatches": count_mismatches, "duplicate_tags": duplicate_tags}


def check_vo_data(vo_data: Dict[str, List[RowStd]],
                  result_dir: str,
                  debug: bool = False,
                  debug_tag: List[str] = None,
                  debug_code: List[str] = None,
                  debug_title_system: List[str] = None,
                  debug_log: List[str] = None,
                  defer_reports: bool = False,
                  ) -> tuple[Dict[str, List[RowStd]], dict]:
    """
    Проверка данных из VO:
    1. Разбиение строк с несколькими тегами на отдельные строки
    2. Проверка на дублирование тегов
    
    Args:
        vo_data: Словарь title_system -> список строк RowStd из VO
        result_dir: Путь к папке результатов
        debug: Флаг отладки
        debug_tag: Тег для точечной отладки (AND с другими фильтрами)
        debug_code: Код для точечной отладки (AND с другими фильтрами)
        debug_title_system: title_system для точечной отладки (AND с другими фильтрами)
        defer_reports: Если True — не печатает таблицы и не пишет Excel,
            а возвращает диагностику в словаре (для агрегации в main-процессе).
        
    Returns:
        (vo_data, diagnostics) — обновлённый словарь + диагностика
        diagnostics: {"duplicate_tags_vo": [...]}
    """
    _empty_diag: dict = {"duplicate_tags_vo": []}

    if not defer_reports:
        print("=" * 80)
        print("ЭТАП 4_1: Проверка данных из VO")
        print("=" * 80)
    
    if not vo_data:
        if not defer_reports:
            print("VO данные пусты, проверка не требуется")
        return {}, _empty_diag
    
    if isinstance(vo_data, list):
        if not defer_reports:
            print("⚠️  ВНИМАНИЕ: vo_data передан как список, преобразуем в словарь")
        vo_data = {"не определен": vo_data}
    
    items_list = list(vo_data.items())
    warnings: List[str] = []
    
    for title_system, vo_rows in tqdm(items_list, total=len(items_list), desc="ЭТАП 4_1 VO", unit="система", ncols=80, disable=defer_reports):
        new_rows: List[RowStd] = []
        for row in vo_rows:
            if row.row_type != RowType.position_row:
                new_rows.append(row)
                continue
            
            tags_list = row.get_tags_list()
            values_value = row.get_value(VALUES)
            try:
                values_int = int(float(values_value)) if values_value is not None else 0
            except (ValueError, TypeError):
                values_int = 0

            if not tags_list:
                if values_int > 1:
                    code_value = row.get_value(CODE) or ""
                    new_rows.extend(RowStd.batch_copy_light(row, values_int, values_override=1))

                    if _should_debug(
                        debug,
                        title_system,
                        code_value,
                        [],
                        debug_tag,
                        debug_code,
                        debug_title_system
                    ):
                        _debug_print(
                            debug_log,
                            f"    [DEBUG][check_vo_data][split_row_no_tag] Разбита строка VO без тегов: CODE={code_value}, VALUES={values_int}"
                        )
                else:
                    new_rows.append(row)
                continue
            
            if len(tags_list) > 1:
                code_value = row.get_value(CODE) or ""
                remaining_values = values_int
                
                for tag in tags_list:
                    val = 1 if remaining_values > 0 else 0
                    if remaining_values > 0:
                        remaining_values -= 1
                    new_row = RowStd.get_row_copy_light(row, tags_override=[tag], values_override=val)
                    new_rows.append(new_row)
                    
                    if _should_debug(
                        debug,
                        title_system,
                        code_value,
                        [tag],
                        debug_tag,
                        debug_code,
                        debug_title_system
                    ):
                        _debug_print(
                            debug_log,
                            f"    [DEBUG][check_vo_data][split_row] Разбита строка VO: CODE={code_value}, TAG={tag}, VALUES={new_row.el[VALUES].value}"
                        )
                if remaining_values > 0:
                    warnings.append(
                        f"Осталось {remaining_values} нераспределенных VALUES в строке {code_value} "
                        f"(title_system: {title_system or 'не определен'})"
                    )
            else:
                new_rows.append(row)
        
        vo_data[title_system] = new_rows
    
    if not defer_reports:
        for w in warnings:
            print(f"Ошибка: {w}")

    duplicate_tags = _check_duplicate_tags_vo(vo_data)

    if defer_reports:
        return vo_data, {"duplicate_tags_vo": duplicate_tags}

    if duplicate_tags:
        _print_duplicate_tags_table_vo(duplicate_tags)
        if result_dir:
            excel_path = _save_duplicate_tags_to_excel_vo(duplicate_tags, result_dir)
            if excel_path:
                print(f"Дублирующиеся теги VO ({len(duplicate_tags)}) -> {os.path.basename(excel_path)}")

    return vo_data, {"duplicate_tags_vo": duplicate_tags}


def save_mto_diagnostics_reports(
    count_mismatches: List[dict],
    duplicate_tags: List[dict],
    result_dir: str,
) -> None:
    """Печатает таблицы и сохраняет Excel-отчёты по MTO-диагностике (агрегированные из worker'ов)."""
    if count_mismatches:
        _print_count_mismatches_table(count_mismatches)
        if result_dir:
            excel_path = _save_count_mismatches_to_excel(count_mismatches, result_dir)
            if excel_path:
                print(f"Несовпадения тегов/VALUES ({len(count_mismatches)}) -> {os.path.basename(excel_path)}")
    if duplicate_tags:
        _print_duplicate_tags_table(duplicate_tags)
        if result_dir:
            excel_path = _save_duplicate_tags_to_excel(duplicate_tags, result_dir)
            if excel_path:
                print(f"Дублирующиеся теги MTO ({len(duplicate_tags)}) -> {os.path.basename(excel_path)}")


def save_vo_diagnostics_reports(
    duplicate_tags_vo: List[dict],
    result_dir: str,
) -> None:
    """Печатает таблицу и сохраняет Excel-отчёт по VO-дубликатам (агрегированные из worker'ов)."""
    if duplicate_tags_vo:
        _print_duplicate_tags_table_vo(duplicate_tags_vo)
        if result_dir:
            excel_path = _save_duplicate_tags_to_excel_vo(duplicate_tags_vo, result_dir)
            if excel_path:
                print(f"Дублирующиеся теги VO ({len(duplicate_tags_vo)}) -> {os.path.basename(excel_path)}")


def _check_duplicate_tags_vo(vo_data: Dict[str, List[RowStd]]) -> List[dict]:
    tag_occurrences = defaultdict(list)
    
    for title_system, vo_rows in vo_data.items():
        for row_idx, row in enumerate(vo_rows):
            if row.row_type != RowType.position_row:
                continue
            
            tags_list = row.get_tags_list()
            if not tags_list:
                continue
            
            for tag in tags_list:
                code_value = row.get_value(CODE) or ""
                name_value = row.get_value(NAME) or ""
                
                tag_occurrences[tag].append({
                    'title_system': title_system or "не определен",
                    'code': str(code_value),
                    'name': str(name_value),
                    'row_index': row_idx + 1
                })
    
    duplicate_tags = []
    for tag, occurrences in tag_occurrences.items():
        if len(occurrences) > 1:
            duplicate_tags.append({
                'tag': tag,
                'count': len(occurrences),
                'occurrences': occurrences
            })
    
    duplicate_tags.sort(key=lambda x: x['count'], reverse=True)
    return duplicate_tags


def _print_duplicate_tags_table_vo(duplicate_tags: List[dict]):
    table = PrettyTable()
    table.field_names = ["№", "Тег", "Кол-во вхождений", "title_system", "CODE"]
    table.border = True
    table.align = "l"
    table.align["№"] = "r"
    table.align["Кол-во вхождений"] = "r"
    
    for idx, dup_info in enumerate(duplicate_tags, 1):
        tag = dup_info['tag']
        count = dup_info['count']
        
        tag_display = tag if len(tag) <= 30 else tag[:27] + "..."
        first_occurrence = dup_info['occurrences'][0]
        title_sys = str(first_occurrence['title_system'])
        if len(title_sys) > 15:
            title_sys = title_sys[:12] + "..."
        
        code = str(first_occurrence['code'])
        if len(code) > 20:
            code = code[:17] + "..."
        
        row = [
            idx,
            tag_display,
            count,
            title_sys,
            code
        ]
        table.add_row(row)
    
    print(table)


def _save_duplicate_tags_to_excel_vo(duplicate_tags: List[dict], result_dir: str) -> str:
    try:
        ensure_result_dir_exists(result_dir)
        
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Дублирующиеся теги VO"
        
        headers = ["№", "Тег", "Кол-во вхождений", "title_system", "CODE", "NAME", "№ строки"]
        ws.append(headers)
        
        header_fill = PatternFill(start_color="FF6B6B", end_color="FF6B6B", fill_type="solid")
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
        
        row_counter = 0
        for dup_info in duplicate_tags:
            tag = dup_info['tag']
            count = dup_info['count']
            for occurrence in dup_info['occurrences']:
                row_counter += 1
                row_data = [
                    row_counter,
                    tag,
                    count,
                    occurrence['title_system'],
                    occurrence['code'],
                    occurrence['name'],
                    occurrence['row_index']
                ]
                ws.append(row_data)
                
                for col_num in range(1, len(headers) + 1):
                    cell = ws.cell(row=ws.max_row, column=col_num)
                    cell.border = border
                    cell.alignment = Alignment(horizontal="left", vertical="center")
                    
                    if col_num in (1, 3, 7):
                        cell.alignment = Alignment(horizontal="right", vertical="center")
                    if col_num == 3 and count > 1:
                        fill_color = "FFC7CE"
                        cell.fill = PatternFill(start_color=fill_color, end_color=fill_color, fill_type="solid")
        
        ws.column_dimensions['A'].width = 8
        ws.column_dimensions['B'].width = 30
        ws.column_dimensions['C'].width = 18
        ws.column_dimensions['D'].width = 20
        ws.column_dimensions['E'].width = 20
        ws.column_dimensions['F'].width = 50
        ws.column_dimensions['G'].width = 12
        
        last_col_letter = get_column_letter(len(headers))
        ws.auto_filter.ref = f'A1:{last_col_letter}{ws.max_row}'
        ws.freeze_panes = 'A2'
        
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        file_name = f"Отчет по дублям тегов - VO_{timestamp}.xlsx"
        file_path = os.path.join(result_dir, file_name)
        
        wb.save(file_path)
        return file_path
        
    except Exception as e:
        print(f"Ошибка при сохранении таблицы дублирующихся тегов VO в Excel: {e}")
        return None


def _print_count_mismatches_table(count_mismatches: List[dict]):
    """
    Выводит таблицу с несовпадениями количества тегов и VALUES (сокращенная версия для консоли)
    
    Args:
        count_mismatches: Список словарей с информацией о несовпадениях
    """
    table = PrettyTable()
    table.field_names = ["№", "title_system", "CODE", "Кол-во тегов", "VALUES", "Несовпадение"]
    table.border = True
    table.align = "l"
    table.align["№"] = "r"
    table.align["Кол-во тегов"] = "r"
    table.align["VALUES"] = "r"
    
    for idx, mismatch in enumerate(count_mismatches, 1):
        # Сокращаем CODE если слишком длинный
        code = str(mismatch['code'])
        if len(code) > 20:
            code = code[:17] + "..."
        
        # Сокращаем title_system если слишком длинный
        title_sys = str(mismatch['title_system'])
        if len(title_sys) > 15:
            title_sys = title_sys[:12] + "..."
        
        # Показываем разницу вместо полного списка тегов
        diff = mismatch['tags_count'] - mismatch['values']
        mismatch_text = f"Разница: {diff:+d}"
        
        row = [
            idx,
            title_sys,
            code,
            mismatch['tags_count'],
            mismatch['values'],
            mismatch_text
        ]
        table.add_row(row)
    
    print(table)


def _save_count_mismatches_to_excel(count_mismatches: List[dict], result_dir: str) -> str:
    """Save MTO tag-count ≠ VALUES rows as compact Excel (same layout as RFP)."""
    return save_tag_count_mismatch_excel(
        count_mismatches,
        result_dir,
        source_label="MTO",
        name_header="MTO_NAME",
        name_key="mto_name",
        ds_header="DS_TITLE",
        ds_key="ds_title",
    )


def _check_duplicate_tags(mto_data: Dict[str, List[RowStd]]) -> List[dict]:
    """
    Проверяет дублирование тегов во всем mto_data
    
    Args:
        mto_data: Словарь title_system -> список строк RowStd из MTO
        
    Returns:
        Список словарей с информацией о дублирующихся тегах
    """
    # Словарь для хранения всех вхождений каждого тега
    tag_occurrences = defaultdict(list)
    
    # Проходим по всем title_system и всем строкам
    for title_system, mto_rows in mto_data.items():
        for row_idx, row in enumerate(mto_rows):
            # Проверяем только строки с типом position_row
            if row.row_type != RowType.position_row:
                continue
            
            # Получаем список тегов
            tags_list = row.get_tags_list()
            
            # Если тегов нет, пропускаем
            if not tags_list:
                continue
            
            # Для каждого тега сохраняем информацию о строке
            for tag in tags_list:
                code_value = row.get_value(CODE) or ""
                mto_name_value = row.get_value(MTO_NAME) or ""
                ds_title_value = row.get_value(DS_TITLE) or ""
                
                tag_occurrences[tag].append({
                    'title_system': title_system or "не определен",
                    'code': str(code_value),
                    'mto_name': str(mto_name_value),
                    'ds_title': str(ds_title_value),
                    'row_index': row_idx + 1  # Для удобства начинаем с 1
                })
    
    # Находим дубликаты (теги, которые встречаются более одного раза)
    duplicate_tags = []
    for tag, occurrences in tag_occurrences.items():
        if len(occurrences) > 1:
            duplicate_tags.append({
                'tag': tag,
                'count': len(occurrences),
                'occurrences': occurrences
            })
    
    # Сортируем по количеству вхождений (по убыванию)
    duplicate_tags.sort(key=lambda x: x['count'], reverse=True)
    
    return duplicate_tags


def _print_duplicate_tags_table(duplicate_tags: List[dict]):
    """
    Выводит таблицу с дублирующимися тегами (сокращенная версия для консоли)
    
    Args:
        duplicate_tags: Список словарей с информацией о дублирующихся тегах
    """
    table = PrettyTable()
    table.field_names = ["№", "Тег", "Кол-во вхождений", "title_system", "CODE"]
    table.border = True
    table.align = "l"
    table.align["№"] = "r"
    table.align["Кол-во вхождений"] = "r"
    
    for idx, dup_info in enumerate(duplicate_tags, 1):
        tag = dup_info['tag']
        count = dup_info['count']
        
        # Сокращаем тег если слишком длинный
        tag_display = tag if len(tag) <= 30 else tag[:27] + "..."
        
        # Берем первый title_system и CODE для краткости
        first_occurrence = dup_info['occurrences'][0]
        title_sys = str(first_occurrence['title_system'])
        if len(title_sys) > 15:
            title_sys = title_sys[:12] + "..."
        
        code = str(first_occurrence['code'])
        if len(code) > 20:
            code = code[:17] + "..."
        
        row = [
            idx,
            tag_display,
            count,
            title_sys,
            code
        ]
        table.add_row(row)
    
    print(table)


def _save_duplicate_tags_to_excel(duplicate_tags: List[dict], result_dir: str) -> str:
    """
    Сохраняет таблицу дублирующихся тегов в Excel файл
    
    Args:
        duplicate_tags: Список словарей с информацией о дублирующихся тегах
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
        ws.title = "Дублирующиеся теги MTO"
        
        # Заголовки
        headers = ["№", "Тег", "Кол-во вхождений", "title_system", "CODE", "MTO_NAME", "DS_TITLE", "№ строки"]
        ws.append(headers)
        
        # Стили для заголовков
        header_fill = PatternFill(start_color="FF6B6B", end_color="FF6B6B", fill_type="solid")
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
        row_counter = 0
        for dup_info in duplicate_tags:
            tag = dup_info['tag']
            count = dup_info['count']
            
            # Для каждого вхождения тега создаем отдельную строку
            for occurrence in dup_info['occurrences']:
                row_counter += 1
                row_data = [
                    row_counter,  # № (общий номер строки в таблице)
                    tag,
                    count,
                    occurrence['title_system'],
                    occurrence['code'],
                    occurrence['mto_name'],
                    occurrence['ds_title'],
                    occurrence['row_index']
                ]
                
                ws.append(row_data)
                
                # Применяем стили к строке
                for col_num in range(1, len(headers) + 1):
                    cell = ws.cell(row=ws.max_row, column=col_num)
                    cell.border = border
                    cell.alignment = Alignment(horizontal="left", vertical="center")
                    
                    # Выравнивание для числовых колонок
                    if col_num == 1 or col_num == 3 or col_num == 8:  # №, Кол-во вхождений, № строки
                        cell.alignment = Alignment(horizontal="right", vertical="center")
                    
                    # Подсветка дубликатов (колонка "Кол-во вхождений")
                    if col_num == 3 and count > 1:
                        fill_color = "FFC7CE"  # Светло-красный
                        cell.fill = PatternFill(start_color=fill_color, end_color=fill_color, fill_type="solid")
        
        # Настраиваем ширину колонок
        ws.column_dimensions['A'].width = 8   # №
        ws.column_dimensions['B'].width = 30  # Тег
        ws.column_dimensions['C'].width = 18  # Кол-во вхождений
        ws.column_dimensions['D'].width = 20  # title_system
        ws.column_dimensions['E'].width = 20  # CODE
        ws.column_dimensions['F'].width = 50  # MTO_NAME
        ws.column_dimensions['G'].width = 20  # DS_TITLE
        ws.column_dimensions['H'].width = 12  # № строки
        
        # Добавляем автофильтр для заголовков столбцов
        last_col_letter = get_column_letter(len(headers))
        ws.auto_filter.ref = f'A1:{last_col_letter}{ws.max_row}'
        
        # Замораживаем первую строку
        ws.freeze_panes = 'A2'
        
        # Сохраняем файл
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        file_name = f"Отчет по дублям тегов - MTO_{timestamp}.xlsx"
        file_path = os.path.join(result_dir, file_name)
        
        wb.save(file_path)
        return file_path
        
    except Exception as e:
        print(f"Ошибка при сохранении таблицы дублирующихся тегов в Excel: {e}")
        return None
