"""
Этап 1: Загрузка данных из RFP файла и проверка тегов
Проверка на дублирование тегов и сравнение количества тегов с VALUES
"""

from __future__ import annotations

from typing import List, Optional
from collections import defaultdict
import os
import sys
import ast
import time
import re

from base.base_class_std_table import STDTable
from base.base_classes import RowStd, TableComments, RowType
from base.base_mto import get_std_from_excel_file
from base.base_excel_out import check_color_out
from base.tables_columns import *
import base.t_comm_initial_classes as t_com_init_cls
from utils.colors import Color
from RFQ.tags_rfp_compare.rfp_tags_utils import (
    ensure_result_dir_exists,
    append_timing_log,
    resolve_units_split_ban_config,
)
from RFQ.tags_rfp_compare.rfp_supply_status import apply_rfp_supply_status_on_rows
from RFQ.tags_rfp_compare.column_optimization import strip_loaded_tags
from RFQ.tags_rfp_compare.tag_count_mismatch_excel import (
    load_parts_lot_tag_mismatches,
    save_tag_count_mismatch_excel,
)
from RFQ.units_convert.models import parse_decimal_quantity, UnitsConversionError
from utils.cache_utils import cache_manager

_RAW_CACHE_ROWTYPE_KEY = "rfp_rt=v3"


def try_exact_int_quantity(value: object) -> int | None:
    """Return a non-negative int when ``value`` is an exact integer quantity."""
    if value is None:
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, int):
        return value if value >= 0 else None
    try:
        quantity = parse_decimal_quantity(value)
    except UnitsConversionError:
        return None
    if quantity < 0:
        return None
    if quantity != quantity.to_integral_value():
        return None
    return int(quantity)


def _raw_cache_extra_key(column_optimization_hash: str | None = None) -> str:
    if column_optimization_hash:
        return f"{column_optimization_hash}|{_RAW_CACHE_ROWTYPE_KEY}"
    return _RAW_CACHE_ROWTYPE_KEY


def step1_load_rfp_raw(
    rfp_path: str,
    result_dir: str,
    column_optimization_hash: str | None = None,
) -> list[RowStd]:
    """Load raw RFP rows before unit conversion, normalize, or split."""
    print("=" * 80)
    print("ЭТАП 1: Загрузка RFP (raw)")

    if not os.path.exists(rfp_path):
        error_msg = f"ОШИБКА: Файл не найден: {rfp_path}"
        print(f"\n{error_msg}")
        print("=" * 80)
        sys.exit(1)

    if not os.path.isfile(rfp_path):
        error_msg = f"ОШИБКА: Указанный путь не является файлом: {rfp_path}"
        print(f"\n{error_msg}")
        print("=" * 80)
        sys.exit(1)

    extra_key = _raw_cache_extra_key(column_optimization_hash)
    load_start = time.perf_counter()
    rfp_t_com = TableComments(
        file_full_path=rfp_path,
        dir_path="-1",
        tabel_class=t_com_init_cls.RFP_AGGREGATED,
    )

    rfp_data = None
    cache_hit = False
    if cache_manager.is_cache_valid(rfp_path, extra_key=extra_key):
        rfp_data = cache_manager.load_from_cache(rfp_path, extra_key=extra_key)
        cache_hit = rfp_data is not None
    if rfp_data is None:
        rfp_data = get_std_from_excel_file(rfp_t_com)
        cache_manager.save_to_cache(rfp_path, rfp_data, extra_key=extra_key)
    append_timing_log(
        result_dir,
        f"step1_load_rfp_raw_file: {time.perf_counter() - load_start:.3f}s (cache_hit={cache_hit})",
    )
    apply_rfp_supply_status_on_rows(rfp_data, file_name=os.path.basename(rfp_path))

    clean_start = time.perf_counter()
    for row in rfp_data:
        if row.row_type == RowType.position_row and row.el.get(DS_TITLE):
            original_value = row.el[DS_TITLE].value
            if original_value and isinstance(original_value, str):
                cleaned_value = original_value.replace(" ", "")
                if cleaned_value != original_value:
                    row.el[DS_TITLE].value = cleaned_value
    append_timing_log(result_dir, f"step1_clean_ds_title_loop: {time.perf_counter() - clean_start:.3f}s")
    print(f"Загружено raw строк: {len(rfp_data)}")
    return rfp_data


def step1_finalize_rfp_data(
    rfp_data: List[RowStd],
    result_dir: str,
    debug: bool = True,
    code_ban_file: str = None,
    skip_split: bool = False,
    units_ban: Optional[List[str]] = None,
    use_rfp_code_ban: bool = True,
    parts_tag_report_path: str | None = None,
) -> tuple[List[RowStd], List[RowStd]]:
    """Normalize, optionally split, and validate already converted RFP rows."""
    print("=" * 80)
    print("ЭТАП 1: Финализация RFP")

    finalize_start = time.perf_counter()
    effective_units_ban = units_ban if units_ban is not None else resolve_units_split_ban_config(None)
    code_ban = _load_code_ban_list(code_ban_file) if use_rfp_code_ban else set()

    normalize_start = time.perf_counter()
    _normalize_rfp_values_columns(rfp_data)
    append_timing_log(result_dir, f"step1_normalize_values_loop: {time.perf_counter() - normalize_start:.3f}s")

    if not skip_split:
        split_start = time.perf_counter()
        rfp_data, split_tags_count, split_no_tags_count = split_rfp_rows(
            rfp_data,
            units_ban=effective_units_ban,
            code_ban=code_ban,
        )
        append_timing_log(result_dir, f"step1_split_rows_loop: {time.perf_counter() - split_start:.3f}s")

    if debug:
        STDTable.print_rows(
            rfp_data,
            column_dict=ColNames.RFP_AGGREAGATED.column_dict,
            count_rows=10,
            row_not_print_list=(RowType.empty_row, RowType.other_row),
        )

    print(f"Финализировано строк: {len(rfp_data)}")
    check_start = time.perf_counter()
    error_rows = _check_rfp_tags_and_values(
        rfp_data,
        result_dir,
        parts_tag_report_path=parts_tag_report_path,
    )
    append_timing_log(result_dir, f"step1_check_rfp_tags_and_values: {time.perf_counter() - check_start:.3f}s")
    append_timing_log(result_dir, f"step1_finalize_rfp_data: {time.perf_counter() - finalize_start:.3f}s")
    return rfp_data, error_rows


def step1_load_rfp_data(
    rfp_path: str,
    result_dir: str,
    debug=True,
    code_ban_file: str = None,
    column_optimization_hash: str = None,
    skip_split: bool = False,
    units_ban: Optional[List[str]] = None,
    use_rfp_code_ban: bool = True,
    parts_tag_report_path: str | None = None,
    load_tags: bool = True,
) -> tuple[List[RowStd], List[RowStd]]:
    """
    Этап 1: Загрузка данных из RFP файла (Excel) и проверка тегов
    
    Args:
        rfp_path: Путь к Excel файлу RFP
        result_dir: Путь к папке результатов для вывода промежуточных результатов
        debug: Флаг отладки
        skip_split: Не выполнять split RFP в step1 (перенос в worker step4 при parallel).
        units_ban: Ед. изм., для которых не раскладывать строки по 1; None — дефолтный список; [] — бан отключён.
        use_rfp_code_ban: Если False, файл code_ban не читается; в split_rfp_rows передаётся пустой code_ban.
        load_tags: Если False, обнулить TAGS до finalize/split (как ``load_tags`` в конфиге).

    Returns:
        Кортеж: (список всех строк RFP, список строк с ошибками проверки)
    """
    step1_start = time.perf_counter()
    rfp_data = step1_load_rfp_raw(
        rfp_path,
        result_dir,
        column_optimization_hash=column_optimization_hash,
    )
    if not load_tags:
        strip_loaded_tags(rfp_data)
    rfp_data, error_rows = step1_finalize_rfp_data(
        rfp_data,
        result_dir,
        debug=debug,
        code_ban_file=code_ban_file,
        skip_split=skip_split,
        units_ban=units_ban,
        use_rfp_code_ban=use_rfp_code_ban,
        parts_tag_report_path=parts_tag_report_path,
    )
    append_timing_log(result_dir, f"step1_load_rfp_data: {time.perf_counter() - step1_start:.3f}s")
    return rfp_data, error_rows


def split_rfp_rows(
    rows: List[RowStd],
    units_ban: List[str] | None = None,
    code_ban: set[str] | None = None,
) -> tuple[List[RowStd], int, int]:
    """
    Разбиение строк RFP:
    - несколько тегов -> одна строка на тег (VALUES по 1 до исчерпания)
    - без тегов и VALUES > 1 -> VALUES строк по 1 (если код/ед.изм. не в ban)
    """
    units_ban_set = {str(u).strip().lower() for u in (units_ban or [])}
    code_ban_set = code_ban or set()

    expanded_rows: List[RowStd] = []
    split_tags_count = 0
    split_no_tags_count = 0

    def _set_order_attrs(target_row: RowStd, source_row: RowStd, split_seq: int = 0) -> None:
        """Сохраняет порядок исходной строки для последующей сборки per-TM результата."""
        orig_idx = getattr(source_row, "_orig_idx", None)
        if isinstance(orig_idx, int):
            setattr(target_row, "_orig_idx", orig_idx)
            setattr(target_row, "_orig_split_seq", split_seq)

    for row in rows:
        if row.row_type != RowType.position_row:
            expanded_rows.append(row)
            continue

        tags_list = row.get_tags_list()
        values_value = row.get_value(VALUES)
        values_int = try_exact_int_quantity(values_value)
        if values_int is None:
            _set_order_attrs(row, row, split_seq=0)
            expanded_rows.append(row)
            continue

        units_value = row.get_value(UNITS) or ""
        units_norm = str(units_value).strip().lower()
        is_unit_banned = units_norm in units_ban_set

        code_value = row.get_value(CODE)
        code_norm = _normalize_code_for_ban(code_value)
        is_code_banned = code_norm in code_ban_set

        if tags_list and len(tags_list) > 1 and not is_unit_banned:
            remaining_values = values_int
            for split_seq, tag in enumerate(tags_list):
                val = 1 if remaining_values > 0 else 0
                if remaining_values > 0:
                    remaining_values -= 1
                new_row = RowStd.get_row_copy_light(row, tags_override=[tag], values_override=val)
                _set_order_attrs(new_row, row, split_seq=split_seq)
                expanded_rows.append(new_row)
            split_tags_count += 1
            continue

        if not tags_list and values_int > 1 and not is_unit_banned and not is_code_banned:
            for split_seq in range(values_int):
                new_row = RowStd.get_row_copy_light(row, values_override=1)
                _set_order_attrs(new_row, row, split_seq=split_seq)
                expanded_rows.append(new_row)
            split_no_tags_count += 1
            continue

        _set_order_attrs(row, row, split_seq=0)
        expanded_rows.append(row)

    return expanded_rows, split_tags_count, split_no_tags_count


def _rfp_count_mismatch_dict(
    row: RowStd,
    tags_list: list[str],
    tags_count: int,
    values_qty: object,
) -> dict:
    """Compact row for ``Отчет по несоответствию тегов - RFP`` (same layout as MTO)."""
    return {
        "title_system": str(row.get_value(DS_TITLE) or "не определен"),
        "code": str(row.get_value(CODE) or ""),
        "name": str(row.get_value(NAME) or ""),
        "ds_name": str(row.get_value(DS_NAME) or ""),
        "tags_count": tags_count,
        "values": values_qty,
        "tags": ", ".join(tags_list),
    }


def _check_rfp_tags_and_values(
    rfp_data: List[RowStd],
    result_dir: str,
    parts_tag_report_path: str | None = None,
) -> List[RowStd]:
    """Check RFP tag uniqueness and tag-count vs VALUES; write Excel reports.

    Compact ``Отчет по несоответствию тегов - RFP_*.xlsx`` matches the MTO
    report. With parts net, tagged rows are already VALUES=1, so lot≠tags
    rows are taken from ``Отчет по тегам - сбор частей.xlsx`` when present.
    """
    tag_occurrences = defaultdict(list)
    duplicate_tag_rows = []
    count_mismatch_rows = []
    net_count_mismatches: list[dict] = []

    for idx, rfp_row in enumerate(rfp_data):
        if rfp_row.row_type != RowType.position_row:
            continue
        rfp_row_tags = rfp_row.get_tags_list()
        if not rfp_row_tags:
            continue
        for tag in rfp_row_tags:
            tag_occurrences[tag].append(idx)

    for idx, rfp_row in enumerate(rfp_data):
        if rfp_row.row_type != RowType.position_row:
            continue
        rfp_row_tags = rfp_row.get_tags_list()
        if not rfp_row_tags:
            continue

        rfp_row_value_1 = rfp_row.get_value(VALUES)
        rfp_row_value_2 = rfp_row.get_value(VALUES_2)
        has_duplicate = False
        has_count_mismatch = False
        error_messages = []

        for tag in rfp_row_tags:
            occurrences = tag_occurrences[tag]
            if len(occurrences) > 1:
                has_duplicate = True
                error_messages.append(
                    f"Тег '{tag}' дублируется (встречается в {len(occurrences)} строках)"
                )

        tags_count = len(rfp_row_tags)
        values_int = try_exact_int_quantity(rfp_row_value_1)
        try:
            value_1 = float(rfp_row_value_1) if rfp_row_value_1 else 0.0
        except (ValueError, TypeError):
            value_1 = 0.0
        try:
            value_2 = float(rfp_row_value_2) if rfp_row_value_2 else 0.0
        except (ValueError, TypeError):
            value_2 = 0.0

        if values_int is not None:
            qty_for_compare = values_int
        else:
            qty_for_compare = value_1
        if tags_count != qty_for_compare:
            has_count_mismatch = True
            error_messages.append(
                f"Количество тегов ({tags_count}) не совпадает с VALUES ({qty_for_compare})"
            )
        if value_2 > 0 and tags_count != value_2:
            has_count_mismatch = True
            error_messages.append(
                f"Количество тегов ({tags_count}) не совпадает с VALUES_2 ({value_2})"
            )

        if has_duplicate or has_count_mismatch:
            error_comment = "; ".join(error_messages)
            if rfp_row.el.get(TAGS):
                rfp_row.el[TAGS].comment += f"\nОшибки: {error_comment}"
            if has_duplicate:
                duplicate_tag_rows.append(RowStd.get_row_copy(rfp_row))
            if has_count_mismatch:
                count_mismatch_rows.append(RowStd.get_row_copy(rfp_row))
                net_count_mismatches.append(
                    _rfp_count_mismatch_dict(
                        rfp_row,
                        rfp_row_tags,
                        tags_count,
                        values_int if values_int is not None else qty_for_compare,
                    )
                )

    ensure_result_dir_exists(result_dir)
    if duplicate_tag_rows:
        _output_duplicate_tags(duplicate_tag_rows, result_dir)

    excel_mismatches = net_count_mismatches
    if parts_tag_report_path and os.path.isfile(parts_tag_report_path):
        parts_mismatches = load_parts_lot_tag_mismatches(parts_tag_report_path)
        if parts_mismatches:
            excel_mismatches = parts_mismatches
    if excel_mismatches:
        file_path = save_tag_count_mismatch_excel(
            excel_mismatches,
            result_dir,
            source_label="RFP",
            name_header="NAME",
            name_key="name",
            ds_header="DS_NAME",
            ds_key="ds_name",
        )
        if file_path:
            print(
                f"Несовпадения количества RFP ({len(excel_mismatches)}) -> "
                f"{os.path.basename(file_path)}"
            )

    return duplicate_tag_rows + count_mismatch_rows


def _parse_int_like_value(raw_value) -> int | None:
    """Convert exact integer-like quantity text to ``int`` without rounding fractions."""
    return try_exact_int_quantity(raw_value)


def _normalize_rfp_values_columns(rfp_data: List[RowStd]) -> int:
    """
    Проходит по строкам RFP и нормализует значения в колонках VALUES и VALUES_2.
    Возвращает количество обновлённых ячеек.
    """
    updated = 0
    for row in rfp_data:
        if row.row_type != RowType.position_row:
            continue

        for col_name in (VALUES, VALUES_2):
            el = row.el.get(col_name)
            if not el:
                continue
            new_val = _parse_int_like_value(el.value)
            if new_val is not None and new_val != el.value:
                el.value = new_val
                updated += 1
    return updated


def _load_code_ban_list(code_ban_file: str) -> set[str]:
    if not code_ban_file:
        return set()
    if not os.path.exists(code_ban_file):
        return set()
    try:
        with open(code_ban_file, "r", encoding="utf-8") as file:
            content = file.read().strip()
        if not content:
            return set()
        if content.startswith("code_ban="):
            content = content.split("=", 1)[1].strip()
        values = ast.literal_eval(content)
        if isinstance(values, list):
            return {_normalize_code_for_ban(v) for v in values if _normalize_code_for_ban(v)}
    except Exception as e:
        print(f"Ошибка чтения code_ban ({code_ban_file}): {e}")
    return set()


def _normalize_code_for_ban(code_value) -> str:
    if code_value is None:
        return ""
    return str(code_value).strip().upper()


def _output_duplicate_tags(duplicate_rows: List[RowStd], result_dir: str):
    """
    Выводит строки с дублирующимися тегами в Excel с сортировкой по тегам и чередующимися цветами
    
    Args:
        duplicate_rows: Список строк с дублирующимися тегами
        result_dir: Путь к папке результатов
    """
    # Сортируем по первому тегу в строке
    def get_first_tag(row):
        tags = row.get_tags_list()
        return tags[0] if tags else ""
    
    sorted_rows = sorted(duplicate_rows, key=get_first_tag)
    
    # Применяем чередующиеся цвета к строкам
    current_color = Color.soft_yellow
    prev_tag = None
    
    for row in sorted_rows:
        tags = row.get_tags_list()
        first_tag = tags[0] if tags else ""
        
        # Если тег изменился, меняем цвет (чередуем между soft_yellow и soft_cyan)
        if prev_tag is not None and first_tag != prev_tag:
            current_color = Color.soft_cyan if current_color == Color.soft_yellow else Color.soft_yellow
        
        # Применяем цвет ко всем элементам строки
        for att in ColNames.column_list:
            if row.el.get(att):
                Color.set_el_color(row.el[att], current_color)
        
        prev_tag = first_tag
    
    # Вывод в Excel
    file_path = check_color_out(
        sorted_rows,
        col_for_out_dict=ColNames.RFP_AGGREAGATED.column_dict,
        out_dir=result_dir,
        file_prefix="Отчет по дублям тегов - RFP",
        row_not_print_list=(RowType.empty_row, RowType.other_row),
        show_row_type=False,
        open_folder=False,
        excel_template_check_color=r"RFQ\tags_rfp_compare\tamplates\RFP_проверка_тегов_ошибки_Сводная RFP_шаблон.xlsx",
        row_index=1
    )
    print(f"Дублирующиеся теги RFP ({len(duplicate_rows)}) -> {os.path.basename(file_path)}")


if __name__ in {"__main__"}:
    # Пример запуска этапа 1
    import sys
    import os
    from rfp_tags_utils import get_result_dir_path, ensure_result_dir_exists
    
    if len(sys.argv) > 1:
        rfp_path = sys.argv[1]
    else:
        # Путь относительно корня проекта
        script_dir = os.path.dirname(os.path.abspath(__file__))
        rfp_path = os.path.join(script_dir, "input", "Сводная RFP.xlsx")
    
    result_dir = get_result_dir_path(rfp_path)
    ensure_result_dir_exists(result_dir)
    
    rfp_data, error_rows = step1_load_rfp_data(rfp_path, result_dir)
    print(f"\nЗагружено строк RFP: {len(rfp_data)}")
    print(f"Найдено ошибок: {len(error_rows)}")

