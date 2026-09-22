"""
Этап анализа и сопоставления данных
Сопоставление строк из RFP со строками из MTO и VO
"""

from typing import List, Dict, Set, Tuple, Optional
from collections import defaultdict
from datetime import datetime
import gc
import os
import sys
import ast
import time
from concurrent.futures import ProcessPoolExecutor, as_completed

from prettytable import PrettyTable
from tqdm import tqdm
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

from base.base_classes import RowStd, RowType
from base.tables_columns import (
    MATCH_STATUS, MATCH_STATUS_VO, TAG_MTO, TAG_VO, DS_TITLE, DS_NUMBER,
    DS_NAME, CODE, CODE_MTO, CODE_VO, VALUES, POSITION_STATUS,
    TAG_EFFECTIVE, VALUES_MTO, VALUES_VO,
    IN_CABINET, UNITS, UNITS_MTO, UNITS_CHECK_STATUS, UNITS_CONVERSION_TRACE,
    NAME, TYPE_MARK, VENDOR
)

from RFQ.ds_compare.ds_units_normalize import (
    merge_units_status_trace,
    normalize_units_text,
)
from RFQ.tags_rfp_compare.rfp_supply_status import is_excluded_from_supply
from RFQ.tags_rfp_compare.step4.step4_1_check_mto_data import (
    save_mto_diagnostics_reports,
    save_vo_diagnostics_reports,
)
from RFQ.tags_rfp_compare.step4.step4_5_check_values_sum import check_values_sum
from RFQ.tags_rfp_compare.step4.step4_6_save_match_result_to_excel import save_match_result_to_excel
from RFQ.tags_rfp_compare.step4.step4_postmerge_ops import apply_mto_ul_quantity_diff
from RFQ.tags_rfp_compare.step4.step4_quantity_balance import (
    QuantityBalanceFatalError,
    _title_key,
    apply_post_split_input_to_snapshot,
    build_input_snapshot,
    format_input_rfp_mto_vo_detail,
    format_input_ul_detail,
    print_input_quantity_balance,
    print_input_ul_quantity_balance,
    run_output_quantity_balance_check,
    run_output_ul_quantity_balance_check,
)
from RFQ.tags_rfp_compare.step4.step4_packing_compare import (
    RfpPackingFatalError,
    apply_zip_smr_pnr_status,
    build_ordered_snapshot,
    compare_rfp_rows_with_packing,
    packing_queue_input_by_title,
    parse_composite_title,
    save_rfp_packing_report,
)
from RFQ.tags_rfp_compare.step4.step4_worker import process_title_mark, process_title_mark_parallel, worker_init
from RFQ.ds_compare.support_finctions import load_replacement_table
from RFQ.packing_list_provider import (
    PackingDataset,
    load_packing_dataset,
    packing_match_key,
    packing_row_key,
    registry_packing_scope,
)
from RFQ.tags_rfp_compare.rfp_tags_utils import (
    DEFAULT_UNITS_SPLIT_BAN,
    ensure_result_dir_exists,
    append_timing_log,
    append_memory_log,
    log_data_structures_memory,
)
from RFQ.tags_rfp_compare.rfp_progress import emit_milestone
from RFQ.tags_rfp_compare.step4.step4_bcc_accum_matrix import (
    BccAccumMatrix,
    save_bcc_accum_matrix_to_excel,
)
from RFQ.tags_rfp_compare.ds_manager_matrix import (
    DsManagerMatrix,
    apply_ds_source_columns,
    build_mto_paths_by_title,
)
from RFQ.tags_rfp_compare.gem_supply_codes import (
    GemSupplyCodes,
    apply_gem_supply_status,
    paint_gem_supply_rows,
)


def _normalize_title_mark(value) -> str:
    if value is None:
        return ""
    return str(value).strip()


def _build_code_base_by_code(rows: List[RowStd]) -> Dict[str, tuple]:
    """Строит словарь {code: (name, type_mark, vendor)} для O(1) lookup.
    Хранит только три нужных поля вместо полных RowStd, чтобы минимизировать
    pickle-объём при передаче в workers через initializer. First match wins.
    """
    out: Dict[str, tuple] = {}
    for r in rows or []:
        c = r.el.get(CODE)
        if c is None:
            continue
        v = c.value
        if v is not None and str(v).strip() and v not in out:
            name_el = r.el.get(NAME)
            type_mark_el = r.el.get(TYPE_MARK)
            vendor_el = r.el.get(VENDOR)
            out[v] = (
                name_el.value if name_el else None,
                type_mark_el.value if type_mark_el else None,
                vendor_el.value if vendor_el else None,
            )
    return out


def _partition_rfp_rows_by_title_mark(
    rfp_data: List[RowStd],
) -> Tuple[List[str], Dict[str, List[RowStd]], List[RowStd]]:
    """
    Разбивает RFP-строки по title_mark, сохраняя исходный порядок внутри каждой группы.
    Строки без DS_TITLE привязываются к последнему встреченному title_mark.
    """
    title_order: List[str] = []
    rows_by_title: Dict[str, List[RowStd]] = defaultdict(list)
    orphan_rows: List[RowStd] = []
    last_title_mark = ""

    for row in rfp_data:
        row_title_mark = _normalize_title_mark(row.get_value(DS_TITLE))
        if row_title_mark:
            last_title_mark = row_title_mark
        elif last_title_mark:
            row_title_mark = last_title_mark

        if row_title_mark:
            if row_title_mark not in rows_by_title:
                title_order.append(row_title_mark)
            rows_by_title[row_title_mark].append(row)
        else:
            orphan_rows.append(row)

    return title_order, dict(rows_by_title), orphan_rows


def _build_title_mark_order(
    rfp_title_order: List[str],
    mto_data: Dict[str, List[RowStd]],
    vo_data: Dict[str, List[RowStd]],
) -> List[str]:
    order: List[str] = []
    seen: Set[str] = set()

    for title_mark in rfp_title_order:
        normalized = _normalize_title_mark(title_mark)
        if normalized and normalized not in seen:
            order.append(normalized)
            seen.add(normalized)

    for source in (mto_data or {}, vo_data or {}):
        for title_mark in source.keys():
            normalized = _normalize_title_mark(title_mark)
            if normalized and normalized not in seen:
                order.append(normalized)
                seen.add(normalized)

    return order


def _restore_original_row_order(result_rows: List[RowStd]) -> List[RowStd]:
    """
    Восстанавливает исходный порядок строк RFP после per-TM merge.
    - строки из исходного RFP (с _orig_idx) сортируются по _orig_idx
    - добавленные строки без _orig_idx остаются в исходном относительном порядке и идут после
    """
    with_idx: List[tuple[int, int, RowStd]] = []
    without_idx: List[RowStd] = []
    for row in result_rows:
        idx = getattr(row, "_orig_idx", None)
        if isinstance(idx, int):
            split_seq = getattr(row, "_orig_split_seq", 0)
            if not isinstance(split_seq, int):
                split_seq = 0
            with_idx.append((idx, split_seq, row))
        else:
            without_idx.append(row)
    with_idx.sort(key=lambda x: (x[0], x[1]))
    return [row for _, _, row in with_idx] + without_idx


def _resolve_packing_dataset_for_step4(
    *,
    include_packing_lists: bool,
    packing_dataset: PackingDataset | None,
) -> PackingDataset | None:
    """Return the caller-supplied dataset or load one for backward-compatible callers."""
    if not include_packing_lists:
        return None
    if packing_dataset is not None:
        return packing_dataset
    return load_packing_dataset()


def step4_analyze_and_match(rfp_data: List[RowStd],
                            mto_data: Dict[str, List[RowStd]],
                            vo_data: Dict[str, List[RowStd]],
                            result_dir: str,
                            replacement_table_file: str = None,
                            rfp_registr_lot_path: str = None,
                            code_ban_file: str = None,
                            debug: bool = True,
                            debug_step4_1: bool = True,
                            debug_step4_2: bool = True,
                            debug_step4_3: bool = True,
                            debug_step4_4: bool = True,
                            collapse_debug: bool = False,
                            unified_debug: bool = True,
                            debug_tag: List[str] = None,
                            debug_code= ["BCC0003424"],
                            debug_title_system=["8445-SOT1"],
                            print_title_systems_table: bool = False,
                            use_multiprocessing: bool = False,
                            split_rfp_in_worker: bool = False,
                            parallel_by_title_mark: bool = False,
                            max_workers: int = 0,
                            timing_log_enabled: bool = True,
                            verbose_progress_messages: bool = True,
                            export_title_systems_comparison_excel: bool = True,
                            code_base_by_code: Dict[str, tuple] | None = None,
                            assign_rfp_mto_code_compare_colors: bool = True,
                            filter_to_mto_titles: bool = False,
                            include_mto_vo_without_rfp_anchor: bool = False,
                            include_packing_lists: bool = True,
                            ul_match_use_mto_tags: bool = True,
                            packing_dataset: PackingDataset | None = None,
                            units_split_ban: Optional[List[str]] = None,
                            use_rfp_code_ban: bool = True,
                            export_bcc_accum_matrix: bool = True,
                            bcc_accum_matrix: BccAccumMatrix | None = None,
                            ds_manager_matrix: DsManagerMatrix | None = None,
                            gem_supply_codes: GemSupplyCodes | None = None,
                            rfp_paths_by_key: dict[str, list[str]] | None = None,
                            registry_packing_index: object | None = None):
    """
    Анализ всех данных для сопоставления строк из RFP со строками из MTO и VO
    
    Args:
        rfp_data: Данные из RFP (список строк RowStd)
        mto_data: Данные из MTO (словарь: title_system -> список строк RowStd)
        vo_data: Данные из VO (словарь: title_system -> список строк RowStd)
        result_dir: Путь к папке результатов
        replacement_table_file: Путь к таблице замен кодов
        rfp_registr_lot_path: Путь к реестру лотов (ДС -> Лот)
        code_ban_file: Путь к файлу code_ban
        debug: Флаг отладки (общий, fallback для debug_step4_1..4)
        debug_step4_1: Отладка check_mto_data, check_vo_data
        debug_step4_2: Отладка match_rfp_with_mto, match_rfp_with_vo
        debug_step4_3: Отладка assign_position_status
        debug_step4_4: Отладка add_unmatched_*, match_* и т.п.
        collapse_debug: Сохранять step4_collapse_debug_*.txt
        unified_debug: Сохранять step4_unified_debug_*.txt
        debug_tag: Список тегов для точечной отладки (AND с другими фильтрами, OR внутри списка)
        debug_code: Список кодов для точечной отладки (AND с другими фильтрами, OR внутри списка)
        debug_title_system: Список title_system для точечной отладки (AND с другими фильтрами, OR внутри списка)
        print_title_systems_table: Выводить в консоль таблицу сравнения title_system (RFP/MTO/VO)
        parallel_by_title_mark: Параллельная обработка title_mark через ProcessPoolExecutor
        max_workers: Количество worker-процессов (<=0 -> auto: cpu_count-1)
        timing_log_enabled: Флаг записи timing_log (передаётся в worker-процессы)
        verbose_progress_messages: Печатать подробные сообщения прогресса 4_3/4_4 в консоль
        units_split_ban: Ед. изм., для которых не раскладывать RFP/MTO в строки по 1 при split;
            None — список по умолчанию; [] — бан отключён (раскладка для всех допустимых случаев).
        use_rfp_code_ban: Если False, файл code_ban не читается и не обновляется в конце step4;
            в worker/step1 split передаётся пустой code_ban (полное раскрытие по кодам).
        include_mto_vo_without_rfp_anchor: Если True, несопоставленные MTO/VO для title_system без
            RFP position-якоря добавляются в конец списка строк TM (шире финальный Excel).
            Возможны дополнительные предупреждения check_values_sum для TS без строк в RFP.
        include_packing_lists: Если True, после quantity balance сопоставить результат
            с общим кэшем упаковочных листов, сверить вход/выход УЛ и сохранить audit-отчёт.
        ul_match_use_mto_tags: Если True, при посадке УЛ после тегов RFP и безтеговых
            слотов сначала искать совпадение с ``TAG_MTO`` той же строки; иначе сразу
            добор по коду (``tag_mismatch``).
        packing_dataset: Уже загруженный набор УЛ из оркестратора; при ``None`` и
            ``include_packing_lists=True`` выполняется fallback ``load_packing_dataset()``.
        export_bcc_accum_matrix: Если True, после основного Excel пишется
            накопительная матрица BCC (отдельный xlsx). Ошибка записи не валит Запуск.
        bcc_accum_matrix: Снимок qty после units_gate. Если None при включённом
            флаге — milestone Skipped.
        ds_manager_matrix: Матрица ДС→МП для колонок DS_MANAGER/PATH_* в Excel.
            None — все DS_NAME считаются без МП (жёлтая подсветка DS_MANAGER).
        gem_supply_codes: Набор кодов ГЭМ; None/пусто — overlay no-op.
        rfp_paths_by_key: Ключ файла ДС (``ДС10``, ``ДС47_13А``) → UNC-пути
            книг в ``RFP_Зиновьев``. None трактуется как пустой dict
            (жёлтая подсветка PATH_RFP при непустом DS_NAME).
        
    Returns:
        Результат анализа (список агрегированных тегов из MTO)
    """
    with registry_packing_scope(registry_packing_index):
        print("=" * 80)
        print("ЭТАП 4: Сопоставление RFP / MTO / VO")
    
        total_mto = sum(len(rows) for rows in mto_data.values()) if mto_data else 0
        total_vo = sum(len(rows) for rows in vo_data.values()) if vo_data else 0
        print(f"RFP: {len(rfp_data)} | MTO: {total_mto} ({len(mto_data)} TS) | VO: {total_vo} ({len(vo_data)} TS)")
    
        _d1 = debug_step4_1
        _d2 = debug_step4_2
        _d3 = debug_step4_3
        _d4 = debug_step4_4
        _any_debug = _d1 or _d2 or _d3 or _d4
        debug_log_lines: List[str] = [] if _any_debug else None
        total_start = time.perf_counter()

        def log_timing(step_name: str, start_time: float) -> None:
            append_timing_log(result_dir, f"step4::{step_name}: {time.perf_counter() - start_time:.3f}s")

        # Сравнение title_system для RFP, MTO, VO (кросс-TM проверка до цикла)
        step_start = time.perf_counter()
        _compare_title_systems_rfp_mto_vo(
            rfp_data,
            mto_data,
            vo_data,
            result_dir,
            print_table=print_title_systems_table,
            export_to_excel=export_title_systems_comparison_excel,
        )
        log_timing("compare_title_systems", step_start)

        quantity_input_snapshot = build_input_snapshot(rfp_data, mto_data, vo_data)
        print_input_quantity_balance(quantity_input_snapshot)
        packing_ordered_snapshot = (
            build_ordered_snapshot(rfp_data) if include_packing_lists else {}
        )

        mto_paths_by_title = build_mto_paths_by_title(mto_data or {})

        split_units_ban = (
            list(units_split_ban) if units_split_ban is not None else list(DEFAULT_UNITS_SPLIT_BAN)
        )
        split_code_ban = (
            set(_load_code_ban_from_file(code_ban_file))
            if use_rfp_code_ban and split_rfp_in_worker
            else set()
        )

        # Снапшот RFP для check_values_sum (до модификации)
        rfp_sums_by_title, rfp_counts_by_title = _extract_rfp_snapshot_for_check(
            rfp_data,
            skip_split=split_rfp_in_worker,
            units_ban=split_units_ban,
            code_ban=split_code_ban,
        )

        step_start = time.perf_counter()
        replacement_table = (
            load_replacement_table(replacement_table_file, print_replacement_table_info=False)
            if replacement_table_file
            else {}
        )
        log_timing("load_replacement_table", step_start)

        lot_map = _load_lot_registry(rfp_registr_lot_path) if rfp_registr_lot_path else {}

        # code_base_by_code загружается в agregate_tags.py параллельно со step2+step3
        # и передаётся сюда готовым. Fallback на пустой dict если не передан.
        if code_base_by_code is None:
            code_base_by_code = {}

        for idx, row in enumerate(rfp_data):
            setattr(row, "_orig_idx", idx)

        rfp_title_order, rfp_rows_by_title, orphan_rows = _partition_rfp_rows_by_title_mark(rfp_data)
        title_marks = _build_title_mark_order(rfp_title_order, mto_data, vo_data)

        if filter_to_mto_titles and mto_data:
            mto_title_set = set(mto_data.keys())
            before = len(title_marks)
            title_marks = [tm for tm in title_marks if tm in mto_title_set]
            n_rows = sum(len(rfp_rows_by_title.get(tm, [])) for tm in title_marks)
            print(f"pre-filter TM: {before} -> {len(title_marks)}, RFP строк в работе: {n_rows}")

        result_rows: List[RowStd] = list(orphan_rows)
        agg_mto_tagged_sums: Dict[str, float] = {}
        agg_mto_tagged_values_by_code: Dict[tuple, float] = defaultdict(float)
        agg_vo_tagged_sums: Dict[str, float] = {}
        agg_vo_codes: Set[str] = set()
        agg_mto_codes_with_tags: Set[str] = set()
        all_mto_count_mismatches: List[dict] = []
        all_mto_duplicate_tags: List[dict] = []
        all_vo_duplicate_tags: List[dict] = []
        post_split_mto_by_title: Dict[str, float] = {}
        post_split_mto_rows_by_title: Dict[str, int] = {}
        post_split_vo_by_title: Dict[str, float] = {}
        post_split_vo_rows_by_title: Dict[str, int] = {}

        append_memory_log(result_dir, "step4_before_tm_loop")
        log_data_structures_memory(
            result_dir,
            "step4_before_tm_loop",
            rfp_data=rfp_data,
            mto_data=mto_data,
            vo_data=vo_data,
        )

        use_parallel_tm = bool(parallel_by_title_mark and len(title_marks) > 1)

        tm_common_kwargs = {
            "result_dir": result_dir,
            "replacement_table": replacement_table,
            "debug_step4_1": _d1,
            "debug_step4_2": _d2,
            "debug_step4_3": _d3,
            "debug_step4_4": _d4,
            "debug_tag": debug_tag,
            "debug_code": debug_code,
            "debug_title_system": debug_title_system,
            "use_multiprocessing": use_multiprocessing,
            "collapse_debug": collapse_debug,
            "split_rfp_in_worker": split_rfp_in_worker,
            "units_ban": split_units_ban,
            "code_ban": split_code_ban,
            "collapse_func": _collapse_rfp_rows_before_export,
            "lot_map": lot_map,
            "code_base_by_code": code_base_by_code,
            # В parallel-by-title-mark workers не пишут timing_log.xlsx:
            # запись выполняется только в main-процессе (step4::tm::* / step4_total).
            "timing_log_enabled": (timing_log_enabled and not use_parallel_tm),
            "verbose_progress_messages": verbose_progress_messages,
            # Подавляем консольный вывод и Excel-отчёты в workers; диагностика
            # собирается через return-dict и сохраняется единым файлом после merge.
            "defer_reports": use_parallel_tm,
            "assign_rfp_mto_code_compare_colors": assign_rfp_mto_code_compare_colors,
            "include_mto_vo_without_rfp_anchor": include_mto_vo_without_rfp_anchor,
        }

        def _merge_tm_result(tm_result: Dict[str, object], tm_idx: int) -> None:
            title = tm_result["title_mark"]
            result_rows.extend(tm_result["result_rows"])
            mto_s = tm_result.get("mto_summary") or {}
            vo_s = tm_result.get("vo_summary") or {}
            if mto_s.get("tagged_sum"):
                agg_mto_tagged_sums[title] = (
                    agg_mto_tagged_sums.get(title, 0.0) + mto_s["tagged_sum"]
                )
            for code, val in (mto_s.get("tagged_values_by_code") or {}).items():
                agg_mto_tagged_values_by_code[(title, code)] += val
            agg_mto_codes_with_tags.update(mto_s.get("codes_with_tags") or set())
            if vo_s.get("tagged_sum"):
                agg_vo_tagged_sums[title] = (
                    agg_vo_tagged_sums.get(title, 0.0) + vo_s["tagged_sum"]
                )
            agg_vo_codes.update(vo_s.get("codes") or set())
            mto_d = tm_result.get("mto_diagnostics") or {}
            vo_d = tm_result.get("vo_diagnostics") or {}
            all_mto_count_mismatches.extend(mto_d.get("count_mismatches", []))
            all_mto_duplicate_tags.extend(mto_d.get("duplicate_tags", []))
            all_vo_duplicate_tags.extend(vo_d.get("duplicate_tags_vo", []))
            tm_key = _title_key(title)
            if "post_split_mto_sum" in tm_result:
                post_split_mto_by_title[tm_key] = float(tm_result["post_split_mto_sum"])
                post_split_mto_rows_by_title[tm_key] = int(tm_result.get("post_split_mto_rows") or 0)
            if "post_split_vo_sum" in tm_result:
                post_split_vo_by_title[tm_key] = float(tm_result["post_split_vo_sum"])
                post_split_vo_rows_by_title[tm_key] = int(tm_result.get("post_split_vo_rows") or 0)
            worker_dbg = tm_result.get("debug_log_lines")
            if worker_dbg and debug_log_lines is not None:
                debug_log_lines.extend(worker_dbg)
            if tm_idx % 10 == 0:
                append_memory_log(result_dir, f"step4_tm_loop_after_{tm_idx}")
                log_data_structures_memory(result_dir, f"step4_tm_loop_after_{tm_idx}", result_rows=result_rows)

        tm_loop_start = time.perf_counter()
        if use_parallel_tm:
            cpu_cnt = os.cpu_count() or 1
            resolved_workers = max_workers if isinstance(max_workers, int) and max_workers > 0 else max(1, cpu_cnt - 1)
            resolved_workers = max(1, min(resolved_workers, len(title_marks)))
            # C.3: LPT scheduling — тяжёлые TM первыми для уменьшения tail latency
            tm_weights = []
            for tm in title_marks:
                n_rfp = len(rfp_rows_by_title.get(tm, []))
                n_mto = len(mto_data.get(tm, []))
                n_vo = len(vo_data.get(tm, []))
                weight = n_rfp + n_mto * 2 + n_vo
                tm_weights.append((tm, weight))
            tm_weights.sort(key=lambda x: x[1], reverse=True)
            sorted_tms = [tw[0] for tw in tm_weights]

            top_w = tm_weights[0][1] if tm_weights else 0
            bot_w = tm_weights[-1][1] if tm_weights else 0
            print(f"Title_mark: {len(title_marks)} | Параллельно: {resolved_workers} workers, weight {top_w}..{bot_w}")
            # Metadata only: count of workers (not a duration).
            append_timing_log(result_dir, f"step4::parallel_by_title_mark::workers_count={resolved_workers}")
            try:
                with ProcessPoolExecutor(
                    max_workers=resolved_workers,
                    initializer=worker_init,
                    initargs=(tm_common_kwargs,),
                ) as executor:
                    future_to_meta = {}
                    for idx, title_mark in enumerate(sorted_tms, 1):
                        future = executor.submit(
                            process_title_mark_parallel,
                            title_mark,
                            rfp_rows_by_title.get(title_mark, []),
                            mto_data.get(title_mark, []),
                            vo_data.get(title_mark, []),
                        )
                        future_to_meta[future] = (idx, title_mark, time.perf_counter())

                    completed_idx = 0
                    total_tms = len(title_marks)
                    gc.disable()
                    with tqdm(total=total_tms, desc="Step4 TM", unit="tm", ncols=80, dynamic_ncols=False) as pbar:
                        for future in as_completed(future_to_meta):
                            idx, title_mark, started = future_to_meta[future]
                            wall_elapsed = time.perf_counter() - started
                            tm_result = future.result()
                            worker_dur = tm_result.get("worker_duration", wall_elapsed)
                            append_timing_log(result_dir, f"step4::tm::{title_mark}: {worker_dur:.3f}s")
                            completed_idx += 1
                            pbar.update(1)
                            _merge_tm_result(tm_result, completed_idx)
                            del tm_result
                    gc.enable()
                    gc.collect()
            except Exception as exc:
                print(f"Параллельная обработка title_mark не удалась, fallback в sequential: {exc}")
                append_timing_log(result_dir, f"step4::parallel_by_title_mark_fallback: {time.perf_counter() - tm_loop_start:.3f}s")
                result_rows = list(orphan_rows)
                agg_mto_tagged_sums.clear()
                agg_mto_tagged_values_by_code.clear()
                agg_vo_tagged_sums.clear()
                agg_vo_codes.clear()
                agg_mto_codes_with_tags.clear()
                all_mto_count_mismatches.clear()
                all_mto_duplicate_tags.clear()
                all_vo_duplicate_tags.clear()
                gc.disable()
                total_tms = len(title_marks)
                with tqdm(total=total_tms, desc="Step4 TM", unit="tm", ncols=80, dynamic_ncols=False) as pbar:
                    for idx, title_mark in enumerate(title_marks, 1):
                        tm_step_start = time.perf_counter()
                        tm_result = process_title_mark(
                            title_mark,
                            rfp_rows_by_title.get(title_mark, []),
                            mto_data.get(title_mark, []),
                            vo_data.get(title_mark, []),
                            debug_log=debug_log_lines,
                            **tm_common_kwargs,
                        )
                        tm_elapsed = time.perf_counter() - tm_step_start
                        append_timing_log(result_dir, f"step4::tm::{title_mark}: {tm_elapsed:.3f}s")
                        pbar.update(1)
                        _merge_tm_result(tm_result, idx)
                        del tm_result
                gc.enable()
                gc.collect()
        else:
            print(f"Title_mark: {len(title_marks)} | Последовательно")
            gc.disable()
            total_tms = len(title_marks)
            with tqdm(total=total_tms, desc="Step4 TM", unit="tm", ncols=80, dynamic_ncols=False) as pbar:
                for idx, title_mark in enumerate(title_marks, 1):
                    tm_step_start = time.perf_counter()
                    tm_result = process_title_mark(
                        title_mark,
                        rfp_rows_by_title.get(title_mark, []),
                        mto_data.get(title_mark, []),
                        vo_data.get(title_mark, []),
                        debug_log=debug_log_lines,
                        **tm_common_kwargs,
                    )
                    tm_elapsed = time.perf_counter() - tm_step_start
                    append_timing_log(result_dir, f"step4::tm::{title_mark}: {tm_elapsed:.3f}s")
                    pbar.update(1)
                    _merge_tm_result(tm_result, idx)
                    del tm_result
            gc.enable()
            gc.collect()

        log_timing("tm_loop_total", tm_loop_start)

        if all_mto_count_mismatches or all_mto_duplicate_tags:
            step_start = time.perf_counter()
            save_mto_diagnostics_reports(all_mto_count_mismatches, all_mto_duplicate_tags, result_dir)
            log_timing("save_mto_diagnostics_reports", step_start)
        if all_vo_duplicate_tags:
            step_start = time.perf_counter()
            save_vo_diagnostics_reports(all_vo_duplicate_tags, result_dir)
            log_timing("save_vo_diagnostics_reports", step_start)

        step_start = time.perf_counter()
        result_rows = _restore_original_row_order(result_rows)
        log_timing("restore_original_row_order", step_start)

        append_memory_log(result_dir, "step4_after_tm_loop")
        log_data_structures_memory(
            result_dir,
            "step4_after_tm_loop",
            result_rows=result_rows,
        )

        # Per-row post-merge ops now run inside workers (step4_postmerge_ops).
        # Only global operations remain in the main process.

        step_start = time.perf_counter()
        if use_rfp_code_ban:
            _save_code_ban_list(
                result_rows,
                result_dir,
                code_ban_file,
                agg_vo_codes,
                agg_mto_codes_with_tags,
                replacement_table,
            )
            log_timing("save_code_ban_list", step_start)

        if filter_to_mto_titles and mto_data:
            mto_title_set = set(mto_data.keys())
            before = len(result_rows)
            result_rows[:] = [r for r in result_rows if r.get_value(DS_TITLE) in mto_title_set]
            print(f"Фильтрация по MTO title_system: {before} -> {len(result_rows)} строк ({len(mto_title_set)} TS)")

        emit_milestone("step4_match", "Done", f"rows={len(result_rows)}")
        emit_milestone("global_checks", "Running")
        try:
            step_start = time.perf_counter()
            _detect_and_save_effective_tag_duplicates(result_rows, result_dir)
            log_timing("detect_effective_tag_duplicates", step_start)

            apply_post_split_input_to_snapshot(
                quantity_input_snapshot,
                post_split_mto_by_title,
                post_split_mto_rows_by_title,
                post_split_vo_by_title,
                post_split_vo_rows_by_title,
            )
            emit_milestone(
                "global_checks",
                "Running",
                format_input_rfp_mto_vo_detail(quantity_input_snapshot),
            )

            balance_report = run_output_quantity_balance_check(
                quantity_input_snapshot,
                result_rows,
                all_mto_count_mismatches,
                result_dir,
            )
            if balance_report.fatal:
                emit_milestone("global_checks", "Error", balance_report.summary)
                balance_report.emit_fatal()

            step_start = time.perf_counter()
            rfp_sums_for_check = rfp_sums_by_title
            rfp_counts_for_check = rfp_counts_by_title
            if filter_to_mto_titles and mto_data:
                _mto_keys = set(mto_data.keys())
                rfp_sums_for_check = {
                    k: v for k, v in rfp_sums_by_title.items() if k in _mto_keys
                }
                rfp_counts_for_check = {
                    k: v for k, v in rfp_counts_by_title.items() if k in _mto_keys
                }
            check_values_sum(
                rfp_sums_for_check,
                rfp_counts_for_check,
                agg_mto_tagged_sums,
                dict(agg_mto_tagged_values_by_code),
                agg_vo_tagged_sums,
                result_rows,
                result_dir,
            )
            log_timing("check_values_sum", step_start)
        except QuantityBalanceFatalError as exc:
            emit_milestone(
                "global_checks",
                "Error",
                exc.summary or str(exc).splitlines()[0],
            )
            raise
        except Exception as exc:
            emit_milestone("global_checks", "Error", str(exc))
            raise
        emit_milestone("global_checks", "Done", balance_report.summary)

        packing_audit = None
        if include_packing_lists:
            emit_milestone("packing_lists", "Running")
            step_start = time.perf_counter()
            try:
                resolved_packing_dataset = _resolve_packing_dataset_for_step4(
                    include_packing_lists=include_packing_lists,
                    packing_dataset=packing_dataset,
                )
                if resolved_packing_dataset is None:
                    raise RuntimeError("Packing dataset is required when include_packing_lists=True")
                if filter_to_mto_titles and mto_data and resolved_packing_dataset.rows:
                    allowed_title_systems = set()
                    for title_mark in mto_data:
                        parsed = parse_composite_title(title_mark)
                        if parsed is not None:
                            allowed_title_systems.add(
                                packing_match_key(*parsed, "")[:2]
                            )
                    resolved_packing_dataset.rows = [
                        row
                        for row in resolved_packing_dataset.rows
                        if packing_row_key(row)[:2] in allowed_title_systems
                    ]
                ul_in_qty, ul_in_rows, ul_in_by_title = packing_queue_input_by_title(
                    resolved_packing_dataset
                )
                print_input_ul_quantity_balance(
                    available=resolved_packing_dataset.available,
                    qty=ul_in_qty,
                    rows=ul_in_rows,
                )
                emit_milestone(
                    "packing_lists",
                    "Running",
                    format_input_ul_detail(
                        available=resolved_packing_dataset.available,
                        qty=ul_in_qty,
                        rows=ul_in_rows,
                    ),
                )
                try:
                    result_rows, packing_audit = compare_rfp_rows_with_packing(
                        result_rows,
                        packing_ordered_snapshot,
                        resolved_packing_dataset,
                        use_mto_tags=ul_match_use_mto_tags,
                    )
                except RfpPackingFatalError as exc:
                    try:
                        save_rfp_packing_report(exc.audit, result_dir)
                    except OSError as report_exc:
                        print(
                            "УЛ Step4 ERROR: не удалось записать фатальный audit-report: "
                            f"{report_exc}"
                        )
                    print(exc.audit.format_short())
                    raise
                for phase_name, phase_seconds in packing_audit.phase_timings.items():
                    append_timing_log(
                        result_dir, f"step4::ul::{phase_name}: {phase_seconds:.3f}s"
                    )
                gem_n = apply_gem_supply_status(result_rows, gem_supply_codes)
                if gem_n:
                    print(f"Поставка ГЭМ: {gem_n} строк")
                zip_n = apply_zip_smr_pnr_status(result_rows)
                if zip_n:
                    print(f"ЗИП для СМР, ПНР: {zip_n} строк")
                try:
                    save_rfp_packing_report(packing_audit, result_dir)
                except OSError as report_exc:
                    print(
                        "УЛ Step4 WARNING: не удалось записать audit-report; "
                        f"Excel будет сохранён: {report_exc}"
                    )
                print(packing_audit.format_short())
                ul_balance_report = run_output_ul_quantity_balance_check(
                    ul_in_qty,
                    ul_in_by_title,
                    result_rows,
                    result_dir,
                    available=resolved_packing_dataset.available,
                )
                if ul_balance_report.fatal:
                    emit_milestone("packing_lists", "Error", ul_balance_report.summary)
                    ul_balance_report.emit_fatal()
                log_timing("compare_packing_lists", step_start)
            except QuantityBalanceFatalError as exc:
                emit_milestone(
                    "packing_lists",
                    "Error",
                    exc.summary or str(exc).splitlines()[0],
                )
                raise
            except Exception as exc:
                emit_milestone("packing_lists", "Error", str(exc))
                raise
            emit_milestone("packing_lists", "Done", ul_balance_report.summary)
        else:
            emit_milestone("packing_lists", "Skipped", "include_packing_lists=false")

        apply_mto_ul_quantity_diff(result_rows)

        step_start = time.perf_counter()
        apply_ds_source_columns(
            result_rows,
            matrix=ds_manager_matrix,
            rfp_paths_by_key=rfp_paths_by_key or {},
            mto_paths_by_title=mto_paths_by_title,
        )
        log_timing("apply_ds_source_columns", step_start)
        paint_gem_supply_rows(result_rows)

        emit_milestone("save_excel", "Running")
        step_start = time.perf_counter()
        try:
            excel_path = save_match_result_to_excel(
                result_rows,
                result_dir,
                debug_log=debug_log_lines,
                packing_audit=packing_audit,
            )
        except Exception as exc:
            emit_milestone("save_excel", "Error", str(exc))
            raise
        emit_milestone("save_excel", "Done", str(excel_path or ""))
        log_timing("save_match_result_to_excel", step_start)

        if not export_bcc_accum_matrix:
            emit_milestone("bcc_accum_matrix", "Skipped", "export_bcc_accum_matrix=false")
        elif bcc_accum_matrix is None:
            emit_milestone("bcc_accum_matrix", "Skipped", "снимок матрицы отсутствует")
        else:
            n_rows = len(bcc_accum_matrix.rows)
            emit_milestone(
                "bcc_accum_matrix",
                "Running",
                f"строк={n_rows}",
            )
            matrix_start = time.perf_counter()
            try:
                matrix_path = save_bcc_accum_matrix_to_excel(
                    bcc_accum_matrix,
                    result_dir,
                )
            except Exception as exc:
                emit_milestone("bcc_accum_matrix", "Error", str(exc))
                print(f"Накопительная матрица BCC: ошибка записи, основной Шаг4 сохранён: {exc}")
                import traceback
                traceback.print_exc()
            else:
                emit_milestone("bcc_accum_matrix", "Done", str(matrix_path or ""))
                log_timing("save_bcc_accum_matrix_to_excel", matrix_start)

        append_timing_log(result_dir, f"step4_total: {time.perf_counter() - total_start:.3f}s")

        if unified_debug and result_dir and debug_log_lines:
            # Сводка по строкам с несколькими шкафами в IN_CABINET (формат "A/B")
            multi_cabinet_rows = []
            for row in result_rows:
                if row.row_type != RowType.position_row:
                    continue
                in_cabinet = str(row.get_value(IN_CABINET) or "").strip()
                if "/" in in_cabinet:
                    multi_cabinet_rows.append({
                        "ds_title": row.get_value(DS_TITLE) or "",
                        "code": row.get_value(CODE) or "",
                        "code_mto": row.el[CODE_MTO].value if row.el.get(CODE_MTO) else "",
                        "code_vo": row.el[CODE_VO].value if row.el.get(CODE_VO) else "",
                        "tag_mto": row.el[TAG_MTO].value if row.el.get(TAG_MTO) else "",
                        "tag_vo": row.el[TAG_VO].value if row.el.get(TAG_VO) else "",
                        "in_cabinet": in_cabinet,
                        "match_status": row.el[MATCH_STATUS].value if row.el.get(MATCH_STATUS) else "",
                        "match_status_vo": row.el[MATCH_STATUS_VO].value if row.el.get(MATCH_STATUS_VO) else "",
                    })
            if multi_cabinet_rows:
                print(f"\n  Внимание: найдено {len(multi_cabinet_rows)} строк с несколькими шкафами в IN_CABINET (формат A/B)")
                debug_log_lines.append("\n")
                debug_log_lines.append("=" * 80 + "\n")
                debug_log_lines.append("СВОДКА: Строки с несколькими шкафами в IN_CABINET (возможное сопоставление из разных шкафов)\n")
                debug_log_lines.append("=" * 80 + "\n")
                for i, r in enumerate(multi_cabinet_rows, 1):
                    debug_log_lines.append(
                        f"  [{i}] DS_TITLE={r['ds_title']}, CODE={r['code']}, "
                        f"CODE_MTO={r['code_mto']}, CODE_VO={r['code_vo']}, "
                        f"TAG_MTO={r['tag_mto']}, TAG_VO={r['tag_vo']}\n"
                    )
                    debug_log_lines.append(
                        f"      IN_CABINET={r['in_cabinet']}, "
                        f"MATCH_STATUS={r['match_status']}, MATCH_STATUS_VO={r['match_status_vo']}\n"
                    )
                debug_log_lines.append("\n")
            try:
                os.makedirs(result_dir, exist_ok=True)
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                file_name = f"step4_unified_debug_{timestamp}.txt"
                file_path = os.path.join(result_dir, file_name)
                with open(file_path, "w", encoding="utf-8") as file:
                    file.writelines(debug_log_lines)
                print(f"Debug log -> {os.path.basename(file_path)}")
            except Exception as exc:
                print(f"Ошибка сохранения unified debug-лога: {exc}")
    
        return result_rows






def _extract_rfp_snapshot_for_check(
    rfp_data: List[RowStd],
    skip_split: bool = False,
    units_ban: List[str] | None = None,
    code_ban: set[str] | None = None,
) -> Tuple[Dict[str, float], Dict[str, int]]:
    """
    Извлекает лёгкий снапшот полей RFP для check_values_sum (до модификации).
    Возвращает (rfp_sums_by_title, rfp_counts_by_title).

    Когда skip_split=True, RFP ещё не раскрыт — нужно эмулировать эффект
    split_rfp_rows на VALUES, чтобы снимок совпадал с результатом.
    Логика повторяет split_rfp_rows: проверяет units_ban и code_ban
    перед применением int()-усечения.
    """
    units_ban_set = {str(u).strip().lower() for u in (units_ban or [])}
    code_ban_set = code_ban or set()

    sums_by_title: Dict[str, float] = defaultdict(float)
    counts_by_title: Dict[str, int] = defaultdict(int)
    for row in rfp_data:
        if row.row_type != RowType.position_row:
            continue
        title_mark = row.get_value(DS_TITLE)
        if not title_mark:
            continue
        values = row.get_value(VALUES)
        try:
            values_float = float(values) if values else 0.0
        except (ValueError, TypeError):
            values_float = 0.0

        if skip_split:
            values_int = int(values_float)
            tags_list = row.get_tags_list()
            units_value = row.get_value(UNITS) or ""
            units_norm = str(units_value).strip().lower()
            is_unit_banned = units_norm in units_ban_set
            code_value = row.get_value(CODE)
            is_code_banned = _normalize_code_for_ban(code_value) in code_ban_set

            if tags_list and len(tags_list) > 1 and not is_unit_banned:
                values_float = float(min(values_int, len(tags_list)))
            elif not tags_list and values_int > 1 and not is_unit_banned and not is_code_banned:
                values_float = float(values_int)

        sums_by_title[title_mark] += values_float
        counts_by_title[title_mark] += 1
    return (dict(sums_by_title), dict(counts_by_title))






def _collapse_rfp_rows_before_export(
    result_rows: List[RowStd],
    result_dir: str = "",
    collapse_debug_enabled: bool = False,
    debug_collapse_ds_title: str = "8950-SOT4",
    debug_collapse_ds_number: str = "465",
) -> List[RowStd]:
    """
    Схлопывание строк перед экспортом.

    Случай 1: строки с DS_NUMBER и без TAGS (RFP тегов):
      Группируем по (DS_TITLE, DS_NUMBER, CODE, CODE_MTO, CODE_VO, IN_CABINET,
      normalize(UNITS), normalize(UNITS_MTO), is_excluded_from_supply).
      Внутри одного DS_NUMBER строки с разными комбинациями кодов образуют
      отдельные под-группы, каждая схлопывается независимо.
      IN_CABINET: при непустом значении позиции из разных шкафов не схлопываются.
      1.а CODE_MTO и CODE_VO пусты → подходит
      1.б CODE_MTO и/или CODE_VO не пусты, TAG_MTO и TAG_VO пустые → подходит
      (если TAG_MTO или TAG_VO не пусты при наличии кодов → строка не подходит)

    Случай 2: строки без DS_NUMBER, TAG_MTO и TAG_VO пустые (новые из MTO/VO):
      Группируем по (DS_TITLE, CODE_MTO, CODE_VO, IN_CABINET,
      normalize(UNITS), normalize(UNITS_MTO), is_excluded_from_supply) — хотя бы один код должен быть.
      IN_CABINET: при непустом значении позиции из разных шкафов не схлопываются.

    Суммирование: VALUES по CODE, VALUES_MTO по CODE_MTO, VALUES_VO по CODE_VO.
    DS_TITLE учитывается в начале формирования словарей.
    """
    # ── DEBUG: инициализация ──
    dbg_lines: List[str] = []
    dbg_title = debug_collapse_ds_title.strip() if debug_collapse_ds_title else ""
    dbg_number = debug_collapse_ds_number.strip() if debug_collapse_ds_number else ""

    def _dbg(msg: str) -> None:
        if collapse_debug_enabled:
            dbg_lines.append(msg + "\n")

    def _is_dbg_row(title_val: str, number_val: str) -> bool:
        """Проверка: строка попадает под debug-фильтр (оба поля должны совпадать)."""
        if not dbg_title or not dbg_number:
            return False
        return title_val.strip() == dbg_title and number_val.strip() == dbg_number

    def _row_debug_str(row: RowStd, idx: int) -> str:
        """Форматирование полной информации по строке для debug."""
        rfp_tags = row.get_tags_list()
        rfp_tags_str = "/".join(rfp_tags) if rfp_tags else "<пусто>"
        tag_mto_v = row.el[TAG_MTO].value if row.el.get(TAG_MTO) else ""
        tag_vo_v = row.el[TAG_VO].value if row.el.get(TAG_VO) else ""
        code_v = row.get_value(CODE) or ""
        code_mto_v = row.el[CODE_MTO].value if row.el.get(CODE_MTO) else ""
        code_vo_v = row.el[CODE_VO].value if row.el.get(CODE_VO) else ""
        values_v = row.get_value(VALUES)
        values_mto_v = row.get_value(VALUES_MTO)
        values_vo_v = row.get_value(VALUES_VO)
        ds_name_v = row.get_value(DS_NAME) or ""
        in_cabinet_v = (row.el[IN_CABINET].value if row.el.get(IN_CABINET) else "") or ""
        return (
            f"  [idx={idx}] DS_TITLE='{row.get_value(DS_TITLE)}' DS_NUMBER='{row.get_value(DS_NUMBER)}' "
            f"DS_NAME='{ds_name_v}' IN_CABINET='{in_cabinet_v}'\n"
            f"    TAGS='{rfp_tags_str}' TAG_MTO='{tag_mto_v}' TAG_VO='{tag_vo_v}'\n"
            f"    CODE='{code_v}' CODE_MTO='{code_mto_v}' CODE_VO='{code_vo_v}'\n"
            f"    VALUES={values_v} VALUES_MTO={values_mto_v} VALUES_VO={values_vo_v}\n"
            f"    row_type={row.row_type}"
        )

    _dbg(f"{'=' * 80}")
    _dbg(f"COLLAPSE DEBUG  |  фильтр: DS_TITLE='{dbg_title}' AND DS_NUMBER='{dbg_number}'")
    _dbg(f"Всего строк на входе: {len(result_rows)}")
    _dbg(f"{'=' * 80}")

    # Проверка: все position_row должны иметь DS_TITLE
    _validate_ds_title_present(result_rows)

    # ── Шаг 1: Группировка строк ──
    _dbg(f"\n{'─' * 60}")
    _dbg("ШАГ 1: Классификация строк и формирование групп")
    _dbg(f"{'─' * 60}")

    # Case 1: по (DS_TITLE, DS_NUMBER, CODE, CODE_MTO, CODE_VO) — под-группы по кодам
    groups_by_number: Dict[tuple, List[int]] = {}
    # Case 2: по (DS_TITLE, CODE_MTO, CODE_VO)
    groups_by_codes: Dict[tuple, List[int]] = {}

    for idx, row in enumerate(result_rows):
        if row.row_type != RowType.position_row:
            continue

        title_system = _normalize_group_value(row.get_value(DS_TITLE))
        if not title_system:
            continue

        position_number = _normalize_group_value(row.get_value(DS_NUMBER))
        in_cabinet_raw = (row.el[IN_CABINET].value if row.el.get(IN_CABINET) else "") or ""
        in_cabinet_key = _normalize_in_cabinet_for_grouping(in_cabinet_raw)
        rfp_tags = row.get_tags_list()
        tag_mto = (row.el[TAG_MTO].value if row.el.get(TAG_MTO) else "") or ""
        tag_vo = (row.el[TAG_VO].value if row.el.get(TAG_VO) else "") or ""
        code_rfp = _normalize_collapse_code(row.get_value(CODE))
        code_mto = _normalize_collapse_code(row.el[CODE_MTO].value if row.el.get(CODE_MTO) else "")
        code_vo = _normalize_collapse_code(row.el[CODE_VO].value if row.el.get(CODE_VO) else "")
        units_rfp = normalize_units_text(row.get_value(UNITS))
        units_mto = normalize_units_text(row.get_value(UNITS_MTO))
        excluded_key = is_excluded_from_supply(row)

        is_dbg = _is_dbg_row(title_system, position_number)

        if is_dbg:
            _dbg(f"\n>>> DEBUG строка idx={idx}:")
            _dbg(_row_debug_str(row, idx))

        # Случай 1: строки с DS_NUMBER и без TAGS (RFP)
        if position_number and not rfp_tags:
            # 1.а: CODE_MTO и CODE_VO пусты → подходит для схлопывания
            if not code_mto and not code_vo:
                grp_key = (
                    title_system, position_number, code_rfp, "", "",
                    in_cabinet_key, units_rfp, units_mto, excluded_key,
                )
                groups_by_number.setdefault(grp_key, []).append(idx)
                if is_dbg:
                    _dbg(f"    -> РЕШЕНИЕ: Случай 1.а (DS_NUMBER есть, TAGS нет, CODE_MTO пусто, CODE_VO пусто, IN_CABINET='{in_cabinet_key}')")
                    _dbg(f"       Ключ под-группы: {grp_key}")
                continue
            # 1.б: CODE_MTO и/или CODE_VO не пустые, TAG_MTO и TAG_VO должны быть пустые
            if not tag_mto and not tag_vo:
                grp_key = (
                    title_system, position_number, code_rfp, code_mto, code_vo,
                    in_cabinet_key, units_rfp, units_mto, excluded_key,
                )
                groups_by_number.setdefault(grp_key, []).append(idx)
                if is_dbg:
                    _dbg(f"    -> РЕШЕНИЕ: Случай 1.б (DS_NUMBER есть, TAGS нет, "
                         f"CODE_MTO='{code_mto}', CODE_VO='{code_vo}', IN_CABINET='{in_cabinet_key}', TAG_MTO пусто, TAG_VO пусто)")
                    _dbg(f"       Ключ под-группы: {grp_key}")
                continue
            # Иначе (есть коды MTO/VO с тегами) → строка не подходит для схлопывания
            if is_dbg:
                _dbg(f"    -> РЕШЕНИЕ: ОТКЛОНЕНА (DS_NUMBER есть, TAGS нет, "
                     f"но TAG_MTO='{tag_mto}' и/или TAG_VO='{tag_vo}' не пустые при наличии кодов)")
            continue

        # Строки с TAGS → не подходят для Case 1
        if position_number and rfp_tags:
            if is_dbg:
                rfp_tags_str = "/".join(rfp_tags)
                _dbg(f"    -> РЕШЕНИЕ: ОТКЛОНЕНА (DS_NUMBER есть, но TAGS не пустые: '{rfp_tags_str}')")
            continue

        # Случай 2: строки без DS_NUMBER, TAG_MTO и TAG_VO пустые
        if not position_number and not tag_mto and not tag_vo:
            if not code_mto and not code_vo:
                if is_dbg:
                    _dbg(f"    -> РЕШЕНИЕ: ОТКЛОНЕНА Случай 2 (DS_NUMBER нет, TAG_MTO/TAG_VO пусты, "
                         f"но CODE_MTO и CODE_VO тоже пусты — нечего группировать)")
                continue
            groups_by_codes.setdefault(
                (title_system, code_mto, code_vo, in_cabinet_key, units_rfp, units_mto, excluded_key),
                [],
            ).append(idx)
            if is_dbg:
                _dbg(f"    -> РЕШЕНИЕ: Случай 2 (DS_NUMBER нет, TAG_MTO пусто, TAG_VO пусто, "
                     f"CODE_MTO='{code_mto}', CODE_VO='{code_vo}', IN_CABINET='{in_cabinet_key}')")
                _dbg(f"       Добавлена в группу groups_by_codes[({title_system}, {code_mto}, {code_vo}, '{in_cabinet_key}')]")
            continue

        # Не попала ни в один случай
        if is_dbg:
            _dbg(f"    -> РЕШЕНИЕ: ОТКЛОНЕНА (не подходит ни под один случай)")

    # ── Шаг 2: Вывод сформированных под-групп для debug ──
    _dbg(f"\n{'─' * 60}")
    _dbg("ШАГ 2: Сформированные под-группы (debug-фильтр)")
    _dbg(f"{'─' * 60}")

    dbg_subgroups_found = 0
    for key, indices in groups_by_number.items():
        # key = (DS_TITLE, DS_NUMBER, CODE, CODE_MTO, CODE_VO, IN_CABINET)
        if key[0].strip() == dbg_title and key[1].strip() == dbg_number:
            dbg_subgroups_found += 1
            _dbg(f"\n  Под-группа {dbg_subgroups_found}: key={key}, {len(indices)} строк, indices={indices}")
            for i in indices:
                _dbg(_row_debug_str(result_rows[i], i))

    if dbg_subgroups_found == 0:
        _dbg(f"\n  groups_by_number: НЕТ под-групп для DS_TITLE='{dbg_title}', DS_NUMBER='{dbg_number}'")
        related_keys = [k for k in groups_by_number.keys() if k[0].strip() == dbg_title]
        if related_keys:
            _dbg(f"  (Все ключи с DS_TITLE='{dbg_title}': {len(related_keys)} шт.)")
            for rk in related_keys[:20]:
                _dbg(f"    {rk}: {len(groups_by_number[rk])} строк")
        else:
            _dbg(f"  (Нет ни одной группы с DS_TITLE='{dbg_title}')")
    else:
        _dbg(f"\n  Итого под-групп для DS_NUMBER='{dbg_number}': {dbg_subgroups_found}")

    _dbg(f"\n  Всего groups_by_number: {len(groups_by_number)} под-групп")
    _dbg(f"  Всего groups_by_codes: {len(groups_by_codes)} групп")

    # ── Шаг 3: Формирование списка групп для схлопывания ──
    _dbg(f"\n{'─' * 60}")
    _dbg("ШАГ 3: Отбор групп для схлопывания (>1 строки)")
    _dbg(f"{'─' * 60}")

    collapse_groups: List[List[int]] = []

    for key, indices in groups_by_number.items():
        is_dbg_group = (key[0].strip() == dbg_title and key[1].strip() == dbg_number)

        if len(indices) <= 1:
            if is_dbg_group:
                _dbg(f"\n  groups_by_number[{key}]: пропуск ({len(indices)} строка, нечего схлопывать)")
            continue

        # Консистентность кодов гарантирована ключом группировки — схлопываем
        collapse_groups.append(indices)
        if is_dbg_group:
            _dbg(f"\n  groups_by_number[{key}]: {len(indices)} строк -> СХЛОПЫВАЕМ")

    for key, indices in groups_by_codes.items():
        is_dbg_group = (key[0].strip() == dbg_title)

        if len(indices) <= 1:
            if is_dbg_group:
                _dbg(f"\n  groups_by_codes[{key}]: пропуск ({len(indices)} строка)")
            continue

        # Консистентность кодов гарантирована ключом группировки — схлопываем
        collapse_groups.append(indices)
        if is_dbg_group:
            _dbg(f"\n  groups_by_codes[{key}]: {len(indices)} строк -> СХЛОПЫВАЕМ")

    _dbg(f"\n  Всего групп для схлопывания: {len(collapse_groups)}")

    if not collapse_groups:
        _dbg(f"\n  Нет групп для схлопывания — возвращаем исходные строки без изменений.")
        if collapse_debug_enabled:
            _save_collapse_debug_log(dbg_lines, result_dir)
        return result_rows

    # ── Шаг 4: Схлопывание ──
    _dbg(f"\n{'─' * 60}")
    _dbg("ШАГ 4: Выполнение схлопывания (суммирование VALUES)")
    _dbg(f"{'─' * 60}")

    # Строим индекс: idx → group_id, и определяем первую строку каждой группы
    index_to_group: Dict[int, int] = {}
    group_first_idx: Dict[int, int] = {}
    for group_id, indices in enumerate(collapse_groups):
        first_idx = min(indices)
        group_first_idx[group_id] = first_idx
        for idx in indices:
            index_to_group[idx] = group_id

    # Формируем итоговый список, суммируя VALUES
    collapsed_rows: List[RowStd] = []
    for idx, row in enumerate(result_rows):
        group_id = index_to_group.get(idx)
        if group_id is None:
            collapsed_rows.append(row)
            continue
        if idx != group_first_idx[group_id]:
            continue  # пропускаем не-первые строки группы

        group_indices = collapse_groups[group_id]
        title_val = _normalize_group_value(row.get_value(DS_TITLE))
        number_val = _normalize_group_value(row.get_value(DS_NUMBER))
        is_dbg_group = _is_dbg_row(title_val, number_val)

        # Суммируем значения по всем строкам группы
        total_values = 0.0
        total_values_mto = 0.0
        total_values_vo = 0.0
        has_code = False
        has_code_mto = False
        has_code_vo = False

        if is_dbg_group:
            code_rfp_val = _normalize_collapse_code(row.get_value(CODE))
            code_mto_val = _normalize_collapse_code(row.el[CODE_MTO].value if row.el.get(CODE_MTO) else "")
            code_vo_val = _normalize_collapse_code(row.el[CODE_VO].value if row.el.get(CODE_VO) else "")
            _dbg(f"\n  Схлопывание под-группы: DS_TITLE='{title_val}', DS_NUMBER='{number_val}', "
                 f"CODE='{code_rfp_val}', CODE_MTO='{code_mto_val}', CODE_VO='{code_vo_val}'")
            _dbg(f"  Первая строка (сохраняется): idx={idx}")
            _dbg(f"  Все индексы под-группы: {group_indices}")

        for group_idx in group_indices:
            group_row = result_rows[group_idx]
            row_code = _normalize_collapse_code(group_row.get_value(CODE))
            row_code_mto = _normalize_collapse_code(group_row.el[CODE_MTO].value if group_row.el.get(CODE_MTO) else "")
            row_code_vo = _normalize_collapse_code(group_row.el[CODE_VO].value if group_row.el.get(CODE_VO) else "")
            row_val = _safe_float(group_row.get_value(VALUES))
            row_val_mto = _safe_float(group_row.get_value(VALUES_MTO))
            row_val_vo = _safe_float(group_row.get_value(VALUES_VO))

            if row_code:
                has_code = True
                total_values += row_val
            if row_code_mto:
                has_code_mto = True
                total_values_mto += row_val_mto
            if row_code_vo:
                has_code_vo = True
                total_values_vo += row_val_vo

            if is_dbg_group:
                _dbg(f"    idx={group_idx}: CODE='{row_code}'(val={row_val}) "
                     f"CODE_MTO='{row_code_mto}'(val_mto={row_val_mto}) "
                     f"CODE_VO='{row_code_vo}'(val_vo={row_val_vo})")

        if has_code:
            row.el[VALUES].value = total_values
        if has_code_mto:
            row.el[VALUES_MTO].value = total_values_mto
        if has_code_vo:
            row.el[VALUES_VO].value = total_values_vo

        merge_units_status_trace(
            row,
            *(result_rows[group_idx] for group_idx in group_indices),
        )

        if is_dbg_group:
            _dbg(f"  ИТОГО: has_code={has_code} total_values={total_values} | "
                 f"has_code_mto={has_code_mto} total_values_mto={total_values_mto} | "
                 f"has_code_vo={has_code_vo} total_values_vo={total_values_vo}")

        collapsed_rows.append(row)

    _dbg(f"\n{'─' * 60}")
    _dbg(f"ИТОГ: строк до={len(result_rows)}, строк после={len(collapsed_rows)}, "
         f"схлопнуто групп={len(collapse_groups)}")
    _dbg(f"{'─' * 60}")

    if collapse_debug_enabled:
        _save_collapse_debug_log(dbg_lines, result_dir)
    return collapsed_rows


def _save_collapse_debug_log(dbg_lines: List[str], result_dir: str) -> None:
    """Сохраняет debug-лог схлопывания в файл."""
    if not result_dir or not dbg_lines:
        return
    try:
        ensure_result_dir_exists(result_dir)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        file_path = os.path.join(result_dir, f"step4_collapse_debug_{timestamp}.txt")
        with open(file_path, "w", encoding="utf-8") as f:
            f.writelines(dbg_lines)
        print(f"Collapse debug-лог сохранен: {file_path}")
    except Exception as exc:
        print(f"Ошибка сохранения collapse debug-лога: {exc}")


def _validate_ds_title_present(result_rows: List[RowStd]) -> None:
    """Проверяет что все position_row имеют DS_TITLE, иначе останавливает выполнение."""
    missing_title_rows: List[RowStd] = []
    for row in result_rows:
        if row.row_type != RowType.position_row:
            continue
        title_system = _normalize_group_value(row.get_value(DS_TITLE))
        if not title_system:
            missing_title_rows.append(row)

    if missing_title_rows:
        print("Ошибка: найдены строки без DS_TITLE. Остановка выполнения.")
        for row in missing_title_rows:
            rfp_tags = row.get_tags_list()
            rfp_tags_str = "/".join(rfp_tags) if rfp_tags else ""
            print(
                f"DS_TITLE='', DS_NUMBER='{row.get_value(DS_NUMBER)}', "
                f"DS_NAME='{row.get_value(DS_NAME)}', CODE='{row.get_value(CODE)}', "
                f"CODE_MTO='{row.get_value(CODE_MTO)}', CODE_VO='{row.get_value(CODE_VO)}', "
                f"TAGS='{rfp_tags_str}', TAG_MTO='{row.el[TAG_MTO].value if row.el.get(TAG_MTO) else ''}', "
                f"TAG_VO='{row.el[TAG_VO].value if row.el.get(TAG_VO) else ''}'"
            )
        sys.exit(1)


def _normalize_collapse_code(value) -> str:
    if value is None:
        return ""
    return str(value).strip().upper()


def _normalize_group_value(value) -> str:
    if value is None:
        return ""
    return str(value).strip()


def _normalize_in_cabinet_for_grouping(value) -> str:
    """
    Нормализует IN_CABINET для ключа группировки.
    Пустое значение → "".
    Непустое: разбиваем по "/", сортируем части, склеиваем — чтобы "A/B" и "B/A"
    давали один ключ (одинаковый набор шкафов).
    """
    if value is None:
        return ""
    s = str(value).strip()
    if not s:
        return ""
    parts = [p.strip() for p in s.split("/") if p.strip()]
    if not parts:
        return ""
    return "/".join(sorted(parts))


def _safe_float(value) -> float:
    try:
        return float(value) if value not in (None, "") else 0.0
    except (ValueError, TypeError):
        return 0.0


def _detect_and_save_effective_tag_duplicates(result_rows: List[RowStd], result_dir: str) -> None:
    """Phase 2: global duplicate detection for TAG_EFFECTIVE.
    Phase 1 (per-row assignment) runs inside workers via step4_postmerge_ops.
    """
    # После того как все TAG_EFFECTIVE выставлены — ищем дубликаты
    tag_to_rows: Dict[str, List[RowStd]] = {}
    for row in result_rows:
        if row.row_type != RowType.position_row:
            continue
        effective_tag = row.el[TAG_EFFECTIVE].value if row.el.get(TAG_EFFECTIVE) else ""
        if not effective_tag:
            continue
        tag_str = str(effective_tag).strip()
        if not tag_str:
            continue
        tag_to_rows.setdefault(tag_str, []).append(row)

    duplicate_tags = {tag: rows for tag, rows in tag_to_rows.items() if len(rows) > 1}
    if not duplicate_tags or not result_dir:
        return

    # 3. Выгружаем дубликаты TAG_EFFECTIVE в Excel
    try:
        ensure_result_dir_exists(result_dir)

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "TAG_EFFECTIVE_дубликаты"

        headers = [
            "TAG_EFFECTIVE",
            "title_system",
            "ДС номер",
            "Наименование ДС",
            "Код RFP",
            "Код VO",
            "TAG_MTO",
            "TAG_VO",
            "Теги RFP",
            "Статус позиции",
            "Статус сопоставления MTO",
            "Статус сопоставления VO",
            "Источник (VO/MTO/RFP)",
        ]
        ws.append(headers)

        header_fill = PatternFill(start_color="dbbcdb", end_color="dbbcdb", fill_type="solid")
        header_font = Font(bold=True, color="000000")
        header_alignment = Alignment(horizontal="center", vertical="center")
        border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )

        for col_num, _ in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col_num)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = header_alignment
            cell.border = border

        for tag_value, rows in sorted(duplicate_tags.items(), key=lambda x: x[0]):
            for row in rows:
                rfp_tags_list = row.get_tags_list()
                rfp_tags_str = "/".join(rfp_tags_list) if rfp_tags_list else ""

                # Источник TAG_EFFECTIVE: VO если есть TAG_VO, иначе MTO если есть TAG_MTO, иначе RFP
                vo_tag = row.el[TAG_VO].value if row.el.get(TAG_VO) else ""
                mto_tag = row.el[TAG_MTO].value if row.el.get(TAG_MTO) else ""
                source = "VO" if vo_tag else ("MTO" if mto_tag else "RFP")
                ws.append([
                    tag_value,
                    row.get_value(DS_TITLE),
                    row.get_value(DS_NUMBER),
                    row.get_value(DS_NAME),
                    row.get_value(CODE),
                    row.get_value(CODE_VO),
                    row.el[TAG_MTO].value if row.el.get(TAG_MTO) else "",
                    row.el[TAG_VO].value if row.el.get(TAG_VO) else "",
                    rfp_tags_str,
                    row.el[POSITION_STATUS].value if row.el.get(POSITION_STATUS) else "",
                    row.el[MATCH_STATUS].value if row.el.get(MATCH_STATUS) else "",
                    row.el[MATCH_STATUS_VO].value if row.el.get(MATCH_STATUS_VO) else "",
                    source,
                ])

        # Применяем стили к данным
        for row_num in range(2, ws.max_row + 1):
            for col_num in range(1, len(headers) + 1):
                cell = ws.cell(row=row_num, column=col_num)
                cell.border = border
                cell.alignment = Alignment(horizontal="left", vertical="center")

        # Настраиваем ширину колонок
        column_widths = {
            "TAG_EFFECTIVE": 23,
            "title_system": 18,
            "ДС номер": 12,
            "Наименование ДС": 50,
            "Код RFP": 18,
            "Код VO": 18,
            "TAG_MTO": 30,
            "TAG_VO": 30,
            "Теги RFP": 30,
            "Статус позиции": 25,
            "Статус сопоставления MTO": 22,
            "Статус сопоставления VO": 22,
            "Источник (VO/MTO/RFP)": 18,
        }
        for col_num, header_label in enumerate(headers, 1):
            col_letter = get_column_letter(col_num)
            ws.column_dimensions[col_letter].width = column_widths.get(header_label, 15)

        # Фильтр и закрепление заголовка
        last_col_letter = get_column_letter(len(headers))
        ws.auto_filter.ref = f'A1:{last_col_letter}{ws.max_row}'
        ws.freeze_panes = 'A2'

        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        file_name = f"Отчет по дублям тегов - TAG_EFFECTIVE_{timestamp}.xlsx"
        file_path = os.path.join(result_dir, file_name)
        wb.save(file_path)

        print(f"Дубликаты TAG_EFFECTIVE ({len(duplicate_tags)}) -> {os.path.basename(file_path)}")
    except Exception as exc:
        print(f"Ошибка при сохранении отчета по дубликатам TAG_EFFECTIVE: {exc}")






def _set_excluded_for_tags_without_codes(result_rows: List[RowStd]) -> None:
    """
    Устанавливает статус 'Исключен' для строк, у которых:
      - effective_tags (TAG_EFFECTIVE) не пустой
      - нет ни MTO кода (CODE_MTO), ни VO кода (CODE_VO)
      - статус 'Исключен' ещё не установлен ранее
    """    
    for row in result_rows:
        if row.row_type != RowType.position_row:
            continue

        effective_tag = row.el[TAG_EFFECTIVE].value if row.el.get(TAG_EFFECTIVE) else ""
        if not effective_tag:
            continue

        position_status = row.el[POSITION_STATUS].value if row.el.get(POSITION_STATUS) else ""
        if position_status == "Исключен":
            continue
        if is_excluded_from_supply(row):
            continue

        code_mto = row.get_value(CODE_MTO)
        code_vo = row.get_value(CODE_VO)

        if not code_mto and not code_vo:
            row.el[POSITION_STATUS].value = "Исключен"
            

    






def _load_lot_registry(rfp_registr_lot_path: str) -> Dict[str, str]:
    if not rfp_registr_lot_path or not os.path.exists(rfp_registr_lot_path):
        print(f"⚠️  Реестр лотов не найден: {rfp_registr_lot_path}")
        return {}
    try:
        wb = openpyxl.load_workbook(rfp_registr_lot_path, data_only=True)
        ws = wb.active
        header = [(_normalize_lot_registry_value(c.value)) for c in next(ws.iter_rows(min_row=1, max_row=1))]
        ds_idx = header.index("ДС") if "ДС" in header else None
        lot_idx = header.index("Лот") if "Лот" in header else None
        if ds_idx is None or lot_idx is None:
            print("⚠️  Реестр лотов: не найдены заголовки 'ДС' и 'Лот'")
            return {}

        lot_map: Dict[str, str] = {}
        for row in ws.iter_rows(min_row=2, values_only=True):
            ds_value = _normalize_lot_registry_value(row[ds_idx] if ds_idx < len(row) else "")
            lot_value = _normalize_lot_registry_value(row[lot_idx] if lot_idx < len(row) else "")
            if ds_value:
                lot_map[ds_value] = lot_value
        return lot_map
    except Exception as exc:
        print(f"⚠️  Ошибка чтения реестра лотов ({rfp_registr_lot_path}): {exc}")
        return {}


def _normalize_lot_registry_value(value) -> str:
    if value is None:
        return ""
    return str(value).strip()


def _save_code_ban_list(
    result_rows: List[RowStd],
    result_dir: str,
    code_ban_file: str,
    vo_codes: Set[str],
    mto_tagged_codes: Set[str],
    replacement_table: Dict[str, List[Tuple[str, str]]]):
    """
    Формирует code_ban — список кодов, для которых запрещено раскрытие строк в step1.
    Из всех кодов RFP вычитаются:
      - коды VO (должны раскрываться)
      - коды-замены для VO из replacement_table (должны раскрываться)
      - коды MTO с тегами (должны раскрываться)
      - RFP коды, сопоставленные с MTO без тегов (должны раскрываться для 1:1 match)
    """
    if not result_dir or not code_ban_file:
        return

    rfp_codes = _collect_rfp_codes(result_rows)
    vo_repl_codes = _collect_replacement_codes_for_vo(vo_codes, replacement_table)
    rfp_mto_no_tags = _collect_rfp_codes_matched_to_mto_no_tags(result_rows)

    split_allowed = vo_codes | vo_repl_codes | mto_tagged_codes | rfp_mto_no_tags
    codes = sorted(rfp_codes - split_allowed)

    # Проверка: если в текущем файле code_ban есть коды из split_allowed — обновляем и перезапускаем
    current_set = set(_load_code_ban_from_file(code_ban_file))
    stale = sorted(current_set & split_allowed)
    if stale:
        print(f"⚠️  В code_ban найдены коды, которые теперь должны раскрываться: {stale}")
        _write_code_ban_file(code_ban_file, codes)
        print(f"code_ban обновлен, требуется перезапуск: {code_ban_file}")
        sys.exit(1)

    # Сохраняем только при наличии изменений
    if set(codes) != current_set:
        _write_code_ban_file(code_ban_file, codes)






def _load_code_ban_from_file(code_ban_file: str) -> list[str]:
    if not code_ban_file or not os.path.exists(code_ban_file):
        return []
    try:
        with open(code_ban_file, "r", encoding="utf-8") as file:
            content = file.read().strip()
        if not content:
            return []
        if content.startswith("code_ban="):
            content = content.split("=", 1)[1].strip()
        values = ast.literal_eval(content)
        if isinstance(values, list):
            return [_normalize_code_for_ban(v) for v in values if _normalize_code_for_ban(v)]
    except Exception as e:
        print(f"Ошибка чтения code_ban ({code_ban_file}): {e}")
    return []


def _write_code_ban_file(file_path: str, codes: list[str]) -> None:
    try:
        with open(file_path, "w", encoding="utf-8") as file:
            file.write(f"code_ban={codes}")
        print(f"code_ban -> {os.path.basename(file_path)}")
    except Exception as e:
        print(f"Ошибка сохранения code_ban: {e}")


def _normalize_code_for_ban(code_value) -> str:
    if code_value is None:
        return ""
    return str(code_value).strip().upper()


def _collect_vo_codes(vo_data: Dict[str, List[RowStd]]) -> set[str]:
    codes = set()
    for _, vo_rows in vo_data.items():
        for vo_row in vo_rows:
            if vo_row.row_type != RowType.position_row:
                continue
            code = vo_row.get_value(CODE)
            if not code:
                continue
            code_str = _normalize_code_for_ban(code)
            if code_str:
                codes.add(code_str)
    return codes


def _collect_rfp_codes(result_rows: List[RowStd]) -> Set[str]:
    codes = set()
    for row in result_rows:
        if row.row_type != RowType.position_row:
            continue
        code = row.get_value(CODE)
        if code:
            code_str = _normalize_code_for_ban(code)
            if code_str:
                codes.add(code_str)
    return codes


def _collect_mto_codes_with_tags(mto_data: Dict[str, List[RowStd]]) -> Set[str]:
    codes = set()
    if not mto_data:
        return codes
    for _, mto_rows in mto_data.items():
        for mto_row in mto_rows:
            if mto_row.row_type != RowType.position_row:
                continue
            if not mto_row.get_tags_list():
                continue
            code = mto_row.get_value(CODE)
            if not code:
                continue
            code_str = _normalize_code_for_ban(code)
            if code_str:
                codes.add(code_str)
    return codes


def _collect_rfp_codes_matched_to_mto_no_tags(result_rows: List[RowStd]) -> Set[str]:
    """
    Собирает RFP коды из строк, сопоставленных с MTO без тегов (CODE_MTO есть, TAG_MTO пустой).
    Эти коды должны раскрываться в step1, чтобы обеспечить 1:1 match и корректное схлопывание.
    """
    codes = set()
    for row in result_rows:
        if row.row_type != RowType.position_row:
            continue
        code_rfp = row.get_value(CODE)
        if not code_rfp:
            continue
        code_mto = row.el[CODE_MTO].value if row.el.get(CODE_MTO) else ""
        tag_mto = row.el[TAG_MTO].value if row.el.get(TAG_MTO) else ""
        if not code_mto or tag_mto:
            continue
        code_str = _normalize_code_for_ban(code_rfp)
        if code_str:
            codes.add(code_str)
    return codes


def _collect_replacement_codes_for_vo(
    vo_codes: Set[str],
    replacement_table: Dict[str, List[Tuple[str, str]]]
) -> Set[str]:
    """
    Собирает все коды, связанные с VO через таблицу замен (в обе стороны):
      - Если VO содержит СТАРЫЙ код → добавляем НОВЫЙ код
      - Если VO содержит НОВЫЙ код → добавляем СТАРЫЙ код
    """
    if not vo_codes or not replacement_table:
        return set()
    replacement_codes: Set[str] = set()
    for old_code, replacements in replacement_table.items():
        old_code_norm = _normalize_code_for_ban(old_code)
        if not old_code_norm:
            continue
        for new_code, _status in replacements:
            new_code_norm = _normalize_code_for_ban(new_code)
            if not new_code_norm:
                continue
            # VO содержит старый код → добавляем новый
            if old_code_norm in vo_codes:
                replacement_codes.add(new_code_norm)
            # VO содержит новый код → добавляем старый
            if new_code_norm in vo_codes:
                replacement_codes.add(old_code_norm)
    return replacement_codes


def _compare_title_systems_rfp_mto_vo(
    rfp_data: List[RowStd],
    mto_data: Dict[str, List[RowStd]],
    vo_data: Dict[str, List[RowStd]],
    result_dir: str,
    print_table: bool = False,
    export_to_excel: bool = True,
):
    if print_table:
        print("\n" + "=" * 80)
        print("СРАВНЕНИЕ title_system: RFP, MTO, VO")
        print("=" * 80)
    
    # Собираем title_system из RFP
    rfp_title_systems: Set[str] = set()
    for row in rfp_data:
        if row.row_type != RowType.position_row:
            continue
        title_system = row.get_value(DS_TITLE)
        if title_system:
            rfp_title_systems.add(str(title_system).strip())
    
    mto_title_systems = set(mto_data.keys()) if mto_data else set()
    vo_title_systems = set(vo_data.keys()) if vo_data else set()
    
    all_title_systems = rfp_title_systems | mto_title_systems | vo_title_systems
    
    table = PrettyTable()
    table.field_names = ["title_system", "В RFP", "В MTO", "В VO", "Результат сравнения", "Статус VO"]
    table.border = True
    table.align = "l"
    
    for title_system in sorted(all_title_systems, key=lambda x: (x is None, x or "")):
        in_rfp = "✓" if title_system in rfp_title_systems else ""
        in_mto = "✓" if title_system in mto_title_systems else ""
        in_vo = "✓" if title_system in vo_title_systems else ""
        vo_status = "Есть в VO" if in_vo else "Нет в VO"
        
        if in_rfp and in_mto and in_vo:
            result = "Совпадают (RFP/MTO/VO)"
        elif in_rfp and in_mto and not in_vo:
            result = "Нет в VO"
        elif in_rfp and in_vo and not in_mto:
            result = "Нет в MTO"
        elif in_mto and in_vo and not in_rfp:
            result = "Нет в RFP"
        elif in_rfp and not in_mto and not in_vo:
            result = "Только RFP"
        elif in_mto and not in_rfp and not in_vo:
            result = "Только MTO"
        elif in_vo and not in_rfp and not in_mto:
            result = "Только VO"
        else:
            result = "Частичное совпадение"
        
        table.add_row([title_system, in_rfp, in_mto, in_vo, result, vo_status])
    
    if print_table:
        print(table)
    
    if result_dir and export_to_excel:
        excel_path = _save_title_systems_to_excel(
            all_title_systems,
            rfp_title_systems,
            mto_title_systems,
            vo_title_systems,
            result_dir
        )
        if excel_path and print_table:
            print(f"Сравнение title_system -> {os.path.basename(excel_path)}")


def _save_title_systems_to_excel(
    all_title_systems: Set[str],
    rfp_title_systems: Set[str],
    mto_title_systems: Set[str],
    vo_title_systems: Set[str],
    result_dir: str
) -> str:
    try:
        ensure_result_dir_exists(result_dir)
        
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Сравнение title_system"
        
        headers = ["title_system", "В RFP", "В MTO", "В VO", "Результат сравнения", "Статус VO"]
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
        
        for title_system in sorted(all_title_systems, key=lambda x: (x is None, x or "")):
            in_rfp = "✓" if title_system in rfp_title_systems else ""
            in_mto = "✓" if title_system in mto_title_systems else ""
            in_vo = "✓" if title_system in vo_title_systems else ""
            vo_status = "Есть в VO" if in_vo else "Нет в VO"
            
            if in_rfp and in_mto and in_vo:
                result = "Совпадают (RFP/MTO/VO)"
                fill_color = "C6EFCE"
            elif in_rfp and in_mto and not in_vo:
                result = "Нет в VO"
                fill_color = "FFEB9C"
            elif in_rfp and in_vo and not in_mto:
                result = "Нет в MTO"
                fill_color = "FFEB9C"
            elif in_mto and in_vo and not in_rfp:
                result = "Нет в RFP"
                fill_color = "FFEB9C"
            elif in_rfp and not in_mto and not in_vo:
                result = "Только RFP"
                fill_color = "FFC7CE"
            elif in_mto and not in_rfp and not in_vo:
                result = "Только MTO"
                fill_color = "FFC7CE"
            elif in_vo and not in_rfp and not in_mto:
                result = "Только VO"
                fill_color = "FFC7CE"
            else:
                result = "Частичное совпадение"
                fill_color = "FFEB9C"
            
            row_data = [title_system, in_rfp, in_mto, in_vo, result, vo_status]
            ws.append(row_data)
            
            for col_num in range(1, len(headers) + 1):
                cell = ws.cell(row=ws.max_row, column=col_num)
                cell.border = border
                cell.alignment = Alignment(horizontal="left", vertical="center")
                if col_num == 5:
                    cell.fill = PatternFill(start_color=fill_color, end_color=fill_color, fill_type="solid")
        
        ws.column_dimensions['A'].width = 30
        ws.column_dimensions['B'].width = 12
        ws.column_dimensions['C'].width = 12
        ws.column_dimensions['D'].width = 12
        ws.column_dimensions['E'].width = 35
        ws.column_dimensions['F'].width = 15
        
        last_col_letter = get_column_letter(len(headers))
        ws.auto_filter.ref = f'A1:{last_col_letter}{ws.max_row}'
        ws.freeze_panes = 'A2'
        
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        file_name = f"Шаг4_RFP_MTO_VO_сравнение_title_system_{timestamp}.xlsx"
        file_path = os.path.join(result_dir, file_name)
        
        wb.save(file_path)
        return file_path
        
    except Exception as e:
        print(f"Ошибка при сохранении таблицы сравнения title_system (RFP/MTO/VO): {e}")
        return None
