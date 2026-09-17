"""
Worker-функции для по-title_mark обработки step4 (sequential).
"""

import time
from typing import Callable, Dict, List, Any, Set, Tuple

from base.base_classes import RowStd, RowType
from base.tables_columns import CODE, VALUES
from RFQ.tags_rfp_compare.step4.step4_1_check_mto_data import check_mto_data, check_vo_data
from RFQ.tags_rfp_compare.step4.step4_2_match_rfp_with_mto import match_rfp_with_mto, match_rfp_with_vo
from RFQ.tags_rfp_compare.step4.step4_3_assign_position_status import assign_position_status
from RFQ.tags_rfp_compare.step4.step4_4_add_unmatched_mto_rows import (
    add_unmatched_mto_rows,
    add_unmatched_vo_rows,
    match_added_vo_rows_with_mto_by_code,
    match_rows_with_vo_and_rfp_code_missing_mto,
    match_unmatched_vo_with_mto_by_reverse_replacement,
    add_unmatched_mto_cabinet_rows,
)
from RFQ.tags_rfp_compare.step4.step4_postmerge_ops import apply_all_postmerge_per_row_ops
from RFQ.tags_rfp_compare.step1_load_rfp import split_rfp_rows
from RFQ.tags_rfp_compare.rfp_tags_utils import set_timing_log_enabled

_shared_kwargs: Dict[str, Any] = {}


def worker_init(shared_kwargs: Dict[str, Any]) -> None:
    """Called once per worker process to store shared (immutable) config."""
    global _shared_kwargs
    _shared_kwargs = shared_kwargs


def process_title_mark_parallel(
    title_mark: str,
    rfp_rows_tm: List[RowStd],
    mto_rows_tm: List[RowStd],
    vo_rows_tm: List[RowStd],
) -> Dict[str, Any]:
    """Thin wrapper for ProcessPoolExecutor: shared kwargs come from initializer globals."""
    return process_title_mark(
        title_mark, rfp_rows_tm, mto_rows_tm, vo_rows_tm,
        debug_log=None,
        **_shared_kwargs,
    )


def process_title_mark(
    title_mark: str,
    rfp_rows_tm: List[RowStd],
    mto_rows_tm: List[RowStd],
    vo_rows_tm: List[RowStd],
    *,
    result_dir: str,
    replacement_table: Dict[str, List[tuple[str, str]]],
    debug_step4_1: bool,
    debug_step4_2: bool,
    debug_step4_3: bool,
    debug_step4_4: bool,
    debug_tag: List[str] | None,
    debug_code: List[str] | None,
    debug_title_system: List[str] | None,
    use_multiprocessing: bool,
    collapse_debug: bool,
    split_rfp_in_worker: bool = False,
    units_ban: List[str] | None = None,
    code_ban: set[str] | None = None,
    collapse_func: Callable[..., List[RowStd]] | None = None,
    lot_map: Dict[str, str] | None = None,
    code_base_by_code: Dict[str, tuple] | None = None,
    debug_log: List[str] | None = None,
    timing_log_enabled: bool = True,
    verbose_progress_messages: bool = True,
    defer_reports: bool = False,
    assign_rfp_mto_code_compare_colors: bool = True,
    include_mto_vo_without_rfp_anchor: bool = False,
) -> Dict[str, Any]:
    """Обрабатывает один title_mark и возвращает локальный результат."""
    t_start = time.perf_counter()
    set_timing_log_enabled(bool(timing_log_enabled))

    _any_debug = debug_step4_1 or debug_step4_2 or debug_step4_3 or debug_step4_4
    _local_debug_log: List[str] | None = None
    if debug_log is None and _any_debug:
        _local_debug_log = []
        debug_log = _local_debug_log

    tm_result_rows: List[RowStd] = list(rfp_rows_tm or [])
    if split_rfp_in_worker and tm_result_rows:
        tm_result_rows, _, _ = split_rfp_rows(
            tm_result_rows,
            units_ban=units_ban,
            code_ban=code_ban,
        )
    tm_mto_data: Dict[str, List[RowStd]] = {title_mark: list(mto_rows_tm or [])}
    tm_vo_data: Dict[str, List[RowStd]] = {title_mark: list(vo_rows_tm or [])}

    tm_mto_data, mto_diag = check_mto_data(
        tm_mto_data,
        result_dir,
        debug=debug_step4_1,
        debug_tag=debug_tag,
        debug_code=debug_code,
        debug_title_system=debug_title_system,
        debug_log=debug_log,
        defer_reports=defer_reports,
        units_ban=units_ban,
    )
    tm_vo_data, vo_diag = check_vo_data(
        tm_vo_data,
        result_dir,
        debug=debug_step4_1,
        debug_tag=debug_tag,
        debug_code=debug_code,
        debug_title_system=debug_title_system,
        debug_log=debug_log,
        defer_reports=defer_reports,
    )

    from RFQ.tags_rfp_compare.step4.step4_quantity_balance import sum_values_in_data_dict

    post_split_mto_sum, post_split_mto_rows = sum_values_in_data_dict(tm_mto_data)
    post_split_vo_sum, post_split_vo_rows = sum_values_in_data_dict(tm_vo_data)

    matched_mto_tags, unmatched_mto_rows = match_rfp_with_mto(
        tm_result_rows,
        tm_mto_data,
        replacement_table=replacement_table,
        result_dir=result_dir,
        debug=debug_step4_2,
        debug_tag=debug_tag,
        debug_code=debug_code,
        debug_title_system=debug_title_system,
        debug_log=debug_log,
    )
    matched_vo_tags, unmatched_vo_rows = match_rfp_with_vo(
        tm_result_rows,
        tm_vo_data,
        replacement_table=replacement_table,
        mto_data=tm_mto_data,
        unmatched_mto_rows=unmatched_mto_rows,
        debug=debug_step4_2,
        debug_tag=debug_tag,
        debug_code=debug_code,
        debug_title_system=debug_title_system,
        debug_log=debug_log,
    )

    assign_position_status(
        tm_result_rows,
        tm_mto_data,
        result_dir=result_dir,
        debug=debug_step4_3,
        debug_tag=debug_tag,
        debug_code=debug_code,
        debug_title_system=debug_title_system,
        debug_log=debug_log,
        verbose_progress_messages=verbose_progress_messages,
    )

    add_unmatched_mto_rows(
        tm_result_rows,
        unmatched_mto_rows,
        tm_mto_data,
        debug=debug_step4_4,
        debug_tag=debug_tag,
        debug_code=debug_code,
        debug_title_system=debug_title_system,
        debug_log=debug_log,
        verbose_progress_messages=verbose_progress_messages,
        include_mto_vo_without_rfp_anchor=include_mto_vo_without_rfp_anchor,
    )
    add_unmatched_vo_rows(
        tm_result_rows,
        unmatched_vo_rows,
        tm_vo_data,
        tm_mto_data,
        replacement_table=replacement_table,
        debug=debug_step4_4,
        debug_tag=debug_tag,
        debug_code=debug_code,
        debug_title_system=debug_title_system,
        debug_log=debug_log,
        verbose_progress_messages=verbose_progress_messages,
        include_mto_vo_without_rfp_anchor=include_mto_vo_without_rfp_anchor,
    )

    mto_by_title, mto_remaining_qty = match_added_vo_rows_with_mto_by_code(
        tm_result_rows,
        tm_mto_data,
        replacement_table=replacement_table,
        debug=debug_step4_4,
        debug_tag=debug_tag,
        debug_code=debug_code,
        debug_title_system=debug_title_system,
        debug_log=debug_log,
        use_multiprocessing=use_multiprocessing,
    )
    match_rows_with_vo_and_rfp_code_missing_mto(
        tm_result_rows,
        tm_mto_data,
        mto_by_title=mto_by_title,
        mto_remaining_qty=mto_remaining_qty,
        debug=debug_step4_4,
        debug_tag=debug_tag,
        debug_code=debug_code,
        debug_title_system=debug_title_system,
        debug_log=debug_log,
        use_multiprocessing=use_multiprocessing,
    )
    match_unmatched_vo_with_mto_by_reverse_replacement(
        tm_result_rows,
        tm_mto_data,
        replacement_table=replacement_table,
        debug=debug_step4_4,
        debug_tag=debug_tag,
        debug_code=debug_code,
        debug_title_system=debug_title_system,
        debug_log=debug_log,
        verbose_progress_messages=verbose_progress_messages,
    )
    add_unmatched_mto_cabinet_rows(
        tm_result_rows,
        tm_mto_data,
        debug=debug_step4_4,
        debug_tag=debug_tag,
        debug_code=debug_code,
        debug_title_system=debug_title_system,
        debug_log=debug_log,
        verbose_progress_messages=verbose_progress_messages,
        include_mto_vo_without_rfp_anchor=include_mto_vo_without_rfp_anchor,
    )

    if collapse_func is not None:
        tm_result_rows = collapse_func(
            tm_result_rows,
            result_dir,
            collapse_debug_enabled=collapse_debug,
        )

    apply_all_postmerge_per_row_ops(
        tm_result_rows,
        replacement_table,
        lot_map=lot_map,
        code_base_by_code=code_base_by_code,
        assign_rfp_mto_code_compare_colors=assign_rfp_mto_code_compare_colors,
    )

    mto_summary, vo_summary = _compute_mto_vo_summaries(
        title_mark, tm_mto_data, tm_vo_data,
    )

    return {
        "title_mark": title_mark,
        "result_rows": tm_result_rows,
        "mto_summary": mto_summary,
        "vo_summary": vo_summary,
        "mto_diagnostics": mto_diag,
        "vo_diagnostics": vo_diag,
        "post_split_mto_sum": post_split_mto_sum,
        "post_split_mto_rows": post_split_mto_rows,
        "post_split_vo_sum": post_split_vo_sum,
        "post_split_vo_rows": post_split_vo_rows,
        "worker_duration": time.perf_counter() - t_start,
        "debug_log_lines": _local_debug_log,
    }


def _normalize_code(code_value) -> str:
    if code_value is None:
        return ""
    return str(code_value).strip().upper()


def _compute_mto_vo_summaries(
    title_mark: str,
    tm_mto_data: Dict[str, List[RowStd]],
    tm_vo_data: Dict[str, List[RowStd]],
) -> Tuple[dict, dict]:
    """Вычисляет лёгкие summary MTO/VO для check_values_sum и _save_code_ban_list.

    Возвращает (mto_summary, vo_summary) — маленькие dict/set вместо
    полных List[RowStd], чтобы не пиклить post-split данные (~120 MB → ~1-2 KB).
    """
    tagged_sum = 0.0
    tagged_values_by_code: Dict[str, float] = {}
    codes_with_tags: Set[str] = set()

    for mto_rows in tm_mto_data.values():
        for row in mto_rows:
            if row.row_type != RowType.position_row:
                continue
            tags = row.get_tags_list()
            code_raw = row.get_value(CODE)
            code_norm = _normalize_code(code_raw)
            if tags:
                val = row.get_value(VALUES)
                try:
                    vf = float(val) if val else 0.0
                except (ValueError, TypeError):
                    vf = 0.0
                tagged_sum += vf
                if code_norm:
                    tagged_values_by_code[code_norm] = (
                        tagged_values_by_code.get(code_norm, 0.0) + vf
                    )
                    codes_with_tags.add(code_norm)

    mto_summary = {
        "tagged_sum": tagged_sum,
        "tagged_values_by_code": tagged_values_by_code,
        "codes_with_tags": codes_with_tags,
    }

    vo_tagged_sum = 0.0
    vo_codes: Set[str] = set()

    for vo_rows in tm_vo_data.values():
        for row in vo_rows:
            if row.row_type != RowType.position_row:
                continue
            code_raw = row.get_value(CODE)
            code_norm = _normalize_code(code_raw)
            if code_norm:
                vo_codes.add(code_norm)
            tags = row.get_tags_list()
            if tags:
                val = row.get_value(VALUES)
                try:
                    vf = float(val) if val else 0.0
                except (ValueError, TypeError):
                    vf = 0.0
                vo_tagged_sum += vf

    vo_summary = {
        "tagged_sum": vo_tagged_sum,
        "codes": vo_codes,
    }

    return mto_summary, vo_summary
