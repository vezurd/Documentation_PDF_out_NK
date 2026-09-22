"""
Задача от 23.12.2025 от Михеева по указанию Нестро
Цель: Подготовить инструмент для сравнения тегов оборудования из RFP с актуальными МТО и вендорской документацией VO

Главный файл для запуска полного цикла проверки тегов.
Этапы разделены на отдельные модули для возможности независимого запуска.
"""

import gc
import io
import os
import sys
import time
import threading
from concurrent.futures import ThreadPoolExecutor
from typing import List, Dict, Any, Optional

from base.base_classes import RowStd
from utils.path import open_dir
from utils.cache_utils import set_cache_dir

# Импорты из модулей этапов (полный путь пакета для работы при импорте из main.py)
from RFQ.tags_rfp_compare.rfp_tags_utils import (
    get_result_dir_path,
    ensure_result_dir_exists,
    append_timing_log,
    append_memory_log,
    append_memory_top_stats,
    log_data_structures_memory,
    finalize_timing_log,
    load_config,
    get_column_optimization_hash,
    resolve_units_split_ban_config,
    set_timing_log_enabled,
    set_memory_log_enabled,
    resolve_rfp_pipeline_mode,
    resolve_effective_rfp_path,
    RFP_PIPELINE_MODE_MTO_VO_ONLY,
    RFP_PIPELINE_MODE_STANDARD,
    RFP_PIPELINE_MODE_WITH_ORPHAN_MTO_VO,
)
from base.base_google import load_base
from RFQ.packing_list_provider import PackingDataset, load_packing_dataset
from RFQ.ds_compare.tsd_packing_load import (
    ensure_tsd_packing_cache_current,
    format_packing_cache_freshness_detail,
)
from RFQ.tags_rfp_compare.step1_load_rfp import (
    step1_finalize_rfp_data,
    step1_load_rfp_raw,
)
from RFQ.tags_rfp_compare.units_gate import run_units_gate
from RFQ.tags_rfp_compare.rfp_supply_status import RfpSupplyStatusError
from RFQ.units_convert.models import UnitsConversionError
from RFQ.tags_rfp_compare.step2_load_mto import step2_load_mto_data
from RFQ.tags_rfp_compare.step3_load_vo import step3_load_vo_data
from base.tables_columns import DS_TITLE
from RFQ.tags_rfp_compare.step4_analyze_and_match import step4_analyze_and_match, _build_code_base_by_code
from RFQ.tags_rfp_compare.step4.step4_bcc_accum_matrix import build_bcc_accum_matrix
from RFQ.tags_rfp_compare.step4.step4_quantity_balance import QuantityBalanceFatalError
from RFQ.tags_rfp_compare.step4.step4_packing_compare import RfpPackingFatalError
from RFQ.tags_rfp_compare.column_optimization import (
    apply_column_optimization,
    apply_load_tags_mode,
)
from RFQ.tags_rfp_compare.rfp_progress import emit_milestone
from RFQ.rfp_parts.parts_net_preflight import ensure_rfp_parts_net_current
from RFQ.rfp_parts.ds_hybrid_preflight import (
    INPUT_MODE_DS_ONLY,
    INPUT_MODE_HYBRID,
    DsBaselineBlockedError,
    DsHybridBlockedError,
    copy_ds_sidecars_to_result_dir,
    ensure_ds_baseline_current,
    ensure_ds_hybrid_current,
    resolve_input_mode,
)
from RFQ.rfp_parts.file_status import (
    DUPLICATE_TAGS_XLSX_NAME,
    EMPTY_CODE_XLSX_NAME,
    copy_empty_code_report_to_result_dir,
)
from RFQ.rfp_parts.ds_checklist import DEFAULT_PARTS_DIR
from RFQ.rfp_parts.ds_id_coverage import check_ds_id_coverage
from RFQ.tags_rfp_compare.ds_manager_matrix import (
    load_ds_manager_matrix,
    collect_parts_ds_paths,
    compare_matrix_to_parts,
)
from RFQ.tags_rfp_compare.ds_manager_roster import sync_ds_roster_sheet
from RFQ.tags_rfp_compare.gem_supply_codes import load_gem_supply_codes


class _ThreadLocalStream:
    """Thread-local stream wrapper: each thread can redirect output to its own buffer."""

    def __init__(self, original):
        self._original = original
        self._local = threading.local()

    def set_stream(self, stream):
        self._local.stream = stream

    def clear_stream(self):
        self._local.stream = None

    @property
    def _target(self):
        return getattr(self._local, 'stream', None) or self._original

    def write(self, data):
        return self._target.write(data)

    def flush(self):
        self._target.flush()

    def isatty(self):
        return False

    @property
    def encoding(self):
        return getattr(self._original, 'encoding', 'utf-8')

    def __getattr__(self, name):
        return getattr(self._original, name)


def _to_list(val):  # -> List[str]
    """Преобразует значение в список строк. Строка разбивается по запятой. Пустой результат -> []."""
    if val is None:
        return []
    if isinstance(val, list):
        return [str(x).strip() for x in val if str(x).strip()]
    s = str(val).strip()
    if not s:
        return []
    return [x.strip() for x in s.split(",") if x.strip()]


def _opt_list(lst: list) -> Optional[List[str]]:
    """Возвращает None для пустого списка (использовать дефолт step4), иначе список."""
    return lst if lst else None


def _load_config_override(config_path: str) -> Dict[str, Any]:
    """Load a supplied main/as-build config against the matching defaults."""
    import json

    from RFQ.tags_rfp_compare.rfp_tags_utils import (
        _deep_merge,
        get_asbuild_config_path,
        get_default_asbuild_config,
        get_default_config,
    )

    resolved_path = os.path.abspath(config_path)
    with open(resolved_path, "r", encoding="utf-8") as config_file:
        loaded = json.load(config_file)
    is_asbuild = os.path.normcase(resolved_path) == os.path.normcase(
        os.path.abspath(get_asbuild_config_path())
    )
    default = get_default_asbuild_config() if is_asbuild else get_default_config()
    merged = _deep_merge(default, loaded)
    if is_asbuild:
        from RFQ.tags_rfp_compare.rfp_tags_utils import _apply_asbuild_rfp_parts_fixed

        _apply_asbuild_rfp_parts_fixed(merged)
    return merged


# ============================================================================
# ГЛАВНАЯ ФУНКЦИЯ
# ============================================================================

def main(config_override: Optional[Dict[str, Any]] = None):
    """Главная функция для запуска процесса сравнения тегов.
    
    :param config_override: Опциональная конфигурация (если None — загружается из файла)
    """
    emit_milestone("prepare", "Running")
    aggregate_start = time.perf_counter()
    config = config_override if config_override is not None else load_config()
    paths = config.get("paths", {})
    step1_cfg = config.get("step1", {})
    step2_cfg = config.get("step2", {})
    step3_cfg = config.get("step3", {})
    step4_cfg = config.get("step4", {})
    rtu_cfg = config.get("rfp_tags_utils", {})
    save_input_fingerprints = rtu_cfg.get("save_input_fingerprints", True)

    rfp_path = paths.get("rfp_path", "")
    rfp_registr_lot_path = paths.get("rfp_registr_lot_path", "")
    mto_path = paths.get("mto_path", "")
    vo_path = paths.get("vo_path", "")
    code_ban_file = paths.get("code_ban_file", "")
    replacement_table_file = paths.get("replacement_table_file", "")
    result_dir_base = paths.get("result_dir_base", "")

    script_dir = os.path.dirname(os.path.abspath(__file__))
    cache_dir = os.path.join(script_dir, "cache")
    set_cache_dir(cache_dir)

    result_dir = get_result_dir_path(result_dir_base)
    ensure_result_dir_exists(result_dir)
    print(f"Папка результатов: {result_dir}\n")
    emit_milestone("prepare", "Done", f"result_dir={result_dir}")

    rfp_parts_cfg = config.get("rfp_parts", {})
    if not isinstance(rfp_parts_cfg, dict):
        rfp_parts_cfg = {}
    rfp_pipeline_mode = resolve_rfp_pipeline_mode(step1_cfg)
    skip_rfp_load = rfp_pipeline_mode == RFP_PIPELINE_MODE_MTO_VO_ONLY
    use_latest_net = bool(rfp_parts_cfg.get("use_latest_net", True))
    load_tags = bool(config.get("load_tags", True))
    include_packing_lists = bool(step4_cfg.get("include_packing_lists", True))
    parts_tag_report_path = None
    if skip_rfp_load:
        emit_milestone("ds_id_check", "Skipped", "mto_vo_only")
    else:
        emit_milestone("ds_id_check", "Running")
        try:
            ds_id_result = check_ds_id_coverage()
        except Exception as exc:
            detail = str(exc).splitlines()[0]
            emit_milestone("ds_id_check", "Error", detail)
            print(f"Соответствие ДС RFP↔УЛ: {detail}")
        else:
            summary = ds_id_result.summary_line()
            if ds_id_result.load_error:
                emit_milestone("ds_id_check", "Error", summary)
            else:
                emit_milestone("ds_id_check", "Done", summary)
            print(f"Соответствие ДС RFP↔УЛ: {summary}")

    ds_manager_matrix = None
    rfp_paths_by_key = collect_parts_ds_paths(DEFAULT_PARTS_DIR)
    matrix_path = str(paths.get("ds_manager_matrix") or "").strip()
    if not matrix_path:
        emit_milestone("ds_mp_check", "Skipped", "paths.ds_manager_matrix пуст")
    else:
        emit_milestone("ds_mp_check", "Running")
        try:
            ds_manager_matrix = load_ds_manager_matrix(matrix_path)
            presence = compare_matrix_to_parts(ds_manager_matrix, DEFAULT_PARTS_DIR)
            print(f"Соответствие ДС RFP↔МП: {presence.summary_line}")
            if ds_manager_matrix.load_error:
                emit_milestone("ds_mp_check", "Error", presence.summary_line)
            else:
                roster = sync_ds_roster_sheet(
                    ds_manager_matrix,
                    parts_dir=DEFAULT_PARTS_DIR,
                )
                print(f"Соответствие ДС RFP↔МП: {roster.summary_line}")
                if roster.write_error:
                    emit_milestone("ds_mp_check", "Error", roster.summary_line)
                else:
                    emit_milestone("ds_mp_check", "Done", roster.summary_line)
        except Exception as exc:
            detail = str(exc).splitlines()[0]
            emit_milestone("ds_mp_check", "Error", detail)
            print(f"Соответствие ДС RFP↔МП: {detail}")

    input_mode = resolve_input_mode(config)
    if skip_rfp_load:
        emit_milestone("parts_preflight", "Skipped", "mto_vo_only")
    elif not use_latest_net:
        emit_milestone("parts_preflight", "Skipped", "use_latest_net=false")
    elif input_mode == INPUT_MODE_DS_ONLY:
        emit_milestone("parts_preflight", "Running", "input_mode=ds_only")
        try:
            net_path, freshness = ensure_ds_baseline_current(
                config=config,
                matrix_path=str(paths.get("units_convert_matrix") or "").strip()
                or None,
                progress=True,
            )
        except DsBaselineBlockedError as exc:
            copy_ds_sidecars_to_result_dir(exc.output_dir, result_dir)
            emit_milestone("parts_preflight", "Error", str(exc).splitlines()[0])
            emit_milestone("complete", "Error", "DS baseline blocked")
            print(str(exc))
            raise SystemExit(1) from exc
        except Exception as exc:
            emit_milestone("parts_preflight", "Error", str(exc))
            emit_milestone("complete", "Error", "DS baseline refresh failed")
            print(
                "Не удалось проверить или пересобрать свод ДС "
                f"(rfp_parts.ds_source_dir → Свод ДС для запуска.xlsx): {exc}"
            )
            raise SystemExit(1) from exc
        action = "rebuilt" if freshness.needs_rebuild else "unchanged"
        print(f"Свод ДС: {action}; {freshness.reason}")
        print(f"{os.path.basename(str(net_path))}: {net_path}")
        copied = copy_ds_sidecars_to_result_dir(
            os.path.dirname(str(net_path)), result_dir
        )
        for dest in copied:
            print(f"Отчёт ДС: {dest.name}")
        emit_milestone(
            "parts_preflight",
            "Done",
            f"{action}; {freshness.reason}",
        )
    elif input_mode == INPUT_MODE_HYBRID:
        emit_milestone("parts_preflight", "Running", "input_mode=hybrid")
        try:
            net_path, freshness = ensure_ds_hybrid_current(
                config=config,
                matrix_path=str(paths.get("units_convert_matrix") or "").strip()
                or None,
                progress=True,
            )
        except DsBaselineBlockedError as exc:
            copy_ds_sidecars_to_result_dir(exc.output_dir, result_dir)
            emit_milestone("parts_preflight", "Error", str(exc).splitlines()[0])
            emit_milestone("complete", "Error", "DS baseline blocked")
            print(str(exc))
            raise SystemExit(1) from exc
        except DsHybridBlockedError as exc:
            copy_ds_sidecars_to_result_dir(exc.output_dir, result_dir)
            emit_milestone(
                "parts_preflight",
                "Error",
                str(exc).splitlines()[0],
            )
            emit_milestone("complete", "Error", "DS-RFP hybrid blocked")
            print(str(exc))
            raise SystemExit(1) from exc
        except Exception as exc:
            emit_milestone("parts_preflight", "Error", str(exc))
            emit_milestone("complete", "Error", "DS-RFP hybrid refresh failed")
            print(
                "Не удалось проверить или пересобрать свод ДС-RFP "
                f"(без fallback на rfp_parts_net.xlsx): {exc}"
            )
            raise SystemExit(1) from exc
        action = "rebuilt" if freshness.needs_rebuild else "unchanged"
        hybrid_detail = freshness.summary or freshness.reason
        print(f"Свод ДС-RFP: {action}; {hybrid_detail}")
        print(f"{os.path.basename(str(net_path))}: {net_path}")
        copied = copy_ds_sidecars_to_result_dir(
            os.path.dirname(str(net_path)), result_dir
        )
        if freshness.baseline_freshness is not None:
            copied.extend(
                copy_ds_sidecars_to_result_dir(
                    freshness.baseline_freshness.output_dir, result_dir
                )
            )
        for dest in copied:
            print(f"Отчёт ДС-RFP: {dest.name}")
        emit_milestone(
            "parts_preflight",
            "Done",
            f"{action}; {hybrid_detail}",
        )
    else:
        emit_milestone("parts_preflight", "Running")
        try:
            net_path, freshness = ensure_rfp_parts_net_current(
                progress=True,
                load_tags=load_tags,
            )
        except RfpSupplyStatusError as exc:
            emit_milestone("parts_preflight", "Error", str(exc).splitlines()[0])
            emit_milestone("complete", "Error", "некорректный статус позиции RFP")
            print(str(exc))
            raise SystemExit(1) from exc
        except Exception as exc:
            emit_milestone("parts_preflight", "Error", str(exc))
            emit_milestone("complete", "Error", "parts net refresh failed")
            print(
                "Не удалось проверить или пересобрать свод частей RFP "
                f"(папка RFP_Зиновьев → "
                f"{'rfp_parts_net.xlsx' if load_tags else 'rfp_parts_net_no_tags.xlsx'}): "
                f"{exc}"
            )
            raise SystemExit(1) from exc
        action = "rebuilt" if freshness.needs_rebuild else "unchanged"
        print(f"Свод частей: {action}; {freshness.reason}")
        print(f"{os.path.basename(str(net_path))}: {net_path}")
        tag_report = os.path.join(os.path.dirname(str(net_path)), DUPLICATE_TAGS_XLSX_NAME)
        if os.path.isfile(tag_report):
            parts_tag_report_path = tag_report
        empty_code_dest = copy_empty_code_report_to_result_dir(
            os.path.dirname(str(net_path)), result_dir
        )
        if empty_code_dest is not None:
            print(f"Позиции без кода RFP: {EMPTY_CODE_XLSX_NAME}")
        emit_milestone(
            "parts_preflight",
            "Done",
            f"{action}; {freshness.reason}",
        )

    if not include_packing_lists:
        emit_milestone("ul_preflight", "Skipped", "include_packing_lists=false")
    else:
        emit_milestone("ul_preflight", "Running")
        try:
            packing_freshness = ensure_tsd_packing_cache_current()
        except FileNotFoundError as exc:
            detail = str(exc).splitlines()[0]
            emit_milestone("ul_preflight", "Error", detail)
            print(f"Свод УЛ: папка ТСД недоступна ({exc})")
        except Exception as exc:
            emit_milestone("ul_preflight", "Error", str(exc).splitlines()[0])
            emit_milestone("complete", "Error", "packing cache refresh failed")
            print(
                "Не удалось проверить или пересобрать свод УЛ "
                f"(папка ТСД → tsd_packing_rows.cache): {exc}"
            )
            raise SystemExit(1) from exc
        else:
            action = "rebuilt" if not packing_freshness.from_cache else "unchanged"
            print(f"Свод УЛ: {action}; {packing_freshness.summary_path}")
            emit_milestone(
                "ul_preflight",
                "Done",
                format_packing_cache_freshness_detail(packing_freshness),
            )

    gem_supply_codes = None
    gem_path = str(paths.get("gem_supply_codes") or "").strip()
    if gem_path:
        gem_supply_codes = load_gem_supply_codes(gem_path)
        if gem_supply_codes.load_error:
            print(f"Коды поставки ГЭМ: {gem_supply_codes.load_error}")
        else:
            print(f"Коды поставки ГЭМ: {len(gem_supply_codes.codes)} кодов")

    set_timing_log_enabled(step4_cfg.get("timing_log", True))  # влияет на step1..4 и finalize
    set_memory_log_enabled(config.get("memory_log", True))  # влияет на step1..4

    total_start = time.perf_counter()
    col_opt_hash_rfp = get_column_optimization_hash(config.get("column_optimization", {}).get("rfp"))
    skip_split_step1 = step1_cfg.get("skip_split", False)
    use_rfp_code_ban = step1_cfg.get("use_rfp_code_ban", True)
    units_split_ban = resolve_units_split_ban_config(config)
    include_cfg = bool(step4_cfg.get("include_mto_vo_without_rfp_anchor", False))
    include_mto_vo_effective = (
        include_cfg if rfp_pipeline_mode == RFP_PIPELINE_MODE_STANDARD else True
    )
    if rfp_pipeline_mode != RFP_PIPELINE_MODE_STANDARD and not include_cfg:
        print(
            f"step1.rfp_pipeline_mode={rfp_pipeline_mode}: "
            "include_mto_vo_without_rfp_anchor принудительно True для step4."
        )

    ul_match_use_mto_tags = bool(step4_cfg.get("ul_match_use_mto_tags", True))
    units_convert_matrix = paths.get("units_convert_matrix", "")

    if skip_rfp_load:
        emit_milestone("rfp_load", "Skipped", "mto_vo_only")
        step_start = time.perf_counter()
        rfp_data, rfp_errors = [], []
        append_timing_log(
            result_dir,
            f"step1_load_rfp_raw: {time.perf_counter() - step_start:.3f}s (skipped, mto_vo_only)",
        )
        print("ЭТАП 1: пропуск загрузки RFP (режим mto_vo_only).")
    else:
        emit_milestone("rfp_load", "Running")
        try:
            effective_rfp = resolve_effective_rfp_path(config)
        except FileNotFoundError as exc:
            emit_milestone("rfp_load", "Error", str(exc))
            emit_milestone("complete", "Error", "RFP source not found")
            print(str(exc))
            raise SystemExit(1) from exc
        rfp_path = effective_rfp.path
        print(f"Источник RFP ({effective_rfp.detail}): {rfp_path}")
        step_start = time.perf_counter()
        try:
            rfp_data = step1_load_rfp_raw(
                rfp_path,
                result_dir,
                column_optimization_hash=col_opt_hash_rfp or None,
            )
            rfp_errors = []
        except RfpSupplyStatusError as exc:
            emit_milestone("rfp_load", "Error", str(exc).splitlines()[0])
            emit_milestone("complete", "Error", "некорректный статус позиции RFP")
            print(str(exc))
            raise SystemExit(1) from exc
        except Exception as exc:
            emit_milestone("rfp_load", "Error", str(exc))
            emit_milestone("complete", "Error", "RFP load failed")
            raise
        emit_milestone("rfp_load", "Done", f"rows={len(rfp_data)}")
        append_timing_log(result_dir, f"step1_load_rfp_raw: {time.perf_counter() - step_start:.3f}s")
    append_memory_log(result_dir, "after_step1_load_rfp_raw")
    log_data_structures_memory(result_dir, "after_step1_load_rfp_raw", rfp_data=rfp_data)

    col_opt_hash_mto = get_column_optimization_hash(config.get("column_optimization", {}).get("mto"))

    mto_load_start = time.perf_counter()
    step_start = mto_load_start
    export_positions_excel_enabled = step2_cfg.get("export_positions_database_excel", False)

    def _load_mto_live():
        return step2_load_mto_data(
            mto_path, result_dir, rfp_data,
            debug=step2_cfg.get("debug", False),
            column_optimization_hash=col_opt_hash_mto or None,
            save_input_fingerprints=save_input_fingerprints,
            export_load_results_excel=step2_cfg.get("export_load_results_excel", True),
            export_positions_database_excel=export_positions_excel_enabled,
            flat_mto_structure=step2_cfg.get("flat_mto_structure", False),
            units_split_ban=units_split_ban,
        )

    def _load_vo_live():
        return step3_load_vo_data(
            vo_path, result_dir,
            debug=step3_cfg.get("debug", False),
            export_to_excel=step3_cfg.get("export_to_excel", False),
            save_input_fingerprints=save_input_fingerprints,
        )

    def _load_mto_vo_and_google():
        if export_positions_excel_enabled:
            print(
                "Step2 export_positions_database_excel=True -> "
                "Step2/Step3 запускаются последовательно (live logs)."
            )
            with ThreadPoolExecutor(max_workers=1) as _gbase_pool:
                _fut_gbase = _gbase_pool.submit(load_base)
                loaded_mto = _load_mto_live()
                loaded_vo = _load_vo_live()
            loaded_google_rows = _fut_gbase.result() or []
            loaded_code_base = _build_code_base_by_code(loaded_google_rows)
            return (
                loaded_mto,
                loaded_vo,
                loaded_code_base,
                loaded_google_rows,
                "step2+step3_sequential_load",
                "after_step2_step3_sequential",
            )

        _mto_buf = io.StringIO()
        _vo_buf = io.StringIO()
        _tls_out = _ThreadLocalStream(sys.stdout)
        _tls_err = _ThreadLocalStream(sys.stderr)
        _saved_out, _saved_err = sys.stdout, sys.stderr
        sys.stdout, sys.stderr = _tls_out, _tls_err

        def _load_mto_buffered():
            _tls_out.set_stream(_mto_buf)
            _tls_err.set_stream(_mto_buf)
            try:
                return _load_mto_live()
            finally:
                _tls_out.clear_stream()
                _tls_err.clear_stream()

        def _load_vo_buffered():
            _tls_out.set_stream(_vo_buf)
            _tls_err.set_stream(_vo_buf)
            try:
                return _load_vo_live()
            finally:
                _tls_out.clear_stream()
                _tls_err.clear_stream()

        _fut_gbase = None
        try:
            # load_base запускается 3-м потоком: его сетевая задержка скрывается за step2+step3
            with ThreadPoolExecutor(max_workers=3) as pool:
                fut_mto = pool.submit(_load_mto_buffered)
                fut_vo = pool.submit(_load_vo_buffered)
                _fut_gbase = pool.submit(load_base)
                loaded_mto = fut_mto.result()
                loaded_vo = fut_vo.result()
        finally:
            sys.stdout, sys.stderr = _saved_out, _saved_err
            for _buf in (_mto_buf, _vo_buf):
                _t = _buf.getvalue()
                if _t:
                    sys.stdout.write(_t)
            sys.stdout.flush()
        loaded_google_rows = (_fut_gbase.result() if _fut_gbase else None) or []
        loaded_code_base = _build_code_base_by_code(loaded_google_rows)
        return (
            loaded_mto,
            loaded_vo,
            loaded_code_base,
            loaded_google_rows,
            "step2+step3_parallel_load",
            "after_step2_step3_parallel",
        )

    emit_milestone("mto_google_load", "Running")
    emit_milestone("vo_load", "Running")
    try:
        (
            mto_data,
            vo_data,
            code_base_by_code,
            google_rows,
            load_timing_label,
            mem_label,
        ) = _load_mto_vo_and_google()
    except Exception as exc:
        emit_milestone("mto_google_load", "Error", str(exc))
        emit_milestone("vo_load", "Error", str(exc))
        emit_milestone("complete", "Error", "MTO/VO/Google load failed")
        raise
    emit_milestone(
        "mto_google_load",
        "Done",
        f"title_systems={len(mto_data)}; google_codes={len(code_base_by_code)}",
    )
    emit_milestone("vo_load", "Done", f"title_systems={len(vo_data)}")

    emit_milestone("units_gate", "Running")
    packing_dataset: PackingDataset | None = None
    try:
        if include_packing_lists:
            step_start = time.perf_counter()
            packing_dataset = load_packing_dataset()
            append_timing_log(
                result_dir,
                f"load_packing_dataset: {time.perf_counter() - step_start:.3f}s",
            )

        apply_load_tags_mode(
            load_tags=load_tags,
            rfp_rows=rfp_data,
            mto_data=mto_data,
            vo_data=vo_data,
            packing_rows=None if packing_dataset is None else packing_dataset.rows,
        )

        step_start = time.perf_counter()
        gate_result = run_units_gate(
            rfp_rows=rfp_data,
            mto_data=mto_data,
            packing_dataset=packing_dataset,
            google_rows=google_rows,
            matrix_path=units_convert_matrix,
            logs_dir=result_dir,
        )
    except UnitsConversionError as exc:
        detail = str(exc).splitlines()[0]
        emit_milestone("units_gate", "Error", detail)
        emit_milestone("complete", "Error", f"конвертация ед. изм.: {detail}")
        print(f"Конвертация ед. изм. прервана до Step4: {exc}")
        raise SystemExit(1) from exc
    except Exception as exc:
        detail = str(exc).splitlines()[0]
        emit_milestone("units_gate", "Error", detail)
        emit_milestone("complete", "Error", detail)
        raise
    append_timing_log(result_dir, f"units_gate: {time.perf_counter() - step_start:.3f}s")
    emit_milestone("units_gate", "Done", gate_result.summary)

    export_bcc_accum_matrix = bool(step4_cfg.get("export_bcc_accum_matrix", True))
    bcc_accum_matrix = None
    if export_bcc_accum_matrix:
        snap_start = time.perf_counter()
        bcc_accum_matrix = build_bcc_accum_matrix(
            rfp_rows=rfp_data,
            mto_data=mto_data,
            packing_dataset=packing_dataset,
            google_rows=google_rows,
        )
        append_timing_log(
            result_dir,
            f"build_bcc_accum_matrix: {time.perf_counter() - snap_start:.3f}s",
        )

    if not skip_rfp_load:
        step_start = time.perf_counter()
        rfp_data, rfp_errors = step1_finalize_rfp_data(
            rfp_data,
            result_dir,
            debug=step1_cfg.get("debug", False),
            code_ban_file=code_ban_file,
            skip_split=skip_split_step1,
            units_ban=units_split_ban,
            use_rfp_code_ban=use_rfp_code_ban,
            parts_tag_report_path=parts_tag_report_path,
        )
        append_timing_log(result_dir, f"step1_finalize_rfp_data: {time.perf_counter() - step_start:.3f}s")

    col_opt = config.get("column_optimization", {})
    if col_opt and rfp_data:
        apply_column_optimization(rfp_data, col_opt.get("rfp", {}))
    append_memory_log(result_dir, "after_step1_finalize_rfp")
    log_data_structures_memory(result_dir, "after_step1_finalize_rfp", rfp_data=rfp_data)

    if col_opt and mto_data:
        mto_cfg = col_opt.get("mto", {})
        if mto_cfg:
            for title_system, rows in mto_data.items():
                if rows:
                    apply_column_optimization(rows, mto_cfg)

    load_elapsed = time.perf_counter() - mto_load_start
    append_timing_log(result_dir, f"{load_timing_label}: {load_elapsed:.3f}s")
    append_memory_log(result_dir, mem_label)
    log_data_structures_memory(result_dir, mem_label, rfp_data=rfp_data, mto_data=mto_data, vo_data=vo_data)

    step_start = time.perf_counter()
    emit_milestone("step4_match", "Running")
    try:
        result = step4_analyze_and_match(
            rfp_data,
            mto_data,
            vo_data,
            result_dir,
            replacement_table_file=replacement_table_file,
            rfp_registr_lot_path=rfp_registr_lot_path,
            code_ban_file=code_ban_file,
            debug=step4_cfg.get("debug", True),
            debug_step4_1=step4_cfg.get("debug_step4_1", True),
            debug_step4_2=step4_cfg.get("debug_step4_2", True),
            debug_step4_3=step4_cfg.get("debug_step4_3", True),
            debug_step4_4=step4_cfg.get("debug_step4_4", True),
            collapse_debug=step4_cfg.get("collapse_debug", False),
            unified_debug=step4_cfg.get("unified_debug", True),
            debug_tag=_opt_list(_to_list(step4_cfg.get("debug_tag"))),
            debug_code=_opt_list(_to_list(step4_cfg.get("debug_code"))),
            debug_title_system=_opt_list(_to_list(step4_cfg.get("debug_title_system"))),
            print_title_systems_table=step4_cfg.get("print_title_systems_table", False),
            export_title_systems_comparison_excel=step4_cfg.get("export_title_systems_comparison_excel", True),
            use_multiprocessing=step4_cfg.get("use_multiprocessing", False),
            split_rfp_in_worker=skip_split_step1,
            parallel_by_title_mark=step4_cfg.get("parallel_by_title_mark", False),
            max_workers=step4_cfg.get("max_workers", 0),
            timing_log_enabled=step4_cfg.get("timing_log", True),
            verbose_progress_messages=step4_cfg.get("verbose_progress_messages", True),
            code_base_by_code=code_base_by_code,
            assign_rfp_mto_code_compare_colors=step4_cfg.get("assign_rfp_mto_code_compare_colors", True),
            filter_to_mto_titles=step4_cfg.get("filter_to_mto_titles", False),
            include_mto_vo_without_rfp_anchor=include_mto_vo_effective,
            include_packing_lists=include_packing_lists,
            ul_match_use_mto_tags=ul_match_use_mto_tags,
            packing_dataset=packing_dataset,
            units_split_ban=units_split_ban,
            use_rfp_code_ban=use_rfp_code_ban,
            export_bcc_accum_matrix=export_bcc_accum_matrix,
            bcc_accum_matrix=bcc_accum_matrix,
            ds_manager_matrix=ds_manager_matrix,
            gem_supply_codes=gem_supply_codes,
            rfp_paths_by_key=rfp_paths_by_key,
        )
    except (QuantityBalanceFatalError, RfpPackingFatalError) as exc:
        detail = getattr(exc, "summary", "") or str(exc)
        emit_milestone("complete", "Error", detail.splitlines()[0])
        raise SystemExit(1) from exc
    except Exception as exc:
        emit_milestone("complete", "Error", str(exc))
        raise
    append_timing_log(result_dir, f"step4_analyze_and_match_total: {time.perf_counter() - step_start:.3f}s")
    append_memory_log(result_dir, "after_step4_analyze_and_match")
    log_data_structures_memory(result_dir, "after_step4_analyze_and_match", result_rows=result, mto_data=mto_data, vo_data=vo_data)
    if step4_cfg.get("memory_top_stats", False):
        append_memory_top_stats(result_dir, "after_step4_analyze_and_match", top_n=step4_cfg.get("memory_top_n", 15))
    append_timing_log(result_dir, f"total_runtime: {time.perf_counter() - total_start:.3f}s")
    finalize_timing_log(result_dir, top_n=rtu_cfg.get("finalize_timing_top_n", 5))
    
    if os.path.exists(result_dir) and os.path.isdir(result_dir) and os.listdir(result_dir):
        open_dir(result_dir)

    # Явная очистка крупных структур данных для освобождения памяти.
    # Критично при запуске из GUI (процесс продолжает жить после return).
    del rfp_data, mto_data, vo_data
    if result is not None:
        result.clear()
    del result
    gc.collect()

    aggregate_elapsed = time.perf_counter() - aggregate_start
    print(f"\nВремя: {aggregate_elapsed:.1f}s")
    emit_milestone("complete", "Done", f"elapsed={aggregate_elapsed:.1f}s")

    return None


if __name__ in {"__main__"}:
    _cfg_override = None
    if len(sys.argv) > 1:
        _cfg_path = os.path.abspath(sys.argv[1])
        try:
            _cfg_override = _load_config_override(_cfg_path)
        except Exception as _e:
            print(f"Ошибка загрузки конфига {_cfg_path}: {_e}")
            raise SystemExit(1) from _e
    main(config_override=_cfg_override)




