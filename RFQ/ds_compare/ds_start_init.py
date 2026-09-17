from pathlib import Path
from typing import Optional

import utils.path
from RFQ.ds_compare.ds_compare_config import (
    InCabinetWatchFilters,
    compare_debug_log_requested,
    load_ds_compare_config,
    normalize_ds_vs_mto_output,
    normalize_in_cabinet_debug,
    resolve_mto_path_from_config,
)
from RFQ.ds_compare.ds_vs_mto_excel_columns import ColumnDef, column_config_from_output_settings
from RFQ.ds_compare.ds_vs_mto_excel_xlsxwriter import (
    save_ds_vs_mto_excel_per_export_mode,
)
from RFQ.ds_compare.mto_check import MtoCheck
from RFQ.ds_compare.report import *
from RFQ.ds_compare.ds_in_cabinet_trace import CabinetInCabinetTrace, spec_key_matches_watch
from RFQ.ds_compare.support_finctions import *
from base.base_classes import RowStd
from base.base_google import load_base


def analyze_ds_specification(
    ds_path: str,
    mto_path=None,
    replacement_table_file=None,
    summ_ds=False,
    flat_mto_structure: Optional[bool] = None,
    mto_use_cache: Optional[bool] = None,
    mto_force_update: Optional[bool] = None,
    ds_vs_mto_xw_columns: list[ColumnDef] | None = None,
    compare_debug_log: bool | str | Path | None = None,
    ds_data_override: list[RowStd] | None = None,
    save_excel: bool = True,
    return_data: bool = False,
    output_prefix: str = "РОБОТ_СРАВНЕНИЕ_",
    show_internal_column_names: bool | None = None,
):
    """
    Задача от 05.09.2025 от Михеева по указанию Кузина
    Цель: Подготовить инструмент для сравнения закупочной спецификации с рабочей спецификацией

    Если ``mto_path`` не передан, корневая папка МТО берётся из ``ds_compare_config.json``
    (``mto_paths`` / ``mto_path_selected_index``), выбор в главном окне справа от кнопки
    «Список ДС vs MTO (файл)».

    ds_vs_mto_xw_columns: optional list of ``ColumnDef`` (see ``ds_vs_mto_excel_columns``)
        for the fast _xw Excel export; None = built-in ``DS_VS_MTO_OUTPUT_COLUMNS_CONFIG``.
    compare_debug_log: None — флаг и фильтры из ``ds_compare_config.json`` → ``in_cabinet_debug``
        (блок в окне настроек ⚙️); True — принудительно включить; False — выключить;
        str/Path — путь к ``ds_mto_compare_debug.log`` (trace — рядом, ``*_in_cabinet_trace.txt``).
        Файлы: ``ds_in_cabinet_trace.txt``, ``ds_mto_compare_debug.log`` (в папке Excel).
    ds_data_override: already loaded/prepared DS rows. Used by grouped MVP mode;
        None keeps the legacy load path.
    save_excel: write the standard DS vs MTO Excel result.
    return_data: return processed DS rows after comparison.
    output_prefix: Excel file prefix for the standard writer when ``save_excel`` is True.
    show_internal_column_names: If True, append internal program column names to
        Excel headers; None = value from ``ds_compare_config.json``.
    """
    print("Analyzing DS Specification v3\n", ds_path)
    if mto_path is None:
        mto_path = resolve_mto_path_from_config()
    if replacement_table_file is None:
        replacement_table_file = r"\\bcc\eng\PrDoc\377_НИПИГАЗ\АГХК\КСБ\РД\Таблички_графики\14_сравнение с ДС\_замены кодов\Code_Replacement_Table.xlsx"

    # Загрузка данных
    if ds_data_override is None:
        ds_data = load_ds_data(ds_path, summ_ds=summ_ds, force_update=True)
    else:
        ds_data = ds_data_override

    cfg = load_ds_compare_config()
    output_cfg = normalize_ds_vs_mto_output(cfg.get("ds_vs_mto_output"))
    if show_internal_column_names is None:
        show_internal_column_names = bool(
            output_cfg.get("show_internal_column_names", False)
        )
    if ds_vs_mto_xw_columns is None:
        ds_vs_mto_xw_columns = column_config_from_output_settings(output_cfg)
    if flat_mto_structure is None:
        flat_mto_structure = bool(cfg.get("flat_mto_structure", False))
    if flat_mto_structure:
        print("DS vs MTO: flat_mto_structure=True (как step2 flat_mto_structure)")

    _use = bool(cfg.get("mto_use_cache", True)) if mto_use_cache is None else mto_use_cache
    _force = bool(cfg.get("mto_force_update", False)) if mto_force_update is None else mto_force_update

    mto_dict, spec_dict, mto_load_audit = load_mto_data(
        ds_data,
        mto_path,
        flat_structure=flat_mto_structure,
        mto_use_cache=_use,
        mto_force_update=_force,
        ds_source_path=ds_path,
    )
    replacement_table = load_replacement_table(replacement_table_file,print_replacement_table_info=False)

    # Проверка МТО
    mto_check_obj = MtoCheck(mto_dict)
    mto_check_obj.check()

    # Проверка на недопустимые значения в колонке количества (формулы Excel, текст и т.п.)
    validate_and_exit_if_errors(ds_data, mto_dict)

    out_dir = utils.path.get_path_from_file_path(ds_path)

    debug_enabled = compare_debug_log_requested(cfg, compare_debug_log)
    cabinet_trace: CabinetInCabinetTrace | None = None
    debug_log_path: Path | None = None
    if debug_enabled:
        if compare_debug_log is True or compare_debug_log is None:
            trace_path = Path(out_dir) / "ds_in_cabinet_trace.txt"
            debug_log_path = Path(out_dir) / "ds_mto_compare_debug.log"
        elif isinstance(compare_debug_log, (str, Path)) and str(compare_debug_log).strip():
            debug_log_path = Path(compare_debug_log)
            trace_path = debug_log_path.parent / (
                debug_log_path.stem + "_in_cabinet_trace.txt"
            )
        else:
            trace_path = Path(out_dir) / "ds_in_cabinet_trace.txt"
            debug_log_path = Path(out_dir) / "ds_mto_compare_debug.log"

        dbg_cfg = normalize_in_cabinet_debug(cfg.get("in_cabinet_debug"))
        watch = InCabinetWatchFilters.from_debug_dict(dbg_cfg)
        cabinet_trace = CabinetInCabinetTrace(trace_path, watch)
        cabinet_trace.open()
        cabinet_trace.log_mto_load_audit(spec_dict, mto_load_audit)
        cabinet_trace.log_mto_inventory(mto_dict)
        for spec_name, audit in mto_load_audit.items():
            if not spec_key_matches_watch(spec_name, watch):
                continue
            path = spec_dict.get(spec_name, "")
            rows = mto_dict.get(spec_name, [])
            cabinet_trace.log_cache_vs_fresh_in_cabinet(
                spec_name,
                path,
                rows,
                from_cache=bool(audit.get("from_cache")),
            )
        cabinet_trace.log_ds_inventory(ds_data, stage="before compare")

    # Сравнение с последовательным распределением
    comparator = DsMtoComparator()
    if debug_log_path is not None:
        comparator.debug_log_path = debug_log_path
        comparator.debug = True
        dbg_cfg = normalize_in_cabinet_debug(cfg.get("in_cabinet_debug"))
        codes = InCabinetWatchFilters.from_debug_dict(dbg_cfg).codes
        if len(codes) == 1:
            comparator.debug_code = next(iter(codes))
        elif codes:
            comparator.debug_code = sorted(codes)[0]
        comparator.cabinet_trace = cabinet_trace
    comparator.compare_specifications(ds_data, mto_dict, replacement_table)

    if cabinet_trace is not None:
        cabinet_trace.log_ds_inventory(ds_data, stage="after compare (before Excel)")
        cabinet_trace.close()
        print(f"IN_CABINET trace: {cabinet_trace.path.resolve()}")

    # # Проверка кодов с гугл базой
    # code_base_data_std = load_base()
    # ds_codr_vs_base_google(ds_data, code_base_data_std)

    # Визуализация результатов (openpyxl из шаблона; затем xlsxwriter — быстрее, компактнее)
    # STDTable.to_excel_ds_vs_mto_spec(
    #     ds_data,
    #     out_dir,
    #     "РОБОТ_СРАВНЕНИЕ_",
    #     show_row_type=False,
    #     open_folder=True,
    # )
    if save_excel:
        save_ds_vs_mto_excel_per_export_mode(
            ds_data,
            out_dir,
            output_prefix,
            suffix="_xw",
            column_config=ds_vs_mto_xw_columns,
            show_internal_column_names=show_internal_column_names,
            excel_export_mode=str(output_cfg.get("excel_export_mode", "both")),
            expand_aggregated_replacement_codes=bool(
                output_cfg.get("expand_aggregated_replacement_codes", True)
            ),
        )
    # ПРОВЕРКА
    check_comparison_results(
        ds_data, spec_dict, mto_use_cache=_use, mto_force_update=_force
    )
    if return_data:
        return ds_data
    return None

if __name__ in {"__main__"}:

    ds_path = r"\\bcc\eng\PrDoc\377_НИПИГАЗ\АГХК\КСБ\РД\Таблички_графики\14_сравнение с ДС\TEST\__TEST.xlsx"
    # mto_path = r"\\bcc\eng\PrDoc\377_НИПИГАЗ\АГХК\КСБ\РД\Таблички_графики\14_сравнение с ДС\TEST"

    analyze_ds_specification(ds_path)
