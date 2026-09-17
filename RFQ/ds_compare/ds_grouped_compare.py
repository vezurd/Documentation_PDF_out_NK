"""Grouped DS vs MTO MVP orchestrator.

This module builds a grouped DS view, passes it through the existing
``analyze_ds_specification`` pipeline, optionally loads RFQ TPK data,
and writes a separate grouped Excel file.
"""

from __future__ import annotations

import os
from collections import defaultdict
from datetime import datetime
from typing import Iterable

import utils.path
from base.base_classes import RowStd, RowType
from base.tables_columns import (
    ANNOTATION_2,
    ANNOTATION_MTO,
    CODE,
    CODE_2,
    DS_NAME,
    DS_NUMBER,
    DS_RFQ,
    DS_SPECIFICATION,
    DS_SYSTEM,
    DS_TITLE,
    IN_CABINET,
    NAME,
    NAME_2,
    NUMBERS_2,
    ROW_TYPE,
    TAGS_2,
    TYPE_MARK,
    TYPE_MARK_2,
    UNITS,
    UNITS_2,
    VALUES,
    VALUES_2,
    VENDOR_2,
)
from RFQ.ds_compare.ds_compare_config import (
    GroupedCompareDict,
    load_ds_compare_config,
    normalize_ds_vs_mto_output,
    normalize_grouped_compare,
    normalize_gui_paths,
    resolve_mto_path_from_config,
)
from RFQ.ds_compare.ds_start_init import analyze_ds_specification
from RFQ.ds_compare.ds_sources_preflight import (
    finalize_packing_manifest,
    refresh_mto_sources,
    refresh_sources_before_grouped,
)
from RFQ.ds_compare.ds_vs_mto_excel_columns import column_config_from_output_settings
from RFQ.ds_compare.ds_vs_mto_excel_xlsxwriter import (
    save_ds_vs_mto_excel_per_export_mode,
)
from RFQ.ds_compare.get_mto_list_from_ds import get_mto_list_from_ds
from RFQ.ds_compare.grouped_compare_cache import (
    build_grouped_compare_extra_key,
    grouped_compare_cache_file_path,
    load_grouped_compare_cache,
    save_grouped_compare_cache,
)
from RFQ.ds_compare.ds_quantity_parse import quantity_as_float
from RFQ.ds_compare.support_finctions import load_ds_data, load_rfq_tpk_data
from RFQ.packing_list_provider import PackingIssue, load_packing_dataset

_last_rfq_quantity_audit: object | None = None
_last_packing_compare_audit: object | None = None

_DEFAULT_REPLACEMENT_TABLE = (
    r"\\bcc\eng\PrDoc\377_НИПИГАЗ\АГХК\КСБ\РД\Таблички_графики\14_сравнение с ДС\_замены кодов"
    r"\Code_Replacement_Table.xlsx"
)
GROUP_KEY_TO_COLUMN: dict[str, str] = {
    "code": CODE,
    "title": DS_TITLE,
    "system": DS_SYSTEM,
    "ds_name": DS_NAME,
}
MTO_GROUP_KEY_TO_COLUMN: dict[str, str] = {
    "code": CODE_2,
    "title": DS_TITLE,
    "system": DS_SYSTEM,
    "ds_name": DS_NAME,
}

DEFAULT_GROUP_KEYS: tuple[str, ...] = ("title", "system", "code")


def get_last_rfq_quantity_audit():
    """Last :class:`~RFQ.ds_compare.ds_rfq_quantity_audit.RfqQuantityAudit` from grouped+RFQ run."""
    return _last_rfq_quantity_audit


def get_last_packing_compare_audit():
    """Return the packing audit from the last grouped run."""
    return _last_packing_compare_audit


_REPRESENTATIVE_COLUMNS: tuple[str, ...] = (
    DS_SPECIFICATION,
    DS_RFQ,
    NAME,
    TYPE_MARK,
)
_MTO_MERGE_TEXT_COLUMNS: tuple[str, ...] = (
    TAGS_2,
    NUMBERS_2,
    ANNOTATION_MTO,
    IN_CABINET,
    UNITS_2,
)
_MTO_REPRESENTATIVE_COLUMNS: tuple[str, ...] = (
    NAME_2,
    TYPE_MARK_2,
    VENDOR_2,
)


def group_keys_from_settings(settings: GroupedCompareDict | dict | None) -> list[str]:
    """Return grouped-compare key ids from config settings."""
    normalized = normalize_grouped_compare(settings)
    keys: list[str] = []
    if normalized.get("group_by_title"):
        keys.append("title")
    if normalized.get("group_by_system"):
        keys.append("system")
    if normalized.get("group_by_code"):
        keys.append("code")
    if normalized.get("group_by_ds_name"):
        keys.append("ds_name")
    return keys or list(DEFAULT_GROUP_KEYS)


def _unique_nonempty(values: Iterable[object]) -> list[str]:
    out: list[str] = []
    seen: set[str] = set()
    for value in values:
        text = str(value or "").strip()
        if not text or text in seen:
            continue
        seen.add(text)
        out.append(text)
    return out


def _short_join(values: list[str], *, limit: int = 12) -> str:
    if len(values) <= limit:
        return "/".join(values)
    return "/".join(values[:limit]) + f"/...(+{len(values) - limit})"


def _group_key(row: RowStd, group_keys: list[str]) -> tuple[str, ...]:
    parts: list[str] = []
    for key in group_keys:
        column = GROUP_KEY_TO_COLUMN[key]
        parts.append(str(row.get_value(column) or "").strip())
    return tuple(parts)


def _mto_group_key(row: RowStd, group_keys: list[str]) -> tuple[str, ...]:
    parts: list[str] = []
    for key in group_keys:
        column = MTO_GROUP_KEY_TO_COLUMN[key]
        parts.append(str(row.get_value(column) or "").strip())
    return tuple(parts)


def _format_source_summary(rows: list[RowStd]) -> str:
    qty_by_ds: dict[str, float] = defaultdict(float)
    row_numbers_by_ds: dict[str, list[str]] = defaultdict(list)
    for row in rows:
        ds_name = str(row.get_value(DS_NAME) or "?").strip() or "?"
        qty_by_ds[ds_name] += quantity_as_float(row.get_value(VALUES))
        row_number = str(row.get_value(DS_NUMBER) or "").strip()
        if row_number and row_number not in row_numbers_by_ds[ds_name]:
            row_numbers_by_ds[ds_name].append(row_number)

    lines = ["Grouped DS rows:"]
    for ds_name in sorted(qty_by_ds):
        nums = row_numbers_by_ds.get(ds_name) or []
        nums_text = f" rows={', '.join(nums)}" if nums else ""
        lines.append(f"- {ds_name}: VALUES={qty_by_ds[ds_name]:g}{nums_text}")
    return "\n".join(lines)


def _format_mto_source_summary(rows: list[RowStd]) -> str:
    lines = ["Grouped MTO-only rows:"]
    for row in rows:
        number = str(row.get_value(NUMBERS_2) or "").strip()
        code = str(row.get_value(CODE_2) or "").strip()
        qty = quantity_as_float(row.get_value(VALUES_2))
        cabinet = str(row.get_value(IN_CABINET) or "").strip()
        suffix = f", IN_CABINET={cabinet}" if cabinet else ""
        lines.append(f"- MTO #{number or '?'}: CODE={code}, VALUES_2={qty:g}{suffix}")
    return "\n".join(lines)


def _cell_values_as_list(row: RowStd, col_name: str) -> list[str]:
    value = row.get_value(col_name)
    if isinstance(value, list):
        return [str(v).strip() for v in value if str(v).strip()]
    text = str(value or "").strip()
    if not text:
        return []
    parts = [p.strip() for p in text.replace("/", ",").split(",")]
    return [p for p in parts if p]


def _merge_unique_cell_values(rows: list[RowStd], col_name: str) -> list[str]:
    out: list[str] = []
    seen: set[str] = set()
    for row in rows:
        for value in _cell_values_as_list(row, col_name):
            if value in seen:
                continue
            seen.add(value)
            out.append(value)
    return out


def _copy_cell_formatting_from(source: RowStd, target: RowStd) -> None:
    """Copy cell colors/comments from ``source`` to ``target``."""
    for col_name, source_el in source.el.items():
        if col_name not in target.el:
            continue
        target.el[col_name].color = source_el.color
        target.el[col_name].comment = source_el.comment


def group_ds_rows(ds_rows: list[RowStd], group_keys: list[str]) -> list[RowStd]:
    """Group DS position rows and sum ``VALUES`` within each selected key."""
    normalized_keys = [k for k in group_keys if k in GROUP_KEY_TO_COLUMN]
    if not normalized_keys:
        normalized_keys = list(DEFAULT_GROUP_KEYS)

    grouped: dict[tuple[str, ...], list[RowStd]] = {}
    order: list[tuple[str, ...]] = []
    for row in ds_rows:
        if not isinstance(row, RowStd) or row.row_type != RowType.position_row:
            continue
        key = _group_key(row, normalized_keys)
        if key not in grouped:
            grouped[key] = []
            order.append(key)
        grouped[key].append(row)

    out_rows: list[RowStd] = []
    for key in order:
        rows = grouped[key]
        first = rows[0]
        out_row = RowStd.get_row_copy(first, first.t_com)
        out_row.row_type = RowType.position_row
        out_row.el[ROW_TYPE].value = RowType.position_row

        total_values = sum(quantity_as_float(row.get_value(VALUES)) for row in rows)
        out_row.el[VALUES].value = total_values
        out_row.el[VALUES].comment = _format_source_summary(rows)

        ds_names = _unique_nonempty(row.get_value(DS_NAME) for row in rows)
        ds_numbers = _unique_nonempty(row.get_value(DS_NUMBER) for row in rows)
        if "ds_name" not in normalized_keys and ds_names:
            out_row.el[DS_NAME].value = _short_join(ds_names)
        if ds_numbers:
            out_row.el[DS_NUMBER].value = _short_join(ds_numbers, limit=20)
            out_row.el[DS_NUMBER].comment = out_row.el[VALUES].comment
        units_values = _unique_nonempty(row.get_value(UNITS) for row in rows)
        if units_values:
            out_row.el[UNITS].value = "/".join(units_values)
            if len(units_values) > 1:
                out_row.el[UNITS].comment = (
                    "Grouped DS rows contain different units:\n"
                    + "\n".join(f"- {v}" for v in units_values)
                )

        for col_name in _REPRESENTATIVE_COLUMNS:
            values = _unique_nonempty(row.get_value(col_name) for row in rows)
            if len(values) > 1:
                out_row.el[col_name].comment = (
                    "Grouped rows contain different values:\n"
                    + "\n".join(f"- {v}" for v in values)
                )

        out_rows.append(out_row)
    return out_rows


def _is_mto_only_new_position(row: RowStd) -> bool:
    if not isinstance(row, RowStd) or row.row_type != RowType.position_row:
        return False
    status = str(row.get_value(ANNOTATION_2) or "").strip()
    return status == "Новая позиция" and not row.get_value(CODE) and bool(row.get_value(CODE_2))


def group_mto_only_new_positions(rows: list[RowStd], group_keys: list[str]) -> list[RowStd]:
    """Group right-side MTO-only rows produced by ``DsMtoComparator`` post-pass."""
    normalized_keys = [k for k in group_keys if k in MTO_GROUP_KEY_TO_COLUMN]
    if not normalized_keys:
        normalized_keys = list(DEFAULT_GROUP_KEYS)

    grouped: dict[tuple[str, ...], list[RowStd]] = {}
    first_index_by_key: dict[tuple[str, ...], int] = {}
    for index, row in enumerate(rows):
        if not _is_mto_only_new_position(row):
            continue
        key = _mto_group_key(row, normalized_keys)
        if key not in grouped:
            grouped[key] = []
            first_index_by_key[key] = index
        grouped[key].append(row)

    if not grouped:
        return rows

    replacement_by_first_index: dict[int, RowStd] = {}
    remove_ids: set[int] = set()
    for key, grouped_rows in grouped.items():
        if len(grouped_rows) <= 1:
            continue
        first = grouped_rows[0]
        out_row = RowStd.get_row_copy(first, first.t_com)
        _copy_cell_formatting_from(first, out_row)
        out_row.row_type = RowType.position_row
        out_row.el[ROW_TYPE].value = RowType.position_row

        total_values = sum(quantity_as_float(row.get_value(VALUES_2)) for row in grouped_rows)
        out_row.el[VALUES_2].value = total_values
        out_row.el[VALUES_2].comment = _format_mto_source_summary(grouped_rows)

        for col_name in _MTO_MERGE_TEXT_COLUMNS:
            values = _merge_unique_cell_values(grouped_rows, col_name)
            if values:
                sep = "/" if col_name in (IN_CABINET, UNITS_2) else ", "
                out_row.el[col_name].value = sep.join(values)
                out_row.el[col_name].comment = out_row.el[VALUES_2].comment

        for col_name in _MTO_REPRESENTATIVE_COLUMNS:
            values = _unique_nonempty(row.get_value(col_name) for row in grouped_rows)
            if len(values) > 1:
                out_row.el[col_name].comment = (
                    "Grouped MTO-only rows contain different values:\n"
                    + "\n".join(f"- {v}" for v in values)
                )

        replacement_by_first_index[first_index_by_key[key]] = out_row
        remove_ids.update(id(row) for row in grouped_rows[1:])

    if not replacement_by_first_index:
        return rows

    out_rows: list[RowStd] = []
    for index, row in enumerate(rows):
        if id(row) in remove_ids:
            continue
        out_rows.append(replacement_by_first_index.get(index, row))
    return out_rows


def _rfq_file_stem(rfq_path: str) -> str:
    """Return RFQ xlsx file name without extension."""
    stem = utils.path.get_file_name_from_full_file_path(rfq_path)
    if stem.lower().endswith(".xlsx"):
        return stem[:-5]
    return stem


def _build_grouped_rfq_output_file_name(rfq_path: str) -> str:
    """Build grouped DS vs MTO vs RFQ Excel name: ``<RFQ>_ДС_МТО__<ГГГГ.ММ.ДД_ЧЧ.ММ>.xlsx``."""
    stamp = datetime.now().strftime("%Y.%m.%d_%H.%M")
    return f"{_rfq_file_stem(rfq_path)}_ДС_МТО__{stamp}.xlsx"


def _resolve_rfq_path(rfq_path: str | None, cfg: dict) -> str | None:
    """Return RFQ file path from argument or GUI config."""
    explicit = str(rfq_path or "").strip()
    if explicit:
        return explicit
    gui_paths = normalize_gui_paths(cfg.get("gui_paths"))
    from_gui = str(gui_paths.get("last_rfq_file", "")).strip()
    return from_gui or None


class GroupedDsMtoRfqOnlyError(ValueError):
    """RFQ-only run requires a valid grouped DS vs MTO compare cache."""


def _log_grouped_row_counts(ds_rows: list[RowStd], grouped_rows: list[RowStd], group_keys: list[str]) -> None:
    print(
        "Grouped DS vs MTO vs RFQ: "
        f"{len(ds_rows)} source rows -> {len(grouped_rows)} grouped rows "
        f"(keys={', '.join(group_keys)})"
    )


def analyze_grouped_ds_specification(
    ds_path: str,
    mto_path: str | None = None,
    rfq_path: str | None = None,
    replacement_table_file: str | None = None,
    summ_ds: bool = True,
    group_keys: list[str] | None = None,
    flat_mto_structure: bool | None = None,
    mto_use_cache: bool | None = None,
    mto_force_update: bool | None = None,
    compare_debug_log: bool | str | None = None,
    rfq_only: bool = False,
) -> str | None:
    """Run grouped DS vs MTO MVP and save a separate grouped Excel file.

    Args:
        rfq_only: If True, skip ``DsMtoComparator`` and require a valid
            ``grouped_cmp_*.cache`` (see ``GroupedDsMtoRfqOnlyError``).
    """
    global _last_packing_compare_audit, _last_rfq_quantity_audit

    # Do not leak an audit from an earlier GUI job when the current run has no
    # RFQ or fails before packing enrichment.
    _last_rfq_quantity_audit = None
    _last_packing_compare_audit = None

    cfg = load_ds_compare_config()
    grouped_cfg = normalize_grouped_compare(cfg.get("grouped_compare"))
    if group_keys is None:
        group_keys = group_keys_from_settings(grouped_cfg)
    if mto_path is None:
        mto_path = resolve_mto_path_from_config(cfg)
    if replacement_table_file is None:
        replacement_table_file = _DEFAULT_REPLACEMENT_TABLE
    if flat_mto_structure is None:
        flat_mto_structure = bool(cfg.get("flat_mto_structure", False))
    if mto_force_update is None:
        mto_force_update = bool(cfg.get("mto_force_update", False))

    preflight = refresh_sources_before_grouped(
        ds_path=ds_path,
        rfq_only=rfq_only,
        cfg=cfg,
    )
    ds_path = preflight.effective_ds_path
    source_manifest = preflight.manifest

    resolved_rfq_path = _resolve_rfq_path(rfq_path, cfg)
    excel_out_dir = utils.path.get_path_from_file_path(ds_path)
    grouped_output_file_name: str | None = None

    cache_extra_key = build_grouped_compare_extra_key(
        group_keys=group_keys,
        flat_mto_structure=bool(flat_mto_structure),
        group_mto_new_positions=bool(grouped_cfg.get("group_mto_new_positions", True)),
        mto_path=str(mto_path or ""),
        replacement_table_file=str(replacement_table_file or ""),
    )
    use_result_cache = bool(grouped_cfg.get("cache_ds_mto_result", False))
    try_cache_read = use_result_cache and not bool(mto_force_update)

    ds_rows: list[RowStd] | None = None
    grouped_rows: list[RowStd] | None = None
    spec_dict: dict | None = None
    compared_rows: list[RowStd] | None = None
    mto_preflight_done = False

    def _ensure_grouped_context(*, force_update: bool) -> tuple[list[RowStd], dict]:
        nonlocal ds_rows, grouped_rows, spec_dict, mto_preflight_done
        if grouped_rows is not None and spec_dict is not None:
            return grouped_rows, spec_dict
        ds_rows = load_ds_data(ds_path, summ_ds=summ_ds, force_update=force_update)
        grouped_rows = group_ds_rows(ds_rows, group_keys)
        _log_grouped_row_counts(ds_rows, grouped_rows, group_keys)
        spec_dict = get_mto_list_from_ds(
            grouped_rows,
            mto_path,
            flat_structure=bool(flat_mto_structure),
            ds_source_path=ds_path,
        )
        if not mto_preflight_done:
            refresh_mto_sources(
                spec_dict,
                source_manifest,
                rfq_only=rfq_only,
            )
            mto_preflight_done = True
        return grouped_rows, spec_dict

    def _try_load_compare_cache() -> list[RowStd] | None:
        if not try_cache_read:
            return None
        _, sd = _ensure_grouped_context(force_update=False)
        cached = load_grouped_compare_cache(
            ds_path,
            cache_extra_key,
            replacement_table_file=str(replacement_table_file or ""),
            spec_dict=sd,
        )
        if cached is not None:
            print(
                "Grouped DS vs MTO: кэш ДС↔MTO — пропуск compare "
                "(ДС без force_update, MTO compare не выполнялся)"
            )
        return cached

    if rfq_only:
        if not resolved_rfq_path or not os.path.isfile(resolved_rfq_path):
            raise GroupedDsMtoRfqOnlyError(
                "RFQ-only: укажите существующий файл RFQ (TPK xlsx)."
            )
        if not use_result_cache:
            raise GroupedDsMtoRfqOnlyError(
                "RFQ-only: включите «Кэшировать результат сравнения ДС vs MTO» "
                "в настройках (вкладка «Настройки»)."
            )
        if bool(mto_force_update):
            raise GroupedDsMtoRfqOnlyError(
                "RFQ-only: отключите «Принудительно обновить кэш MTO» для этого запуска."
            )
        if not grouped_compare_cache_file_path(ds_path, cache_extra_key).exists():
            raise GroupedDsMtoRfqOnlyError(
                "RFQ-only: файл кэша ДС vs MTO не найден. Сначала выполните полное "
                "«ДС vs MTO vs RFQ» с включённым кэшем результата сравнения."
            )
        compared_rows = _try_load_compare_cache()
        if compared_rows is None:
            raise GroupedDsMtoRfqOnlyError(
                "RFQ-only: нет актуального кэша ДС vs MTO для этого файла ДС, "
                "пресета МТО и ключей группировки. Сначала выполните полное "
                "«ДС vs MTO vs RFQ» (достаточно один раз без смены ДС/MTO)."
            )
    else:
        compared_rows = _try_load_compare_cache()

    if compared_rows is None:
        if grouped_rows is None:
            _ensure_grouped_context(force_update=not try_cache_read)
        assert grouped_rows is not None and spec_dict is not None

        compared_rows = analyze_ds_specification(
            ds_path,
            mto_path=mto_path,
            replacement_table_file=replacement_table_file,
            summ_ds=summ_ds,
            flat_mto_structure=flat_mto_structure,
            mto_use_cache=mto_use_cache,
            mto_force_update=mto_force_update,
            compare_debug_log=compare_debug_log,
            ds_data_override=grouped_rows,
            save_excel=False,
            return_data=True,
        )

        if grouped_cfg.get("group_mto_new_positions", True):
            before_count = len(compared_rows or grouped_rows)
            compared_rows = group_mto_only_new_positions(
                compared_rows or grouped_rows, group_keys
            )
            after_count = len(compared_rows)
            if after_count != before_count:
                print(
                    "Grouped DS vs MTO: "
                    f"MTO-only rows post-grouped {before_count} -> {after_count}"
                )

        if use_result_cache and compared_rows is not None:
            save_grouped_compare_cache(
                ds_path,
                cache_extra_key,
                compared_rows,
                replacement_table_file=str(replacement_table_file or ""),
                spec_dict=spec_dict,
            )

    if resolved_rfq_path:
        if not os.path.isfile(resolved_rfq_path):
            print(f"RFQ TPK: файл не найден — {resolved_rfq_path}")
        else:
            excel_out_dir = utils.path.get_path_from_file_path(resolved_rfq_path)
            grouped_output_file_name = _build_grouped_rfq_output_file_name(resolved_rfq_path)
            rfq_rows = load_rfq_tpk_data(
                resolved_rfq_path,
                dump_non_position_dir=excel_out_dir,
            )
            from RFQ.ds_compare.ds_rfq_grouped_compare import compare_ds_mto_with_rfq

            compared_rows, rfq_stats = compare_ds_mto_with_rfq(
                compared_rows or grouped_rows,
                rfq_rows,
                group_keys,
            )
            from RFQ.ds_compare.ds_rfq_quantity_audit import compute_rfq_quantity_audit

            _last_rfq_quantity_audit = compute_rfq_quantity_audit(
                resolved_rfq_path,
                rfq_rows,
                compared_rows or grouped_rows,
            )

            print(
                "Grouped DS vs MTO vs RFQ: "
                f"RFQ grouped={rfq_stats.grouped_rfq_rows}, "
                f"matched={rfq_stats.rfq_matched}, "
                f"not_in_rfq={rfq_stats.rfq_not_found}, "
                f"rfq_only={rfq_stats.rfq_only_added}, "
                f"rfq_dup_skip={rfq_stats.rfq_duplicate_skipped}, "
                f"status: совпадение={rfq_stats.status_match}, "
                f"увеличение={rfq_stats.status_increase}, "
                f"недопоставка={rfq_stats.status_shortfall}"
            )
            print(_last_rfq_quantity_audit.format_detail())

    from RFQ.ds_compare.ds_packing_grouped_compare import (
        compare_grouped_rows_with_packing,
        save_packing_compare_report,
    )
    from RFQ.ds_compare.ds_units_normalize import highlight_compared_units_columns

    final_rows = compared_rows or grouped_rows or []
    packing_dataset = load_packing_dataset()
    finalize_packing_manifest(source_manifest, packing_dataset)
    final_rows, packing_audit = compare_grouped_rows_with_packing(
        final_rows,
        packing_dataset,
    )
    units_stats = highlight_compared_units_columns(final_rows)
    packing_audit.stats.units_match = units_stats.all_match
    packing_audit.stats.units_mismatch = units_stats.mismatch
    try:
        save_packing_compare_report(packing_audit, excel_out_dir)
    except OSError as exc:
        report_issue = PackingIssue(
            code="compare_report_write",
            message=f"Не удалось записать отчёт сравнения УЛ ({exc})",
            action="Проверьте доступ на запись в папку итогового Excel",
        )
        packing_audit.issues.append(report_issue)
        print(f"УЛ WARNING: {report_issue.format_line()}")
    _last_packing_compare_audit = packing_audit
    print(packing_audit.format_short())
    print(
        "Grouped units DS/MTO/RFQ/UL: "
        f"match={units_stats.all_match}, mismatch={units_stats.mismatch}"
    )

    output_cfg = normalize_ds_vs_mto_output(cfg.get("ds_vs_mto_output"))
    return save_ds_vs_mto_excel_per_export_mode(
        final_rows,
        excel_out_dir,
        "РОБОТ_СРАВНЕНИЕ_GROUPED",
        suffix="_xw",
        output_file_name=grouped_output_file_name,
        column_config=column_config_from_output_settings(output_cfg),
        show_internal_column_names=bool(
            output_cfg.get("show_internal_column_names", False)
        ),
        excel_export_mode=str(output_cfg.get("excel_export_mode", "both")),
        expand_aggregated_replacement_codes=bool(
            output_cfg.get("expand_aggregated_replacement_codes", True)
        ),
        source_manifest=source_manifest,
    )