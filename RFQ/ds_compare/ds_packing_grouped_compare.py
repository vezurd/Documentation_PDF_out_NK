"""Grouped DS/RFQ delivery comparison with packing-list quantities."""

from __future__ import annotations

import math
from collections import defaultdict
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Iterable

from base.base_classes import CheckElement, RowStd, RowType, TableComments
from base.tables_columns import (
    CODE,
    CODE_2,
    DS_SYSTEM,
    DS_TITLE,
    NAME,
    RFQ_CODE,
    RFQ_VALUES,
    ROW_TYPE,
    TAGS,
    TYPE_MARK,
    UL_CODE,
    UL_COMPARE_STATUS,
    UL_DATA_STATUS,
    UL_NAME,
    UL_ORDERED_VALUES,
    UL_REMAINING_VALUES,
    UL_SOURCE_FILES,
    UL_TAGS,
    UL_TYPE_MARK,
    UL_UNITS,
    UL_VALUES,
    UL_VENDOR,
    UNITS,
    VALUES,
    VENDOR,
)
from RFQ.ds_compare.ds_quantity_parse import parse_quantity_strict, quantity_as_float
from RFQ.ds_compare.ds_units_normalize import normalize_units_text
from RFQ.packing_list_provider import (
    PackingDataset,
    PackingIssue,
    PackingIssueSeverity,
    PackingQualityLevel,
    packing_codes_for_row,
    packing_match_key,
    packing_row_key,
    packing_row_source,
)
from utils.colors import Color

_FLOAT_EPS = 1e-9
_MISS_SAMPLE_LIMIT = 10

STATUS_COMPLETE = "Поставка комплектна"
STATUS_SHORTFALL = "Недопоставка"
STATUS_OVERDELIVERY = "Перепоставка"
STATUS_NOT_IN_PACKING = "Нет в УЛ"
STATUS_DATA_PROBLEM = "Проблема данных УЛ"
STATUS_UNAVAILABLE = "УЛ недоступны"
STATUS_PACKING_ONLY = "Только в УЛ"
STATUS_ACCOUNTED_ABOVE = "УЛ учтён выше"

DATA_OK = "OK"
DATA_PARTIAL = "Частичные данные УЛ"
DATA_ROW_PROBLEM = "Проблема строки УЛ"
DATA_UNAVAILABLE = "УЛ недоступны"

_UL_COLUMNS: tuple[str, ...] = (
    UL_ORDERED_VALUES,
    UL_VALUES,
    UL_UNITS,
    UL_REMAINING_VALUES,
    UL_COMPARE_STATUS,
    UL_DATA_STATUS,
    UL_CODE,
    UL_SOURCE_FILES,
    UL_NAME,
    UL_TYPE_MARK,
    UL_VENDOR,
    UL_TAGS,
)


@dataclass
class PackingCompareStats:
    """Counters from one grouped packing enrichment."""

    compared_position_rows: int = 0
    packing_rows: int = 0
    grouped_packing_keys: int = 0
    matched_keys: int = 0
    matched_rows: int = 0
    not_in_packing: int = 0
    packing_only_added: int = 0
    duplicate_target_rows: int = 0
    invalid_quantity_rows: int = 0
    invalid_key_rows: int = 0
    complete: int = 0
    shortfall: int = 0
    overdelivery: int = 0
    data_problem: int = 0
    units_match: int = 0
    units_mismatch: int = 0


@dataclass
class PackingCompareAudit:
    """User-facing result and diagnostics for packing enrichment."""

    dataset: PackingDataset
    stats: PackingCompareStats
    issues: list[PackingIssue] = field(default_factory=list)
    miss_samples: list[str] = field(default_factory=list)
    report_path: str | None = None

    @property
    def quality(self) -> PackingQualityLevel:
        """Combined provider and matcher quality."""
        if self.dataset.quality == PackingQualityLevel.UNAVAILABLE:
            return PackingQualityLevel.UNAVAILABLE
        return PackingQualityLevel.PARTIAL if self.issues else PackingQualityLevel.OK

    def format_short(self) -> str:
        """One-line GUI/job-monitor summary."""
        s = self.stats
        quality = {
            PackingQualityLevel.OK: "OK",
            PackingQualityLevel.PARTIAL: "ВНИМАНИЕ: частичные данные",
            PackingQualityLevel.UNAVAILABLE: "ОШИБКА: УЛ недоступны",
        }[self.quality]
        report = f" Отчёт: {self.report_path}." if self.report_path else ""
        return (
            f"УЛ [{quality}]: matched={s.matched_keys}, нет_в_УЛ={s.not_in_packing}, "
            f"только_УЛ={s.packing_only_added}, проблем={len(self.issues)}.{report}"
        )


@dataclass
class _PackingGroup:
    key: tuple[str, str, str]
    rows: list[RowStd] = field(default_factory=list)
    valid_quantity: float = 0.0
    valid_quantity_rows: int = 0
    units: list[str] = field(default_factory=list)
    issues: list[PackingIssue] = field(default_factory=list)
    sources: list[str] = field(default_factory=list)

    @property
    def first(self) -> RowStd:
        return self.rows[0]


def _unique_text(values: Iterable[object]) -> list[str]:
    result: list[str] = []
    for value in values:
        if isinstance(value, list):
            candidates = [str(item).strip() for item in value]
        else:
            candidates = [str(value or "").strip()]
        for text in candidates:
            if text and text not in result:
                result.append(text)
    return result


def _join_values(values: Iterable[object], *, limit: int = 20) -> str:
    unique = _unique_text(values)
    if len(unique) <= limit:
        return "; ".join(unique)
    return "; ".join(unique[:limit]) + f"; …(+{len(unique) - limit})"


def _source_text(row: RowStd) -> str:
    source_file, sheet, excel_row = packing_row_source(row)
    parts = [part for part in (source_file, sheet) if part]
    if excel_row is not None:
        parts.append(f"строка {excel_row}")
    return " · ".join(parts) or "(источник не записан)"


def _row_issue(
    row: RowStd,
    *,
    code: str,
    message: str,
    field_name: str = "",
    raw_value: object = None,
    severity: PackingIssueSeverity = PackingIssueSeverity.ERROR,
    action: str = "",
) -> PackingIssue:
    source_file, sheet, excel_row = packing_row_source(row)
    return PackingIssue(
        code=code,
        message=message,
        severity=severity,
        source_file=source_file,
        sheet=sheet,
        excel_row=excel_row,
        field=field_name,
        raw_value=raw_value,
        action=action,
    )


def _build_packing_groups(
    dataset: PackingDataset,
    stats: PackingCompareStats,
) -> tuple[dict[tuple[str, str, str], _PackingGroup], list[RowStd], list[PackingIssue]]:
    groups: dict[tuple[str, str, str], _PackingGroup] = {}
    invalid_key_rows: list[RowStd] = []
    issues: list[PackingIssue] = []

    for row in dataset.rows:
        stats.packing_rows += 1
        key = packing_row_key(row)
        if not all(key):
            missing = [
                name
                for name, value in zip(("DS_TITLE", "DS_SYSTEM", "CODE"), key)
                if not value
            ]
            issue = _row_issue(
                row,
                code="packing_key_missing",
                message=f"Нельзя сопоставить строку: пустые части ключа {', '.join(missing)}",
                field_name=",".join(missing),
                action="Исправьте спецификацию/титул/марку/код в исходном УЛ",
            )
            issues.append(issue)
            invalid_key_rows.append(row)
            stats.invalid_key_rows += 1
            continue

        group = groups.setdefault(key, _PackingGroup(key=key))
        group.rows.append(row)
        source = _source_text(row)
        if source not in group.sources:
            group.sources.append(source)

        raw_quantity = row.el[VALUES].value if VALUES in row.el else None
        try:
            if raw_quantity is None or (
                isinstance(raw_quantity, str) and not raw_quantity.strip()
            ):
                raise ValueError("empty quantity")
            quantity = parse_quantity_strict(raw_quantity, context=source)
            if not math.isfinite(quantity) or quantity < 0:
                raise ValueError(f"non-finite or negative quantity {quantity!r}")
        except ValueError as exc:
            issue = _row_issue(
                row,
                code="packing_quantity",
                message=f"Количество УЛ не является допустимым неотрицательным числом ({exc})",
                field_name=VALUES,
                raw_value=raw_quantity,
                action="Исправьте количество в исходном упаковочном листе",
            )
            group.issues.append(issue)
            issues.append(issue)
            stats.invalid_quantity_rows += 1
            continue

        tags_count = len(
            {
                str(tag).strip()
                for tag in row.get_tags_list()
                if str(tag).strip()
            }
        )
        if tags_count > quantity:
            issue = _row_issue(
                row,
                code="packing_tags_exceed_quantity",
                message=(
                    f"Количество тегов УЛ ({tags_count}) превышает "
                    f"количество позиции ({quantity:g})"
                ),
                field_name=f"{TAGS},{VALUES}",
                raw_value={"tags": tags_count, "quantity": raw_quantity},
                action="Исправьте теги или количество в исходном упаковочном листе",
            )
            group.issues.append(issue)
            issues.append(issue)

        units = normalize_units_text(row.get_value(UNITS))
        if not units:
            issue = _row_issue(
                row,
                code="packing_units_missing",
                message="У позиции УЛ не заполнена единица измерения",
                field_name=UNITS,
                action="Заполните единицу измерения в исходном УЛ",
            )
            group.issues.append(issue)
            issues.append(issue)
        elif units not in group.units:
            group.units.append(units)

        # Identical rows across files/sheets are trusted as separate deliveries;
        # quantities are summed without packing_suspicious_duplicate warnings.
        group.valid_quantity += quantity
        group.valid_quantity_rows += 1

    for group in groups.values():
        if len(group.units) > 1:
            issue = _row_issue(
                group.first,
                code="packing_units_conflict",
                message=f"Для одного ключа УЛ указаны разные единицы: {' / '.join(group.units)}",
                field_name=UNITS,
                action="Приведите единицы к одному значению либо разделите коды",
            )
            group.issues.append(issue)
            issues.append(issue)

    stats.grouped_packing_keys = len(groups)
    return groups, invalid_key_rows, issues


def _ensure_ul_columns(row: RowStd) -> None:
    for column in _UL_COLUMNS:
        if column not in row.el:
            row.el[column] = CheckElement(None)


def _clear_ul_columns(row: RowStd) -> None:
    _ensure_ul_columns(row)
    for column in _UL_COLUMNS:
        row.el[column].value = None
        row.el[column].color = Color.no
        row.el[column].comment = ""


def _set_cell(
    row: RowStd,
    column: str,
    value: object,
    *,
    color: str = Color.no,
    comment: str = "",
) -> None:
    _ensure_ul_columns(row)
    row.el[column].value = value
    row.el[column].color = color
    row.el[column].comment = comment


def _ordered_quantity(row: RowStd) -> float:
    return quantity_as_float(row.get_value(VALUES)) + quantity_as_float(
        row.get_value(RFQ_VALUES)
    )


def _data_status_for_dataset(dataset: PackingDataset) -> tuple[str, str]:
    if dataset.quality == PackingQualityLevel.UNAVAILABLE:
        return DATA_UNAVAILABLE, Color.red
    if dataset.quality == PackingQualityLevel.PARTIAL:
        return DATA_PARTIAL, Color.red
    return DATA_OK, Color.green


def _status_for_remaining(remaining: float) -> tuple[str, str]:
    if abs(remaining) < _FLOAT_EPS:
        return STATUS_COMPLETE, Color.green
    if remaining > 0:
        return STATUS_SHORTFALL, Color.yellow
    return STATUS_OVERDELIVERY, Color.soft_cyan


def _fill_group_details(row: RowStd, group: _PackingGroup) -> None:
    first = group.first
    details = (
        (CODE, UL_CODE),
        (NAME, UL_NAME),
        (TYPE_MARK, UL_TYPE_MARK),
        (VENDOR, UL_VENDOR),
        (TAGS, UL_TAGS),
    )
    for source_column, target_column in details:
        value = _join_values(item.get_value(source_column) for item in group.rows)
        _set_cell(row, target_column, value or None)
    _set_cell(row, UL_SOURCE_FILES, "; ".join(group.sources))
    units = " / ".join(group.units)
    _set_cell(row, UL_UNITS, units or None)


def _group_comment(group: _PackingGroup, target_rows: list[RowStd]) -> str:
    lines = [
        f"Ключ УЛ: {group.key!r}",
        f"Строк УЛ: {len(group.rows)}; строк результата: {len(target_rows)}",
        f"Сумма валидных количеств УЛ: {group.valid_quantity:g}",
        "Источники:",
    ]
    lines.extend(f"- {source}" for source in group.sources)
    if group.issues:
        lines.append("Проблемы:")
        lines.extend(f"- {issue.format_line()}" for issue in group.issues)
    return "\n".join(lines)


def _create_packing_only_row(
    group: _PackingGroup,
    *,
    data_status: tuple[str, str],
) -> RowStd:
    row = RowStd.get_std_check_row({}, TableComments())
    row.row_type = RowType.position_row
    row.el[ROW_TYPE].value = RowType.position_row
    _clear_ul_columns(row)
    for column in (DS_TITLE, DS_SYSTEM):
        row.el[column].value = group.first.get_value(column)
        Color.set_el_color(row.el[column], Color.soft_cyan)
    _fill_group_details(row, group)
    _set_cell(row, UL_ORDERED_VALUES, 0.0, color=Color.soft_cyan)
    delivered = group.valid_quantity if group.valid_quantity_rows else None
    remaining = -group.valid_quantity if group.valid_quantity_rows else None
    _set_cell(row, UL_VALUES, delivered, color=Color.soft_cyan)
    _set_cell(row, UL_REMAINING_VALUES, remaining, color=Color.soft_cyan)
    comment = _group_comment(group, [row])
    if group.issues:
        row.el[UL_VALUES].color = Color.red
        row.el[UL_REMAINING_VALUES].color = Color.red
        _set_cell(
            row,
            UL_COMPARE_STATUS,
            STATUS_DATA_PROBLEM,
            color=Color.red,
            comment=comment,
        )
        _set_cell(row, UL_DATA_STATUS, DATA_ROW_PROBLEM, color=Color.red, comment=comment)
    else:
        _set_cell(
            row,
            UL_COMPARE_STATUS,
            STATUS_PACKING_ONLY,
            color=Color.soft_cyan,
            comment=comment,
        )
        _set_cell(row, UL_DATA_STATUS, data_status[0], color=data_status[1])
    for column in (UL_CODE, UL_SOURCE_FILES, UL_NAME, UL_TYPE_MARK, UL_VENDOR, UL_TAGS, UL_UNITS):
        if row.el[column].value:
            row.el[column].color = Color.red if group.issues else Color.soft_cyan
    return row


def _create_invalid_key_row(
    source_row: RowStd,
    issue: PackingIssue,
) -> RowStd:
    group = _PackingGroup(key=("", "", ""), rows=[source_row], sources=[_source_text(source_row)])
    row = _create_packing_only_row(
        group,
        data_status=(DATA_ROW_PROBLEM, Color.red),
    )
    _set_cell(row, UL_VALUES, None, color=Color.red)
    _set_cell(row, UL_REMAINING_VALUES, None, color=Color.red)
    _set_cell(
        row,
        UL_COMPARE_STATUS,
        STATUS_DATA_PROBLEM,
        color=Color.red,
        comment=issue.format_line(),
    )
    _set_cell(
        row,
        UL_DATA_STATUS,
        DATA_ROW_PROBLEM,
        color=Color.red,
        comment=issue.format_line(),
    )
    return row


def _candidate_key(
    row: RowStd,
    groups: dict[tuple[str, str, str], _PackingGroup],
) -> tuple[tuple[str, str, str] | None, bool]:
    title = row.get_value(DS_TITLE)
    system = row.get_value(DS_SYSTEM)
    ds_code = packing_match_key(title, system, row.get_value(CODE))[2]
    candidate_codes = packing_codes_for_row(row)
    if RFQ_CODE in row.el:
        rfq_code = packing_match_key(title, system, row.get_value(RFQ_CODE))[2]
        if rfq_code and rfq_code not in candidate_codes:
            candidate_codes.append(rfq_code)
    for code in candidate_codes:
        key = packing_match_key(title, system, code)
        if key in groups:
            return key, bool(ds_code and code != ds_code)
    return None, False


def compare_grouped_rows_with_packing(
    compared_rows: list[RowStd],
    dataset: PackingDataset,
) -> tuple[list[RowStd], PackingCompareAudit]:
    """Enrich grouped DS/MTO/RFQ rows with cumulative packing delivery data.

    Invalid or unavailable packing data never aborts the DS Excel export.
    Instead, affected cells are red and every issue remains in the audit.
    """
    stats = PackingCompareStats()
    issues = list(dataset.issues)
    miss_samples: list[str] = []

    for row in compared_rows:
        _clear_ul_columns(row)
        if row.row_type == RowType.position_row:
            stats.compared_position_rows += 1

    data_status = _data_status_for_dataset(dataset)
    if not dataset.available:
        comment = "\n".join(issue.format_line() for issue in dataset.issues)
        for row in compared_rows:
            if row.row_type != RowType.position_row:
                continue
            _set_cell(row, UL_ORDERED_VALUES, _ordered_quantity(row))
            _set_cell(
                row,
                UL_COMPARE_STATUS,
                STATUS_UNAVAILABLE,
                color=Color.red,
                comment=comment,
            )
            _set_cell(
                row,
                UL_DATA_STATUS,
                DATA_UNAVAILABLE,
                color=Color.red,
                comment=comment,
            )
        return compared_rows, PackingCompareAudit(
            dataset=dataset,
            stats=stats,
            issues=issues,
        )

    groups, invalid_key_rows, matcher_issues = _build_packing_groups(dataset, stats)
    issues.extend(matcher_issues)

    selected_rows: dict[tuple[str, str, str], list[tuple[RowStd, bool]]] = defaultdict(list)
    for row in compared_rows:
        if row.row_type != RowType.position_row:
            continue
        key, by_mto_code = _candidate_key(row, groups)
        if key is None:
            ordered = _ordered_quantity(row)
            _set_cell(row, UL_ORDERED_VALUES, ordered)
            _set_cell(
                row,
                UL_COMPARE_STATUS,
                STATUS_NOT_IN_PACKING,
                color=Color.yellow,
            )
            _set_cell(row, UL_DATA_STATUS, data_status[0], color=data_status[1])
            stats.not_in_packing += 1
            if len(miss_samples) < _MISS_SAMPLE_LIMIT:
                miss_samples.append(
                    f"title={row.get_value(DS_TITLE)!r}, "
                    f"system={row.get_value(DS_SYSTEM)!r}, "
                    f"codes={packing_codes_for_row(row)!r}"
                )
            continue
        selected_rows[key].append((row, by_mto_code))

    consumed_keys: set[tuple[str, str, str]] = set()
    for key, targets_with_flags in selected_rows.items():
        group = groups[key]
        targets = [item[0] for item in targets_with_flags]
        primary, primary_by_mto_code = targets_with_flags[0]
        ordered = sum(_ordered_quantity(row) for row in targets)
        delivered = group.valid_quantity if group.valid_quantity_rows else None
        remaining = (
            ordered - group.valid_quantity
            if group.valid_quantity_rows
            else None
        )
        comment = _group_comment(group, targets)
        if primary_by_mto_code:
            comment = f"УЛ matched by MTO code {key[2]!r}\n{comment}"

        _fill_group_details(primary, group)
        _set_cell(primary, UL_ORDERED_VALUES, ordered, comment=comment)
        _set_cell(primary, UL_VALUES, delivered, comment=comment)
        _set_cell(primary, UL_REMAINING_VALUES, remaining, comment=comment)
        if group.issues:
            primary.el[UL_VALUES].color = Color.red
            primary.el[UL_REMAINING_VALUES].color = Color.red
            _set_cell(
                primary,
                UL_COMPARE_STATUS,
                STATUS_DATA_PROBLEM,
                color=Color.red,
                comment=comment,
            )
            _set_cell(
                primary,
                UL_DATA_STATUS,
                DATA_ROW_PROBLEM,
                color=Color.red,
                comment=comment,
            )
            stats.data_problem += 1
        else:
            assert remaining is not None
            status, status_color = _status_for_remaining(remaining)
            _set_cell(
                primary,
                UL_COMPARE_STATUS,
                status,
                color=status_color,
                comment=comment,
            )
            _set_cell(primary, UL_DATA_STATUS, data_status[0], color=data_status[1])
            if status == STATUS_COMPLETE:
                stats.complete += 1
            elif status == STATUS_SHORTFALL:
                stats.shortfall += 1
            else:
                stats.overdelivery += 1

        for duplicate, _by_mto_code in targets_with_flags[1:]:
            _set_cell(
                duplicate,
                UL_COMPARE_STATUS,
                STATUS_DATA_PROBLEM if group.issues else STATUS_ACCOUNTED_ABOVE,
                color=Color.red if group.issues else Color.no,
                comment=(
                    f"Количество УЛ по ключу {key!r} показано только на первой "
                    "строке, чтобы не вычесть поставку повторно.\n"
                    + comment
                ),
            )
            if group.issues:
                _set_cell(
                    duplicate,
                    UL_DATA_STATUS,
                    DATA_ROW_PROBLEM,
                    color=Color.red,
                    comment=comment,
                )
            else:
                _set_cell(duplicate, UL_DATA_STATUS, data_status[0], color=data_status[1])
            stats.duplicate_target_rows += 1

        consumed_keys.add(key)
        stats.matched_keys += 1
        stats.matched_rows += len(targets)

    for key, group in groups.items():
        if key in consumed_keys:
            continue
        compared_rows.append(_create_packing_only_row(group, data_status=data_status))
        stats.packing_only_added += 1
        if group.issues:
            stats.data_problem += 1

    key_issues = [issue for issue in matcher_issues if issue.code == "packing_key_missing"]
    for source_row, issue in zip(invalid_key_rows, key_issues):
        compared_rows.append(_create_invalid_key_row(source_row, issue))
        stats.packing_only_added += 1
        stats.data_problem += 1

    return compared_rows, PackingCompareAudit(
        dataset=dataset,
        stats=stats,
        issues=issues,
        miss_samples=miss_samples,
    )


def save_packing_compare_report(
    audit: PackingCompareAudit,
    out_dir: str | Path,
) -> str:
    """Write a complete UTF-8 audit next to the grouped Excel output."""
    output_dir = Path(out_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y.%m.%d_%H.%M.%S")
    path = output_dir / f"ul_compare_report_{stamp}.txt"
    meta = audit.dataset.meta
    stats = audit.stats
    lines = [
        "=== СРАВНЕНИЕ С УПАКОВОЧНЫМИ ЛИСТАМИ ===",
        f"quality: {audit.quality.value}",
        f"cache_path: {audit.dataset.cache_path}",
        f"cache_version: {meta.version if meta else ''}",
        f"cache_created_at: {meta.created_at if meta else ''}",
        f"source_root: {meta.root if meta else ''}",
        f"cache_fingerprint: {meta.fingerprint if meta else ''}",
        "",
        "COUNTERS:",
    ]
    for name, value in vars(stats).items():
        lines.append(f"  {name}: {value}")
    if meta is not None:
        lines.append("")
        lines.append("LOADER COUNTERS:")
        for name, value in sorted(meta.counters.items()):
            lines.append(f"  {name}: {value}")
        lines.append("")
        lines.append("SOURCE REPORTS:")
        if meta.report_paths:
            for name, value in sorted(meta.report_paths.items()):
                lines.append(f"  {name}: {value}")
        else:
            lines.append("  (not recorded)")
    lines.extend(["", f"ISSUES ({len(audit.issues)}):"])
    if audit.issues:
        lines.extend(f"  {index}. {issue.format_line()}" for index, issue in enumerate(audit.issues, 1))
    else:
        lines.append("  (none)")
    lines.extend(["", "SAMPLE KEYS NOT FOUND IN UL:"])
    if audit.miss_samples:
        lines.extend(f"  - {sample}" for sample in audit.miss_samples)
    else:
        lines.append("  (none)")
    path.write_text("\n".join(lines) + "\n", encoding="utf-8")
    audit.report_path = str(path)
    return str(path)
