"""Per-item packing-list allocation for the RFP Step4 result."""

from __future__ import annotations

import math
import re
import time
from collections import defaultdict, deque
from contextlib import contextmanager
from dataclasses import dataclass, field, replace
from datetime import datetime
from pathlib import Path
from typing import Callable, Iterable, Iterator

from base.base_classes import CheckElement, RowStd, RowType, TableComments
from base.tables_columns import (
    CODE,
    CODE_MTO,
    CODE_VO,
    DS_ACTUAL,
    DS_NAME,
    DS_SYSTEM,
    DS_TITLE,
    MATCH_STATUS,
    MATCH_STATUS_VO,
    MTO_CODE_STRUCK,
    NAME,
    ROW_TYPE,
    TAG_MTO,
    TAG_VO,
    TAGS,
    TYPE_MARK,
    UL_CODE,
    UL_COMPARE_STATUS,
    UL_DATA_STATUS,
    UL_NAME,
    UL_ORDERED_VALUES,
    UL_REMAINING_VALUES,
    UL_SOURCE_FILES,
    UL_TAG_MATCH_STATUS,
    UL_TAGS,
    UL_TYPE_MARK,
    UL_UNITS,
    UL_VALUES,
    UL_VENDOR,
    UNITS,
    UNITS_CHECK_STATUS,
    UNITS_CONVERSION_TRACE,
    VALUES,
    VALUES_MTO,
    VALUES_VO,
    VENDOR,
)
from RFQ.ds_compare.ds_quantity_parse import parse_quantity_strict, read_raw_quantity
from RFQ.ds_compare.ds_units_normalize import join_unique_units_text, normalize_units_text
from RFQ.packing_list_provider import (
    PackingDataset,
    PackingIssue,
    PackingIssueSeverity,
    PackingQualityLevel,
    normalize_packing_key_part,
    packing_match_key,
    packing_row_actual_ds,
    packing_row_key,
    packing_row_source,
    rfp_packing_match_key,
    rfp_row_actual_ds,
    supply_ul_folders,
)
from RFQ.rfp_parts.ds_identity import parse_ul_folder_ds_identity
from RFQ.tags_rfp_compare.rfp_supply_status import (
    CANONICAL_EXCLUDED_FROM_SUPPLY,
    is_excluded_from_supply,
)
from utils.colors import Color

# Same system token as AgccFilenamePatterns.TITLE_SYSTEM: SOO1, SKUD, SKUD.1, POS3.
_COMPOSITE_TITLE_RE = re.compile(
    r"(?P<title>[0-9]+)-(?P<system>[A-Za-zА-Яа-я]{2,5}\d?(?:\.\d)?)",
    re.IGNORECASE,
)
_EPS = 1e-9
_MISS_SAMPLE_LIMIT = 20

STATUS_COMPLETE = "Поставка комплектна"
STATUS_SHORTFALL = "Недопоставка"
STATUS_OVERDELIVERY = "Перепоставка"
STATUS_NOT_IN_PACKING = "Нет в УЛ"  # legacy alias; Step4 Excel no longer writes it
STATUS_OPEN_UPD = "Неотгружено по УЛ"
STATUS_GEM_SUPPLY = "Поставка ГЭМ"
STATUS_RFP_ONLY_EXTRA = "Только в RFP (лишняя)"
STATUS_MTO_ONLY = "Только в МТО (Недопоставка)"
STATUS_MTO_DELIVERED = "Есть в МТО, Поставлено, Нет в RFP"
STATUS_MTO_DELIVERED_SHORT = "Есть в МТО, Недопоставлено, Нет в RFP"
STATUS_VO_ONLY = "Только в РКД"
STATUS_VO_DELIVERED = "Есть в РКД, Поставлено, Нет в RFP"
STATUS_VO_DELIVERED_SHORT = "Есть в РКД, Недопоставлено, Нет в RFP"
STATUS_DATA_PROBLEM = "Проблема данных УЛ"
STATUS_UNAVAILABLE = "УЛ недоступны"
STATUS_PACKING_ONLY = "Только в УЛ"
STATUS_ZIP_SMR_PNR = "ЗИП для СМР, ПНР"
STATUS_EXCLUDED_FROM_SUPPLY = CANONICAL_EXCLUDED_FROM_SUPPLY
ZIP_SMR_PNR_NAME_RE = re.compile(
    r"комплект\s+зип\s+для\s*смр"
    r"|зип\s+для\s*смр"
    r"|зип\s*,\s*для\s+проведения\s+смр",
    re.IGNORECASE,
)
STATUS_TAG_MISMATCH_COLLECTED = "Собрано по несопоставленным тегам"
STATUS_TAG_MATCHED_VIA_MTO = "Собрано по тегу МТО"
# Marker in UL_COMPARE_STATUS.comment; Excel writes the cloud only for this
# plus shortfall / data-problem (see should_write_excel_comment).
FALLBACK_COMMENT_MARKER = "Fallback сопоставления:"

_MATCH_ADDED_MTO = "Добавлен из МТО"
_MATCH_ADDED_VO = "Добавлен из VO"

DATA_OK = "OK"
DATA_PARTIAL = "Частичные данные УЛ"
DATA_ROW_PROBLEM = "Проблема строки УЛ"
DATA_UNAVAILABLE = "УЛ недоступны"

_UL_COLUMNS = (
    UL_ORDERED_VALUES,
    UL_VALUES,
    UL_UNITS,
    UL_REMAINING_VALUES,
    UL_COMPARE_STATUS,
    UL_DATA_STATUS,
    UL_TAG_MATCH_STATUS,
    UL_CODE,
    UL_SOURCE_FILES,
    UL_NAME,
    UL_TYPE_MARK,
    UL_VENDOR,
    UL_TAGS,
)


@dataclass
class RfpPackingStats:
    """Counters from one RFP packing-list enrichment."""

    result_position_rows: int = 0
    packing_rows: int = 0
    packing_units: int = 0
    allocated_units: int = 0
    matched_rows: int = 0
    not_in_packing: int = 0
    packing_only_added: int = 0
    invalid_quantity_rows: int = 0
    invalid_key_rows: int = 0
    fallback_code_matches: int = 0
    complete: int = 0
    shortfall: int = 0
    overdelivery: int = 0
    unit_mismatches: int = 0
    complete_zero_allocated: int = 0
    complete_zero_fallback: int = 0
    complete_zero_empty_code: int = 0
    queue_keys: int = 0
    max_queue_len: int = 0


@dataclass
class RfpPackingAudit:
    """Diagnostics and summary for RFP packing comparison."""

    dataset: PackingDataset
    stats: RfpPackingStats
    issues: list[PackingIssue] = field(default_factory=list)
    miss_samples: list[str] = field(default_factory=list)
    report_path: str | None = None
    phase_timings: dict[str, float] = field(default_factory=dict)

    @property
    def quality(self) -> PackingQualityLevel:
        """Return combined provider and matcher quality."""
        if self.dataset.quality == PackingQualityLevel.UNAVAILABLE:
            return PackingQualityLevel.UNAVAILABLE
        return PackingQualityLevel.PARTIAL if self.issues else PackingQualityLevel.OK

    def format_short(self) -> str:
        """Return a compact user-facing summary."""
        label = {
            PackingQualityLevel.OK: "OK",
            PackingQualityLevel.PARTIAL: "ВНИМАНИЕ: частичные данные",
            PackingQualityLevel.UNAVAILABLE: "ОШИБКА: УЛ недоступны",
        }[self.quality]
        report = f" Отчёт: {self.report_path}." if self.report_path else ""
        return (
            f"УЛ Step4 [{label}]: allocated={self.stats.allocated_units}, "
            f"нет_в_УЛ={self.stats.not_in_packing}, "
            f"только_УЛ={self.stats.packing_only_added}, "
            f"проблем={len(self.issues)}.{report}"
        )


class RfpPackingFatalError(RuntimeError):
    """Fatal packing data error which must stop Step4 before Excel export."""

    def __init__(self, audit: RfpPackingAudit):
        self.audit = audit
        fatal_locations = {
            (issue.source_file, issue.sheet, issue.excel_row)
            for issue in audit.issues
            if issue.code == "packing_tags_exceed_quantity"
        }
        fatal_count = len(fatal_locations)
        super().__init__(
            f"В УЛ найдено строк, где tags > quantity: {fatal_count}. "
            "Step4 остановлен до сохранения Excel."
        )


@dataclass(frozen=True)
class _PackingUnit:
    key: tuple[str, str, str, str]
    source_row: RowStd
    tag: str
    units: str
    source: str
    units_check_status: str = ""
    units_conversion_trace: str = ""
    qty: float = 1.0
    tag_mismatch: bool = False
    via_mto: bool = False


@dataclass
class _RowAllocationState:
    """Mutable packing allocation for one result row across global passes.

    ``delivered`` is always equal to ``_delivered_qty(allocated)``.
    Append allocated slots only through ``take()``.
    """

    row: RowStd
    key: tuple[str, str, str, str]
    fallback: str
    ordered: float
    allocated: list[_PackingUnit] = field(default_factory=list)
    delivered: float = 0.0
    rfp_tags: list[str] = field(default_factory=list)
    mto_tags: list[str] = field(default_factory=list)
    vo_tags: list[str] = field(default_factory=list)
    queue_keys: tuple[tuple[str, str, str, str], ...] = ()

    def take(self, unit: _PackingUnit) -> None:
        """Append an allocated unit and keep the running delivered total in sync."""
        self.allocated.append(unit)
        self.delivered += _unit_qty(unit)

    def remaining_qty(self) -> float:
        return self.ordered - self.delivered


def parse_composite_title(value: object) -> tuple[str, str] | None:
    """Parse a Step4 title such as ``8950-SOO1`` or ``7421-SKUD.1``.

    The mark may include a trailing digit (``SOO1``) and a ``.N`` suffix
    (``SKUD.1``), matching ``AgccFilenamePatterns.TITLE_SYSTEM``.

    Args:
        value: Raw ``DS_TITLE`` value.

    Returns:
        ``(title, system)`` or ``None`` when the complete value is malformed.
    """
    text = str(value or "").strip()
    match = _COMPOSITE_TITLE_RE.fullmatch(text)
    if match is None:
        return None
    return match.group("title"), match.group("system")


def build_ordered_snapshot(
    rfp_rows: Iterable[RowStd],
) -> dict[tuple[str, str, str, str], float]:
    """Sum original RFP ordered quantities by ``(actual, title, system, code)``.

    Only source ``position_row`` rows with a parsed actual DS, a valid
    composite title, code, and non-negative quantity contribute to the
    snapshot. Rows without actual DS are skipped. Validation diagnostics
    are produced later against result rows so this helper remains a plain dict
    contract suitable for orchestration before split/matching.
    """
    ordered: dict[tuple[str, str, str, str], float] = defaultdict(float)
    for row in rfp_rows:
        if row.row_type != RowType.position_row:
            continue
        if is_excluded_from_supply(row):
            continue
        actual = rfp_row_actual_ds(row)
        if actual is None:
            continue
        title_system = parse_composite_title(row.get_value(DS_TITLE))
        if title_system is None:
            continue
        code = row.get_value(CODE)
        key = rfp_packing_match_key(actual, *title_system, code)
        if not all(key):
            continue
        try:
            quantity = parse_quantity_strict(read_raw_quantity(row, VALUES))
        except ValueError:
            continue
        if math.isfinite(quantity) and quantity >= 0:
            ordered[key] += quantity
    return dict(ordered)


@contextmanager
def _phase(audit: RfpPackingAudit, name: str) -> Iterator[None]:
    """Accumulate wall time of one matcher phase into ``audit.phase_timings``."""
    started = time.perf_counter()
    try:
        yield
    finally:
        audit.phase_timings[name] = (
            audit.phase_timings.get(name, 0.0) + time.perf_counter() - started
        )


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
    if column not in row.el:
        row.el[column] = CheckElement(None)
    row.el[column].value = value
    row.el[column].color = color
    row.el[column].comment = comment


PACKING_ONLY_SEQUENTIAL_MARKER = "RFP_не_найден"


def _format_ds_actual_label(actual: int | None) -> str:
    """Return display ``ДС{n}`` for an actual DS number."""
    return f"ДС{actual}" if actual is not None else ""


def packing_only_sequential_label(actual: int | None) -> str:
    """Return leftover sequential label ``ДС{n}_RFP_не_найден``.

    Args:
        actual: UL folder actual DS number, or ``None`` for GF / unparsed.

    Returns:
        ``ДС{n}_RFP_не_найден`` when ``actual`` is set, otherwise the marker
        alone so GF leftovers stay visible in Excel and the stats sheet.
    """
    label = _format_ds_actual_label(actual)
    if label:
        return f"{label}_{PACKING_ONLY_SEQUENTIAL_MARKER}"
    return PACKING_ONLY_SEQUENTIAL_MARKER


def is_packing_only_sequential(ds_name: object) -> bool:
    """Return True when ``ds_name`` is the leftover «RFP not found» marker."""
    text = str(ds_name or "").strip()
    return text == PACKING_ONLY_SEQUENTIAL_MARKER or text.endswith(
        f"_{PACKING_ONLY_SEQUENTIAL_MARKER}"
    )


def _set_ds_actual(
    row: RowStd,
    actual: int | None,
    *,
    missing_comment: bool = False,
) -> None:
    """Fill ``DS_ACTUAL`` with ``ДС{n}``, or empty plus a missing-actual comment."""
    comment = "нет фактического ДС" if actual is None and missing_comment else ""
    _set_cell(row, DS_ACTUAL, _format_ds_actual_label(actual), comment=comment)


def _fill_packing_only_ds_identity(row: RowStd, source_row: RowStd) -> None:
    """Fill leftover identity: actual from the UL folder, sequential marker.

    ``DS_ACTUAL`` is ``ДС{n}`` from the folder name (same parser as the
    RFP · ДС ↔ УЛ tab). ``DS_NAME`` is ``ДС{n}_RFP_не_найден`` so the
    stats sheet shows leftover rows separately from RFP sequential labels.
    GF folders without a DS number keep actual empty and sequential
    ``RFP_не_найден``.

    Args:
        row: Newly created leftover result row.
        source_row: Original packing cache row (``ANNOTATION`` has the folder).
    """
    actual = packing_row_actual_ds(source_row)
    _set_ds_actual(row, actual)
    _set_cell(row, DS_NAME, packing_only_sequential_label(actual))


def _unique_text(values: Iterable[object]) -> list[str]:
    result: list[str] = []
    for value in values:
        candidates = value if isinstance(value, list) else [value]
        for candidate in candidates:
            text = str(candidate or "").strip()
            if text and text not in result:
                result.append(text)
    return result


def _joined(values: Iterable[object], *, limit: int = 40) -> str:
    unique = _unique_text(values)
    if len(unique) <= limit:
        return "; ".join(unique)
    return "; ".join(unique[:limit]) + f"; …(+{len(unique) - limit})"


_SOURCE_SEP = " · "


def _format_ul_source_files(values: Iterable[object], *, limit: int = 40) -> str:
    """Join unique packing sources with line breaks, grouped by file.

    One source stays a single ``file · sheet · строка N`` line. Several sources
    from the same file print the path once, then indented sheet/row lines.
    """
    unique = _unique_text(values)
    extra = max(0, len(unique) - limit)
    lines = _group_ul_source_lines(unique[:limit])
    if extra:
        lines.append(f"…(+{extra})")
    return "\n".join(lines)


def _group_ul_source_lines(sources: list[str]) -> list[str]:
    if not sources:
        return []
    if len(sources) == 1:
        return list(sources)

    groups: dict[str, list[str]] = {}
    order: list[str] = []
    for source in sources:
        parts = [part.strip() for part in source.split(_SOURCE_SEP) if part.strip()]
        file_name = parts[0] if parts else source
        rest = _SOURCE_SEP.join(parts[1:])
        if file_name not in groups:
            groups[file_name] = []
            order.append(file_name)
        if rest and rest not in groups[file_name]:
            groups[file_name].append(rest)

    lines: list[str] = []
    for file_name in order:
        lines.append(file_name)
        lines.extend(f"  {detail}" for detail in groups[file_name])
    return lines


def _data_status(dataset: PackingDataset) -> tuple[str, str]:
    if dataset.quality == PackingQualityLevel.UNAVAILABLE:
        return DATA_UNAVAILABLE, Color.red
    if dataset.quality == PackingQualityLevel.PARTIAL:
        return DATA_PARTIAL, Color.red
    return DATA_OK, Color.green


def _strict_packing_quantity(row: RowStd) -> float:
    """Return a finite non-negative packing quantity.

    Integer rows stay piece-wise in the queue. A genuine fraction (typical
    after units-gate coef≠1) is kept as a remainder slot instead of dropping
    the delivery as a data problem.
    """
    raw = read_raw_quantity(row, VALUES)
    if raw is None or (isinstance(raw, str) and not raw.strip()):
        raise ValueError("empty quantity")
    quantity = parse_quantity_strict(raw, context=_source_text(row))
    if not math.isfinite(quantity) or quantity < 0:
        raise ValueError(f"non-finite or negative quantity {quantity!r}")
    return float(quantity)


def _unit_qty(unit: _PackingUnit) -> float:
    qty = float(unit.qty)
    return qty if math.isfinite(qty) and qty > 0 else 0.0


def _delivered_qty(units: Iterable[_PackingUnit]) -> float:
    return float(sum(_unit_qty(unit) for unit in units))


def _whole_and_remainder(quantity: float) -> tuple[int, float]:
    whole = int(math.floor(quantity + _EPS))
    remainder = quantity - whole
    if remainder < _EPS:
        return whole, 0.0
    return whole, remainder


def _packing_row_folder_raw(row: RowStd) -> str:
    """Return the first ANNOTATION path segment (physical UL folder name)."""
    annotation, _sheet, _excel_row = packing_row_source(row)
    normalized = annotation.replace("\\", "/")
    parts = Path(normalized).parts
    return parts[0] if parts else ""


def _packing_row_folder_key(row: RowStd) -> str:
    """Return the normalized physical UL folder key for a packing row."""
    return normalize_packing_key_part(_packing_row_folder_raw(row))


def _dataset_folders_by_parsed_actual(
    dataset: PackingDataset,
) -> dict[int, frozenset[str]]:
    """Map parsed folder actual DS → folder keys present in this dataset."""
    grouped: dict[int, set[str]] = defaultdict(set)
    for row in dataset.rows:
        raw = _packing_row_folder_raw(row)
        parsed = parse_ul_folder_ds_identity(raw).actual
        if parsed is None:
            continue
        grouped[parsed].add(normalize_packing_key_part(raw))
    return {actual: frozenset(keys) for actual, keys in grouped.items()}


def _eligible_folder_keys(
    actual: int,
    name_folders: dict[int, frozenset[str]],
) -> frozenset[str]:
    """Return folder keys this supply may take from in block A.

    ``supply_ul_folders`` is ``None`` when there is no planting index or
    ``actual`` is not a known source: name-match every folder in this
    dataset whose parsed actual equals ``actual``. A frozenset (including
    empty) is the registry set only — physical folders are not added just
    because their names parse as the same number.
    """
    registry = supply_ul_folders(actual)
    if registry is None:
        return name_folders.get(actual, frozenset())
    return registry


def packing_queue_input_by_title(
    dataset: PackingDataset,
) -> tuple[float, int, dict[str, float]]:
    """Sum VALUES that enter unit queues, grouped by RFP-style display title.

    Skip rules must stay in sync with ``_build_unit_queues``: missing
    title/system/code, invalid (empty/non-finite/negative) quantity, and tags
    exceeding quantity are excluded (they never increment ``packing_units``).
    Empty actual DS or empty folder does not skip a row. Genuine fractions
    after units conversion are queued and counted. Unavailable datasets
    yield zeros.

    Returns:
        Total queued quantity, number of queued source rows, per-title sums.
    """
    by_title: dict[str, float] = defaultdict(float)
    total = 0.0
    rows = 0
    if not dataset.available:
        return 0.0, 0, {}
    for row in dataset.rows:
        title_system_code = packing_row_key(row)
        key = rfp_packing_match_key(
            _packing_row_folder_key(row),
            *title_system_code,
        )
        if not all(key[1:]):
            continue
        try:
            quantity = _strict_packing_quantity(row)
        except ValueError:
            continue
        tags = [
            str(tag).strip()
            for tag in row.get_tags_list()
            if str(tag).strip()
        ]
        if len(tags) > quantity:
            continue
        display = _packing_only_display_title(row).strip()
        title_key = display if display else "<без title_system>"
        by_title[title_key] += float(quantity)
        total += float(quantity)
        rows += 1
    return total, rows, dict(by_title)


def _build_unit_queues(
    dataset: PackingDataset,
    stats: RfpPackingStats,
) -> tuple[
    dict[tuple[str, str, str, str], deque[_PackingUnit]],
    list[tuple[RowStd, PackingIssue]],
    list[PackingIssue],
]:
    queues: dict[tuple[str, str, str, str], deque[_PackingUnit]] = defaultdict(deque)
    problem_rows: list[tuple[RowStd, PackingIssue]] = []
    issues: list[PackingIssue] = []

    for row in dataset.rows:
        stats.packing_rows += 1
        title_system_code = packing_row_key(row)
        key = rfp_packing_match_key(
            _packing_row_folder_key(row),
            *title_system_code,
        )
        if not all(key[1:]):
            issue = _row_issue(
                row,
                code="packing_key_missing",
                message="Нельзя сопоставить строку: пустая часть ключа title/system/code",
                field_name=f"{DS_TITLE},{DS_SYSTEM},{CODE}",
                action="Исправьте титул, марку или код в исходном УЛ",
            )
            issues.append(issue)
            problem_rows.append((row, issue))
            stats.invalid_key_rows += 1
            continue
        try:
            quantity = _strict_packing_quantity(row)
        except ValueError as exc:
            issue = _row_issue(
                row,
                code="packing_quantity",
                message=f"Количество УЛ должно быть неотрицательным конечным числом ({exc})",
                field_name=VALUES,
                raw_value=read_raw_quantity(row, VALUES),
                action="Исправьте количество в исходном упаковочном листе",
            )
            issues.append(issue)
            problem_rows.append((row, issue))
            stats.invalid_quantity_rows += 1
            continue

        tags = [
            str(tag).strip()
            for tag in row.get_tags_list()
            if str(tag).strip()
        ]
        if len(tags) > quantity:
            issues.append(
                _row_issue(
                    row,
                    code="packing_tags_exceed_quantity",
                    message=(
                        f"Количество тегов УЛ ({len(tags)}) превышает "
                        f"количество позиции ({quantity})"
                    ),
                    field_name=f"{TAGS},{VALUES}",
                    raw_value={"tags": len(tags), "quantity": quantity},
                    action="Исправьте теги или количество в исходном упаковочном листе",
                )
            )
            continue

        units = normalize_units_text(row.get_value(UNITS))
        if not units:
            issues.append(
                _row_issue(
                    row,
                    code="packing_units_missing",
                    message="У позиции УЛ не заполнена единица измерения",
                    field_name=UNITS,
                    action="Заполните единицу измерения в исходном УЛ",
                )
            )
        source = _source_text(row)
        whole, remainder = _whole_and_remainder(quantity)
        common = dict(
            key=key,
            source_row=row,
            units=units,
            source=source,
            units_check_status=str(row.el.get(UNITS_CHECK_STATUS, CheckElement(None)).value or "").strip(),
            units_conversion_trace=str(
                row.el.get(UNITS_CONVERSION_TRACE, CheckElement(None)).value or ""
            ).strip(),
        )
        for index in range(whole):
            queues[key].append(
                _PackingUnit(
                    tag=tags[index] if index < len(tags) else "",
                    qty=1.0,
                    **common,
                )
            )
            stats.packing_units += 1
        if remainder > 0:
            queues[key].append(_PackingUnit(tag="", qty=remainder, **common))
            stats.packing_units += 1
    stats.queue_keys = len(queues)
    stats.max_queue_len = max((len(queue) for queue in queues.values()), default=0)
    return dict(queues), problem_rows, issues


def _candidate_key(
    row: RowStd,
    queues: dict[tuple[str, str, str, str], deque[_PackingUnit]],
    eligible_folders: frozenset[str],
) -> tuple[
    tuple[str, str, str, str] | None,
    str,
    tuple[tuple[str, str, str, str], ...],
]:
    actual = rfp_row_actual_ds(row)
    if actual is None:
        return None, "", ()
    parsed = parse_composite_title(row.get_value(DS_TITLE))
    if parsed is None:
        return None, "malformed_title", ()
    title, system = parsed
    struck_values = {
        normalize_packing_key_part(row.get_value(MTO_CODE_STRUCK)),
        normalize_packing_key_part(
            getattr(row.el.get(CODE_MTO), "struck_value", "")
        ),
    }
    seen: set[str] = set()
    for column, fallback_name in (
        (CODE, ""),
        (CODE_MTO, "CODE_MTO"),
        (CODE_VO, "CODE_VO"),
    ):
        code = normalize_packing_key_part(row.get_value(column))
        if not code or code in seen:
            continue
        seen.add(code)
        if column == CODE_MTO and code in struck_values:
            continue
        queue_keys = tuple(
            sorted(
                rfp_packing_match_key(folder, title, system, code)
                for folder in eligible_folders
                if rfp_packing_match_key(folder, title, system, code) in queues
            )
        )
        if queue_keys:
            key = rfp_packing_match_key(actual, title, system, code)
            return key, fallback_name, queue_keys
    return None, "", ()


def _row_capacity(row: RowStd) -> float:
    for column in (VALUES, VALUES_MTO, VALUES_VO):
        if column not in row.el:
            continue
        raw = read_raw_quantity(row, column)
        try:
            value = parse_quantity_strict(raw)
        except ValueError:
            continue
        if math.isfinite(value) and value > 0:
            return value
    return 0.0


def _positive_qty(row: RowStd, column: str) -> bool:
    if column not in row.el:
        return False
    raw = read_raw_quantity(row, column)
    try:
        value = parse_quantity_strict(raw)
    except ValueError:
        return False
    return math.isfinite(value) and value > 0


def _has_rfp(row: RowStd) -> bool:
    match_status = str(row.get_value(MATCH_STATUS) or "").strip()
    match_vo = str(row.get_value(MATCH_STATUS_VO) or "").strip()
    if match_status == _MATCH_ADDED_MTO or match_vo == _MATCH_ADDED_VO:
        return False
    if str(row.get_value(CODE) or "").strip():
        return True
    return _positive_qty(row, VALUES)


def _has_mto(row: RowStd) -> bool:
    if str(row.get_value(CODE_MTO) or "").strip():
        return True
    return _positive_qty(row, VALUES_MTO)


def _has_vo(row: RowStd) -> bool:
    match_vo = str(row.get_value(MATCH_STATUS_VO) or "").strip()
    if match_vo == _MATCH_ADDED_VO:
        return True
    if str(row.get_value(CODE_VO) or "").strip():
        return True
    return _positive_qty(row, VALUES_VO)


def _presence_status(row: RowStd) -> tuple[str, str]:
    """Return UL compare status when no packing units were allocated."""
    has_rfp = _has_rfp(row)
    has_mto = _has_mto(row)
    if has_rfp and has_mto:
        return STATUS_OPEN_UPD, Color.yellow
    if has_rfp and not has_mto:
        return STATUS_RFP_ONLY_EXTRA, Color.yellow
    if has_mto and not has_rfp:
        return STATUS_MTO_ONLY, Color.yellow
    if _has_vo(row):
        return STATUS_VO_ONLY, Color.yellow
    return STATUS_RFP_ONLY_EXTRA, Color.yellow


def _extract_unit_at(
    queue: deque[_PackingUnit],
    index: int,
    remaining: float,
) -> _PackingUnit:
    """Take up to ``remaining`` qty from ``queue[index]``, leave leftover in place."""
    queue.rotate(-index)
    unit = queue.popleft()
    available = _unit_qty(unit)
    take = min(available, remaining)
    leftover = available - take
    if leftover > _EPS:
        queue.appendleft(replace(unit, qty=leftover))
        taken = replace(unit, qty=take)
    else:
        taken = unit if abs(take - available) <= _EPS else replace(unit, qty=take)
    queue.rotate(index)
    return taken


def _coerce_tag_values(value: object) -> list[str]:
    """Normalize TAG_MTO / similar cell values to a unique tag list."""
    if value is None or value == "":
        return []
    if isinstance(value, list):
        return _unique_text(value)
    return _unique_text([value])


def _reserved_tags_for_state(
    state: _RowAllocationState,
    states: list[_RowAllocationState],
) -> set[str]:
    """Return sibling RFP tags still needed on the same folder queue.

    ``states`` must already be the rows whose ``queue_keys`` contain the
    folder deque being scanned. Siblings are not grouped by ``state.key``
    (that would miss another DS sharing the folder).
    """
    reserved: set[str] = set()
    for other in states:
        if other is state:
            continue
        if other.remaining_qty() <= _EPS:
            continue
        reserved.update(other.rfp_tags)
    return reserved


def _row_leftover_code_key(
    row: RowStd,
    code_column: str,
) -> tuple[str, str, str] | None:
    """Build leftover lookup ``(title, system, code)`` from a result row.

    Args:
        row: Step4 result row.
        code_column: ``CODE_MTO`` for block B, ``CODE_VO`` for block C.

    Returns:
        Normalized 3-tuple, or ``None`` when ``DS_TITLE`` is malformed.
    """
    parsed = parse_composite_title(row.get_value(DS_TITLE))
    if parsed is None:
        return None
    title, system = parsed
    return packing_match_key(title, system, row.get_value(code_column))


def _reserved_leftover_tags(
    state: _RowAllocationState,
    siblings: list[_RowAllocationState],
    tags_fn: Callable[[_RowAllocationState], list[str]],
) -> set[str]:
    """Return sibling own-tags still needed on the same leftover 3-tuple.

    Groups by leftover ``(title, system, code)``, not by the 4-tuple A key,
    so an H1 MTO row with actual in ``state.key`` still reserves against a
    new B row keyed with empty actual.

    Args:
        state: Participant currently taking a blind leftover slot.
        siblings: Participants sharing the leftover 3-tuple of ``state``.
        tags_fn: Own tags to reserve (``TAG_MTO`` or ``TAG_VO``).

    Returns:
        Tags that must stay in the leftover deque for siblings.
    """
    reserved: set[str] = set()
    for other in siblings:
        if other is state:
            continue
        if other.remaining_qty() <= _EPS:
            continue
        reserved.update(tags_fn(other))
    return reserved


def _order_leftover_participants(
    participants: list[_RowAllocationState],
    row_index_by_id: dict[int, int],
    tags_fn: Callable[[_RowAllocationState], list[str]],
) -> list[_RowAllocationState]:
    """Return participants in original row order, tagged rows first."""
    return sorted(
        participants,
        key=lambda state: (
            not bool(tags_fn(state)),
            row_index_by_id.get(id(state.row), 10**9),
        ),
    )


def _allocate_queue_phase(
    queue: deque[_PackingUnit],
    state: _RowAllocationState,
    phase: str,
    *,
    reserved_tags: set[str] | frozenset[str] = frozenset(),
    own_tags: list[str] | None = None,
) -> None:
    """Pop packing units for one row in a single global matching phase.

    Args:
        queue: Remaining units for this phase's lookup key.
        state: Mutable allocation state; taken units are appended in place.
        phase: ``rfp_tags``, ``untagged``, ``mto_tags``, or ``blind``.
        reserved_tags: Sibling tags that must stay in the queue.
        own_tags: Tags treated as this row's own for ``rfp_tags`` / ``blind``.
            Defaults to ``state.rfp_tags`` (block A). Block B/C pass MTO/VO
            tags; do not use phase ``mto_tags`` there (it sets ``via_mto``).
    """
    own = state.rfp_tags if own_tags is None else own_tags
    own_set = set(own)

    def remaining() -> float:
        return state.remaining_qty()

    if phase == "rfp_tags":
        wanted = deque(own)
        while wanted and remaining() > _EPS:
            tag = wanted.popleft()
            match_index = next(
                (index for index, unit in enumerate(queue) if unit.tag == tag),
                None,
            )
            if match_index is None:
                continue
            taken = _extract_unit_at(queue, match_index, remaining())
            state.take(taken)
        return

    if phase == "untagged":
        while queue and remaining() > _EPS:
            match_index = next(
                (index for index, unit in enumerate(queue) if not unit.tag),
                None,
            )
            if match_index is None:
                break
            taken = _extract_unit_at(queue, match_index, remaining())
            state.take(taken)
        return

    if phase == "mto_tags":
        rfp_tag_set = set(state.rfp_tags)
        wanted_mto = deque(
            tag
            for tag in state.mto_tags
            if tag not in rfp_tag_set and tag not in reserved_tags
        )
        while wanted_mto and remaining() > _EPS:
            tag = wanted_mto.popleft()
            match_index = next(
                (index for index, unit in enumerate(queue) if unit.tag == tag),
                None,
            )
            if match_index is None:
                continue
            taken = replace(
                _extract_unit_at(queue, match_index, remaining()),
                via_mto=True,
            )
            state.take(taken)
        return

    if phase == "blind":
        while queue and remaining() > _EPS:
            match_index = next(
                (
                    index
                    for index, unit in enumerate(queue)
                    if not unit.tag or unit.tag not in reserved_tags
                ),
                None,
            )
            if match_index is None:
                break
            taken = _extract_unit_at(queue, match_index, remaining())
            if own_set and taken.tag and taken.tag not in own_set:
                taken = replace(taken, tag_mismatch=True)
            state.take(taken)
        return

    raise ValueError(f"unknown allocation phase: {phase!r}")


def _allocate_block_a_phase(
    a_states: list[_RowAllocationState],
    queues: dict[tuple[str, str, str, str], deque[_PackingUnit]],
    phase: str,
    *,
    states_by_queue_key: dict[
        tuple[str, str, str, str], list[_RowAllocationState]
    ]
    | None = None,
) -> None:
    """Run one global A phase, scanning each row's folder deques in order.

    Does not concatenate deques. Stops walking folders when remaining qty
    hits ``_EPS``. A3/A4 reservation is computed live per folder key.
    """
    reserve = phase in {"mto_tags", "blind"}
    for state in a_states:
        for queue_key in state.queue_keys:
            if state.remaining_qty() <= _EPS:
                break
            queue = queues.get(queue_key)
            if not queue:
                continue
            reserved: set[str] | frozenset[str] = frozenset()
            if reserve and states_by_queue_key is not None:
                reserved = _reserved_tags_for_state(
                    state, states_by_queue_key.get(queue_key, [])
                )
            _allocate_queue_phase(queue, state, phase, reserved_tags=reserved)


def _drain_leftover_queues(
    queues: dict[tuple[str, str, str, str], deque[_PackingUnit]],
) -> dict[tuple[str, str, str], deque[_PackingUnit]]:
    """Move remaining A-queue units onto leftover deques keyed by 3-tuple.

    FIFO: 4-tuple queues are walked in insertion order; each unit is
    ``popleft`` onto ``leftover_queues[key[1:]]``. Does not sort and does
    not invent a 4-tuple with empty actual inside A queues.
    """
    leftover_queues: dict[tuple[str, str, str], deque[_PackingUnit]] = defaultdict(
        deque
    )
    for key, queue in queues.items():
        code_key = key[1:]
        while queue:
            leftover_queues[code_key].append(queue.popleft())
    return leftover_queues


def _allocate_leftover_block(
    participants: list[_RowAllocationState],
    leftover_queues: dict[tuple[str, str, str], deque[_PackingUnit]],
    *,
    code_column: str,
    tags_fn: Callable[[_RowAllocationState], list[str]],
) -> None:
    """Run leftover own-tag / untagged / blind phases for one non-RFP block.

    Args:
        participants: B or C rows in tagged-first original order.
        leftover_queues: Remaining units keyed by ``(title, system, code)``.
        code_column: ``CODE_MTO`` (B) or ``CODE_VO`` (C).
        tags_fn: Own tags for this block (never sets ``via_mto``).
    """
    code_keys: dict[int, tuple[str, str, str] | None] = {
        id(state): _row_leftover_code_key(state.row, code_column)
        for state in participants
    }
    by_code_key: dict[tuple[str, str, str], list[_RowAllocationState]] = defaultdict(
        list
    )
    for state in participants:
        code_key = code_keys[id(state)]
        if code_key is not None:
            by_code_key[code_key].append(state)

    for state in participants:
        code_key = code_keys[id(state)]
        queue = leftover_queues.get(code_key) if code_key is not None else None
        if not queue:
            continue
        _allocate_queue_phase(queue, state, "rfp_tags", own_tags=tags_fn(state))
    for state in participants:
        code_key = code_keys[id(state)]
        queue = leftover_queues.get(code_key) if code_key is not None else None
        if not queue:
            continue
        _allocate_queue_phase(queue, state, "untagged", own_tags=tags_fn(state))
    for state in participants:
        code_key = code_keys[id(state)]
        queue = leftover_queues.get(code_key) if code_key is not None else None
        if not queue:
            continue
        reserved = (
            frozenset()
            if code_key is None
            else _reserved_leftover_tags(state, by_code_key[code_key], tags_fn)
        )
        _allocate_queue_phase(
            queue,
            state,
            "blind",
            reserved_tags=reserved,
            own_tags=tags_fn(state),
        )


def _fill_unit_details(row: RowStd, units: list[_PackingUnit]) -> None:
    if not units:
        return
    _set_cell(row, UL_CODE, _joined(unit.source_row.get_value(CODE) for unit in units))
    _set_cell(row, UL_SOURCE_FILES, _format_ul_source_files(unit.source for unit in units))
    _set_cell(row, UL_NAME, _joined(unit.source_row.get_value(NAME) for unit in units))
    _set_cell(
        row,
        UL_TYPE_MARK,
        _joined(unit.source_row.get_value(TYPE_MARK) for unit in units),
    )
    _set_cell(row, UL_VENDOR, _joined(unit.source_row.get_value(VENDOR) for unit in units))
    _set_cell(row, UL_TAGS, _joined(unit.tag for unit in units))
    _set_cell(row, UL_UNITS, _joined(unit.units for unit in units))
    _merge_packing_units_trace(row, units)


def _merge_packing_units_trace(row: RowStd, units: list[_PackingUnit]) -> None:
    """Attach UL source status/trace without overwriting existing RFP/MTO traces."""
    for column in (UNITS_CHECK_STATUS, UNITS_CONVERSION_TRACE):
        if column not in row.el:
            row.el[column] = CheckElement(None)
    if not any(
        unit.units_check_status or unit.units_conversion_trace for unit in units
    ):
        return
    for column, attribute in (
        (UNITS_CHECK_STATUS, "units_check_status"),
        (UNITS_CONVERSION_TRACE, "units_conversion_trace"),
    ):
        merged = join_unique_units_text(
            [row.get_value(column), *(getattr(unit, attribute) for unit in units)]
        )
        row.el[column].value = merged or None


def _status_for_remaining(remaining: float) -> tuple[str, str]:
    if abs(remaining) < _EPS:
        return STATUS_COMPLETE, Color.green
    if remaining > 0:
        return STATUS_SHORTFALL, Color.yellow
    return STATUS_OVERDELIVERY, Color.soft_cyan


def _status_for_non_rfp_allocated(row: RowStd, remaining: float) -> tuple[str, str]:
    """Return UL status for MTO/VO-only rows that received packing units.

    Args:
        row: Result row without RFP presence.
        remaining: Ordered qty minus delivered packing qty.

    Returns:
        Status text and fill color. Remaining below zero falls back to
        ``_status_for_remaining`` (allocation is capped at remaining qty).
    """
    if remaining < -_EPS:
        return _status_for_remaining(remaining)
    has_mto = _has_mto(row)
    if abs(remaining) <= _EPS:
        if has_mto:
            return STATUS_MTO_DELIVERED, Color.match_matched
        if _has_vo(row):
            return STATUS_VO_DELIVERED, Color.match_matched
        return _status_for_remaining(remaining)
    if has_mto:
        return STATUS_MTO_DELIVERED_SHORT, Color.yellow
    if _has_vo(row):
        return STATUS_VO_DELIVERED_SHORT, Color.yellow
    return _status_for_remaining(remaining)


def _apply_non_rfp_allocated_ds_actual(
    row: RowStd,
    allocated: list[_PackingUnit],
) -> None:
    """Set ``DS_ACTUAL`` from allocated UL folders; never invent ``DS_NAME``."""
    actuals = {packing_row_actual_ds(unit.source_row) for unit in allocated}
    if len(actuals) == 1:
        _set_ds_actual(row, next(iter(actuals)), missing_comment=False)
        return
    _set_ds_actual(row, None, missing_comment=False)


def _allocation_comment(
    key: tuple[str, str, str, str],
    units: list[_PackingUnit],
    *,
    fallback: str = "",
    extra: Iterable[str] = (),
) -> str:
    lines = [
        f"Ключ УЛ: {key!r}",
        f"Распределено: {_delivered_qty(units):g} (слотов: {len(units)})",
    ]
    if fallback:
        lines.append(f"{FALLBACK_COMMENT_MARKER} {fallback}={key[3]!r}")
    sources = _unique_text(unit.source for unit in units)
    if sources:
        lines.append("Источники:")
        lines.extend(f"- {source}" for source in sources)
    lines.extend(extra)
    return "\n".join(lines)


def _fill_tag_match_status(row: RowStd, units: list[_PackingUnit]) -> None:
    if any(unit.tag_mismatch for unit in units):
        _set_cell(
            row,
            UL_TAG_MATCH_STATUS,
            STATUS_TAG_MISMATCH_COLLECTED,
            color=Color.yellow,
        )
    elif any(unit.via_mto for unit in units):
        _set_cell(
            row,
            UL_TAG_MATCH_STATUS,
            STATUS_TAG_MATCHED_VIA_MTO,
            color=Color.soft_yellow,
        )


def _fill_result_row(
    row: RowStd,
    key: tuple[str, str, str, str],
    ordered: float,
    allocated: list[_PackingUnit],
    dataset_status: tuple[str, str],
    audit: RfpPackingAudit,
    *,
    fallback: str = "",
) -> None:
    delivered = _delivered_qty(allocated)
    remaining = ordered - delivered
    unit_problem = False
    row_units = normalize_units_text(row.get_value(UNITS))
    packing_units = _unique_text(unit.units for unit in allocated)
    extra: list[str] = []
    if allocated and row_units and any(unit and unit != row_units for unit in packing_units):
        unit_problem = True
        audit.stats.unit_mismatches += 1
        issue = PackingIssue(
            code="packing_units_mismatch",
            message=(
                f"Единица RFP {row_units!r} не совпадает с УЛ "
                f"{' / '.join(packing_units)!r}"
            ),
            field=UNITS,
            raw_value={"rfp": row_units, "packing": packing_units},
            action="Проверьте единицы измерения RFP и упаковочного листа",
        )
        audit.issues.append(issue)
        extra.append(issue.format_line())
    comment = _allocation_comment(key, allocated, fallback=fallback, extra=extra)
    _fill_unit_details(row, allocated)
    _fill_tag_match_status(row, allocated)
    if allocated and not _has_rfp(row):
        _apply_non_rfp_allocated_ds_actual(row, allocated)
    _set_cell(row, UL_ORDERED_VALUES, float(ordered))
    if not allocated:
        status, color = _presence_status(row)
        if ordered:
            _set_cell(row, UL_REMAINING_VALUES, remaining)
        _set_cell(row, UL_COMPARE_STATUS, status, color=color, comment=comment)
        _set_cell(
            row,
            UL_DATA_STATUS,
            dataset_status[0],
            color=dataset_status[1],
        )
        audit.stats.not_in_packing += 1
        return
    _set_cell(row, UL_VALUES, delivered)
    if unit_problem:
        _set_cell(row, UL_REMAINING_VALUES, "")
    else:
        _set_cell(row, UL_REMAINING_VALUES, remaining)
    if unit_problem:
        row.el[UL_UNITS].color = Color.yellow
        _set_cell(
            row,
            UL_COMPARE_STATUS,
            STATUS_DATA_PROBLEM,
            color=Color.red,
            comment=comment,
        )
        _set_cell(
            row,
            UL_DATA_STATUS,
            DATA_ROW_PROBLEM,
            color=Color.red,
        )
        return
    if not _has_rfp(row):
        status, color = _status_for_non_rfp_allocated(row, remaining)
    else:
        status, color = _status_for_remaining(remaining)
    _set_cell(row, UL_COMPARE_STATUS, status, color=color, comment=comment)
    _set_cell(
        row,
        UL_DATA_STATUS,
        dataset_status[0],
        color=dataset_status[1],
    )
    if status == STATUS_COMPLETE:
        audit.stats.complete += 1
        if not allocated:
            audit.stats.complete_zero_allocated += 1
            if fallback:
                audit.stats.complete_zero_fallback += 1
            if not str(row.get_value(CODE) or "").strip():
                audit.stats.complete_zero_empty_code += 1
    elif status == STATUS_SHORTFALL:
        audit.stats.shortfall += 1
    elif status == STATUS_OVERDELIVERY:
        audit.stats.overdelivery += 1


def _packing_only_group_key(unit: _PackingUnit) -> tuple[str, ...]:
    row = unit.source_row
    return (
        *unit.key,
        unit.tag,
        unit.units,
        str(row.get_value(NAME) or "").strip(),
        str(row.get_value(TYPE_MARK) or "").strip(),
        str(row.get_value(VENDOR) or "").strip(),
    )


def _packing_only_display_title(source: RowStd) -> str:
    """Return RFP-style ``8950-SOO1`` for a packing leftover row.

    The TSD loader stores title and system in separate cells. Step4 Excel
    exports only ``DS_TITLE`` as «Титул/Марка», so leftovers must recombine
    the pair or the mark disappears from the sheet.
    """
    title = str(source.get_value(DS_TITLE) or "").strip()
    system = str(source.get_value(DS_SYSTEM) or "").strip()
    if title and system and parse_composite_title(title) is None:
        return f"{title}-{system}"
    return title


def _create_packing_only_row(
    units: list[_PackingUnit],
    dataset_status: tuple[str, str],
) -> RowStd:
    first = units[0]
    row = RowStd.get_std_check_row({}, TableComments())
    row.row_type = RowType.position_row
    row.el[ROW_TYPE].value = RowType.position_row
    _clear_ul_columns(row)
    row.el[DS_TITLE].value = _packing_only_display_title(first.source_row)
    row.el[DS_SYSTEM].value = first.source_row.get_value(DS_SYSTEM)
    for column in (DS_TITLE, DS_SYSTEM):
        row.el[column].color = Color.soft_cyan
    _fill_packing_only_ds_identity(row, first.source_row)
    _fill_unit_details(row, units)
    delivered = _delivered_qty(units)
    comment = _allocation_comment(first.key, units)
    _set_cell(row, UL_ORDERED_VALUES, 0.0, color=Color.soft_cyan)
    _set_cell(row, UL_VALUES, delivered, color=Color.soft_cyan)
    _set_cell(
        row,
        UL_REMAINING_VALUES,
        -delivered,
        color=Color.soft_cyan,
    )
    _set_cell(
        row,
        UL_COMPARE_STATUS,
        STATUS_PACKING_ONLY,
        color=Color.soft_cyan,
        comment=comment,
    )
    _set_cell(row, UL_DATA_STATUS, dataset_status[0], color=dataset_status[1])
    for column in (UL_CODE, UL_SOURCE_FILES, UL_NAME, UL_TYPE_MARK, UL_VENDOR, UL_TAGS, UL_UNITS):
        if row.el[column].value:
            row.el[column].color = Color.soft_cyan
    return row


def ul_name_is_zip_smr_pnr(name: object) -> bool:
    """Return True when the UL name denotes a ZIP kit for SMR/PNR."""
    collapsed = " ".join(str(name or "").split())
    return ZIP_SMR_PNR_NAME_RE.search(collapsed) is not None


def apply_zip_smr_pnr_status(rows: list[RowStd]) -> int:
    """Overlay UL compare status for ZIP-for-SMR/PNR packing rows.

    Mutates ``UL_COMPARE_STATUS`` in place (value + color only; comment and
    quantity columns are left unchanged). For leftover ``STATUS_PACKING_ONLY``
    rows, cyan highlight is recast to ``Color.match_matched`` on the same
    columns as ``_create_packing_only_row`` except ``UL_DATA_STATUS``.

    Args:
        rows: Step4 result rows to scan.

    Returns:
        Number of rows whose UL status was changed.
    """
    changed = 0
    for row in rows:
        if row.row_type != RowType.position_row:
            continue
        if not ul_name_is_zip_smr_pnr(row.get_value(UL_NAME)):
            continue
        status = str(row.get_value(UL_COMPARE_STATUS) or "").strip()
        if status in {
            STATUS_DATA_PROBLEM,
            STATUS_UNAVAILABLE,
            STATUS_ZIP_SMR_PNR,
            STATUS_EXCLUDED_FROM_SUPPLY,
        }:
            continue
        if UL_COMPARE_STATUS not in row.el:
            row.el[UL_COMPARE_STATUS] = CheckElement(None)
        was_packing_only = status == STATUS_PACKING_ONLY
        el = row.el[UL_COMPARE_STATUS]
        el.value = STATUS_ZIP_SMR_PNR
        el.color = Color.match_matched
        if was_packing_only:
            for column in (
                DS_TITLE,
                DS_SYSTEM,
                UL_ORDERED_VALUES,
                UL_VALUES,
                UL_REMAINING_VALUES,
                UL_COMPARE_STATUS,
            ):
                if column in row.el:
                    row.el[column].color = Color.match_matched
            for column in (
                UL_CODE,
                UL_SOURCE_FILES,
                UL_NAME,
                UL_TYPE_MARK,
                UL_VENDOR,
                UL_TAGS,
                UL_UNITS,
            ):
                if column in row.el and row.el[column].value:
                    row.el[column].color = Color.match_matched
        changed += 1
    return changed


def _create_invalid_key_row(source: RowStd, issue: PackingIssue) -> RowStd:
    unit = _PackingUnit(
        key=("", "", "", ""),
        source_row=source,
        tag=_joined(source.get_tags_list()),
        units=normalize_units_text(source.get_value(UNITS)),
        source=_source_text(source),
        units_check_status=str(source.el.get(UNITS_CHECK_STATUS, CheckElement(None)).value or "").strip(),
        units_conversion_trace=str(
            source.el.get(UNITS_CONVERSION_TRACE, CheckElement(None)).value or ""
        ).strip(),
    )
    row = _create_packing_only_row([unit], (DATA_ROW_PROBLEM, Color.red))
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
    )
    return row


def compare_rfp_rows_with_packing(
    result_rows: list[RowStd],
    ordered_by_key: dict[tuple[str, str, str, str], float],
    dataset: PackingDataset,
    *,
    use_mto_tags: bool = True,
) -> tuple[list[RowStd], RfpPackingAudit]:
    """Allocate packing deliveries once across the current Step4 result.

    Block A runs four global passes over rows that have an actual DS, using
    queues keyed by physical UL folder ``(folder_key, title, system, code)``:
    own RFP tags, untagged slots, optional MTO tags, then blind same-code
    fill. A row takes from eligible folders for its supply id
    (``rfp_row_actual_ds``): registry folder keys when the planting index
    knows that id, otherwise every folder in this dataset whose parsed
    name actual equals the id. Remaining units are reindexed by
    ``(title, system, code)``. Block B plants that leftover onto MTO-only
    rows (including MTO+VO on the same row). Block C plants whatever is
    left onto VO-only rows. Units still unused become packing-only leftover
    rows. Folder queues are not cloned; two DS numbers debit separate
    ordered snapshots.

    Args:
        result_rows: Fully matched and already collapsed Step4 rows.
        ordered_by_key: Snapshot built from original RFP position rows.
        dataset: Validated shared packing cache dataset.
        use_mto_tags: If True, after RFP tags and untagged slots, match UL
            tags against ``TAG_MTO`` of the same row (skipping tags still
            reserved by a sibling on the same folder queue) before blind
            same-code fill.

    Returns:
        The enriched list and a complete audit object.

    Raises:
        RfpPackingFatalError: If any packing row has more tags than quantity.
    """
    stats = RfpPackingStats()
    audit = RfpPackingAudit(dataset=dataset, stats=stats, issues=list(dataset.issues))
    excluded_ids: set[int] = set()
    with _phase(audit, "prepare_rows"):
        for row in result_rows:
            _clear_ul_columns(row)
            if row.row_type == RowType.position_row:
                stats.result_position_rows += 1
                _set_ds_actual(
                    row,
                    rfp_row_actual_ds(row),
                    missing_comment=_has_rfp(row),
                )
                if is_excluded_from_supply(row):
                    _set_cell(
                        row,
                        UL_COMPARE_STATUS,
                        STATUS_EXCLUDED_FROM_SUPPLY,
                        color=Color.yellow,
                    )
                    excluded_ids.add(id(row))

    dataset_status = _data_status(dataset)
    if not dataset.available:
        comment = "\n".join(issue.format_line() for issue in dataset.issues)
        for row in result_rows:
            if row.row_type != RowType.position_row:
                continue
            if id(row) in excluded_ids:
                continue
            parsed = parse_composite_title(row.get_value(DS_TITLE))
            actual = rfp_row_actual_ds(row)
            if actual is None or parsed is None:
                ordered = 0.0
            else:
                key = rfp_packing_match_key(actual, *parsed, row.get_value(CODE))
                ordered = ordered_by_key.get(key, 0.0)
            _set_cell(row, UL_ORDERED_VALUES, ordered)
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
            )
        return result_rows, audit

    with _phase(audit, "build_unit_queues"):
        queues, problem_rows, matcher_issues = _build_unit_queues(dataset, stats)
        audit.issues.extend(matcher_issues)
    if any(issue.code == "packing_tags_exceed_quantity" for issue in matcher_issues):
        raise RfpPackingFatalError(audit)

    remaining_ordered = dict(ordered_by_key)
    name_folders = _dataset_folders_by_parsed_actual(dataset)
    position_rows = [
        row for row in result_rows if row.row_type == RowType.position_row
    ]
    row_index_by_id = {id(row): index for index, row in enumerate(position_rows)}
    a_states: list[_RowAllocationState] = []
    with _phase(audit, "a_states_build"):
        allocation_order = sorted(
            enumerate(position_rows),
            key=lambda item: (not bool(item[1].get_tags_list()), item[0]),
        )
        for _original_index, row in allocation_order:
            if id(row) in excluded_ids:
                continue
            actual = rfp_row_actual_ds(row)
            eligible = (
                _eligible_folder_keys(actual, name_folders)
                if actual is not None
                else frozenset()
            )
            key, fallback, queue_keys = _candidate_key(row, queues, eligible)
            if key is None:
                parsed = parse_composite_title(row.get_value(DS_TITLE))
                if parsed is None:
                    audit.issues.append(
                        PackingIssue(
                            code="rfp_composite_title",
                            message=(
                                f"Некорректный DS_TITLE Step4 "
                                f"{row.get_value(DS_TITLE)!r}; ожидается полный формат 8950-SOO1"
                            ),
                            field=DS_TITLE,
                            raw_value=row.get_value(DS_TITLE),
                            action="Исправьте составной титул/марку в RFP",
                        )
                    )
                    ordered = _row_capacity(row) if row.get_value(CODE) else 0.0
                    _set_cell(row, UL_ORDERED_VALUES, ordered)
                    if ordered:
                        _set_cell(row, UL_REMAINING_VALUES, ordered)
                    _set_cell(
                        row,
                        UL_COMPARE_STATUS,
                        STATUS_DATA_PROBLEM,
                        color=Color.red,
                    )
                    _set_cell(
                        row,
                        UL_DATA_STATUS,
                        DATA_ROW_PROBLEM,
                        color=Color.red,
                    )
                    stats.not_in_packing += 1
                    if len(audit.miss_samples) < _MISS_SAMPLE_LIMIT:
                        audit.miss_samples.append(
                            f"title={row.get_value(DS_TITLE)!r}, "
                            f"codes={[row.get_value(c) for c in (CODE, CODE_MTO, CODE_VO)]!r}"
                        )
                continue

            capacity = _row_capacity(row)
            parsed = parse_composite_title(row.get_value(DS_TITLE))
            actual = rfp_row_actual_ds(row)
            primary_order_key = (
                rfp_packing_match_key(actual, *parsed, row.get_value(CODE))
                if parsed is not None and actual is not None
                else key
            )
            if not all(primary_order_key):
                primary_order_key = key
            snapshot_remaining = remaining_ordered.get(primary_order_key, 0.0)
            rfp_code = str(row.get_value(CODE) or "").strip()
            if snapshot_remaining > 0:
                ordered = min(capacity, snapshot_remaining)
                remaining_ordered[primary_order_key] = max(
                    0.0, snapshot_remaining - ordered
                )
            elif not rfp_code and fallback:
                ordered = capacity
            else:
                ordered = 0.0
            a_states.append(
                _RowAllocationState(
                    row=row,
                    key=key,
                    fallback=fallback,
                    ordered=ordered,
                    rfp_tags=_unique_text(row.get_tags_list()),
                    mto_tags=_coerce_tag_values(row.get_value(TAG_MTO)),
                    vo_tags=_coerce_tag_values(row.get_value(TAG_VO)),
                    queue_keys=queue_keys,
                )
            )
        states_by_queue_key: dict[
            tuple[str, str, str, str], list[_RowAllocationState]
        ] = defaultdict(list)
        for state in a_states:
            for queue_key in state.queue_keys:
                states_by_queue_key[queue_key].append(state)

    with _phase(audit, "a1_rfp_tags"):
        _allocate_block_a_phase(a_states, queues, "rfp_tags")
    with _phase(audit, "a2_untagged"):
        _allocate_block_a_phase(a_states, queues, "untagged")
    if use_mto_tags:
        with _phase(audit, "a3_mto_tags"):
            _allocate_block_a_phase(
                a_states,
                queues,
                "mto_tags",
                states_by_queue_key=states_by_queue_key,
            )
    with _phase(audit, "a4_blind"):
        _allocate_block_a_phase(
            a_states,
            queues,
            "blind",
            states_by_queue_key=states_by_queue_key,
        )

    with _phase(audit, "drain_leftover"):
        leftover_queues = _drain_leftover_queues(queues)
    a_row_ids = {id(state.row) for state in a_states}

    new_b_states: list[_RowAllocationState] = []
    with _phase(audit, "block_b"):
        b_participants: list[_RowAllocationState] = []
        for state in a_states:
            if (
                not _has_rfp(state.row)
                and _has_mto(state.row)
                and state.remaining_qty() > _EPS
            ):
                b_participants.append(state)
        for row in position_rows:
            if id(row) in a_row_ids or id(row) in excluded_ids:
                continue
            parsed = parse_composite_title(row.get_value(DS_TITLE))
            if parsed is None:
                continue
            if not _has_rfp(row) and _has_mto(row):
                title, system = parsed
                state = _RowAllocationState(
                    row=row,
                    key=rfp_packing_match_key(
                        "", title, system, row.get_value(CODE_MTO)
                    ),
                    fallback="",
                    ordered=_row_capacity(row),
                    mto_tags=_coerce_tag_values(row.get_value(TAG_MTO)),
                    vo_tags=_coerce_tag_values(row.get_value(TAG_VO)),
                )
                new_b_states.append(state)
                b_participants.append(state)
        b_participants = _order_leftover_participants(
            b_participants, row_index_by_id, lambda state: state.mto_tags
        )
        _allocate_leftover_block(
            b_participants,
            leftover_queues,
            code_column=CODE_MTO,
            tags_fn=lambda state: state.mto_tags,
        )

    new_c_states: list[_RowAllocationState] = []
    with _phase(audit, "block_c"):
        c_participants: list[_RowAllocationState] = []
        for state in a_states:
            if (
                not _has_rfp(state.row)
                and not _has_mto(state.row)
                and _has_vo(state.row)
                and state.remaining_qty() > _EPS
            ):
                c_participants.append(state)
        for row in position_rows:
            if id(row) in a_row_ids or id(row) in excluded_ids:
                continue
            parsed = parse_composite_title(row.get_value(DS_TITLE))
            if parsed is None:
                continue
            if not _has_rfp(row) and not _has_mto(row) and _has_vo(row):
                title, system = parsed
                state = _RowAllocationState(
                    row=row,
                    key=rfp_packing_match_key(
                        "", title, system, row.get_value(CODE_VO)
                    ),
                    fallback="",
                    ordered=_row_capacity(row),
                    vo_tags=_coerce_tag_values(row.get_value(TAG_VO)),
                )
                new_c_states.append(state)
                c_participants.append(state)
        c_participants = _order_leftover_participants(
            c_participants, row_index_by_id, lambda state: state.vo_tags
        )
        _allocate_leftover_block(
            c_participants,
            leftover_queues,
            code_column=CODE_VO,
            tags_fn=lambda state: state.vo_tags,
        )

    covered_ids = (
        excluded_ids
        | a_row_ids
        | {id(state.row) for state in new_b_states}
        | {id(state.row) for state in new_c_states}
    )
    presence_states: list[_RowAllocationState] = []
    with _phase(audit, "presence_states"):
        for row in position_rows:
            if id(row) in covered_ids:
                continue
            parsed = parse_composite_title(row.get_value(DS_TITLE))
            if parsed is None:
                continue
            ordered = _row_capacity(row) if row.get_value(CODE) else 0.0
            actual = rfp_row_actual_ds(row)
            key = rfp_packing_match_key(actual or "", *parsed, row.get_value(CODE))
            presence_states.append(
                _RowAllocationState(row=row, key=key, fallback="", ordered=ordered)
            )
            if len(audit.miss_samples) < _MISS_SAMPLE_LIMIT:
                audit.miss_samples.append(
                    f"title={row.get_value(DS_TITLE)!r}, "
                    f"codes={[row.get_value(c) for c in (CODE, CODE_MTO, CODE_VO)]!r}"
                )

    fill_states = a_states + new_b_states + new_c_states + presence_states
    with _phase(audit, "fill_rows"):
        for state in fill_states:
            _fill_result_row(
                state.row,
                state.key,
                state.ordered,
                state.allocated,
                dataset_status,
                audit,
                fallback=state.fallback,
            )
            stats.allocated_units += len(state.allocated)
            if state.allocated:
                stats.matched_rows += 1
            if state.fallback:
                stats.fallback_code_matches += 1

    with _phase(audit, "leftover_dump"):
        leftover_groups: dict[tuple[str, ...], list[_PackingUnit]] = defaultdict(list)
        for queue in leftover_queues.values():
            while queue:
                unit = queue.popleft()
                leftover_groups[_packing_only_group_key(unit)].append(unit)
        for units in leftover_groups.values():
            result_rows.append(_create_packing_only_row(units, dataset_status))
            stats.packing_only_added += 1
            stats.overdelivery += 1

    with _phase(audit, "invalid_key_rows"):
        for source, issue in problem_rows:
            result_rows.append(_create_invalid_key_row(source, issue))
            stats.packing_only_added += 1

    return result_rows, audit


def save_rfp_packing_report(
    audit: RfpPackingAudit,
    out_dir: str | Path,
) -> str:
    """Write the full UTF-8 Step4 packing audit.

    Args:
        audit: Completed audit, including fatal pre-export audits.
        out_dir: Existing or desired result directory.

    Returns:
        Path to the written report.
    """
    output_dir = Path(out_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y.%m.%d_%H.%M.%S")
    path = output_dir / f"ul_compare_report_{stamp}.txt"
    meta = audit.dataset.meta
    lines = [
        "=== RFP STEP4: СРАВНЕНИЕ С УПАКОВОЧНЫМИ ЛИСТАМИ ===",
        f"quality: {audit.quality.value}",
        f"cache_path: {audit.dataset.cache_path}",
        f"cache_version: {meta.version if meta else ''}",
        f"cache_created_at: {meta.created_at if meta else ''}",
        f"source_root: {meta.root if meta else ''}",
        f"cache_fingerprint: {meta.fingerprint if meta else ''}",
        "",
        "COUNTERS:",
    ]
    lines.extend(f"  {name}: {value}" for name, value in vars(audit.stats).items())
    lines.extend(["", "PHASES (seconds):"])
    if audit.phase_timings:
        lines.extend(
            f"  {name}: {seconds:.3f}"
            for name, seconds in sorted(
                audit.phase_timings.items(), key=lambda item: -item[1]
            )
        )
    else:
        lines.append("  (none)")
    if meta is not None:
        lines.extend(["", "LOADER COUNTERS:"])
        lines.extend(
            f"  {name}: {value}" for name, value in sorted(meta.counters.items())
        )
        lines.extend(["", "SOURCE REPORTS:"])
        lines.extend(
            f"  {name}: {value}"
            for name, value in sorted(meta.report_paths.items())
        )
    lines.extend(["", f"ISSUES ({len(audit.issues)}):"])
    if audit.issues:
        lines.extend(
            f"  {index}. {issue.format_line()}"
            for index, issue in enumerate(audit.issues, 1)
        )
    else:
        lines.append("  (none)")
    lines.extend(["", "SAMPLE KEYS NOT FOUND IN UL:"])
    lines.extend(f"  - {sample}" for sample in audit.miss_samples)
    if not audit.miss_samples:
        lines.append("  (none)")
    path.write_text("\n".join(lines) + "\n", encoding="utf-8")
    audit.report_path = str(path)
    return str(path)


# Compatibility aliases for orchestration drafts and external smoke fixtures.
Step4PackingCompareAudit = RfpPackingAudit
compare_step4_rows_with_packing = compare_rfp_rows_with_packing
save_step4_packing_compare_report = save_rfp_packing_report
