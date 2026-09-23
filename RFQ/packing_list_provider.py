"""Shared, validated access to the packing-list cache.

The provider is intentionally independent from DS and RFP comparison code so
the same cache contract can be reused by both pipelines.
"""

from __future__ import annotations

import contextlib
import os
import pickle
import re
from contextvars import ContextVar
from dataclasses import dataclass, field
from enum import Enum
from pathlib import Path
from typing import Any, Iterator, Mapping

import base.t_comm_initial_classes as t_com_init_cls
from base.base_classes import CheckElement, RowStd, RowType, TableComments
from base.tables_columns import (
    ANNOTATION,
    CODE,
    CODE_2,
    DS_NAME,
    DS_SYSTEM,
    DS_TITLE,
    NAME,
    ROW_TYPE,
    TAGS,
    TITLE,
    TYPE_MARK,
    UNITS,
    VALUES,
    VENDOR,
)
from RFQ.rfp_parts.ds_identity import parse_rfp_ds_identity, parse_ul_folder_ds_identity

DEFAULT_PACKING_OUTPUT_DIR = Path(
    r"\\bcc\eng\PrDoc\377_НИПИГАЗ\АГХК\КСБ\RFP_MTO_VO\_RFP\УЛ сводный файл"
)
PACKING_CACHE_NAME = "tsd_packing_rows.cache"
# v6: slim RowStd (only compare columns; no full ColNames / UNC TableComments).
PACKING_CACHE_VERSION = "v6"

# Columns persisted in the packing pickle (same row set as summary xlsx positions).
# SPECIFICATION_NAME stays in the human summary only — compare uses DS_TITLE/DS_SYSTEM.
PACKING_CACHE_COLUMNS: tuple[str, ...] = (
    DS_TITLE,
    DS_SYSTEM,
    CODE,
    VALUES,
    UNITS,
    TAGS,
    NAME,
    TYPE_MARK,
    VENDOR,
    ANNOTATION,
    TITLE,
    ROW_TYPE,
)

_WHITESPACE_RE = re.compile(r"[\s\u00a0]+")

# Shared lightweight TableComments for cached packing rows (no source UNC paths).
_SLIM_PACKING_T_COM = TableComments(
    dir_path="-1",
    file_full_path="",
    tabel_class=t_com_init_cls.TsdPacking,
)


def slim_packing_row(row: RowStd) -> RowStd:
    """Build a cache-sized ``RowStd`` with only packing-compare columns.

    Avoids pickling ~135 empty ``CheckElement`` slots and per-sheet UNC
    ``TableComments``. Callers must pass a ``position_row``.

    Args:
        row: Enriched packing position row (may be a full ``RowStd``).

    Returns:
        New slim ``RowStd`` suitable for ``pickle.dump``.
    """
    out = object.__new__(RowStd)
    out.el = {}
    for key in PACKING_CACHE_COLUMNS:
        src = row.el.get(key) if hasattr(row, "el") else None
        value = src.value if src is not None else None
        out.el[key] = CheckElement(value)
    out.row_type = RowType.position_row
    out.el[ROW_TYPE].value = RowType.position_row
    out.t_com = _SLIM_PACKING_T_COM
    excel_row = getattr(row, "_packing_source_row", None)
    if isinstance(excel_row, int) and excel_row > 0:
        out._packing_source_row = excel_row
    return out


def slim_packing_rows(rows: list[RowStd]) -> list[RowStd]:
    """Slim every packing position row for cache persistence."""
    return [slim_packing_row(row) for row in rows]


class PackingQualityLevel(str, Enum):
    """Overall usability of a packing dataset."""

    OK = "ok"
    PARTIAL = "partial"
    UNAVAILABLE = "unavailable"


class PackingIssueSeverity(str, Enum):
    """Severity of one cache or source-data issue."""

    WARNING = "warning"
    ERROR = "error"


@dataclass(frozen=True)
class PackingIssue:
    """One actionable cache or packing-row problem."""

    code: str
    message: str
    severity: PackingIssueSeverity = PackingIssueSeverity.ERROR
    source_file: str = ""
    sheet: str = ""
    excel_row: int | None = None
    field: str = ""
    raw_value: object = None
    action: str = ""

    def location(self) -> str:
        """Return a compact file/sheet/row location."""
        parts: list[str] = []
        if self.source_file:
            parts.append(self.source_file)
        if self.sheet:
            parts.append(self.sheet)
        if self.excel_row is not None:
            parts.append(f"строка {self.excel_row}")
        return " · ".join(parts)

    def format_line(self) -> str:
        """Return a human-readable report line."""
        prefix = "ERROR" if self.severity == PackingIssueSeverity.ERROR else "WARN"
        location = f" [{self.location()}]" if self.location() else ""
        field = f" field={self.field}" if self.field else ""
        raw = f" value={self.raw_value!r}" if self.raw_value is not None else ""
        action = f" Исправление: {self.action}" if self.action else ""
        return f"{prefix} {self.code}{location}{field}{raw}: {self.message}.{action}"


@dataclass(frozen=True)
class PackingCacheMeta:
    """Validated metadata saved with the packing rows."""

    version: str
    created_at: str
    root: str
    fingerprint: str
    counters: dict[str, int] = field(default_factory=dict)
    critical_ok: bool = False
    critical_remarks: tuple[str, ...] = ()
    report_paths: dict[str, str] = field(default_factory=dict)


@dataclass
class PackingDataset:
    """Packing rows plus explicit quality and diagnostics."""

    rows: list[RowStd]
    meta: PackingCacheMeta | None
    quality: PackingQualityLevel
    issues: list[PackingIssue]
    cache_path: str

    @property
    def available(self) -> bool:
        """Whether rows can be used, possibly with partial-quality warnings."""
        return self.quality != PackingQualityLevel.UNAVAILABLE

    @property
    def report_paths(self) -> dict[str, str]:
        """Source loader reports recorded in cache metadata."""
        return dict(self.meta.report_paths) if self.meta is not None else {}

    def format_short(self) -> str:
        """Return a short user-facing quality verdict."""
        if self.quality == PackingQualityLevel.OK:
            return f"УЛ: кэш OK, строк={len(self.rows)}"
        if self.quality == PackingQualityLevel.PARTIAL:
            return (
                f"УЛ: частичные/непроверенные данные, строк={len(self.rows)}, "
                f"проблем={len(self.issues)}"
            )
        return f"УЛ: данные недоступны, проблем={len(self.issues)}"


def default_packing_cache_path() -> Path:
    """Return the default network cache path without creating directories."""
    return DEFAULT_PACKING_OUTPUT_DIR / PACKING_CACHE_NAME


@dataclass(frozen=True, slots=True)
class RegistryPackingIndex:
    """Actual DS numbers taken from the registry, not from folder or file names.

    ``known_source_ids`` are every active digit-only source. ``ds_to_folders``
    holds those ids that name at least one UL folder. ``folder_owners`` maps
    a normalized folder key to every active source that names it.
    ``rfp_number_to_actual`` maps an RFP number to a source only when exactly
    one active source names that number. Folder keys use
    ``normalize_packing_key_part`` (whitespace removed, casefold).
    """

    known_source_ids: frozenset[int] = field(default_factory=frozenset)
    ds_to_folders: dict[int, frozenset[str]] = field(default_factory=dict)
    folder_owners: dict[str, frozenset[int]] = field(default_factory=dict)
    rfp_number_to_actual: dict[int, int] = field(default_factory=dict)


_REGISTRY_PACKING_INDEX: ContextVar[RegistryPackingIndex | None] = ContextVar(
    "registry_packing_index",
    default=None,
)


@contextlib.contextmanager
def registry_packing_scope(
    index: RegistryPackingIndex | None,
) -> Iterator[None]:
    """Use registry actual numbers for UL folders and RFP rows in this block."""

    token = _REGISTRY_PACKING_INDEX.set(index)
    try:
        yield
    finally:
        _REGISTRY_PACKING_INDEX.reset(token)


def build_registry_planting_index(document) -> RegistryPackingIndex:
    """Build planting maps from active registry rows.

    Shared UL folders keep every owner. An RFP number maps to an actual DS
    only when exactly one active source names it. ``validation.is_ok`` is
    ignored so a shared RFP error does not drop folder links. History and
    Disabled rows are not active. Non-digit source ids such as ``4905_1``
    are omitted.

    Args:
        document: Loaded registry document. Only ``active_rows`` are read.

    Returns:
        Index over digit-only actual DS numbers.
    """

    known: set[int] = set()
    folders_by_ds: dict[int, set[str]] = {}
    owners_by_folder: dict[str, set[int]] = {}
    rfp_owners: dict[int, set[int]] = {}
    for row in document.active_rows:
        if not row.source_id.isdigit():
            continue
        actual = int(row.source_id)
        known.add(actual)
        for rel in row.relations:
            if rel.ul_folder:
                key = normalize_packing_key_part(rel.ul_folder)
                if key:
                    folders_by_ds.setdefault(actual, set()).add(key)
                    owners_by_folder.setdefault(key, set()).add(actual)
            if rel.rfp_key.isdigit():
                rfp_owners.setdefault(int(rel.rfp_key), set()).add(actual)
    return RegistryPackingIndex(
        known_source_ids=frozenset(known),
        ds_to_folders={
            ds_id: frozenset(keys) for ds_id, keys in folders_by_ds.items()
        },
        folder_owners={
            key: frozenset(ids) for key, ids in owners_by_folder.items()
        },
        rfp_number_to_actual={
            number: next(iter(ids))
            for number, ids in rfp_owners.items()
            if len(ids) == 1
        },
    )


def supply_ul_folders(actual: int | None) -> frozenset[str] | None:
    """Return UL folder keys named by ``actual`` in the current planting index.

    Args:
        actual: Actual DS number, or ``None``.

    Returns:
        ``None`` when there is no index, ``actual`` is ``None``, or the
        number is not a known digit source (the caller should name-match).
        An empty frozenset when the source is known and names no UL folder.
        The folder set otherwise.
    """

    index = _REGISTRY_PACKING_INDEX.get()
    if index is None or actual is None or actual not in index.known_source_ids:
        return None
    return index.ds_to_folders.get(actual, frozenset())


def normalize_packing_key_part(value: object) -> str:
    """Normalize a title, mark, or code for cross-pipeline matching."""
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    return _WHITESPACE_RE.sub("", str(value).strip()).casefold()


def packing_match_key(
    title: object,
    system: object,
    code: object,
) -> tuple[str, str, str]:
    """Build the common ``(title, system, code)`` packing lookup key."""
    return (
        normalize_packing_key_part(title),
        normalize_packing_key_part(system),
        normalize_packing_key_part(code),
    )


def rfp_packing_match_key(
    actual: object,
    title: object,
    system: object,
    code: object,
) -> tuple[str, str, str, str]:
    """Build Step4 packing key ``(actual_ds, title, system, code)``.

    Args:
        actual: Actual DS number (int or str). Normalized like other key
            parts; ``24`` becomes ``"24"``, never ``"ДС24"``.
        title: Composite title number.
        system: System/mark token.
        code: Position code.

    Returns:
        Four normalized strings. Empty ``actual`` stays ``""``.
    """
    return (
        normalize_packing_key_part(actual),
        *packing_match_key(title, system, code),
    )


def packing_row_key(row: RowStd, *, code_column: str = CODE) -> tuple[str, str, str]:
    """Build a packing lookup key from a ``RowStd``."""
    return packing_match_key(
        row.get_value(DS_TITLE),
        row.get_value(DS_SYSTEM),
        row.get_value(code_column),
    )


def rfp_row_actual_ds(row: RowStd) -> int | None:
    """Return RFP actual DS number from ``DS_NAME``, or None.

    Args:
        row: Step4 / RFP result row.

    Returns:
        ``DsIdentity.actual`` from ``DS_NAME``, or ``None`` when the cell is
        empty or unparsed.
    """
    if DS_NAME not in row.el:
        return None
    ds_name = str(row.get_value(DS_NAME) or "").strip()
    if not ds_name:
        return None
    ident = parse_rfp_ds_identity(ds_name)
    if ident.kind != "compound":
        # Stored ``DS_NAME`` is a prefix without the file-name ``. `` that
        # the compound regex requires; retry so ``ДС92_24Б`` still splits.
        retry = parse_rfp_ds_identity(ds_name + ". ")
        if retry.compound:
            ident = retry
    actual = ident.actual
    index = _REGISTRY_PACKING_INDEX.get()
    if index is not None and actual is not None:
        override = index.rfp_number_to_actual.get(actual)
        if override is not None:
            return override
    return actual


def packing_row_actual_ds(row: RowStd) -> int | None:
    """Return UL folder actual DS from first ANNOTATION path segment, or None.

    Args:
        row: Packing cache row.

    Returns:
        ``DsIdentity.actual`` from the first path segment of ``ANNOTATION``.
        ``None`` when there is no segment, the folder is госфин (``kind=gf``),
        or the name does not parse.
    """
    if ANNOTATION not in row.el:
        return None
    annotation, _sheet, _excel_row = packing_row_source(row)
    normalized = annotation.replace("\\", "/")
    parts = Path(normalized).parts
    folder = parts[0] if parts else ""
    index = _REGISTRY_PACKING_INDEX.get()
    if index is not None and folder:
        owners = index.folder_owners.get(normalize_packing_key_part(folder))
        if owners is not None and len(owners) == 1:
            return next(iter(owners))
    return parse_ul_folder_ds_identity(folder).actual


def packing_codes_for_row(row: RowStd) -> list[str]:
    """Return normalized DS and MTO codes, each at most once."""
    codes: list[str] = []
    for column in (CODE, CODE_2):
        if column not in row.el:
            continue
        code = normalize_packing_key_part(row.get_value(column))
        if code and code not in codes:
            codes.append(code)
    return codes


def packing_row_source(row: RowStd) -> tuple[str, str, int | None]:
    """Return source file, sheet, and Excel row stored by the TSD loader."""
    source_file = str(row.get_value(ANNOTATION) or "").strip()
    sheet = str(row.get_value(TITLE) or "").strip()
    excel_row = getattr(row, "_packing_source_row", None)
    if not isinstance(excel_row, int) or excel_row <= 0:
        excel_row = None
    return source_file, sheet, excel_row


def _unavailable_dataset(path: Path, issue: PackingIssue) -> PackingDataset:
    return PackingDataset(
        rows=[],
        meta=None,
        quality=PackingQualityLevel.UNAVAILABLE,
        issues=[issue],
        cache_path=str(path),
    )


def _int_counters(raw: object) -> dict[str, int]:
    counters: dict[str, int] = {}
    if not isinstance(raw, Mapping):
        return counters
    for key, value in raw.items():
        try:
            counters[str(key)] = int(value)
        except (TypeError, ValueError):
            continue
    return counters


def _meta_from_mapping(raw: Mapping[str, Any]) -> PackingCacheMeta:
    critical = raw.get("critical")
    critical_map = critical if isinstance(critical, Mapping) else {}
    remarks_raw = critical_map.get("remarks", raw.get("critical_remarks", []))
    remarks = (
        tuple(str(item) for item in remarks_raw)
        if isinstance(remarks_raw, (list, tuple))
        else ()
    )
    reports_raw = raw.get("report_paths")
    reports = (
        {str(key): str(value) for key, value in reports_raw.items()}
        if isinstance(reports_raw, Mapping)
        else {}
    )
    return PackingCacheMeta(
        version=str(raw.get("version", "")),
        created_at=str(raw.get("created_at", "")),
        root=str(raw.get("root", "")),
        fingerprint=str(raw.get("fingerprint", "")),
        counters=_int_counters(raw.get("counters")),
        critical_ok=bool(critical_map.get("ok", raw.get("critical_ok", False))),
        critical_remarks=remarks,
        report_paths=reports,
    )


def _issue_from_mapping(raw: object) -> PackingIssue | None:
    if not isinstance(raw, Mapping):
        return None
    severity_raw = str(raw.get("severity", PackingIssueSeverity.ERROR.value))
    severity = (
        PackingIssueSeverity.WARNING
        if severity_raw == PackingIssueSeverity.WARNING.value
        else PackingIssueSeverity.ERROR
    )
    excel_row: int | None
    try:
        excel_row = int(raw["excel_row"]) if raw.get("excel_row") is not None else None
    except (TypeError, ValueError):
        excel_row = None
    return PackingIssue(
        code=str(raw.get("code", "loader_issue")),
        message=str(raw.get("message", "Проблема исходной строки УЛ")),
        severity=severity,
        source_file=str(raw.get("source_file", "")),
        sheet=str(raw.get("sheet", "")),
        excel_row=excel_row,
        field=str(raw.get("field", "")),
        raw_value=raw.get("raw_value"),
        action=str(raw.get("action", "")),
    )


def _validate_rows(rows: list[object]) -> tuple[list[RowStd], list[PackingIssue]]:
    valid: list[RowStd] = []
    issues: list[PackingIssue] = []
    for index, row in enumerate(rows, start=1):
        if not isinstance(row, RowStd):
            issues.append(
                PackingIssue(
                    code="cache_row_type",
                    message=f"Элемент кэша #{index} имеет тип {type(row).__name__}, ожидался RowStd",
                    action="Повторите «Прочитать и проанализировать»",
                )
            )
            continue
        if row.row_type != RowType.position_row:
            source_file, sheet, excel_row = packing_row_source(row)
            issues.append(
                PackingIssue(
                    code="cache_non_position_row",
                    message="В data кэша попала строка, не являющаяся позицией",
                    source_file=source_file,
                    sheet=sheet,
                    excel_row=excel_row,
                    action="Повторите чтение УЛ и проверьте формат исходника",
                )
            )
            continue
        source_file, sheet, excel_row = packing_row_source(row)
        missing = [
            column
            for column in (DS_TITLE, DS_SYSTEM, CODE)
            if column not in row.el or not str(row.get_value(column) or "").strip()
        ]
        if missing:
            issues.append(
                PackingIssue(
                    code="packing_key_missing",
                    message=f"У позиции отсутствуют части ключа: {', '.join(missing)}",
                    source_file=source_file,
                    sheet=sheet,
                    excel_row=excel_row,
                    field=",".join(missing),
                    action="Исправьте титул, марку или код и повторите чтение УЛ",
                )
            )
        valid.append(row)
    return valid, issues


def _deduplicate_issues(issues: list[PackingIssue]) -> list[PackingIssue]:
    result: list[PackingIssue] = []
    seen: set[tuple[object, ...]] = set()
    for issue in issues:
        key = (
            issue.code,
            issue.message,
            issue.source_file,
            issue.sheet,
            issue.excel_row,
            issue.field,
            repr(issue.raw_value),
        )
        if key in seen:
            continue
        seen.add(key)
        result.append(issue)
    return result


def load_packing_dataset(
    cache_path: str | os.PathLike[str] | None = None,
    *,
    expected_version: str = PACKING_CACHE_VERSION,
    expected_root: str | None = None,
    expected_fingerprint: str | None = None,
) -> PackingDataset:
    """Load and validate a packing cache without silently suppressing failures.

    Args:
        cache_path: Pickle path; default is the network packing output folder.
        expected_version: Required cache schema version.
        expected_root: Optional source root required by a compatibility caller.
        expected_fingerprint: Optional exact source-tree fingerprint.

    Returns:
        A dataset in ``ok``, ``partial``, or ``unavailable`` state. Expected
        cache/data problems are returned as structured issues, not raised.
    """
    path = Path(cache_path) if cache_path is not None else default_packing_cache_path()
    if not path.is_file():
        return _unavailable_dataset(
            path,
            PackingIssue(
                code="cache_missing",
                message=f"Кэш упаковочных листов не найден: {path}",
                action="Откройте «ДС · Упаковочные листы» и выполните «Прочитать и проанализировать»",
            ),
        )

    try:
        with path.open("rb") as stream:
            payload = pickle.load(stream)
    except Exception as exc:
        return _unavailable_dataset(
            path,
            PackingIssue(
                code="cache_corrupt",
                message=f"Кэш упаковочных листов не читается ({type(exc).__name__}: {exc})",
                action="Повторите «Прочитать и проанализировать»",
            ),
        )

    if not isinstance(payload, Mapping):
        return _unavailable_dataset(
            path,
            PackingIssue(
                code="cache_schema",
                message="Корень кэша должен быть словарём с meta и data",
                action="Повторите «Прочитать и проанализировать»",
            ),
        )
    meta_raw = payload.get("meta")
    rows_raw = payload.get("data")
    if not isinstance(meta_raw, Mapping) or not isinstance(rows_raw, list):
        return _unavailable_dataset(
            path,
            PackingIssue(
                code="cache_schema",
                message="В кэше отсутствует корректная структура meta/data",
                action="Повторите «Прочитать и проанализировать»",
            ),
        )

    meta = _meta_from_mapping(meta_raw)
    if meta.version != expected_version:
        return PackingDataset(
            rows=[],
            meta=meta,
            quality=PackingQualityLevel.UNAVAILABLE,
            issues=[
                PackingIssue(
                    code="cache_version",
                    message=(
                        f"Версия кэша {meta.version or '<пусто>'}, "
                        f"требуется {expected_version}"
                    ),
                    action="Повторите «Прочитать и проанализировать»",
                )
            ],
            cache_path=str(path),
        )
    if not meta.created_at or not meta.root or not meta.fingerprint:
        return PackingDataset(
            rows=[],
            meta=meta,
            quality=PackingQualityLevel.UNAVAILABLE,
            issues=[
                PackingIssue(
                    code="cache_meta",
                    message="В meta кэша отсутствуют created_at, root или fingerprint",
                    action="Повторите «Прочитать и проанализировать»",
                )
            ],
            cache_path=str(path),
        )
    if expected_root is not None and os.path.normcase(
        os.path.normpath(meta.root)
    ) != os.path.normcase(os.path.normpath(str(expected_root))):
        return PackingDataset(
            rows=[],
            meta=meta,
            quality=PackingQualityLevel.UNAVAILABLE,
            issues=[
                PackingIssue(
                    code="cache_root",
                    message=f"Кэш создан для другой папки: {meta.root}",
                    action="Повторите чтение для выбранной папки УЛ",
                )
            ],
            cache_path=str(path),
        )
    if expected_fingerprint is not None and meta.fingerprint != expected_fingerprint:
        return PackingDataset(
            rows=[],
            meta=meta,
            quality=PackingQualityLevel.UNAVAILABLE,
            issues=[
                PackingIssue(
                    code="cache_stale",
                    message="Состав или время изменения исходных УЛ изменились после создания кэша",
                    action="Повторите «Прочитать и проанализировать»",
                )
            ],
            cache_path=str(path),
        )

    rows, issues = _validate_rows(rows_raw)
    required_meta_types: tuple[tuple[str, type], ...] = (
        ("counters", Mapping),
        ("critical", Mapping),
        ("report_paths", Mapping),
        ("loader_issues", list),
    )
    for key, expected_type in required_meta_types:
        if not isinstance(meta_raw.get(key), expected_type):
            issues.append(
                PackingIssue(
                    code="cache_meta_schema",
                    message=(
                        f"Поле meta.{key} имеет неверный тип "
                        f"{type(meta_raw.get(key)).__name__}"
                    ),
                    action="Повторите «Прочитать и проанализировать»",
                )
            )
    critical_raw = meta_raw.get("critical")
    if isinstance(critical_raw, Mapping) and not isinstance(
        critical_raw.get("ok"), bool
    ):
        issues.append(
            PackingIssue(
                code="cache_meta_schema",
                message="Поле meta.critical.ok должно быть boolean",
                action="Повторите «Прочитать и проанализировать»",
            )
        )
    loader_issues = meta_raw.get("loader_issues", [])
    if isinstance(loader_issues, list):
        issues.extend(
            issue
            for issue in (_issue_from_mapping(raw) for raw in loader_issues)
            if issue is not None
        )
    if not meta.critical_ok:
        uncovered_remarks = [
            remark
            for remark in meta.critical_remarks
            if not any(issue.message and issue.message in remark for issue in issues)
        ]
        if uncovered_remarks:
            issues.extend(
                PackingIssue(
                    code="loader_critical_remark",
                    message=remark,
                    action="Откройте tsd_packing_critical.txt и исправьте исходный файл",
                )
                for remark in uncovered_remarks
            )
        elif not loader_issues:
            issues.append(
                PackingIssue(
                    code="loader_critical",
                    message="Загрузка УЛ завершилась с критичными замечаниями",
                    action="Откройте tsd_packing_critical.txt и исправьте исходные файлы",
                )
            )
    expected_position_rows = meta.counters.get("position_rows")
    if expected_position_rows is None:
        issues.append(
            PackingIssue(
                code="cache_meta_schema",
                message="В meta.counters отсутствует position_rows",
                action="Повторите «Прочитать и проанализировать»",
            )
        )
    elif expected_position_rows != len(rows_raw):
        issues.append(
            PackingIssue(
                code="cache_row_count",
                message=(
                    f"meta.position_rows={expected_position_rows}, "
                    f"фактически data={len(rows_raw)}"
                ),
                action="Повторите «Прочитать и проанализировать»",
            )
        )

    issues = _deduplicate_issues(issues)
    quality = PackingQualityLevel.PARTIAL if issues else PackingQualityLevel.OK
    return PackingDataset(
        rows=rows,
        meta=meta,
        quality=quality,
        issues=issues,
        cache_path=str(path),
    )
