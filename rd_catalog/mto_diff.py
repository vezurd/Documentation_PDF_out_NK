"""Side-effect-free semantic comparison of current RD and robot MTO files."""

from __future__ import annotations

import hashlib
import json
import math
import re
from collections import Counter, defaultdict
from collections.abc import Callable, Iterable, Mapping, Sequence
from dataclasses import dataclass, replace
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Any

from rd_catalog.models import (
    FileKind,
    MtoContentStatus,
    MtoIssueKind,
    MtoReadinessStatus,
    OverlayResult,
    ParsedFile,
    ParseStatus,
)


MTO_COMPARE_ALGORITHM_VERSION = 1
CANONICAL_MTO_FIELDS = (
    "CODE",
    "UNITS",
    "VALUES",
    "TAGS",
    "NAME",
    "VENDOR",
    "TYPE_MARK",
)

RowLoader = Callable[[str | Path], Iterable[Any]]
_TAG_SPLIT_RE = re.compile(r"[,;\n\r]+")
_EMPTY_TEXT_VALUES = frozenset({"", "none", "null", "nan"})


@dataclass(frozen=True, slots=True)
class CanonicalMtoRow:
    """One material MTO row normalized for semantic comparison."""

    code: str
    units: str
    values: str
    tags: tuple[str, ...]
    name: str
    vendor: str
    type_mark: str

    def as_dict(self) -> dict[str, str | list[str]]:
        """Return the canonical public field mapping."""

        return {
            "CODE": self.code,
            "UNITS": self.units,
            "VALUES": self.values,
            "TAGS": list(self.tags),
            "NAME": self.name,
            "VENDOR": self.vendor,
            "TYPE_MARK": self.type_mark,
        }

    def sort_key(self) -> tuple[Any, ...]:
        """Return a stable order used only for serialization and display."""

        return (
            self.code,
            self.units,
            self.values,
            self.tags,
            self.name,
            self.vendor,
            self.type_mark,
        )


@dataclass(frozen=True, slots=True)
class CanonicalMtoDocument:
    """Canonical material rows and their versioned semantic fingerprint."""

    rows: tuple[CanonicalMtoRow, ...]
    fingerprint: str


@dataclass(frozen=True, slots=True)
class MtoFieldDifference:
    """One changed canonical field."""

    field: str
    rd_value: str | list[str]
    robot_value: str | list[str]


@dataclass(frozen=True, slots=True)
class MtoRowChange:
    """An unambiguously matched row with field-level differences."""

    code: str
    rd_row: CanonicalMtoRow
    robot_row: CanonicalMtoRow
    fields: tuple[MtoFieldDifference, ...]


@dataclass(frozen=True, slots=True)
class MtoStructuredDiff:
    """Multiset-preserving semantic difference."""

    added: tuple[CanonicalMtoRow, ...] = ()
    removed: tuple[CanonicalMtoRow, ...] = ()
    changed: tuple[MtoRowChange, ...] = ()

    @property
    def is_empty(self) -> bool:
        """Return whether both canonical multisets are equal."""

        return not (self.added or self.removed or self.changed)

    def stats(self, rd_rows: int, robot_rows: int) -> dict[str, int]:
        """Return compact counters suitable for SQLite and GUI cards."""

        return {
            "rd_rows": rd_rows,
            "robot_rows": robot_rows,
            "added": len(self.added),
            "removed": len(self.removed),
            "changed": len(self.changed),
        }

    def as_dict(self) -> dict[str, Any]:
        """Return JSON-safe structured details."""

        return {
            "added": [row.as_dict() for row in self.added],
            "removed": [row.as_dict() for row in self.removed],
            "changed": [
                {
                    "code": change.code,
                    "rd_row": change.rd_row.as_dict(),
                    "robot_row": change.robot_row.as_dict(),
                    "fields": [
                        {
                            "field": field.field,
                            "rd": field.rd_value,
                            "robot": field.robot_value,
                        }
                        for field in change.fields
                    ],
                }
                for change in self.changed
            ],
        }


@dataclass(frozen=True, slots=True)
class MtoPair:
    """One current RD MTO and all robot candidates sharing its key."""

    key: tuple[str, str]
    rd_file: ParsedFile
    robot_candidates: tuple[ParsedFile, ...]


@dataclass(frozen=True, slots=True)
class MtoPairingResult:
    """Pairing output including robot files with no current RD counterpart."""

    pairs: tuple[MtoPair, ...]
    robot_extras: tuple[ParsedFile, ...]


@dataclass(frozen=True, slots=True)
class MtoComparisonResult:
    """Complete comparison result consumed by persistence and the future GUI."""

    key: tuple[str, str]
    rd_file: ParsedFile
    robot_file: ParsedFile | None
    readiness: MtoReadinessStatus
    content_status: MtoContentStatus
    issues: tuple[MtoIssueKind, ...]
    revision_equal: bool
    mtime_suspicious: bool
    rd_fingerprint: str | None = None
    robot_fingerprint: str | None = None
    diff: MtoStructuredDiff = MtoStructuredDiff()
    rd_rows: int = 0
    robot_rows: int = 0
    error: str | None = None
    cache_hit: bool = False

    @property
    def stats(self) -> dict[str, int]:
        """Return row and difference counters."""

        return self.diff.stats(self.rd_rows, self.robot_rows)


def _normalize_text(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and math.isnan(value):
        return ""
    text = str(value).replace("_x000D_", "").replace("\xa0", " ")
    text = " ".join(text.split())
    return "" if text.casefold() in _EMPTY_TEXT_VALUES else text


def _normalize_number(value: Any) -> str:
    text = _normalize_text(value)
    if not text or text in {"-", "—"}:
        return ""
    numeric = text.replace(" ", "").replace(",", ".")
    try:
        decimal = Decimal(numeric)
    except InvalidOperation:
        return text
    if not decimal.is_finite():
        return ""
    if decimal == 0:
        return "0"
    normalized = decimal.normalize()
    return format(normalized, "f")


def _normalize_tags(value: Any) -> tuple[str, ...]:
    if value is None:
        return ()
    if isinstance(value, str):
        raw_values: Iterable[Any] = _TAG_SPLIT_RE.split(value)
    elif isinstance(value, Mapping):
        raw_values = value.values()
    elif isinstance(value, Iterable):
        raw_values = value
    else:
        raw_values = (value,)
    normalized = [_normalize_text(item) for item in raw_values]
    return tuple(sorted(item for item in normalized if item))


def _row_value(row: Any, field: str) -> Any:
    field_lower = field.casefold()
    if isinstance(row, Mapping):
        for key in (field, field_lower):
            if key in row:
                return row[key]
        return None
    elements = getattr(row, "el", None)
    if isinstance(elements, Mapping):
        element = elements.get(field_lower)
        if element is None:
            element = elements.get(field)
        return getattr(element, "value", element)
    for name in (field_lower, field):
        if hasattr(row, name):
            return getattr(row, name)
    return None


def _is_material_row(row: Any) -> bool:
    row_type = (
        row.get("row_type") or row.get("ROW_TYPE")
        if isinstance(row, Mapping)
        else getattr(row, "row_type", None)
    )
    if row_type is not None:
        row_type_value = getattr(row_type, "value", row_type)
        return str(row_type_value) == "position_row"
    return any(_normalize_text(_row_value(row, field)) for field in CANONICAL_MTO_FIELDS)


def canonicalize_mto_rows(rows: Iterable[Any]) -> tuple[CanonicalMtoRow, ...]:
    """Normalize every material row without aggregating duplicate codes.

    Args:
        rows: Legacy ``RowStd`` objects or field mappings.

    Returns:
        Canonical rows in input order. Multiset comparison does not depend on
        this order.
    """

    result: list[CanonicalMtoRow] = []
    for row in rows:
        if not _is_material_row(row):
            continue
        result.append(
            CanonicalMtoRow(
                code=_normalize_text(_row_value(row, "CODE")),
                units=_normalize_text(_row_value(row, "UNITS")),
                values=_normalize_number(_row_value(row, "VALUES")),
                tags=_normalize_tags(_row_value(row, "TAGS")),
                name=_normalize_text(_row_value(row, "NAME")),
                vendor=_normalize_text(_row_value(row, "VENDOR")),
                type_mark=_normalize_text(_row_value(row, "TYPE_MARK")),
            )
        )
    return tuple(result)


def semantic_fingerprint(rows: Iterable[CanonicalMtoRow]) -> str:
    """Hash a canonical row multiset together with the algorithm version."""

    payload = {
        "algorithm_version": MTO_COMPARE_ALGORITHM_VERSION,
        "fields": list(CANONICAL_MTO_FIELDS),
        "rows": [
            row.as_dict()
            for row in sorted(tuple(rows), key=CanonicalMtoRow.sort_key)
        ],
    }
    encoded = json.dumps(
        payload,
        ensure_ascii=False,
        separators=(",", ":"),
        sort_keys=True,
    ).encode("utf-8")
    return hashlib.sha256(encoded).hexdigest()


def load_mto_rows(path: str | Path) -> list[Any]:
    """Load one MTO with the project's active loader.

    Heavy legacy modules are deliberately imported only when this function is
    called. The adapter supplies a direct file path, does not use RFQ caches,
    and never invokes report-writing comparison code.

    Args:
        path: Existing MTO XLSX path.

    Returns:
        Legacy ``RowStd`` rows.

    Raises:
        RuntimeError: If the loader returns no row collection.
        Exception: Any workbook or legacy normalization error.
    """

    from base import t_comm_initial_classes
    from base.base_classes import TableComments
    from base.base_mto import get_mto_std_from_file

    comments = TableComments(
        file_full_path=str(path),
        dir_path="-1",
        tabel_class=t_comm_initial_classes.MTO,
    )
    rows = get_mto_std_from_file(comments, dbg=0)
    if rows is None:
        raise RuntimeError(f"MTO loader returned no rows: {path}")
    return list(rows)


def load_canonical_mto(
    path: str | Path,
    *,
    loader: RowLoader | None = None,
) -> CanonicalMtoDocument:
    """Load and fingerprint one MTO without writing reports or source files."""

    loaded = (loader or load_mto_rows)(path)
    rows = canonicalize_mto_rows(loaded)
    return CanonicalMtoDocument(rows=rows, fingerprint=semantic_fingerprint(rows))


def _field_differences(
    rd_row: CanonicalMtoRow,
    robot_row: CanonicalMtoRow,
) -> tuple[MtoFieldDifference, ...]:
    rd_values = rd_row.as_dict()
    robot_values = robot_row.as_dict()
    return tuple(
        MtoFieldDifference(field, rd_values[field], robot_values[field])
        for field in CANONICAL_MTO_FIELDS
        if rd_values[field] != robot_values[field]
    )


def compare_canonical_rows(
    rd_rows: Sequence[CanonicalMtoRow],
    robot_rows: Sequence[CanonicalMtoRow],
) -> MtoStructuredDiff:
    """Compare canonical rows as multisets and build unambiguous row changes."""

    rd_counter = Counter(rd_rows)
    robot_counter = Counter(robot_rows)
    common = rd_counter & robot_counter
    rd_remaining = list((rd_counter - common).elements())
    robot_remaining = list((robot_counter - common).elements())
    consumed_rd: set[int] = set()
    consumed_robot: set[int] = set()
    changes: list[MtoRowChange] = []

    rd_by_code: dict[str, list[int]] = defaultdict(list)
    robot_by_code: dict[str, list[int]] = defaultdict(list)
    for index, row in enumerate(rd_remaining):
        if row.code:
            rd_by_code[row.code].append(index)
    for index, row in enumerate(robot_remaining):
        if row.code:
            robot_by_code[row.code].append(index)

    for code in sorted(rd_by_code.keys() & robot_by_code.keys()):
        rd_indexes = rd_by_code[code]
        robot_indexes = robot_by_code[code]
        if len(rd_indexes) == len(robot_indexes) == 1:
            rd_index = rd_indexes[0]
            robot_index = robot_indexes[0]
            rd_row = rd_remaining[rd_index]
            robot_row = robot_remaining[robot_index]
            fields = _field_differences(rd_row, robot_row)
            if fields:
                changes.append(MtoRowChange(code, rd_row, robot_row, fields))
                consumed_rd.add(rd_index)
                consumed_robot.add(robot_index)

    # A CODE-only change has no stable CODE for pairing. Preserve it as changed
    # only when both directions have exactly one otherwise-identical candidate.
    rd_candidates: dict[int, list[int]] = defaultdict(list)
    robot_candidates: dict[int, list[int]] = defaultdict(list)
    for rd_index, rd_row in enumerate(rd_remaining):
        if rd_index in consumed_rd:
            continue
        for robot_index, robot_row in enumerate(robot_remaining):
            if robot_index in consumed_robot:
                continue
            fields = _field_differences(rd_row, robot_row)
            if len(fields) == 1 and fields[0].field == "CODE":
                rd_candidates[rd_index].append(robot_index)
                robot_candidates[robot_index].append(rd_index)
    for rd_index in sorted(rd_candidates):
        candidates = rd_candidates[rd_index]
        if len(candidates) != 1:
            continue
        robot_index = candidates[0]
        if len(robot_candidates[robot_index]) != 1:
            continue
        rd_row = rd_remaining[rd_index]
        robot_row = robot_remaining[robot_index]
        changes.append(
            MtoRowChange(
                robot_row.code or rd_row.code,
                rd_row,
                robot_row,
                _field_differences(rd_row, robot_row),
            )
        )
        consumed_rd.add(rd_index)
        consumed_robot.add(robot_index)

    removed = tuple(
        sorted(
            (
                row
                for index, row in enumerate(rd_remaining)
                if index not in consumed_rd
            ),
            key=CanonicalMtoRow.sort_key,
        )
    )
    added = tuple(
        sorted(
            (
                row
                for index, row in enumerate(robot_remaining)
                if index not in consumed_robot
            ),
            key=CanonicalMtoRow.sort_key,
        )
    )
    changes.sort(key=lambda item: (item.code, item.rd_row.sort_key()))
    return MtoStructuredDiff(added=added, removed=removed, changed=tuple(changes))


def _mto_key(file: ParsedFile) -> tuple[str, str] | None:
    if not file.title_system or not file.discipline_block:
        return None
    return (file.title_system.casefold(), file.discipline_block.casefold())


def pair_current_mto(
    rd_overlay: OverlayResult,
    robot_files: Iterable[ParsedFile],
) -> MtoPairingResult:
    """Pair current RD MTO files with robot candidates by the prescribed key."""

    if rd_overlay.file_kind is not FileKind.MTO_XLSX:
        raise ValueError("RD overlay must contain MTO_XLSX files")
    robot_by_key: dict[tuple[str, str], list[ParsedFile]] = defaultdict(list)
    all_robot_files = tuple(
        file for file in robot_files if file.file_kind is FileKind.MTO_XLSX
    )
    for file in all_robot_files:
        key = _mto_key(file)
        if key is not None:
            robot_by_key[key].append(file)

    pairs: list[MtoPair] = []
    used_robot_paths: set[str] = set()
    for rd_file in sorted(rd_overlay.current.values(), key=lambda file: file.path_key):
        key = _mto_key(rd_file)
        if key is None:
            key = ("", "")
        candidates = tuple(
            sorted(robot_by_key.get(key, ()), key=lambda file: file.path_key)
        )
        used_robot_paths.update(file.path_key for file in candidates)
        pairs.append(MtoPair(key, rd_file, candidates))
    extras = tuple(
        sorted(
            (
                file
                for file in all_robot_files
                if file.path_key not in used_robot_paths
            ),
            key=lambda file: file.path_key,
        )
    )
    return MtoPairingResult(tuple(pairs), extras)


def _revision_equal(rd_file: ParsedFile, robot_file: ParsedFile) -> bool:
    if not rd_file.revision or not robot_file.revision:
        return False
    return (
        rd_file.revision.casefold(),
        (rd_file.appendix or "").casefold(),
    ) == (
        robot_file.revision.casefold(),
        (robot_file.appendix or "").casefold(),
    )


def blocked_mto_result(
    pair: MtoPair,
    issue: MtoIssueKind,
    *,
    robot_file: ParsedFile | None = None,
    error: str | None = None,
) -> MtoComparisonResult:
    """Build a typed BLOCKED result for an uncheckable pair."""

    return MtoComparisonResult(
        key=pair.key,
        rd_file=pair.rd_file,
        robot_file=robot_file,
        readiness=MtoReadinessStatus.BLOCKED,
        content_status=MtoContentStatus.NOT_COMPARED,
        issues=(issue,),
        revision_equal=False,
        mtime_suspicious=False,
        error=error,
    )


def compare_loaded_pair(
    pair: MtoPair,
    rd_document: CanonicalMtoDocument,
    robot_document: CanonicalMtoDocument,
) -> MtoComparisonResult:
    """Compare already loaded canonical documents and evaluate readiness."""

    robot_file = pair.robot_candidates[0]
    try:
        diff = compare_canonical_rows(rd_document.rows, robot_document.rows)
    except Exception as exc:
        return MtoComparisonResult(
            key=pair.key,
            rd_file=pair.rd_file,
            robot_file=robot_file,
            readiness=MtoReadinessStatus.BLOCKED,
            content_status=MtoContentStatus.COMPARE_ERROR,
            issues=(MtoIssueKind.COMPARE_ERROR,),
            revision_equal=False,
            mtime_suspicious=False,
            rd_fingerprint=rd_document.fingerprint,
            robot_fingerprint=robot_document.fingerprint,
            error=str(exc),
        )

    revision_equal = _revision_equal(pair.rd_file, robot_file)
    mtime_suspicious = robot_file.mtime_ns < pair.rd_file.mtime_ns
    issues: list[MtoIssueKind] = []
    if not diff.is_empty:
        issues.append(MtoIssueKind.CONTENT_DIFF)
    if not revision_equal:
        issues.append(MtoIssueKind.REVISION_MISMATCH)
    if mtime_suspicious:
        issues.append(MtoIssueKind.MTIME_SUSPICION)
    return MtoComparisonResult(
        key=pair.key,
        rd_file=pair.rd_file,
        robot_file=robot_file,
        readiness=(
            MtoReadinessStatus.READY
            if not issues
            else MtoReadinessStatus.WARNING
        ),
        content_status=(
            MtoContentStatus.EQUAL if diff.is_empty else MtoContentStatus.DIFF
        ),
        issues=tuple(issues),
        revision_equal=revision_equal,
        mtime_suspicious=mtime_suspicious,
        rd_fingerprint=rd_document.fingerprint,
        robot_fingerprint=robot_document.fingerprint,
        diff=diff,
        rd_rows=len(rd_document.rows),
        robot_rows=len(robot_document.rows),
    )


def compare_mto_pair(
    pair: MtoPair,
    *,
    loader: RowLoader | None = None,
) -> MtoComparisonResult:
    """Load and compare one pair, converting all expected failures to BLOCKED."""

    if pair.rd_file.parse_status is not ParseStatus.PARSED or _mto_key(pair.rd_file) is None:
        return blocked_mto_result(
            pair,
            MtoIssueKind.RD_UNPARSED,
            error=pair.rd_file.parse_error or "Current RD MTO is unparsed",
        )
    if not pair.robot_candidates:
        return blocked_mto_result(pair, MtoIssueKind.ROBOT_MISSING)
    if len(pair.robot_candidates) > 1:
        return blocked_mto_result(
            pair,
            MtoIssueKind.ROBOT_DUPLICATE,
            error=f"{len(pair.robot_candidates)} robot candidates",
        )
    robot_file = pair.robot_candidates[0]
    if robot_file.parse_status is not ParseStatus.PARSED or _mto_key(robot_file) is None:
        return blocked_mto_result(
            pair,
            MtoIssueKind.ROBOT_UNPARSED,
            robot_file=robot_file,
            error=robot_file.parse_error or "Robot MTO is unparsed",
        )

    try:
        rd_document = load_canonical_mto(pair.rd_file.path, loader=loader)
        robot_document = load_canonical_mto(robot_file.path, loader=loader)
    except Exception as exc:
        return blocked_mto_result(
            pair,
            MtoIssueKind.LOAD_ERROR,
            robot_file=robot_file,
            error=str(exc),
        )
    return compare_loaded_pair(pair, rd_document, robot_document)


def compare_current_mto(
    rd_overlay: OverlayResult,
    robot_files: Iterable[ParsedFile],
    *,
    loader: RowLoader | None = None,
) -> tuple[MtoComparisonResult, ...]:
    """Pair and compare every current RD MTO."""

    pairing = pair_current_mto(rd_overlay, robot_files)
    return tuple(compare_mto_pair(pair, loader=loader) for pair in pairing.pairs)


def comparison_to_dict(result: MtoComparisonResult) -> dict[str, Any]:
    """Serialize a result for the ``mto_comparison`` cache."""

    return {
        "key": list(result.key),
        "readiness": result.readiness.value,
        "content_status": result.content_status.value,
        "issues": [issue.value for issue in result.issues],
        "revision_equal": result.revision_equal,
        "mtime_suspicious": result.mtime_suspicious,
        "rd_fingerprint": result.rd_fingerprint,
        "robot_fingerprint": result.robot_fingerprint,
        "diff": result.diff.as_dict(),
        "rd_rows": result.rd_rows,
        "robot_rows": result.robot_rows,
        "error": result.error,
    }


def _row_from_dict(data: Mapping[str, Any]) -> CanonicalMtoRow:
    return CanonicalMtoRow(
        code=str(data.get("CODE") or ""),
        units=str(data.get("UNITS") or ""),
        values=str(data.get("VALUES") or ""),
        tags=tuple(str(value) for value in data.get("TAGS") or ()),
        name=str(data.get("NAME") or ""),
        vendor=str(data.get("VENDOR") or ""),
        type_mark=str(data.get("TYPE_MARK") or ""),
    )


def comparison_from_dict(
    data: Mapping[str, Any],
    *,
    rd_file: ParsedFile,
    robot_file: ParsedFile | None,
    cache_hit: bool = True,
) -> MtoComparisonResult:
    """Rehydrate a cached result against current parsed file records."""

    diff_data = data.get("diff") or {}
    changes = tuple(
        MtoRowChange(
            code=str(item.get("code") or ""),
            rd_row=_row_from_dict(item.get("rd_row") or {}),
            robot_row=_row_from_dict(item.get("robot_row") or {}),
            fields=tuple(
                MtoFieldDifference(
                    field=str(field.get("field") or ""),
                    rd_value=field.get("rd", ""),
                    robot_value=field.get("robot", ""),
                )
                for field in item.get("fields") or ()
            ),
        )
        for item in diff_data.get("changed") or ()
    )
    diff = MtoStructuredDiff(
        added=tuple(_row_from_dict(item) for item in diff_data.get("added") or ()),
        removed=tuple(
            _row_from_dict(item) for item in diff_data.get("removed") or ()
        ),
        changed=changes,
    )
    key_data = data.get("key") or ("", "")
    return MtoComparisonResult(
        key=(str(key_data[0]), str(key_data[1])),
        rd_file=rd_file,
        robot_file=robot_file,
        readiness=MtoReadinessStatus(str(data["readiness"])),
        content_status=MtoContentStatus(str(data["content_status"])),
        issues=tuple(MtoIssueKind(str(issue)) for issue in data.get("issues") or ()),
        revision_equal=bool(data.get("revision_equal")),
        mtime_suspicious=bool(data.get("mtime_suspicious")),
        rd_fingerprint=data.get("rd_fingerprint"),
        robot_fingerprint=data.get("robot_fingerprint"),
        diff=diff,
        rd_rows=int(data.get("rd_rows") or 0),
        robot_rows=int(data.get("robot_rows") or 0),
        error=data.get("error"),
        cache_hit=cache_hit,
    )


def mark_cache_hit(result: MtoComparisonResult) -> MtoComparisonResult:
    """Return an otherwise identical result marked as cache-backed."""

    return replace(result, cache_hit=True)
