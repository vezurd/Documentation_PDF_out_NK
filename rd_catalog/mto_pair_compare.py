"""Qt-free symmetric MTO pair comparison for export source vs destination.

Compares the material-row multiset (``row_type == position_row``, seven
canonical fields) of two XLSX files. Pair identity in SQLite is canonical
``(min(file_id), max(file_id))`` so ``(A, B)`` and ``(B, A)`` share one
``mto_pair_comparison`` row and are never computed twice.

Persistence is an optimisation, not a precondition of the export gate.
Pairs whose both files are in ``file_entry`` are stored and reused.
Pairs that cannot be persisted (custom export folders) are still computed;
their verdicts are session-scoped. ``is_drained`` means every pair has been
attempted in this session.

This module does not import Qt. ``compare_pair`` is the only function that
reads workbooks; scans and pipeline rebuilds must not call it.
"""

from __future__ import annotations

from collections.abc import Callable, Mapping, Sequence
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from rd_catalog.db import (
    CatalogDatabase,
    canonical_mto_pair_ids,
    mto_file_stat_signature,
)
from rd_catalog.models import FileRecord, ParsedFile, MtoContentStatus, make_path_key
from rd_catalog.mto_diff import (
    MTO_COMPARE_ALGORITHM_VERSION,
    RowLoader,
    compare_canonical_rows,
    load_canonical_mto,
)

CancelCallback = Callable[[], bool]
ProgressCallback = Callable[[int, int], None]
_STATE_ADD = "add"
_STATE_NO_SOURCE = "no_source"


@dataclass(frozen=True, slots=True)
class MtoFilePair:
    """One export source file and the file already in the destination folder."""

    left_path: str
    right_path: str
    title: str
    mark: str


@dataclass(frozen=True, slots=True)
class PairComparePlan:
    """Split of export pairs into cache misses versus reusable verdicts."""

    pending: tuple[MtoFilePair, ...]
    cached: tuple[MtoFilePair, ...]


@dataclass(frozen=True, slots=True)
class PairComparisonResult:
    """Pure compare of two MTO workbooks; nothing is written to SQLite."""

    content_status: str
    left_fingerprint: str | None = None
    right_fingerprint: str | None = None
    left_rows: int = 0
    right_rows: int = 0
    added: int = 0
    removed: int = 0
    changed: int = 0
    error: str | None = None
    diff: dict[str, Any] | None = None

    @property
    def stats(self) -> dict[str, int]:
        """Return compact counters for SQLite and the GUI."""

        return {
            "left_rows": self.left_rows,
            "right_rows": self.right_rows,
            "added": self.added,
            "removed": self.removed,
            "changed": self.changed,
        }


@dataclass(frozen=True, slots=True)
class PairCompareOutcome:
    """Result of one :func:`execute_pair_compare` call.

    ``written`` counts rows upserted into ``mto_pair_comparison``.
    ``session_verdicts`` is keyed by canonically ordered
    ``(left_path_key, right_path_key)`` and holds every pair computed in
    this call: persisted successes, persisted read errors, and unpersisted
    custom-target pairs. The parent process must keep this map for the
    export gate; the child cannot share memory with the GUI.

    Attributes:
        written: Number of SQLite upserts (including ``not_compared``).
        session_verdicts: This-session computations by canonical path key.
    """

    written: int
    session_verdicts: Mapping[tuple[str, str], PairComparisonResult]


@dataclass(frozen=True, slots=True)
class PairPoolStatus:
    """Exact pool counter for the export gate and the remaining-pairs label.

    ``is_drained`` means every pair has been *attempted in this session*.
    Persistence is an optimisation, not a precondition: custom-target files
    are outside ``file_entry`` and are never stored, but a session verdict
    still counts as attempted.

    ``pending`` is pairs with neither a cache-valid material verdict
    (``content_equal`` / ``content_diff``) nor a session verdict.
    ``failed`` is this-session ``not_compared`` attempts (read errors).
    A stored error row without a session verdict is ``pending`` so it is
    retried; the gate is not held by a stale Excel lock.

    ``unpersisted`` is the number of pairs that cannot be stored because at
    least one file is outside the catalog (typically a custom export
    folder). Those pairs are recomputed every session. The UI can say so.

    Attributes:
        total: Unique comparable pairs in the current export pool.
        compared: Material ``content_equal`` / ``content_diff`` (cache or
            session).
        pending: Not yet attempted in this session.
        failed: This-session ``not_compared`` attempts.
        unpersisted: Pairs with no ``file_entry`` identity for both files.
    """

    total: int
    compared: int
    pending: int
    failed: int
    unpersisted: int = 0

    @property
    def is_drained(self) -> bool:
        """Return True when every pair has been attempted this session."""

        return self.pending == 0


def pairs_from_export_selections(
    selections: Sequence[Any],
) -> tuple[MtoFilePair, ...]:
    """Build the compare pool from export preview rows.

    One pair per selection that has both ``source_path`` and
    ``existing_target_path``. ``add`` and ``no_source`` rows are excluded.
    Duplicate path pairs (including swapped paths) are collapsed.

    Args:
        selections: Rows from :func:`rd_catalog.mto_export.resolve_export_selections`.

    Returns:
        Unique source/destination path pairs in selection order.
    """

    result: list[MtoFilePair] = []
    seen: set[tuple[str, str]] = set()
    for selection in selections:
        if selection.state in {_STATE_ADD, _STATE_NO_SOURCE}:
            continue
        left = str(selection.source_path or "").strip()
        right = str(selection.existing_target_path or "").strip()
        if not left or not right:
            continue
        key = _canonical_path_keys(left, right)
        if key in seen:
            continue
        seen.add(key)
        result.append(
            MtoFilePair(
                left_path=left,
                right_path=right,
                title=selection.title,
                mark=selection.mark,
            )
        )
    return tuple(result)


def build_pair_compare_plan(
    database: CatalogDatabase,
    *,
    pairs: Sequence[MtoFilePair],
    records: Sequence[FileRecord | ParsedFile],
) -> PairComparePlan:
    """Split pairs into cache misses and signature-valid cached verdicts.

    Pair identity is canonical ``file_entry.id`` order when both files are
    catalogued. A pair whose files are not both in ``file_entry`` is always
    pending (it cannot be cached). Stored ``not_compared`` rows are not
    cache-valid and return to ``pending`` so a transient Excel lock is retried.

    Args:
        database: Catalog SQLite.
        pairs: Export source/destination pairs.
        records: Catalog file rows used for ids and current signatures.

    Returns:
        Plan with cache-valid material verdicts omitted from ``pending``.
    """

    index = _index_records(records)
    pending: list[MtoFilePair] = []
    cached: list[MtoFilePair] = []
    seen_ids: set[tuple[int, int]] = set()
    for pair in _unique_pairs(pairs, index):
        resolved = _resolve_pair_files(database, pair, index)
        if resolved is None:
            pending.append(pair)
            continue
        left_id, right_id, left_file, right_file = resolved
        id_key = canonical_mto_pair_ids(left_id, right_id)
        if id_key in seen_ids:
            continue
        seen_ids.add(id_key)
        if database.mto_pair_comparison_valid(left=left_file, right=right_file):
            cached.append(pair)
        else:
            pending.append(pair)
    return PairComparePlan(pending=tuple(pending), cached=tuple(cached))


def compare_pair(
    left_path: str | Path,
    right_path: str | Path,
    *,
    loader: RowLoader | None = None,
) -> PairComparisonResult:
    """Read two XLSX files and compare material-row multisets.

    Never raises into the queue: any loader failure becomes ``not_compared``
    plus an ``error`` string. Fingerprint equality is ``content_equal``.
    Formatting, row order, and non-``position_row`` rows are ignored.

    Args:
        left_path: First workbook path.
        right_path: Second workbook path.
        loader: Optional test adapter matching ``mto_diff.RowLoader``.

    Returns:
        Verdict, fingerprints, and added/removed/changed counts.
    """

    try:
        left_doc = load_canonical_mto(left_path, loader=loader)
        right_doc = load_canonical_mto(right_path, loader=loader)
    except Exception as exc:
        return PairComparisonResult(
            content_status=MtoContentStatus.NOT_COMPARED.value,
            error=f"{type(exc).__name__}: {exc}",
        )
    if left_doc.fingerprint == right_doc.fingerprint:
        return PairComparisonResult(
            content_status=MtoContentStatus.EQUAL.value,
            left_fingerprint=left_doc.fingerprint,
            right_fingerprint=right_doc.fingerprint,
            left_rows=len(left_doc.rows),
            right_rows=len(right_doc.rows),
        )
    structured = compare_canonical_rows(left_doc.rows, right_doc.rows)
    counts = structured.stats(len(left_doc.rows), len(right_doc.rows))
    return PairComparisonResult(
        content_status=MtoContentStatus.DIFF.value,
        left_fingerprint=left_doc.fingerprint,
        right_fingerprint=right_doc.fingerprint,
        left_rows=counts["rd_rows"],
        right_rows=counts["robot_rows"],
        added=counts["added"],
        removed=counts["removed"],
        changed=counts["changed"],
        diff=structured.as_dict(),
    )


def execute_pair_compare(
    database: CatalogDatabase,
    *,
    plan: PairComparePlan,
    records: Sequence[FileRecord | ParsedFile],
    cancel: CancelCallback | None = None,
    progress: ProgressCallback | None = None,
    loader: RowLoader | None = None,
    algorithm_version: int | None = None,
) -> PairCompareOutcome:
    """Compare pending pairs; persist catalogued ones, keep the rest in session.

    Batch size is one pair so cancel is responsive and each catalogued
    verdict is durable as soon as it is computed. Cache-valid pairs in
    ``plan.cached`` are not recomputed.

    A pair that cannot be resolved to two ``file_entry`` ids is still
    compared. Its verdict is returned in ``session_verdicts`` and is not
    written. Persistence is an optimisation, not a precondition of the
    export gate.

    Catalogued read errors are upserted as ``not_compared`` (history / UI
    hint) and also returned in ``session_verdicts`` so this session can
    drain. They are not cache-valid.

    Args:
        database: Catalog SQLite.
        plan: Output of :func:`build_pair_compare_plan`.
        records: Catalog file rows for ids and signatures.
        cancel: Cooperative cancellation predicate, checked between pairs.
        progress: Optional ``(completed, total)`` callback over pending.
        loader: Optional test adapter; production leaves this ``None``.
        algorithm_version: Stored algorithm version; defaults to
            ``MTO_COMPARE_ALGORITHM_VERSION``.

    Returns:
        Upsert count and this-session verdicts keyed by canonical path keys.
    """

    version = (
        MTO_COMPARE_ALGORITHM_VERSION
        if algorithm_version is None
        else int(algorithm_version)
    )
    index = _index_records(records)
    total = len(plan.pending)
    written = 0
    attempted = 0
    session_verdicts: dict[tuple[str, str], PairComparisonResult] = {}
    seen_ids: set[tuple[int, int]] = set()
    seen_paths: set[tuple[str, str]] = set()
    for pair in plan.pending:
        if cancel is not None and cancel():
            break
        session_key = _canonical_path_keys(pair.left_path, pair.right_path)
        resolved = _resolve_pair_files(database, pair, index)
        if resolved is None:
            if session_key in seen_paths:
                continue
            seen_paths.add(session_key)
            compare_left, compare_right = _ordered_paths(
                pair.left_path, pair.right_path
            )
            result = compare_pair(compare_left, compare_right, loader=loader)
            session_verdicts[session_key] = result
            attempted += 1
            if progress is not None:
                progress(attempted, total)
            continue
        left_id, right_id, left_file, right_file = resolved
        if int(left_id) > int(right_id):
            left_id, right_id = right_id, left_id
            left_file, right_file = right_file, left_file
        id_key = (left_id, right_id)
        if id_key in seen_ids:
            continue
        seen_ids.add(id_key)
        if database.mto_pair_comparison_valid(left=left_file, right=right_file):
            attempted += 1
            if progress is not None:
                progress(attempted, total)
            continue
        result = compare_pair(left_file.path, right_file.path, loader=loader)
        database.upsert_mto_pair_comparison(
            left_id,
            right_id,
            algorithm_version=version,
            content_status=result.content_status,
            left_fingerprint=result.left_fingerprint,
            right_fingerprint=result.right_fingerprint,
            left_signature=mto_file_stat_signature(left_file),
            right_signature=mto_file_stat_signature(right_file),
            stats=result.stats,
            diff=result.diff or {},
            error=result.error,
        )
        written += 1
        session_verdicts[session_key] = result
        attempted += 1
        if progress is not None:
            progress(attempted, total)
    return PairCompareOutcome(written=written, session_verdicts=session_verdicts)


def pair_pool_status(
    database: CatalogDatabase,
    *,
    pairs: Sequence[MtoFilePair],
    records: Sequence[FileRecord | ParsedFile],
    session_verdicts: Mapping[tuple[str, str], str] | None = None,
) -> PairPoolStatus:
    """Return exact remaining / compared / failed counts for the export gate.

    Uses one :meth:`CatalogDatabase.list_mto_pair_comparisons` query for
    catalogued pairs. ``is_drained`` is True when every unique pair has been
    attempted *in this session*: a cache-valid material verdict, or a
    ``session_verdicts`` entry. Persistence is an optimisation, not a
    precondition.

    Unresolvable pairs (custom target folders) count in ``unpersisted``.
    Present in ``session_verdicts`` they are ``compared`` or ``failed``;
    absent they stay ``pending``. Stored ``not_compared`` rows without a
    session verdict are ``pending`` so they are retried.

    Args:
        database: Catalog SQLite.
        pairs: Current export compare pool.
        records: Catalog file rows for ids and current signatures.
        session_verdicts: Optional this-session statuses keyed by canonical
            ``(left_path_key, right_path_key)``. Values are
            ``content_equal``, ``content_diff``, or ``not_compared``.

    Returns:
        Pool counters. ``total`` is the number of unique pairs.
    """

    index = _index_records(records)
    unique = _unique_pairs(pairs, index)
    id_pairs: list[tuple[int, int]] = []
    resolved_files: list[
        tuple[MtoFilePair, int, int, FileRecord | ParsedFile, FileRecord | ParsedFile]
    ] = []
    unresolved_pairs: list[MtoFilePair] = []
    seen_ids: set[tuple[int, int]] = set()
    for pair in unique:
        resolved = _resolve_pair_files(database, pair, index)
        if resolved is None:
            unresolved_pairs.append(pair)
            continue
        left_id, right_id, left_file, right_file = resolved
        id_key = canonical_mto_pair_ids(left_id, right_id)
        if id_key in seen_ids:
            continue
        seen_ids.add(id_key)
        id_pairs.append(id_key)
        resolved_files.append((pair, left_id, right_id, left_file, right_file))
    rows = database.list_mto_pair_comparisons(id_pairs)
    compared = 0
    failed = 0
    pending = 0

    def consume_session(pair: MtoFilePair) -> str | None:
        return _session_status(session_verdicts, pair)

    for pair in unresolved_pairs:
        status = consume_session(pair)
        if status in _MATERIAL_STATUSES:
            compared += 1
        elif status == MtoContentStatus.NOT_COMPARED.value:
            failed += 1
        else:
            pending += 1
    for pair, left_id, right_id, left_file, right_file in resolved_files:
        status = consume_session(pair)
        if status in _MATERIAL_STATUSES:
            compared += 1
            continue
        if status == MtoContentStatus.NOT_COMPARED.value:
            failed += 1
            continue
        row = rows.get(canonical_mto_pair_ids(left_id, right_id))
        if (
            row is not None
            and _row_signatures_match(row, left_id, right_id, left_file, right_file)
            and str(row.get("content_status") or "") in _MATERIAL_STATUSES
        ):
            compared += 1
        else:
            pending += 1
    return PairPoolStatus(
        total=len(id_pairs) + len(unresolved_pairs),
        compared=compared,
        pending=pending,
        failed=failed,
        unpersisted=len(unresolved_pairs),
    )


_MATERIAL_STATUSES = frozenset(
    {MtoContentStatus.EQUAL.value, MtoContentStatus.DIFF.value}
)


def _canonical_path_keys(left_path: str, right_path: str) -> tuple[str, str]:
    first = make_path_key(left_path)
    second = make_path_key(right_path)
    if first <= second:
        return first, second
    return second, first


def _ordered_paths(left_path: str, right_path: str) -> tuple[str, str]:
    if make_path_key(left_path) <= make_path_key(right_path):
        return left_path, right_path
    return right_path, left_path


def _session_status(
    session_verdicts: Mapping[tuple[str, str], str] | None,
    pair: MtoFilePair,
) -> str | None:
    if not session_verdicts:
        return None
    key = _canonical_path_keys(pair.left_path, pair.right_path)
    status = session_verdicts.get(key)
    if status is not None:
        return str(status)
    swapped = (key[1], key[0])
    extra = session_verdicts.get(swapped)
    return str(extra) if extra is not None else None


def _index_records(
    records: Sequence[FileRecord | ParsedFile],
) -> dict[str, FileRecord | ParsedFile]:
    index: dict[str, FileRecord | ParsedFile] = {}
    for record in records:
        index[make_path_key(record.path)] = record
        index[str(record.path_key)] = record
    return index


def _unique_pairs(
    pairs: Sequence[MtoFilePair],
    index: Mapping[str, FileRecord | ParsedFile],
) -> tuple[MtoFilePair, ...]:
    unique: list[MtoFilePair] = []
    seen_ids: set[tuple[int, int]] = set()
    seen_paths: set[tuple[str, str]] = set()
    for pair in pairs:
        left = index.get(make_path_key(pair.left_path))
        right = index.get(make_path_key(pair.right_path))
        if isinstance(left, FileRecord) and isinstance(right, FileRecord):
            key = canonical_mto_pair_ids(left.id, right.id)
            if key in seen_ids:
                continue
            seen_ids.add(key)
            unique.append(pair)
            continue
        path_key = _canonical_path_keys(pair.left_path, pair.right_path)
        if path_key in seen_paths:
            continue
        seen_paths.add(path_key)
        unique.append(pair)
    return tuple(unique)


def _lookup_record(
    index: Mapping[str, FileRecord | ParsedFile],
    path: str,
) -> FileRecord | ParsedFile | None:
    return index.get(make_path_key(path))


def _resolve_pair_files(
    database: CatalogDatabase,
    pair: MtoFilePair,
    index: Mapping[str, FileRecord | ParsedFile],
) -> tuple[int, int, FileRecord | ParsedFile, FileRecord | ParsedFile] | None:
    left_file = _lookup_record(index, pair.left_path)
    right_file = _lookup_record(index, pair.right_path)
    if left_file is None or right_file is None:
        return None
    try:
        left_id = (
            int(left_file.id)
            if isinstance(left_file, FileRecord)
            else database._file_id_for_path_key(left_file.path_key)
        )
        right_id = (
            int(right_file.id)
            if isinstance(right_file, FileRecord)
            else database._file_id_for_path_key(right_file.path_key)
        )
    except KeyError:
        return None
    return left_id, right_id, left_file, right_file


def _row_signatures_match(
    row: Mapping[str, Any],
    left_id: int,
    right_id: int,
    left_file: FileRecord | ParsedFile,
    right_file: FileRecord | ParsedFile,
) -> bool:
    left_sig = mto_file_stat_signature(left_file)
    right_sig = mto_file_stat_signature(right_file)
    if int(left_id) > int(right_id):
        left_sig, right_sig = right_sig, left_sig
    return row.get("left_signature") == left_sig and row.get(
        "right_signature"
    ) == right_sig


__all__ = [
    "MtoFilePair",
    "PairCompareOutcome",
    "PairComparePlan",
    "PairComparisonResult",
    "PairPoolStatus",
    "build_pair_compare_plan",
    "compare_pair",
    "execute_pair_compare",
    "pair_pool_status",
    "pairs_from_export_selections",
]
