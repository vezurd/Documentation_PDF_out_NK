"""Qt-free MTO pair-compare job, for in-process tests and a child process."""

from __future__ import annotations

import traceback
from collections.abc import Callable, Iterable, Mapping, Sequence
from dataclasses import dataclass, field
from typing import Any

from rd_catalog.config import CatalogConfig
from rd_catalog.db import CatalogDatabase
from rd_catalog.mto_pair_compare import (
    MtoFilePair,
    PairComparePlan,
    PairComparisonResult,
    build_pair_compare_plan,
    execute_pair_compare,
)
from rd_catalog.scan_job import config_from_job_dict

LogCallback = Callable[[str], None]
CancelCallback = Callable[[], bool]


@dataclass(slots=True)
class MtoPairCompareJobOutcome:
    """Compact result of one deferred export-pair compare job.

    ``session_verdicts`` is keyed by canonical
    ``(left_path_key, right_path_key)``. It includes persisted catalogued
    pairs and unpersisted custom-target pairs. The child emits this map on
    the JSON-line ``done`` payload so the GUI can drain the export gate
    without sharing memory.
    """

    comparison_count: int = 0
    planned: int = 0
    skipped_cache_hits: int = 0
    cancelled: bool = False
    failure: str | None = None
    failure_traceback: str | None = None
    remaining_pairs: tuple[MtoFilePair, ...] = ()
    session_verdicts: dict[tuple[str, str], PairComparisonResult] = field(
        default_factory=dict
    )


def pairs_from_payload(raw: object) -> list[MtoFilePair]:
    """Decode ``MtoFilePair`` rows from a job JSON list.

    Args:
        raw: Job ``pairs`` value.

    Returns:
        Decoded pairs; unknown shapes are skipped.
    """

    if not isinstance(raw, list):
        return []
    result: list[MtoFilePair] = []
    for item in raw:
        if not isinstance(item, dict):
            continue
        left = str(item.get("left_path") or "").strip()
        right = str(item.get("right_path") or "").strip()
        if not left or not right:
            continue
        result.append(
            MtoFilePair(
                left_path=left,
                right_path=right,
                title=str(item.get("title") or ""),
                mark=str(item.get("mark") or ""),
            )
        )
    return result


def session_verdicts_to_payload(
    verdicts: Mapping[tuple[str, str], PairComparisonResult],
) -> list[dict[str, Any]]:
    """Serialize session verdicts for the JSON-line ``done`` payload.

    Args:
        verdicts: This-session results keyed by canonical path keys.

    Returns:
        JSON-ready list of objects with ``left_path_key`` / ``right_path_key``.
    """

    payload: list[dict[str, Any]] = []
    for (left_key, right_key), result in verdicts.items():
        payload.append(
            {
                "left_path_key": left_key,
                "right_path_key": right_key,
                "content_status": result.content_status,
                "left_fingerprint": result.left_fingerprint,
                "right_fingerprint": result.right_fingerprint,
                "left_rows": result.left_rows,
                "right_rows": result.right_rows,
                "added": result.added,
                "removed": result.removed,
                "changed": result.changed,
                "error": result.error,
            }
        )
    return payload


def session_verdicts_from_payload(
    raw: object,
) -> dict[tuple[str, str], PairComparisonResult]:
    """Decode session verdicts from a job or JSON-line payload.

    Args:
        raw: ``done.session_verdicts`` list from the child process.

    Returns:
        Canonical ``(left_path_key, right_path_key)`` map. Unknown shapes
        are skipped.
    """

    if not isinstance(raw, list):
        return {}
    result: dict[tuple[str, str], PairComparisonResult] = {}
    for item in raw:
        if not isinstance(item, dict):
            continue
        left_key = str(item.get("left_path_key") or "")
        right_key = str(item.get("right_path_key") or "")
        if not left_key or not right_key:
            continue
        if left_key > right_key:
            left_key, right_key = right_key, left_key
        left_fp = item.get("left_fingerprint")
        right_fp = item.get("right_fingerprint")
        error = item.get("error")
        result[(left_key, right_key)] = PairComparisonResult(
            content_status=str(item.get("content_status") or ""),
            left_fingerprint=None if left_fp is None else str(left_fp),
            right_fingerprint=None if right_fp is None else str(right_fp),
            left_rows=int(item.get("left_rows") or 0),
            right_rows=int(item.get("right_rows") or 0),
            added=int(item.get("added") or 0),
            removed=int(item.get("removed") or 0),
            changed=int(item.get("changed") or 0),
            error=None if error is None else str(error),
        )
    return result


def pairs_to_payload(pairs: Sequence[MtoFilePair]) -> list[dict[str, str]]:
    """Serialize pairs for a job JSON file."""

    return [
        {
            "left_path": pair.left_path,
            "right_path": pair.right_path,
            "title": pair.title,
            "mark": pair.mark,
        }
        for pair in pairs
    ]


def execute_mto_pair_compare(
    config: CatalogConfig,
    *,
    pairs: Iterable[MtoFilePair],
    batch_size: int = 1,
    loader=None,
    is_cancelled: CancelCallback | None = None,
    progress: Callable[[int, int], None] | None = None,
    log: LogCallback | None = None,
) -> MtoPairCompareJobOutcome:
    """Compare export MTO pairs in batches with cooperative cancellation.

    Catalogued pairs use the WAL cache. Uncatalogued destination files are
    still compared; their verdicts are returned in ``session_verdicts``.

    Args:
        config: Resolved catalog configuration.
        pairs: Source/destination pairs to consider.
        batch_size: Number of pending pairs per cancel checkpoint. Production
            uses ``1``.
        loader: Optional ``RowLoader`` test adapter.
        is_cancelled: Cooperative cancellation predicate.
        progress: Optional callback receiving completed and planned counts.
        log: Optional human-readable status lines.

    Returns:
        Comparison totals, session verdicts, cancellation state, and any
        captured failure.
    """

    def emit(message: str) -> None:
        if log is not None:
            log(message)

    def cancelled() -> bool:
        return bool(is_cancelled and is_cancelled())

    outcome = MtoPairCompareJobOutcome()
    batch_size = max(1, int(batch_size))
    materialized = tuple(pairs)
    database = CatalogDatabase(config.db_path)
    try:
        records = database.list_files()
        plan = build_pair_compare_plan(
            database, pairs=materialized, records=records
        )
        outcome.planned = len(plan.pending)
        outcome.skipped_cache_hits = len(plan.cached)
        emit(
            f"MTO-пары: план {outcome.planned} пар, "
            f"пропущено по кешу: {outcome.skipped_cache_hits}"
        )
        if not plan.pending:
            return outcome

        processed = 0
        remaining = list(plan.pending)
        session: dict[tuple[str, str], PairComparisonResult] = {}
        for batch_start in range(0, len(plan.pending), batch_size):
            if cancelled():
                outcome.cancelled = True
                break
            batch = plan.pending[batch_start : batch_start + batch_size]
            batch_plan = PairComparePlan(pending=tuple(batch), cached=())
            batch_outcome = execute_pair_compare(
                database,
                plan=batch_plan,
                records=records,
                cancel=is_cancelled,
                loader=loader,
            )
            session.update(batch_outcome.session_verdicts)
            processed += len(batch_outcome.session_verdicts)
            outcome.comparison_count = processed
            remaining = list(plan.pending[batch_start + len(batch) :])
            if progress is not None:
                progress(processed, outcome.planned)
            if cancelled():
                outcome.cancelled = True
                break
        outcome.remaining_pairs = tuple(remaining)
        outcome.session_verdicts = session
    except Exception as exc:
        outcome.failure = f"{type(exc).__name__}: {exc}"
        outcome.failure_traceback = traceback.format_exc()
    return outcome


def config_to_pair_job_dict(
    config: CatalogConfig,
    *,
    pairs: Sequence[MtoFilePair] = (),
    cancel_path: str | None = None,
    batch_size: int = 1,
) -> dict[str, Any]:
    """Serialize an MTO pair-compare request for the child-process CLI."""

    return {
        "rd_root": str(config.rd_root),
        "sq_root": str(config.sq_root),
        "robot_root": str(config.robot_root),
        "runtime_dir": str(config.runtime_dir),
        "db_path": str(config.db_path),
        "robot_flat_structure": config.robot_flat_structure,
        "skip_dirs": list(config.skip_dirs),
        "google_kits_spreadsheet_id": config.google_kits_spreadsheet_id,
        "google_kits_sheet_name": config.google_kits_sheet_name,
        "google_issuance_spreadsheet_id": config.google_issuance_spreadsheet_id,
        "google_issuance_sheet_name": config.google_issuance_sheet_name,
        "pairs": pairs_to_payload(pairs),
        "cancel_path": cancel_path,
        "batch_size": batch_size,
    }


__all__ = [
    "MtoPairCompareJobOutcome",
    "config_from_job_dict",
    "config_to_pair_job_dict",
    "execute_mto_pair_compare",
    "pairs_from_payload",
    "pairs_to_payload",
    "session_verdicts_from_payload",
    "session_verdicts_to_payload",
]
