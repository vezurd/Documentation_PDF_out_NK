"""Qt-free MTO content compare job, for in-process tests and a child process."""

from __future__ import annotations

import traceback
from collections.abc import Callable, Iterable
from dataclasses import dataclass
from typing import Any

from rd_catalog.config import CatalogConfig
from rd_catalog.db import CatalogDatabase
from rd_catalog.mto_compare_queue import build_mto_compare_plan
from rd_catalog.pipeline import official_rd_mto_overlay
from rd_catalog.scan_job import config_from_job_dict

LogCallback = Callable[[str], None]
CancelCallback = Callable[[], bool]


@dataclass(slots=True)
class MtoCompareJobOutcome:
    """Compact result of one deferred MTO compare job."""

    comparison_count: int = 0
    planned: int = 0
    skipped_cache_hits: int = 0
    cancelled: bool = False
    failure: str | None = None
    failure_traceback: str | None = None
    remaining_keys: tuple[tuple[str, str], ...] = ()


def _keys_from_payload(raw: object) -> list[tuple[str, str]] | None:
    if raw is None:
        return None
    return [(str(a), str(b)) for a, b in raw]  # type: ignore[misc]


def execute_mto_compare(
    config: CatalogConfig,
    *,
    document_keys: Iterable[tuple[str, str]] | None = None,
    priority_keys: Iterable[tuple[str, str]] = (),
    batch_size: int = 1,
    loader=None,
    is_cancelled: CancelCallback | None = None,
    progress: Callable[[int, int], None] | None = None,
    log: LogCallback | None = None,
) -> MtoCompareJobOutcome:
    """Compare persisted MTO pairs in batches with cooperative cancellation.

    Files must already have been stored by a prior scan. Planning uses the
    WAL-backed cache to skip unchanged pairs.

    Args:
        config: Resolved catalog configuration.
        document_keys: Optional MTO keys to restrict work; ``None`` runs full
            incremental compare.
        priority_keys: Keys from the latest scan walk; sorted first in the plan.
        batch_size: Number of plan pairs per ``compare_and_store_mto`` call.
        loader: Optional ``RowLoader`` test adapter.
        is_cancelled: Cooperative cancellation predicate.
        progress: Optional callback receiving completed and planned pair counts.
        log: Optional human-readable status lines.

    Returns:
        Comparison totals, cancellation state, and any captured failure.
    """

    def emit(message: str) -> None:
        if log is not None:
            log(message)

    def cancelled() -> bool:
        return bool(is_cancelled and is_cancelled())

    outcome = MtoCompareJobOutcome()
    batch_size = max(1, int(batch_size))
    scope_keys = None if document_keys is None else set(document_keys)
    database = CatalogDatabase(config.db_path)
    try:
        rd_overlay = official_rd_mto_overlay(database)
        robot_files = database.list_present_robot_mto_files()
        plan = build_mto_compare_plan(
            database,
            rd_overlay,
            robot_files,
            priority_keys=priority_keys,
            scope_keys=scope_keys,
        )
        outcome.planned = len(plan.pairs)
        outcome.skipped_cache_hits = plan.skipped_cache_hits
        emit(
            f"MTO-сверка: план {outcome.planned} пар, "
            f"пропущено по кешу: {outcome.skipped_cache_hits}"
        )
        if progress is not None:
            progress(0, outcome.planned)
        if not plan.pairs:
            return outcome

        processed_keys: set[tuple[str, str]] = set()
        plan_keys = plan.document_keys
        for batch_start in range(0, len(plan.pairs), batch_size):
            if cancelled():
                outcome.cancelled = True
                break
            batch = plan.pairs[batch_start : batch_start + batch_size]
            keys = {pair.key for pair in batch}
            results = database.compare_and_store_mto(
                rd_overlay,
                robot_files,
                loader=loader,
                is_cancelled=is_cancelled,
                document_keys=keys,
            )
            outcome.comparison_count += len(results)
            for result in results:
                processed_keys.add(result.key)
            if progress is not None:
                progress(len(processed_keys), outcome.planned)
            if cancelled():
                outcome.cancelled = True
                break

        outcome.remaining_keys = tuple(
            key for key in plan_keys if key not in processed_keys
        )
    except Exception as exc:
        outcome.failure = f"{type(exc).__name__}: {exc}"
        outcome.failure_traceback = traceback.format_exc()
    return outcome


def config_to_mto_job_dict(
    config: CatalogConfig,
    *,
    document_keys: list[tuple[str, str]] | None = None,
    priority_keys: Iterable[tuple[str, str]] = (),
    cancel_path: str | None = None,
    batch_size: int = 1,
) -> dict[str, Any]:
    """Serialize an MTO compare request for ``rd_catalog.mto_compare_cli``."""

    document_keys_payload = (
        None if document_keys is None else [list(key) for key in document_keys]
    )
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
        "document_keys": document_keys_payload,
        "priority_keys": [list(key) for key in priority_keys],
        "cancel_path": cancel_path,
        "batch_size": batch_size,
    }


__all__ = [
    "MtoCompareJobOutcome",
    "config_from_job_dict",
    "config_to_mto_job_dict",
    "execute_mto_compare",
    "_keys_from_payload",
]
