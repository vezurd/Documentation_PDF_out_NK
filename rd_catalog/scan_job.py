"""Qt-free catalog scan + MTO persist, for in-process tests and a child process."""

from __future__ import annotations

import traceback
from collections.abc import Callable, Iterable
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from rd_catalog.config import CatalogConfig
from rd_catalog.db import CatalogDatabase
from rd_catalog.models import ParsedFile, ScanProgress, ScanSummary, SourceKind
from rd_catalog.perf_log import perf_span
from rd_catalog.scan import ScanSubtree, normalize_scan_subtrees, scan_catalog


LogCallback = Callable[[str], None]
ProgressCallback = Callable[[ScanProgress], None]
CancelCallback = Callable[[], bool]


@dataclass(slots=True)
class ScanJobOutcome:
    """Compact result of one scan job after persistence."""

    run_id: int | None = None
    comparison_count: int = 0
    failure: str | None = None
    failure_traceback: str | None = None
    summary: ScanSummary | None = None
    touched_mto_keys: frozenset[tuple[str, str]] = field(default_factory=frozenset)
    mto_compare_skipped: bool = True
    compare_enqueue_scope: str = "none"


def _mto_document_keys(files: Iterable[ParsedFile]) -> set[tuple[str, str]]:
    keys: set[tuple[str, str]] = set()
    for file in files:
        key = file.document_key
        if isinstance(key, tuple) and len(key) == 2:
            keys.add((str(key[0]), str(key[1])))
    return keys


def _subtree_arg(subtree: ScanSubtree | None) -> str | tuple[str, ...] | None:
    folders = normalize_scan_subtrees(subtree)
    if not folders:
        return None
    if len(folders) == 1:
        return folders[0]
    return folders


def _subtree_log(subtree: ScanSubtree | None) -> str:
    return "; ".join(normalize_scan_subtrees(subtree))


def execute_scan(
    config: CatalogConfig,
    sources: Iterable[SourceKind],
    *,
    robot_subtree: str | Path | None = None,
    rd_subtree: ScanSubtree | None = None,
    sq_subtree: str | Path | None = None,
    is_cancelled: CancelCallback | None = None,
    progress: ProgressCallback | None = None,
    log: LogCallback | None = None,
) -> ScanJobOutcome:
    """Scan sources and persist the walk plus overlay only.

    MTO content compare is deferred to a separate job; this function never
    calls ``compare_and_store_mto``.

    Args:
        config: Resolved catalog configuration.
        sources: Source subset to scan.
        robot_subtree: Optional folder under ``robot_root`` for a scoped robot
            rescan.
        rd_subtree: Optional folder or folders under ``rd_root`` for a scoped
            RD rescan after an SQ→RD move or a kit point rescan. Pass the
            gate folder (``Для передачи``), not only the new ``NN_рев``
            package. A mixed-title kit may add foreign title/mark folders.
            Overlay is rebuilt from all present RD files.
        sq_subtree: Optional folder under ``sq_root`` for a scoped SQ rescan.
        is_cancelled: Cooperative cancellation predicate.
        progress: Optional filesystem-walk callback.
        log: Optional human-readable status lines.

    Returns:
        Persistence identifiers, deferred-compare metadata, and any captured
        failure. ``summary`` is set when the walk itself completed.
    """

    def emit(message: str) -> None:
        if log is not None:
            log(message)

    def cancelled() -> bool:
        return bool(is_cancelled and is_cancelled())

    outcome = ScanJobOutcome()
    source_tuple = tuple(sources)
    robot_text = str(robot_subtree) if robot_subtree is not None else None
    rd_arg = _subtree_arg(rd_subtree)
    rd_text = _subtree_log(rd_subtree)
    sq_text = str(sq_subtree) if sq_subtree is not None else None
    database = CatalogDatabase(config.db_path)
    with perf_span("scan.execute", sources=",".join(s.value for s in source_tuple), rd=rd_text or "", sq=sq_text or "", robot=robot_text or ""):
        try:
            labels = ", ".join(source.value.upper() for source in source_tuple)
            scoped = [
                f"{name} {path}"
                for name, path in (
                    ("RD", rd_text),
                    ("SQ", sq_text),
                    ("робот", robot_text),
                )
                if path
            ]
            if scoped:
                emit(f"Сканирование источников: {labels} · {'; '.join(scoped)}")
            else:
                emit(f"Сканирование источников: {labels}")
            with perf_span("scan.walk"):
                summary = scan_catalog(
                    config,
                    sources=source_tuple,
                    ignored_path_keys=database.get_ignored_path_keys(),
                    progress=progress,
                    is_cancelled=cancelled,
                    robot_subtree=robot_text,
                    rd_subtree=rd_arg,
                    sq_subtree=sq_text,
                )
            outcome.summary = summary
            with perf_span("scan.store"):
                outcome.run_id = database.store_scan(summary)
            outcome.touched_mto_keys = frozenset(_mto_document_keys(summary.files))
            emit(
                f"Скан #{outcome.run_id}: {summary.status.value}; "
                f"файлов: {len(summary.files)}, ошибок: {len(summary.errors)}"
            )

            comparison_sources = {
                source
                for source in source_tuple
                if source in (SourceKind.RD, SourceKind.ROBOT)
            }
            requested_results = [
                summary.sources.get(source) for source in comparison_sources
            ]
            can_compare = bool(
                comparison_sources
                and all(
                    result is not None
                    and not result.errors
                    and not result.cancelled
                    for result in requested_results
                )
                and not cancelled()
            )
            if can_compare:
                if rd_text or robot_text:
                    outcome.compare_enqueue_scope = "subtree"
                else:
                    outcome.compare_enqueue_scope = "full"
                emit(
                    "Сверка содержимого MTO отложена — "
                    "будет выполнена отдельной задачей."
                )
            elif comparison_sources:
                emit(
                    "MTO-сверка пропущена: обновляемый RD/robot источник "
                    "завершён не полностью или отменён."
                )
        except Exception as exc:
            outcome.failure = f"{type(exc).__name__}: {exc}"
            outcome.failure_traceback = traceback.format_exc()
    return outcome


def config_to_job_dict(
    config: CatalogConfig,
    sources: Iterable[SourceKind],
    *,
    robot_subtree: str | None = None,
    rd_subtree: ScanSubtree | None = None,
    sq_subtree: str | None = None,
    cancel_path: str | None = None,
) -> dict[str, Any]:
    """Serialize a scan request for ``rd_catalog.scan_cli``."""

    rd_payload = _subtree_arg(rd_subtree)
    if isinstance(rd_payload, tuple):
        rd_payload = list(rd_payload)
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
        "sources": [source.value for source in sources],
        "robot_subtree": robot_subtree,
        "rd_subtree": rd_payload,
        "sq_subtree": sq_subtree,
        "cancel_path": cancel_path,
    }


def config_from_job_dict(payload: dict[str, Any]) -> CatalogConfig:
    """Rebuild configuration from a ``scan_cli`` job file."""

    skip_raw = payload.get("skip_dirs") or ()
    return CatalogConfig(
        rd_root=Path(str(payload["rd_root"])),
        sq_root=Path(str(payload["sq_root"])),
        robot_root=Path(str(payload["robot_root"])),
        runtime_dir=Path(str(payload["runtime_dir"])),
        db_path=Path(str(payload["db_path"])),
        robot_flat_structure=bool(payload.get("robot_flat_structure", True)),
        skip_dirs=tuple(str(item) for item in skip_raw),
        google_kits_spreadsheet_id=str(
            payload.get("google_kits_spreadsheet_id") or ""
        ),
        google_kits_sheet_name=str(payload.get("google_kits_sheet_name") or ""),
        google_issuance_spreadsheet_id=str(
            payload.get("google_issuance_spreadsheet_id") or ""
        ),
        google_issuance_sheet_name=str(
            payload.get("google_issuance_sheet_name") or ""
        ),
    )
