"""Read-only AN dump walker. Persists ``an_mto_file`` only (Qt-free)."""

from __future__ import annotations

import os
from collections.abc import Callable
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

from rd_catalog.an_index import AnMtoFile, parse_an_mto_file
from rd_catalog.config import CatalogConfig
from rd_catalog.db import CatalogDatabase
from rd_catalog.perf_log import perf_span
from rd_catalog.skip_dirs import is_skipped_dir_name


@dataclass(frozen=True, slots=True)
class AnScanProgress:
    """Filesystem-walk callback payload (no ``SourceKind``)."""

    path: str
    files_seen: int
    message: str = ""


@dataclass(frozen=True, slots=True)
class AnScanOutcome:
    """Result of one AN dump walk.

    ``files_seen`` counts non-lock ``.xlsx`` candidates. ``accepted`` is the
    number of names that ``parse_an_mto_file`` kept. ``skipped`` is
    ``files_seen - accepted`` (junk workbooks and stat failures).
    """

    files: tuple[AnMtoFile, ...]
    files_seen: int
    accepted: int
    skipped: int
    cancelled: bool
    failure: str | None
    scanned_at: str
    errors: tuple[str, ...] = ()


def _utc_now() -> str:
    return datetime.now(timezone.utc).isoformat()


def _is_cancelled(callback: Callable[[], bool] | None) -> bool:
    return bool(callback and callback())


def _is_excel_lock_name(name: str) -> bool:
    return Path(name).name.startswith("~$")


def _is_xlsx_name(name: str) -> bool:
    return Path(name).suffix.casefold() == ".xlsx"


def _an_root_disabled(root: Path) -> bool:
    """Return whether ``an_root`` is the empty/disabled CatalogConfig sentinel.

    ``Path('')`` becomes ``Path('.')`` on Windows; ``parts`` stays empty.
    """

    if not os.fspath(root).strip():
        return True
    return not root.parts


def _outcome(
    *,
    files: tuple[AnMtoFile, ...] = (),
    files_seen: int = 0,
    skipped: int = 0,
    cancelled: bool = False,
    failure: str | None = None,
    scanned_at: str,
    errors: tuple[str, ...] = (),
) -> AnScanOutcome:
    return AnScanOutcome(
        files=files,
        files_seen=files_seen,
        accepted=len(files),
        skipped=skipped,
        cancelled=cancelled,
        failure=failure,
        scanned_at=scanned_at,
        errors=errors,
    )


def an_job_from_config(
    config: CatalogConfig,
    cancel_path: str | Path | None = None,
) -> dict[str, Any]:
    """Serialize an AN dump request for ``rd_catalog.an_scan_cli``.

    Writes every ``CatalogConfig`` path field so the child does not need
    ``config_from_job_dict`` (that helper does not know ``an_root``).

    Args:
        config: Resolved catalog configuration.
        cancel_path: Optional cooperative-cancel flag file.

    Returns:
        JSON-serializable job mapping.
    """

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
        "an_root": "" if _an_root_disabled(config.an_root) else str(config.an_root),
        "cancel_path": str(cancel_path) if cancel_path is not None else None,
    }


def config_from_an_job(payload: dict[str, Any]) -> CatalogConfig:
    """Rebuild configuration from an ``an_scan_cli`` job file.

    Missing RD/SQ/robot paths become ``.`` placeholders. ``an_root`` may be
    empty (AN dump disabled).

    Args:
        payload: Decoded job object.

    Returns:
        Catalog configuration for this walk.

    Raises:
        KeyError: If ``db_path`` is missing.
        TypeError: If ``skip_dirs`` items are not strings.
    """

    skip_raw = payload.get("skip_dirs") or ()
    raw_an = str(payload.get("an_root") or "")
    return CatalogConfig(
        rd_root=Path(str(payload.get("rd_root") or ".")),
        sq_root=Path(str(payload.get("sq_root") or ".")),
        robot_root=Path(str(payload.get("robot_root") or ".")),
        runtime_dir=Path(str(payload.get("runtime_dir") or ".")),
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
        an_root=Path(raw_an) if raw_an.strip() else Path(""),
    )


def scan_an_dump(
    config: CatalogConfig,
    *,
    persist: bool = True,
    progress: Callable[[AnScanProgress], None] | None = None,
    is_cancelled: Callable[[], bool] | None = None,
    log: Callable[[str], None] | None = None,
) -> AnScanOutcome:
    """Walk ``config.an_root`` and optionally replace the AN snapshot.

    Read-only on the dump. Writes only the local SQLite ``an_mto_file``
    table. Does not touch ``file_entry``, overlay, or ``SourceKind``.

    An empty or missing root sets ``failure`` and does **not** call
    ``replace_an_snapshot``. A mid-walk cancel also leaves the last
    snapshot unchanged.

    Args:
        config: Resolved catalog configuration (``an_root``, ``skip_dirs``,
            ``db_path``).
        persist: When true, replace the snapshot after a complete walk.
        progress: Optional per-xlsx-candidate callback.
        is_cancelled: Cooperative cancellation predicate.
        log: Optional human-readable status lines.

    Returns:
        Walk counts, accepted files, and any failure string.
    """

    with perf_span("an.scan_dump"):
        scanned_at = _utc_now()

        def emit(message: str) -> None:
            if log is not None:
                log(message)

        def report(path: str, files_seen: int, message: str = "") -> None:
            if progress is not None:
                progress(
                    AnScanProgress(path=path, files_seen=files_seen, message=message)
                )

        root = config.an_root
        if _an_root_disabled(root):
            failure = "AN dump is disabled: an_root is empty"
            emit(failure)
            return _outcome(failure=failure, scanned_at=scanned_at)
        root_text = os.fspath(root)
        if not os.path.isdir(root_text):
            failure = f"AN root is not a directory: {root_text}"
            emit(failure)
            return _outcome(failure=failure, scanned_at=scanned_at)

        emit(f"Scanning AN dump: {root_text}")
        accepted: list[AnMtoFile] = []
        errors: list[str] = []
        files_seen = 0
        skipped = 0
        cancelled = False

        def on_walk_error(error: OSError) -> None:
            location = error.filename or root_text
            errors.append(f"{location}: {error}")

        for current_root, dirs, names in os.walk(root_text, onerror=on_walk_error):
            if _is_cancelled(is_cancelled):
                cancelled = True
                break
            dirs[:] = [
                name
                for name in dirs
                if not is_skipped_dir_name(name, config.skip_dirs)
            ]
            for name in names:
                if _is_cancelled(is_cancelled):
                    cancelled = True
                    break
                if _is_excel_lock_name(name) or not _is_xlsx_name(name):
                    continue
                files_seen += 1
                path = os.path.join(current_root, name)
                try:
                    stat = os.stat(path, follow_symlinks=False)
                except OSError as exc:
                    errors.append(f"{path}: {exc}")
                    skipped += 1
                    report(path, files_seen)
                    continue
                parsed = parse_an_mto_file(
                    path,
                    size=stat.st_size,
                    mtime_ns=stat.st_mtime_ns,
                )
                if parsed is None:
                    skipped += 1
                    report(path, files_seen)
                    continue
                accepted.append(parsed)
                report(path, files_seen)
            if cancelled:
                break

        files = tuple(accepted)
        if cancelled:
            emit("AN dump cancelled; snapshot left unchanged")
            return _outcome(
                files=files,
                files_seen=files_seen,
                skipped=skipped,
                cancelled=True,
                scanned_at=scanned_at,
                errors=tuple(errors),
            )

        if persist:
            try:
                database = CatalogDatabase(config.db_path)
                database.initialize()
                database.replace_an_snapshot(files, scanned_at=scanned_at)
            except Exception as exc:
                failure = f"{type(exc).__name__}: {exc}"
                emit(failure)
                return _outcome(
                    files=files,
                    files_seen=files_seen,
                    skipped=skipped,
                    failure=failure,
                    scanned_at=scanned_at,
                    errors=tuple(errors),
                )

        emit(
            f"AN dump finished: accepted {len(files)}, "
            f"seen {files_seen}, skipped {skipped}"
        )
        return _outcome(
            files=files,
            files_seen=files_seen,
            skipped=skipped,
            scanned_at=scanned_at,
            errors=tuple(errors),
        )
