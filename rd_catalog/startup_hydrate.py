"""Off-GUI snapshot for the first Комплекты paint after window show.

The catalog window must stay responsive after ``show()``. SQLite ``list_*``,
Google snapshots, the kit matrix, worklist rows, and the Auto MTO index run
here on a ``QThread``. Qt widgets are filled on the GUI thread afterwards.
"""

from __future__ import annotations

import traceback
from dataclasses import dataclass, replace
from datetime import datetime, timezone
from typing import Any, Mapping, Sequence

from PySide6.QtCore import QThread, Signal

from rd_catalog.an_index import AnMtoFile
from rd_catalog.config import CatalogConfig
from rd_catalog.customer_pi_auto_mto import AutoMtoFile, list_auto_mto_files_by_kit
from rd_catalog.db import CatalogDatabase
from rd_catalog.google_kits import load_cached_google_kits
from rd_catalog.issuance_review import latest_effective_issuance_kits
from rd_catalog.kits import (
    GoogleKit,
    IssuanceKit,
    KitMatrixRow,
    build_kit_matrix,
    kit_identity_key,
    mto_content_equal_by_kit,
    annulled_folders_from_pipelines,
    working_folders_from_pipelines,
)
from rd_catalog.models import FileKind, FileRecord
from rd_catalog.mto_export import ExportPin, load_export_pins
from rd_catalog.parse import record_has_canonical_layout
from rd_catalog.perf_log import perf_span
from rd_catalog.transfer_review_compare import path_pair_labels_from_cache
from rd_catalog.pipeline import (
    PIPELINE_STATUS_ALGORITHM_VERSION,
    FolderTreeHint,
    KitPipelineRow,
    MtoWorklistRow,
    ingest_google_snapshot,
    list_folder_tree_hints,
    list_kit_pipelines,
    list_mto_worklist,
    official_detected_current_ids,
    pipeline_algorithm_needs_rebuild,
    rebuild_pipeline,
    records_in_contour,
)

_NOT_COMPARED_ERROR = "Сверка RD ↔ robot ещё не выполнена"


@dataclass(frozen=True, slots=True)
class GoogleMonitorState:
    """In-memory Google / issuance snapshot for the kits matrix."""

    kits: tuple[GoogleKit, ...] = ()
    issuance_kits: tuple[IssuanceKit, ...] = ()
    issuance_sends: tuple[IssuanceKit, ...] = ()
    source: str = ""
    fetched_at: str = ""
    warning: str | None = None
    loaded: bool = False
    from_sqlite: bool = False


@dataclass(frozen=True, slots=True)
class StartupSnapshot:
    """Read-only first-paint payload for ``CatalogWindow``."""

    records: list[FileRecord]
    detected_current_ids: set[int]
    mto_rows: list[dict[str, Any]]
    collision_rows: list[dict[str, Any]]
    google: GoogleMonitorState
    kit_rows: tuple[KitMatrixRow, ...]
    kit_pipelines: dict[tuple[str, str], KitPipelineRow]
    worklist_rows: tuple[MtoWorklistRow, ...]
    folder_hints: dict[tuple[str, str, str], FolderTreeHint]
    auto_mto_by_kit: dict[tuple[str, str], tuple[AutoMtoFile, ...]]
    an_by_kit: dict[tuple[str, str], tuple[AnMtoFile, ...]]
    mto_content_by_kit: dict[tuple[str, str], bool]
    export_pins: dict[tuple[str, str], ExportPin]
    last_scan: dict[str, Any] | None
    contour_records: tuple[FileRecord, ...] = ()
    official_current_ids: set[int] | None = None
    log_lines: tuple[str, ...] = ()


def overlay_current_ids(
    overlay_rows: Sequence[Mapping[str, Any]],
    records_by_id: Mapping[int, FileRecord],
    rd_root: str,
) -> set[int]:
    """Return overlay file ids that belong to the catalog contour."""

    detected: set[int] = set()
    for row in overlay_rows:
        file_id = int(row["file_id"])
        record = records_by_id.get(file_id)
        if record is None:
            continue
        if record_has_canonical_layout(record, rd_root):
            detected.add(file_id)
    return detected


def build_mto_display_rows(
    database: CatalogDatabase,
    records_by_id: Mapping[int, FileRecord],
    overlay_rows: Sequence[Mapping[str, Any]],
    *,
    pending_error: str,
) -> list[dict[str, Any]]:
    """Return MTO comparison rows plus synthetic not-compared placeholders."""

    rows = database.list_mto_comparisons()
    rows.extend(database.list_robot_extra_mto())
    compared_ids = {
        int(row["rd_file_id"])
        for row in rows
        if row.get("rd_file_id") is not None
    }
    for overlay in overlay_rows:
        file_id = int(overlay["file_id"])
        if (
            overlay.get("file_kind") != FileKind.MTO_XLSX.value
            or file_id in compared_ids
        ):
            continue
        record = records_by_id.get(file_id)
        if record is None:
            continue
        data = record.data
        rows.append(
            {
                "rd_file_id": file_id,
                "robot_file_id": None,
                "status": "blocked",
                "title_system": data.get("title_system"),
                "title": data.get("title"),
                "mark": data.get("mark"),
                "discipline_block": data.get("discipline_block"),
                "rd_path": record.path,
                "rd_path_key": record.path_key,
                "rd_name": data.get("name"),
                "rd_revision": data.get("revision"),
                "rd_appendix": data.get("appendix"),
                "rd_transfer_revision": data.get("transfer_revision"),
                "rd_transfer_appendix": data.get("transfer_appendix"),
                "rd_is_as_build": data.get("transfer_is_as_build"),
                "rd_review_state": record.review_state.value,
                "rd_present": record.present,
                "rd_mtime_ns": data.get("mtime_ns"),
                "robot_path": None,
                "robot_present": False,
                "diff": {
                    "content_status": "not_compared",
                    "issues": ["not_compared"],
                    "error": pending_error,
                },
                "stats": {},
            }
        )
    rows.sort(
        key=lambda row: (
            str(row.get("title_system") or ""),
            str(row.get("discipline_block") or ""),
        )
    )
    return rows


def load_google_for_monitor(
    database: CatalogDatabase,
    runtime_dir: str,
) -> GoogleMonitorState:
    """Prefer SQLite Google snapshots; fall back to JSON export caches."""

    try:
        kits = tuple(database.list_google_kits())
        sends = tuple(database.list_issuance_sends())
    except Exception:
        kits = ()
        sends = ()
    if kits or sends:
        info = database.google_load_info() or {}
        warning = info.get("warning")
        return GoogleMonitorState(
            kits=kits,
            issuance_kits=latest_effective_issuance_kits(database),
            issuance_sends=sends,
            source=str(info.get("source") or "sqlite"),
            fetched_at=str(info.get("loaded_at") or ""),
            warning=str(warning) if warning else None,
            loaded=True,
            from_sqlite=True,
        )
    cached = load_cached_google_kits(runtime_dir)
    if cached is None:
        return GoogleMonitorState()
    return GoogleMonitorState(
        kits=cached.kits,
        issuance_kits=cached.issuance_kits,
        issuance_sends=cached.issuance_sends,
        source=cached.source,
        fetched_at=cached.fetched_at,
        warning=cached.warning,
        loaded=True,
        from_sqlite=False,
    )


def load_startup_snapshot(
    database: CatalogDatabase,
    config: CatalogConfig,
) -> StartupSnapshot:
    """Read SQLite (and JSON cache if needed) for the first kits paint.

    Does not touch Qt. Rebuilds derived tables when ``kit_pipeline`` is
    empty while files or Google data exist, or when stored
    ``algorithm_version`` is older than ``PIPELINE_STATUS_ALGORITHM_VERSION``.
    """

    with perf_span("startup.load_snapshot"):
        logs: list[str] = []
        purged = database.purge_noncanonical_rd_files(
            config.rd_root, skip_if_clean=True
        )
        if purged:
            logs.append(f"РД: удалены неканонические строки: {purged}")
        records = database.list_files()
        records_by_id = {record.id: record for record in records}
        overlay_rows = database.current_overlay()
        detected = overlay_current_ids(
            overlay_rows, records_by_id, str(config.rd_root)
        )
        google = load_google_for_monitor(database, str(config.runtime_dir))
        pipelines: tuple[KitPipelineRow, ...] = ()
        try:
            pipelines = list_kit_pipelines(database)
        except Exception as exc:
            logs.append(f"Pipeline list: {type(exc).__name__}: {exc}")
        needs_empty_hatch = not pipelines and bool(
            records or google.kits or google.issuance_kits or google.issuance_sends
        )
        needs_stale_hatch = pipeline_algorithm_needs_rebuild(pipelines)
        if needs_empty_hatch or needs_stale_hatch:
            try:
                if needs_empty_hatch and not google.from_sqlite and (
                    google.kits or google.issuance_sends
                ):
                    ingest_google_snapshot(
                        database,
                        google.kits,
                        google.issuance_sends,
                        loaded_at=google.fetched_at
                        or datetime.now(timezone.utc).isoformat(),
                        source=google.source or "cache",
                        warning=google.warning,
                    )
                rebuild_pipeline(
                    database,
                    records=records,
                    detected_current_ids=detected,
                    rd_root=config.rd_root,
                )
                pipelines = list_kit_pipelines(database)
                google = replace(
                    google,
                    issuance_kits=latest_effective_issuance_kits(database),
                )
                if needs_stale_hatch:
                    logs.append(
                        "Pipeline: пересчёт статусов "
                        f"(algorithm {PIPELINE_STATUS_ALGORITHM_VERSION})"
                    )
            except Exception as exc:
                logs.append(f"Pipeline: {type(exc).__name__}: {exc}")
        mto_rows = build_mto_display_rows(
            database,
            records_by_id,
            overlay_rows,
            pending_error=_NOT_COMPARED_ERROR,
        )
        try:
            collision_rows = database.list_current_collisions()
        except Exception as exc:
            collision_rows = []
            logs.append(f"Коллизии: {type(exc).__name__}: {exc}")
        contour = tuple(records_in_contour(records, config.rd_root))
        try:
            worklist_rows = list_mto_worklist(
                database, records=contour, rd_root=None
            )
        except Exception as exc:
            worklist_rows = ()
            logs.append(f"MTO перечень: {type(exc).__name__}: {exc}")
        try:
            folder_hints = list_folder_tree_hints(database)
        except Exception as exc:
            folder_hints = {}
            logs.append(f"Подписи дерева: {type(exc).__name__}: {exc}")
        official_ids = official_detected_current_ids(contour, detected, pipelines)
        kit_rows = build_kit_matrix(
            google.kits,
            contour,
            official_ids,
            issuance_kits=google.issuance_kits,
            rd_root=None,
            mto_compare_by_pair=path_pair_labels_from_cache(config.runtime_dir),
            working_folders_by_kit=working_folders_from_pipelines(pipelines),
            annulled_folders_by_kit=annulled_folders_from_pipelines(pipelines),
        )
        auto_mto: dict[tuple[str, str], tuple[AutoMtoFile, ...]] = {}
        try:
            auto_mto = list_auto_mto_files_by_kit()
        except FileNotFoundError:
            auto_mto = {}
        except Exception as exc:
            logs.append(f"Авто МТО: {type(exc).__name__}: {exc}")
        an_by_kit: dict[tuple[str, str], tuple[AnMtoFile, ...]] = {}
        try:
            an_by_kit = database.list_an_files_by_kit()
        except Exception as exc:
            an_by_kit = {}
            logs.append(f"АН: {type(exc).__name__}: {exc}")
        try:
            last_scan = database.last_scan_info(successful_only=True)
        except Exception:
            last_scan = None
        pins = {
            kit_identity_key(pin.title, pin.mark): pin
            for pin in load_export_pins(config)
        }
        return StartupSnapshot(
            records=records,
            detected_current_ids=detected,
            mto_rows=mto_rows,
            collision_rows=collision_rows,
            google=google,
            kit_rows=kit_rows,
            kit_pipelines={
                kit_identity_key(item.title, item.mark): item for item in pipelines
            },
            worklist_rows=worklist_rows,
            folder_hints=folder_hints,
            auto_mto_by_kit=auto_mto,
            an_by_kit=an_by_kit,
            mto_content_by_kit=mto_content_equal_by_kit(mto_rows),
            export_pins=pins,
            last_scan=last_scan,
            contour_records=contour,
            official_current_ids=official_ids,
            log_lines=tuple(logs),
        )


class StartupHydrateThread(QThread):
    """Load the first-paint snapshot away from the GUI thread."""

    log = Signal(str)

    def __init__(self, config: CatalogConfig, parent=None) -> None:
        """Store an immutable hydrate request.

        Args:
            config: Resolved catalog configuration (SQLite path, RD root).
            parent: Optional Qt parent.
        """

        super().__init__(parent)
        self._config = config
        self.snapshot: StartupSnapshot | None = None
        self.failure: str | None = None

    def run(self) -> None:
        """Read SQLite / caches; never touch UNC sources or Qt widgets."""

        try:
            database = CatalogDatabase(self._config.db_path)
            self.snapshot = load_startup_snapshot(database, self._config)
            for line in self.snapshot.log_lines:
                self.log.emit(line)
        except Exception as exc:
            self.failure = f"{type(exc).__name__}: {exc}"
            self.log.emit(f"{self.failure}\n{traceback.format_exc()}")
