"""Schema-versioned SQLite persistence for the RD catalog."""

from __future__ import annotations

import hashlib
import json
import re
import sqlite3
from collections.abc import Callable, Collection, Iterable, Iterator, Mapping, Sequence
from contextlib import contextmanager
from dataclasses import dataclass
from datetime import date, datetime, timezone
from pathlib import Path
from typing import TYPE_CHECKING, Any

from rd_catalog.an_index import AnMtoFile
from rd_catalog.kits import (
    GoogleKit,
    IssuanceKit,
    KitEvent,
    format_revision,
    kit_identity_key,
    parse_history_line,
    parse_sheet_revision,
)
from rd_catalog.models import (
    FileKind,
    FileRecord,
    MtoContentStatus,
    MtoIssueKind,
    OverlayCollision,
    OverlayResult,
    ParseStatus,
    ParsedFile,
    ReviewState,
    ScanRunStatus,
    ScanSummary,
    SourceKind,
    SourceScanResult,
    TransferMetadata,
    make_path_key,
)
from rd_catalog.overlay import OVERLAY_ALGORITHM_VERSION, build_rd_overlays, revision_rank
from rd_catalog.parse import (
    has_canonical_rd_issued_path,
    matches_agcc_filename,
    parse_transfer_folder,
    path_is_as_build,
    transfer_name_is_void,
)
from rd_catalog.path_actions import path_is_under
from rd_catalog.perf_log import perf_span
from rd_catalog.skip_dirs import path_has_skipped_dir

if TYPE_CHECKING:
    from rd_catalog.mto_diff import MtoComparisonResult, RowLoader


SCHEMA_VERSION = 13
_FILE_ID_CHUNK = 400
_OVERRIDE_DATE_RE = re.compile(r"^(\d{1,2})\.(\d{1,2})\.(\d{4})$")
FILE_MTIME_OVERRIDE_REASONS = frozenset(
    {"code_a", "code_b", "code_c", "manual", "folder_mean"}
)
_FILE_MTIME_OVERRIDE_KINDS = frozenset(
    {
        FileKind.PDF.value,
        FileKind.MTO_XLSX.value,
        FileKind.SOURCE_EDITABLE.value,
    }
)


def _json_str_tuple(raw: object) -> tuple[str, ...]:
    try:
        parsed = json.loads(str(raw or "[]"))
    except (TypeError, ValueError):
        return ()
    if not isinstance(parsed, list):
        return ()
    return tuple(str(item) for item in parsed if str(item).strip())


def _json_int_tuple(raw: object) -> tuple[int, ...]:
    try:
        parsed = json.loads(str(raw or "[]"))
    except (TypeError, ValueError):
        return ()
    if not isinstance(parsed, list):
        return ()
    values: list[int] = []
    for item in parsed:
        try:
            values.append(int(item))
        except (TypeError, ValueError):
            continue
    return tuple(values)


def _working_flag_sequence(
    transfer_name: str, sequence: int | None = None
) -> int | None:
    if sequence is not None:
        try:
            return int(sequence)
        except (TypeError, ValueError):
            pass
    folder = str(transfer_name or "").strip()
    if not folder:
        return None
    return parse_transfer_folder(folder, under_gate=True).sequence


def _better_revision_text(left: str, right: str) -> str:
    if not right:
        return left
    if not left:
        return right
    left_rank = revision_rank(*parse_sheet_revision(left))
    right_rank = revision_rank(*parse_sheet_revision(right))
    return right if right_rank > left_rank else left


_NONCANONICAL_RD_CLEAN_KEY = "noncanonical_rd_clean"


def _id_chunks(ids: Sequence[int]) -> Iterator[tuple[int, ...]]:
    values = tuple(int(item) for item in ids)
    for start in range(0, len(values), _FILE_ID_CHUNK):
        yield values[start : start + _FILE_ID_CHUNK]


DEFAULT_BUSY_TIMEOUT_MS = 5_000
LIQUIDITY_DECISIONS = frozenset({"confirmed_ok", "confirmed_illiquid"})
ISSUANCE_REVIEW_KINDS = frozenset({"send", "orphan", "manual"})
ISSUANCE_REVIEW_SOURCES = frozenset(
    {"issuance", "google_f", "rd", "robot", "auto_mto", "manual"}
)
ISSUANCE_REVIEW_DECISIONS = frozenset(
    {"", "active", "legalized", "annulled", "erroneous", "duplicate"}
)
ISSUANCE_REVIEW_MATCH_STATES = frozenset(
    {"", "matched", "unmatched", "ambiguous"}
)
_MTO_PAIR_LIST_CHUNK = 200


def canonical_mto_pair_ids(left_file_id: int, right_file_id: int) -> tuple[int, int]:
    """Order a symmetric MTO file pair by ascending ``file_entry.id``.

    The comparison is direction-independent: ``(A, B)`` and ``(B, A)`` must
    share one ``mto_pair_comparison`` row. Callers that store fingerprints or
    ``(path_key, size, mtime_ns)`` signatures must apply the same permutation.

    Args:
        left_file_id: First file id (any order).
        right_file_id: Second file id (any order).

    Returns:
        ``(min_id, max_id)``.
    """

    first = int(left_file_id)
    second = int(right_file_id)
    if first <= second:
        return first, second
    return second, first


def mto_file_stat_signature(file: FileRecord | ParsedFile) -> dict[str, Any]:
    """Return the cache signature ``path_key`` / ``size`` / ``mtime_ns``.

    Args:
        file: A persisted ``FileRecord`` or an in-memory ``ParsedFile``.

    Returns:
        JSON-safe mapping used as ``left_signature`` / ``right_signature``.
    """

    if isinstance(file, ParsedFile):
        return {
            "path_key": str(file.path_key),
            "size": int(file.size),
            "mtime_ns": int(file.mtime_ns),
        }
    size = file.data.get("size")
    mtime_ns = file.data.get("disk_mtime_ns")
    if mtime_ns in (None, ""):
        mtime_ns = file.data.get("mtime_ns")
    return {
        "path_key": str(file.path_key),
        "size": int(size or 0),
        "mtime_ns": int(mtime_ns or 0),
    }


def _signature_mapping(raw: Any) -> dict[str, Any]:
    if isinstance(raw, str):
        try:
            raw = json.loads(raw)
        except json.JSONDecodeError:
            return {"path_key": "", "size": 0, "mtime_ns": 0}
    if isinstance(raw, Mapping):
        return {
            "path_key": str(raw.get("path_key") or ""),
            "size": int(raw.get("size") or 0),
            "mtime_ns": int(raw.get("mtime_ns") or 0),
        }
    if isinstance(raw, (list, tuple)) and len(raw) >= 3:
        return {
            "path_key": str(raw[0]),
            "size": int(raw[1] or 0),
            "mtime_ns": int(raw[2] or 0),
        }
    return {"path_key": "", "size": 0, "mtime_ns": 0}


def _dump_signature(raw: Any) -> str:
    return json.dumps(_signature_mapping(raw), ensure_ascii=False, sort_keys=True)


def _signatures_equal(stored: Any, current: Any) -> bool:
    return _signature_mapping(stored) == _signature_mapping(current)


@dataclass(frozen=True, slots=True)
class KitPackageRow:
    """One issued package folder or grey issuance-only stub."""

    title: str
    mark: str
    source: str
    sequence: int | None = None
    transfer_name: str | None = None
    package_path: str = ""
    revision_text: str = ""
    max_mtime_ns: int | None = None
    pdf_count: int = 0
    editable_count: int = 0
    mto_revision_text: str = ""
    overlay_current_count: int = 0
    is_grey: bool = False
    is_current: bool = False
    is_as_build: bool = False
    id: int | None = None


@dataclass(frozen=True, slots=True)
class KitCycleRow:
    """One official send cycle linking issuance, package, and F events.

    On write, ``package_id`` is treated as a client-side id matching
    :attr:`KitPackageRow.id` in the same ``replace_kit_derived`` call and is
    remapped to the new SQLite row id.
    """

    title: str
    mark: str
    revision_text: str = ""
    match_reason: str = ""
    send_id: int | None = None
    package_id: int | None = None
    tdo_event_id: int | None = None
    code_event_id: int | None = None
    id: int | None = None


@dataclass(frozen=True, slots=True)
class KitPipelineRow:
    """Derived current pipeline status for one ``(title, mark)`` kit."""

    title: str
    mark: str
    status: str
    code: str | None = None
    code_origin: str | None = None
    working_revision_text: str = ""
    official_revision_text: str = ""
    suspicious: bool = False
    algorithm_version: int = 1
    code_stale: bool = False
    code_revision_text: str = ""
    code_date: str = ""
    tdo_date: str = ""
    review_as_build: bool = False
    working_as_build: bool = False
    working_transfer_names: tuple[str, ...] = ()
    working_sequences: tuple[int, ...] = ()
    annulled_transfer_names: tuple[str, ...] = ()
    annulled_sequences: tuple[int, ...] = ()


@dataclass(frozen=True, slots=True)
class FileMtimeOverrideRow:
    """User catalog date for one catalog file; survives scan while fingerprint matches.

    Identity is ``path_key``. ``fingerprint`` is SHA-256 of
    ``path_key|size|disk_mtime_ns`` from ``file_entry`` at decision time.
    Scan still writes the disk ``mtime_ns``; ``list_files`` substitutes
    ``override_mtime_ns`` only when the fingerprint still matches.
    """

    path_key: str
    size: int
    disk_mtime_ns: int
    fingerprint: str
    override_date: str
    override_mtime_ns: int
    reason: str = "code_a"
    decided_at: str = ""
    id: int | None = None


@dataclass(frozen=True, slots=True)
class KitWorkingFlagRow:
    """User mark that one issued folder is working, not official.

    Identity is ``(title, mark, transfer_name)``. Survives scan and Google
    rebuilds. Auto-detected working revisions are not stored here.
    """

    title: str
    mark: str
    revision_text: str
    transfer_name: str = ""
    sequence: int | None = None
    decided_at: str = ""
    id: int | None = None


@dataclass(frozen=True, slots=True)
class KitAnnulledFlagRow:
    """User mark that one issued folder is annulled, not in contests.

    Identity is ``(title, mark, transfer_name)``. Survives scan and Google
    rebuilds. Distinct from ``issuance_review.decision == "annulled"``.
    """

    title: str
    mark: str
    revision_text: str
    transfer_name: str = ""
    sequence: int | None = None
    decided_at: str = ""
    id: int | None = None


@dataclass(frozen=True, slots=True)
class LiquidityReviewRow:
    """User liquidity decision; survives Google ingest and pipeline rebuild."""

    title: str
    mark: str
    transfer_name: str
    revision_text: str
    decision: str
    sequence: int | None = None
    comment: str | None = None
    evidence_mtime_ns: int | None = None
    decided_at: str = ""
    id: int | None = None


@dataclass(frozen=True, slots=True)
class IssuanceReviewRow:
    """User issuance decision; survives Google ingest and pipeline rebuild."""

    title: str
    mark: str
    kind: str
    source: str
    decision: str
    revision_text: str
    send_date: str
    send_date_sortable: str
    send_transmittal: str
    incoming_control_date: str
    incoming_control_date_sortable: str
    confirm_transmittal: str
    sheet_status: str
    note: str
    comment: str | None
    identity_fingerprint: str
    evidence_fingerprint: str
    match_state: str
    source_path: str
    path_key: str
    decided_at: str
    id: int | None = None


@dataclass(frozen=True, slots=True)
class KitRevisionRow:
    """One derived heatmap cell for a kit filename revision."""

    title: str
    mark: str
    revision_text: str
    pipeline_status: str
    letters: str = ""
    is_as_build: bool = False
    is_current: bool = False
    is_current_ifc: bool = False
    has_mto: bool = False
    problem_kinds_json: str = "[]"
    package_ids_json: str = "[]"
    algorithm_version: int = 1
    id: int | None = None


def _utc_now() -> str:
    return datetime.now(timezone.utc).isoformat()


def file_mtime_fingerprint(path_key: str, size: int, disk_mtime_ns: int) -> str:
    """Return SHA-256 of the scan identity used to keep an MTO date override.

    Args:
        path_key: Case-insensitive path identity.
        size: ``file_entry.size`` at decision time.
        disk_mtime_ns: Disk ``file_entry.mtime_ns`` at decision time.

    Returns:
        Hex digest. Empty ``path_key`` still hashes the numeric parts.
    """

    payload = (
        f"{str(path_key or '').casefold()}|{int(size)}|{int(disk_mtime_ns)}"
    )
    return hashlib.sha256(payload.encode("utf-8")).hexdigest()


def parse_override_date(text: str) -> date:
    """Parse a catalog override date ``DD.MM.YYYY``.

    Args:
        text: Date from F «код A» or a typed value.

    Returns:
        Local calendar date.

    Raises:
        ValueError: Empty or malformed date.
    """

    raw = str(text or "").strip()
    matched = _OVERRIDE_DATE_RE.fullmatch(raw)
    if matched is None:
        raise ValueError(f"Override date must be DD.MM.YYYY, got {raw!r}")
    day, month, year = (int(matched.group(1)), int(matched.group(2)), int(matched.group(3)))
    try:
        return date(year, month, day)
    except ValueError as exc:
        raise ValueError(f"Override date must be DD.MM.YYYY, got {raw!r}") from exc


def override_mtime_ns_from_date(text: str) -> int:
    """Return local-noon nanoseconds for an override calendar date.

    Args:
        text: ``DD.MM.YYYY``.

    Returns:
        Nanoseconds since epoch at 12:00 local time.
    """

    parsed = parse_override_date(text)
    stamp = datetime(parsed.year, parsed.month, parsed.day, 12, 0, 0)
    return int(stamp.timestamp() * 1_000_000_000)


def apply_mtime_override_to_data(
    path_key: str,
    data: dict[str, Any],
    override: FileMtimeOverrideRow | None,
) -> None:
    """Set catalog ``mtime_ns`` from a matching override; keep disk mtime aside.

    Args:
        path_key: File identity.
        data: ``FileRecord.data`` (mutated).
        override: Stored row, or ``None`` to restore the disk mtime.
    """

    if data.get("disk_mtime_ns") in (None, ""):
        data["disk_mtime_ns"] = int(data.get("mtime_ns") or 0)
    disk_mtime = int(data.get("disk_mtime_ns") or 0)
    size = int(data.get("size") or 0)
    data.pop("mtime_override_applied", None)
    data.pop("mtime_override_stale", None)
    data.pop("mtime_override_date", None)
    data.pop("mtime_override_reason", None)
    if override is None:
        data["mtime_ns"] = disk_mtime
        return
    data["mtime_override_date"] = override.override_date
    data["mtime_override_reason"] = override.reason
    expected = file_mtime_fingerprint(path_key, size, disk_mtime)
    if expected != override.fingerprint:
        data["mtime_ns"] = disk_mtime
        data["mtime_override_stale"] = True
        return
    data["mtime_ns"] = int(override.override_mtime_ns)
    data["mtime_override_applied"] = True


def _mtime_override_from_row(row: sqlite3.Row) -> FileMtimeOverrideRow:
    return FileMtimeOverrideRow(
        path_key=str(row["path_key"]),
        size=int(row["size"] or 0),
        disk_mtime_ns=int(row["disk_mtime_ns"] or 0),
        fingerprint=str(row["fingerprint"] or ""),
        override_date=str(row["override_date"] or ""),
        override_mtime_ns=int(row["override_mtime_ns"] or 0),
        reason=str(row["reason"] or "code_a"),
        decided_at=str(row["decided_at"] or ""),
        id=int(row["id"]) if row["id"] is not None else None,
    )


class CatalogDatabase:
    """Short-connection SQLite repository under an injected local path.

    Constructing this class has no filesystem side effects. Call
    :meth:`initialize` explicitly to create or migrate the database.
    Connections use WAL so the GUI can read while a deferred MTO worker
    writes ``mto_comparison`` or ``mto_pair_comparison``.
    """

    def __init__(
        self,
        path: str | Path,
        *,
        busy_timeout_ms: int = DEFAULT_BUSY_TIMEOUT_MS,
    ) -> None:
        """Store database connection settings without opening a file.

        Args:
            path: Injected SQLite path, normally from ``CatalogConfig``.
            busy_timeout_ms: SQLite lock wait timeout.
        """

        self.path = Path(path)
        self.busy_timeout_ms = busy_timeout_ms

    @contextmanager
    def _connection(self) -> Iterator[sqlite3.Connection]:
        connection = sqlite3.connect(self.path, timeout=self.busy_timeout_ms / 1000)
        connection.row_factory = sqlite3.Row
        connection.execute(f"PRAGMA busy_timeout = {self.busy_timeout_ms}")
        connection.execute("PRAGMA foreign_keys = ON")
        connection.execute("PRAGMA journal_mode=WAL")
        try:
            yield connection
        finally:
            connection.close()

    def initialize(self) -> None:
        """Create the parent directory and apply versioned schema migrations."""

        self.path.parent.mkdir(parents=True, exist_ok=True)
        with self._connection() as connection, connection:
            connection.execute(
                "CREATE TABLE IF NOT EXISTS schema_meta "
                "(key TEXT PRIMARY KEY, value TEXT NOT NULL)"
            )
            row = connection.execute(
                "SELECT value FROM schema_meta WHERE key = 'schema_version'"
            ).fetchone()
            version = int(row["value"]) if row else 0
            if version > SCHEMA_VERSION:
                raise RuntimeError(
                    f"Database schema {version} is newer than supported {SCHEMA_VERSION}"
                )
            if version < 1:
                self._migrate_0_to_1(connection)
                version = 1
            if version < 2:
                self._migrate_1_to_2(connection)
                version = 2
            if version < 3:
                self._migrate_2_to_3(connection)
                version = 3
            if version < 4:
                self._migrate_3_to_4(connection)
                version = 4
            if version < 5:
                self._migrate_4_to_5(connection)
                version = 5
            if version < 6:
                self._migrate_5_to_6(connection)
                version = 6
            if version < 7:
                self._migrate_6_to_7(connection)
                version = 7
            if version < 8:
                self._migrate_7_to_8(connection)
                version = 8
            if version < 9:
                self._migrate_8_to_9(connection)
                version = 9
            if version < 10:
                self._migrate_9_to_10(connection)
                version = 10
            if version < 11:
                self._migrate_10_to_11(connection)
                version = 11
            if version < 12:
                self._migrate_11_to_12(connection)
                version = 12
            if version < 13:
                self._migrate_12_to_13(connection)
            self._ensure_kit_pipeline_columns(connection)

    @staticmethod
    def _migrate_0_to_1(connection: sqlite3.Connection) -> None:
        connection.executescript(
            """
            CREATE TABLE scan_root (
                id INTEGER PRIMARY KEY,
                source TEXT NOT NULL UNIQUE,
                root_path TEXT NOT NULL,
                root_path_key TEXT NOT NULL COLLATE NOCASE,
                last_successful_run_id INTEGER
            );

            CREATE TABLE scan_run (
                id INTEGER PRIMARY KEY,
                started_at TEXT NOT NULL,
                completed_at TEXT,
                status TEXT NOT NULL,
                is_baseline INTEGER NOT NULL DEFAULT 0,
                error_json TEXT NOT NULL DEFAULT '[]'
            );

            CREATE TABLE file_entry (
                id INTEGER PRIMARY KEY,
                path TEXT NOT NULL,
                path_key TEXT NOT NULL COLLATE NOCASE UNIQUE,
                source TEXT NOT NULL,
                file_kind TEXT NOT NULL,
                name TEXT NOT NULL,
                size INTEGER NOT NULL,
                mtime_ns INTEGER NOT NULL,
                first_seen_run_id INTEGER NOT NULL REFERENCES scan_run(id),
                last_seen_run_id INTEGER NOT NULL REFERENCES scan_run(id),
                present INTEGER NOT NULL DEFAULT 1,
                review_state TEXT NOT NULL,
                parse_status TEXT NOT NULL,
                contract TEXT,
                title_system TEXT,
                title TEXT,
                mark TEXT,
                discipline_block TEXT,
                core_stem TEXT,
                revision TEXT,
                appendix TEXT,
                language TEXT,
                extension TEXT,
                transfer_sequence INTEGER,
                transfer_revision TEXT,
                transfer_appendix TEXT,
                transfer_is_as_build INTEGER NOT NULL DEFAULT 0,
                transfer_name TEXT,
                parse_error TEXT
            );
            CREATE INDEX file_entry_source_present_idx
                ON file_entry(source, present);
            CREATE INDEX file_entry_document_idx
                ON file_entry(title_system, discipline_block, core_stem);

            CREATE TABLE overlay_state (
                document_key TEXT NOT NULL COLLATE NOCASE,
                file_kind TEXT NOT NULL,
                file_id INTEGER NOT NULL REFERENCES file_entry(id),
                detected_current INTEGER NOT NULL DEFAULT 1,
                algorithm_version INTEGER NOT NULL,
                scan_run_id INTEGER NOT NULL REFERENCES scan_run(id),
                PRIMARY KEY(document_key, file_kind)
            );

            CREATE TABLE review_event (
                id INTEGER PRIMARY KEY,
                file_id INTEGER NOT NULL REFERENCES file_entry(id),
                action TEXT NOT NULL,
                comment TEXT,
                created_at TEXT NOT NULL
            );
            CREATE INDEX review_event_file_idx
                ON review_event(file_id, created_at);

            CREATE TABLE mto_comparison (
                id INTEGER PRIMARY KEY,
                rd_file_id INTEGER NOT NULL REFERENCES file_entry(id),
                robot_file_id INTEGER REFERENCES file_entry(id),
                rd_fingerprint TEXT,
                robot_fingerprint TEXT,
                algorithm_version INTEGER NOT NULL,
                status TEXT NOT NULL,
                stats_json TEXT NOT NULL DEFAULT '{}',
                diff_json TEXT NOT NULL DEFAULT '{}',
                compared_at TEXT NOT NULL,
                UNIQUE(rd_file_id, robot_file_id, algorithm_version)
            );
            CREATE UNIQUE INDEX mto_comparison_pair_idx
                ON mto_comparison(
                    rd_file_id, COALESCE(robot_file_id, -1), algorithm_version
                );
            """
        )
        connection.execute(
            "INSERT OR REPLACE INTO schema_meta(key, value) VALUES"
            " ('schema_version', ?)",
            ("1",),
        )

    @staticmethod
    def _migrate_1_to_2(connection: sqlite3.Connection) -> None:
        connection.executescript(
            """
            CREATE TABLE current_collision (
                id INTEGER PRIMARY KEY,
                scope TEXT NOT NULL,
                source TEXT NOT NULL,
                kind TEXT NOT NULL,
                message TEXT NOT NULL,
                document_key TEXT,
                path_keys_json TEXT NOT NULL DEFAULT '[]',
                scan_run_id INTEGER NOT NULL REFERENCES scan_run(id)
            );
            CREATE INDEX current_collision_scope_source_idx
                ON current_collision(scope, source);
            """
        )
        connection.execute(
            "INSERT OR REPLACE INTO schema_meta(key, value) VALUES"
            " ('schema_version', ?)",
            ("2",),
        )

    @staticmethod
    def _migrate_2_to_3(connection: sqlite3.Connection) -> None:
        connection.executescript(
            """
            CREATE TABLE google_kit (
                id INTEGER PRIMARY KEY,
                title TEXT NOT NULL,
                mark TEXT NOT NULL,
                mark_raw TEXT NOT NULL DEFAULT '',
                title_system TEXT NOT NULL DEFAULT '',
                sheet_revision TEXT,
                sheet_appendix TEXT,
                sheet_revision_text TEXT NOT NULL DEFAULT '',
                status_sheet TEXT NOT NULL DEFAULT '',
                comment_raw TEXT NOT NULL DEFAULT '',
                row_index INTEGER NOT NULL DEFAULT 0,
                UNIQUE(title COLLATE NOCASE, mark COLLATE NOCASE)
            );
            CREATE INDEX google_kit_identity_idx
                ON google_kit(title COLLATE NOCASE, mark COLLATE NOCASE);

            CREATE TABLE google_event (
                id INTEGER PRIMARY KEY,
                kit_id INTEGER NOT NULL REFERENCES google_kit(id) ON DELETE CASCADE,
                seq INTEGER NOT NULL,
                raw TEXT NOT NULL DEFAULT '',
                event_date TEXT,
                stage TEXT NOT NULL DEFAULT 'other',
                stage_label TEXT NOT NULL DEFAULT '',
                revision TEXT,
                appendix TEXT,
                transmittals_json TEXT NOT NULL DEFAULT '[]',
                parsed INTEGER NOT NULL DEFAULT 0,
                UNIQUE(kit_id, seq)
            );
            CREATE INDEX google_event_kit_seq_idx
                ON google_event(kit_id, seq);

            CREATE TABLE issuance_send (
                id INTEGER PRIMARY KEY,
                title TEXT NOT NULL,
                mark TEXT NOT NULL,
                mark_raw TEXT NOT NULL DEFAULT '',
                title_system TEXT NOT NULL DEFAULT '',
                revision TEXT,
                appendix TEXT,
                revision_text TEXT NOT NULL DEFAULT '',
                status TEXT NOT NULL DEFAULT '',
                send_date TEXT NOT NULL DEFAULT '',
                send_date_sortable TEXT NOT NULL DEFAULT '',
                send_transmittal TEXT NOT NULL DEFAULT '',
                incoming_control_date TEXT NOT NULL DEFAULT '',
                incoming_control_date_sortable TEXT NOT NULL DEFAULT '',
                confirm_transmittal TEXT NOT NULL DEFAULT '',
                note_raw TEXT NOT NULL DEFAULT '',
                row_index INTEGER NOT NULL DEFAULT 0
            );
            CREATE INDEX issuance_send_kit_idx
                ON issuance_send(title COLLATE NOCASE, mark COLLATE NOCASE);

            CREATE TABLE google_load (
                id INTEGER PRIMARY KEY CHECK (id = 1),
                loaded_at TEXT NOT NULL,
                source TEXT NOT NULL,
                warning TEXT
            );

            CREATE TABLE kit_package (
                id INTEGER PRIMARY KEY,
                title TEXT NOT NULL,
                mark TEXT NOT NULL,
                source TEXT NOT NULL,
                sequence INTEGER,
                transfer_name TEXT,
                package_path TEXT NOT NULL DEFAULT '',
                revision_text TEXT NOT NULL DEFAULT '',
                max_mtime_ns INTEGER,
                pdf_count INTEGER NOT NULL DEFAULT 0,
                editable_count INTEGER NOT NULL DEFAULT 0,
                mto_revision_text TEXT NOT NULL DEFAULT '',
                overlay_current_count INTEGER NOT NULL DEFAULT 0,
                is_grey INTEGER NOT NULL DEFAULT 0,
                is_current INTEGER NOT NULL DEFAULT 0
            );
            CREATE INDEX kit_package_kit_idx
                ON kit_package(title COLLATE NOCASE, mark COLLATE NOCASE);

            CREATE TABLE kit_cycle (
                id INTEGER PRIMARY KEY,
                title TEXT NOT NULL,
                mark TEXT NOT NULL,
                send_id INTEGER REFERENCES issuance_send(id) ON DELETE SET NULL,
                package_id INTEGER REFERENCES kit_package(id) ON DELETE SET NULL,
                tdo_event_id INTEGER REFERENCES google_event(id) ON DELETE SET NULL,
                code_event_id INTEGER REFERENCES google_event(id) ON DELETE SET NULL,
                revision_text TEXT NOT NULL DEFAULT '',
                match_reason TEXT NOT NULL DEFAULT ''
            );
            CREATE INDEX kit_cycle_kit_idx
                ON kit_cycle(title COLLATE NOCASE, mark COLLATE NOCASE);

            CREATE TABLE kit_pipeline (
                title TEXT NOT NULL,
                mark TEXT NOT NULL,
                status TEXT NOT NULL DEFAULT '',
                code TEXT,
                code_origin TEXT,
                working_revision_text TEXT NOT NULL DEFAULT '',
                official_revision_text TEXT NOT NULL DEFAULT '',
                suspicious INTEGER NOT NULL DEFAULT 0,
                algorithm_version INTEGER NOT NULL DEFAULT 1,
                PRIMARY KEY(title COLLATE NOCASE, mark COLLATE NOCASE)
            );

            CREATE TABLE kit_liquidity_review (
                id INTEGER PRIMARY KEY,
                title TEXT NOT NULL,
                mark TEXT NOT NULL,
                transfer_name TEXT NOT NULL DEFAULT '',
                revision_text TEXT NOT NULL DEFAULT '',
                sequence INTEGER,
                decision TEXT NOT NULL,
                comment TEXT,
                evidence_mtime_ns INTEGER,
                decided_at TEXT NOT NULL,
                UNIQUE(
                    title COLLATE NOCASE,
                    mark COLLATE NOCASE,
                    transfer_name COLLATE NOCASE,
                    revision_text COLLATE NOCASE
                )
            );
            CREATE INDEX kit_liquidity_kit_idx
                ON kit_liquidity_review(title COLLATE NOCASE, mark COLLATE NOCASE);
            """
        )
        connection.execute(
            "INSERT OR REPLACE INTO schema_meta(key, value) VALUES"
            " ('schema_version', ?)",
            ("3",),
        )

    @staticmethod
    def _table_column_names(
        connection: sqlite3.Connection, table: str
    ) -> set[str]:
        rows = connection.execute(f"PRAGMA table_info({table})").fetchall()
        return {str(row[1]) for row in rows}

    @staticmethod
    def _ensure_kit_pipeline_columns(connection: sqlite3.Connection) -> None:
        columns = CatalogDatabase._table_column_names(connection, "kit_pipeline")
        if not columns:
            return
        if "tdo_date" not in columns:
            connection.execute(
                "ALTER TABLE kit_pipeline "
                "ADD COLUMN tdo_date TEXT NOT NULL DEFAULT ''"
            )
        if "working_transfer_names_json" not in columns:
            connection.execute(
                "ALTER TABLE kit_pipeline "
                "ADD COLUMN working_transfer_names_json TEXT NOT NULL DEFAULT '[]'"
            )
        if "working_sequences_json" not in columns:
            connection.execute(
                "ALTER TABLE kit_pipeline "
                "ADD COLUMN working_sequences_json TEXT NOT NULL DEFAULT '[]'"
            )
        if "working_as_build" not in columns:
            connection.execute(
                "ALTER TABLE kit_pipeline "
                "ADD COLUMN working_as_build INTEGER NOT NULL DEFAULT 0"
            )
        if "annulled_transfer_names_json" not in columns:
            connection.execute(
                "ALTER TABLE kit_pipeline "
                "ADD COLUMN annulled_transfer_names_json "
                "TEXT NOT NULL DEFAULT '[]'"
            )
        if "annulled_sequences_json" not in columns:
            connection.execute(
                "ALTER TABLE kit_pipeline "
                "ADD COLUMN annulled_sequences_json "
                "TEXT NOT NULL DEFAULT '[]'"
            )

    @staticmethod
    def _migrate_3_to_4(connection: sqlite3.Connection) -> None:
        columns = CatalogDatabase._table_column_names(connection, "kit_package")
        if "is_as_build" not in columns:
            connection.execute(
                "ALTER TABLE kit_package "
                "ADD COLUMN is_as_build INTEGER NOT NULL DEFAULT 0"
            )
        connection.executescript(
            """
            CREATE TABLE IF NOT EXISTS kit_revision_cell (
                id INTEGER PRIMARY KEY,
                title TEXT NOT NULL,
                mark TEXT NOT NULL,
                revision_text TEXT NOT NULL,
                pipeline_status TEXT NOT NULL DEFAULT '',
                letters TEXT NOT NULL DEFAULT '',
                is_as_build INTEGER NOT NULL DEFAULT 0,
                is_current INTEGER NOT NULL DEFAULT 0,
                is_current_ifc INTEGER NOT NULL DEFAULT 0,
                has_mto INTEGER NOT NULL DEFAULT 0,
                problem_kinds_json TEXT NOT NULL DEFAULT '[]',
                package_ids_json TEXT NOT NULL DEFAULT '[]',
                algorithm_version INTEGER NOT NULL DEFAULT 1,
                UNIQUE(
                    title COLLATE NOCASE,
                    mark COLLATE NOCASE,
                    revision_text COLLATE NOCASE
                )
            );
            CREATE INDEX IF NOT EXISTS kit_revision_cell_kit_idx
                ON kit_revision_cell(title COLLATE NOCASE, mark COLLATE NOCASE);
            """
        )
        connection.execute(
            "INSERT OR REPLACE INTO schema_meta(key, value) VALUES"
            " ('schema_version', ?)",
            ("4",),
        )

    @staticmethod
    def _migrate_4_to_5(connection: sqlite3.Connection) -> None:
        columns = CatalogDatabase._table_column_names(connection, "kit_pipeline")
        if "code_stale" not in columns:
            connection.execute(
                "ALTER TABLE kit_pipeline "
                "ADD COLUMN code_stale INTEGER NOT NULL DEFAULT 0"
            )
        if "code_revision_text" not in columns:
            connection.execute(
                "ALTER TABLE kit_pipeline "
                "ADD COLUMN code_revision_text TEXT NOT NULL DEFAULT ''"
            )
        if "code_date" not in columns:
            connection.execute(
                "ALTER TABLE kit_pipeline "
                "ADD COLUMN code_date TEXT NOT NULL DEFAULT ''"
            )
        if "review_as_build" not in columns:
            connection.execute(
                "ALTER TABLE kit_pipeline "
                "ADD COLUMN review_as_build INTEGER NOT NULL DEFAULT 0"
            )
        connection.execute(
            "INSERT OR REPLACE INTO schema_meta(key, value) VALUES"
            " ('schema_version', ?)",
            ("5",),
        )

    @staticmethod
    def _migrate_5_to_6(connection: sqlite3.Connection) -> None:
        connection.executescript(
            """
            CREATE TABLE IF NOT EXISTS mto_pair_comparison (
                id INTEGER PRIMARY KEY,
                left_file_id INTEGER NOT NULL REFERENCES file_entry(id),
                right_file_id INTEGER NOT NULL REFERENCES file_entry(id),
                algorithm_version INTEGER NOT NULL,
                content_status TEXT NOT NULL,
                left_fingerprint TEXT,
                right_fingerprint TEXT,
                left_signature TEXT NOT NULL,
                right_signature TEXT NOT NULL,
                stats_json TEXT NOT NULL DEFAULT '{}',
                diff_json TEXT NOT NULL DEFAULT '{}',
                error TEXT,
                compared_at TEXT NOT NULL
            );
            CREATE UNIQUE INDEX IF NOT EXISTS mto_pair_comparison_idx
                ON mto_pair_comparison(
                    left_file_id, right_file_id, algorithm_version
                );
            """
        )
        connection.execute(
            "INSERT OR REPLACE INTO schema_meta(key, value) VALUES"
            " ('schema_version', ?)",
            ("6",),
        )

    @staticmethod
    def _migrate_6_to_7(connection: sqlite3.Connection) -> None:
        connection.executescript(
            """
            CREATE TABLE issuance_review (
                id INTEGER PRIMARY KEY,
                title TEXT NOT NULL,
                mark TEXT NOT NULL,
                kind TEXT NOT NULL,
                source TEXT NOT NULL,
                decision TEXT NOT NULL DEFAULT '',
                revision_text TEXT NOT NULL DEFAULT '',
                send_date TEXT NOT NULL DEFAULT '',
                send_date_sortable TEXT NOT NULL DEFAULT '',
                send_transmittal TEXT NOT NULL DEFAULT '',
                incoming_control_date TEXT NOT NULL DEFAULT '',
                incoming_control_date_sortable TEXT NOT NULL DEFAULT '',
                confirm_transmittal TEXT NOT NULL DEFAULT '',
                sheet_status TEXT NOT NULL DEFAULT '',
                note TEXT NOT NULL DEFAULT '',
                comment TEXT,
                identity_fingerprint TEXT NOT NULL,
                evidence_fingerprint TEXT NOT NULL DEFAULT '',
                match_state TEXT NOT NULL DEFAULT '',
                source_path TEXT NOT NULL DEFAULT '',
                path_key TEXT NOT NULL DEFAULT '',
                decided_at TEXT NOT NULL
            );
            CREATE UNIQUE INDEX issuance_review_identity_idx
                ON issuance_review(kind, identity_fingerprint);
            CREATE INDEX issuance_review_kit_idx
                ON issuance_review(title COLLATE NOCASE, mark COLLATE NOCASE);
            """
        )
        connection.execute(
            "INSERT OR REPLACE INTO schema_meta(key, value) VALUES"
            " ('schema_version', ?)",
            ("7",),
        )

    @staticmethod
    def _migrate_7_to_8(connection: sqlite3.Connection) -> None:
        connection.executescript(
            """
            CREATE TABLE an_mto_file (
                id INTEGER PRIMARY KEY,
                path TEXT NOT NULL UNIQUE,
                path_key TEXT NOT NULL,
                title TEXT NOT NULL,
                mark TEXT NOT NULL,
                revision_text TEXT NOT NULL DEFAULT '',
                core_stem TEXT NOT NULL DEFAULT '',
                discipline_block TEXT NOT NULL DEFAULT '',
                name TEXT NOT NULL,
                parent_dir TEXT NOT NULL DEFAULT '',
                mtime_ns INTEGER NOT NULL DEFAULT 0,
                size INTEGER NOT NULL DEFAULT 0,
                scanned_at TEXT NOT NULL
            );
            CREATE INDEX an_mto_file_kit_idx
                ON an_mto_file(title COLLATE NOCASE, mark COLLATE NOCASE);
            """
        )
        connection.execute(
            "INSERT OR REPLACE INTO schema_meta(key, value) VALUES"
            " ('schema_version', ?)",
            ("8",),
        )

    @staticmethod
    def _migrate_8_to_9(connection: sqlite3.Connection) -> None:
        connection.executescript(
            """
            CREATE TABLE kit_working_flag (
                id INTEGER PRIMARY KEY,
                title TEXT NOT NULL,
                mark TEXT NOT NULL,
                revision_text TEXT NOT NULL,
                transfer_name TEXT NOT NULL DEFAULT '',
                decided_at TEXT NOT NULL DEFAULT '',
                UNIQUE (
                    title COLLATE NOCASE,
                    mark COLLATE NOCASE,
                    revision_text COLLATE NOCASE
                )
            );
            CREATE INDEX kit_working_flag_kit_idx
                ON kit_working_flag(title COLLATE NOCASE, mark COLLATE NOCASE);
            """
        )
        connection.execute(
            "INSERT OR REPLACE INTO schema_meta(key, value) VALUES"
            " ('schema_version', ?)",
            ("9",),
        )

    @staticmethod
    def _migrate_9_to_10(connection: sqlite3.Connection) -> None:
        connection.executescript(
            """
            CREATE TABLE kit_working_flag_v10 (
                id INTEGER PRIMARY KEY,
                title TEXT NOT NULL,
                mark TEXT NOT NULL,
                revision_text TEXT NOT NULL,
                transfer_name TEXT NOT NULL DEFAULT '',
                sequence INTEGER,
                decided_at TEXT NOT NULL DEFAULT '',
                UNIQUE (
                    title COLLATE NOCASE,
                    mark COLLATE NOCASE,
                    transfer_name COLLATE NOCASE
                )
            );
            """
        )
        rows = connection.execute("SELECT * FROM kit_working_flag").fetchall()
        seen: set[tuple[str, str, str]] = set()
        for row in rows:
            title = str(row["title"] or "").strip()
            mark = str(row["mark"] or "").strip()
            revision = str(row["revision_text"] or "").strip()
            folder = str(row["transfer_name"] or "").strip() or revision
            if not title or not mark or not folder:
                continue
            key = (title.casefold(), mark.casefold(), folder.casefold())
            if key in seen:
                continue
            seen.add(key)
            sequence = _working_flag_sequence(folder)
            connection.execute(
                """
                INSERT INTO kit_working_flag_v10(
                    title, mark, revision_text, transfer_name, sequence, decided_at
                ) VALUES (?, ?, ?, ?, ?, ?)
                """,
                (
                    title,
                    mark,
                    revision,
                    folder,
                    sequence,
                    str(row["decided_at"] or ""),
                ),
            )
        connection.execute("DROP TABLE kit_working_flag")
        connection.execute(
            "ALTER TABLE kit_working_flag_v10 RENAME TO kit_working_flag"
        )
        connection.execute(
            """
            CREATE INDEX kit_working_flag_kit_idx
                ON kit_working_flag(title COLLATE NOCASE, mark COLLATE NOCASE)
            """
        )
        pipeline_columns = CatalogDatabase._table_column_names(
            connection, "kit_pipeline"
        )
        if (
            pipeline_columns
            and "working_transfer_names_json" not in pipeline_columns
        ):
            connection.execute(
                "ALTER TABLE kit_pipeline "
                "ADD COLUMN working_transfer_names_json TEXT NOT NULL DEFAULT '[]'"
            )
        if pipeline_columns and "working_sequences_json" not in pipeline_columns:
            connection.execute(
                "ALTER TABLE kit_pipeline "
                "ADD COLUMN working_sequences_json TEXT NOT NULL DEFAULT '[]'"
            )
        connection.execute(
            "INSERT OR REPLACE INTO schema_meta(key, value) VALUES"
            " ('schema_version', ?)",
            ("10",),
        )

    @staticmethod
    def _migrate_10_to_11(connection: sqlite3.Connection) -> None:
        connection.executescript(
            """
            CREATE TABLE file_mtime_override (
                id INTEGER PRIMARY KEY,
                path_key TEXT NOT NULL COLLATE NOCASE,
                size INTEGER NOT NULL,
                disk_mtime_ns INTEGER NOT NULL,
                fingerprint TEXT NOT NULL,
                override_date TEXT NOT NULL,
                override_mtime_ns INTEGER NOT NULL,
                reason TEXT NOT NULL DEFAULT 'code_a',
                decided_at TEXT NOT NULL DEFAULT '',
                UNIQUE (path_key COLLATE NOCASE)
            );
            CREATE INDEX file_mtime_override_fp_idx
                ON file_mtime_override(fingerprint);
            """
        )
        connection.execute(
            "INSERT OR REPLACE INTO schema_meta(key, value) VALUES"
            " ('schema_version', ?)",
            ("11",),
        )

    @staticmethod
    def _migrate_11_to_12(connection: sqlite3.Connection) -> None:
        connection.executescript(
            """
            CREATE TABLE kit_annulled_flag (
                id INTEGER PRIMARY KEY,
                title TEXT NOT NULL,
                mark TEXT NOT NULL,
                revision_text TEXT NOT NULL,
                transfer_name TEXT NOT NULL DEFAULT '',
                sequence INTEGER,
                decided_at TEXT NOT NULL DEFAULT '',
                UNIQUE (
                    title COLLATE NOCASE,
                    mark COLLATE NOCASE,
                    transfer_name COLLATE NOCASE
                )
            );
            CREATE INDEX kit_annulled_flag_kit_idx
                ON kit_annulled_flag(
                    title COLLATE NOCASE, mark COLLATE NOCASE
                );
            """
        )
        connection.execute(
            "INSERT OR REPLACE INTO schema_meta(key, value) VALUES"
            " ('schema_version', ?)",
            ("12",),
        )

    @staticmethod
    def _migrate_12_to_13(connection: sqlite3.Connection) -> None:
        connection.executescript(
            """
            CREATE TABLE rd_dump_mto_file (
                id INTEGER PRIMARY KEY,
                path TEXT NOT NULL UNIQUE,
                path_key TEXT NOT NULL,
                title TEXT NOT NULL,
                mark TEXT NOT NULL,
                revision_text TEXT NOT NULL DEFAULT '',
                core_stem TEXT NOT NULL DEFAULT '',
                discipline_block TEXT NOT NULL DEFAULT '',
                name TEXT NOT NULL,
                parent_dir TEXT NOT NULL DEFAULT '',
                mtime_ns INTEGER NOT NULL DEFAULT 0,
                size INTEGER NOT NULL DEFAULT 0,
                scanned_at TEXT NOT NULL
            );
            CREATE INDEX rd_dump_mto_file_kit_idx
                ON rd_dump_mto_file(title COLLATE NOCASE, mark COLLATE NOCASE);
            """
        )
        connection.execute(
            "INSERT OR REPLACE INTO schema_meta(key, value) VALUES"
            " ('schema_version', ?)",
            ("13",),
        )

    def schema_version(self) -> int:
        """Return the initialized database schema version.

        Returns:
            Integer schema version.
        """

        with self._connection() as connection:
            row = connection.execute(
                "SELECT value FROM schema_meta WHERE key = 'schema_version'"
            ).fetchone()
        return int(row["value"]) if row else 0

    @staticmethod
    def _file_values(file: ParsedFile) -> tuple[Any, ...]:
        transfer = file.transfer
        transfer_is_as_build = bool(transfer and transfer.is_as_build) or path_is_as_build(
            file.path
        )
        return (
            file.path,
            file.path_key,
            file.source.value,
            file.file_kind.value,
            file.name,
            file.size,
            file.mtime_ns,
            file.parse_status.value,
            file.contract,
            file.title_system,
            file.title,
            file.mark,
            file.discipline_block,
            file.core_stem,
            file.revision,
            file.appendix,
            file.language,
            file.extension,
            transfer.sequence if transfer else None,
            transfer.revision if transfer else None,
            transfer.appendix if transfer else None,
            int(transfer_is_as_build),
            transfer.original_name if transfer else None,
            file.parse_error,
        )

    @staticmethod
    def _has_successful_scan(connection: sqlite3.Connection) -> bool:
        return (
            connection.execute(
                "SELECT 1 FROM scan_run WHERE status = ? LIMIT 1",
                (ScanRunStatus.SUCCESS.value,),
            ).fetchone()
            is not None
        )

    def store_scan(self, summary: ScanSummary) -> int:
        """Persist a scan in one transaction with baseline/missing semantics.

        The first successful scan is a baseline and creates no pending rows.
        Later additions, stat changes, and missing files become pending.
        A cancelled walk or a failed source/subtree root skips missing
        detection. Nested walk errors still mark files missing outside those
        error paths, so a permission blip on one folder cannot freeze stale
        paths for the rest of RD. After missing detection, RD rows whose
        path is not the issued ``title/mark/gate/NN_`` layout are deleted
        (legacy leftover; canonical ``present=0`` is kept). Does not touch
        ``an_mto_file`` or ``rd_dump_mto_file``.

        Args:
            summary: Completed or partial scanner output.

        Returns:
            New scan-run identifier.
        """

        started_at = _utc_now()
        with self._connection() as connection, connection:
            had_success = self._has_successful_scan(connection)
            is_baseline = summary.status is ScanRunStatus.SUCCESS and not had_success
            error_payload = [
                {
                    "source": error.source.value,
                    "path": error.path,
                    "message": error.message,
                }
                for error in summary.errors
            ]
            cursor = connection.execute(
                """
                INSERT INTO scan_run(
                    started_at, completed_at, status, is_baseline, error_json
                ) VALUES (?, ?, ?, ?, ?)
                """,
                (
                    started_at,
                    _utc_now(),
                    summary.status.value,
                    int(is_baseline),
                    json.dumps(error_payload, ensure_ascii=False),
                ),
            )
            run_id = int(cursor.lastrowid)

            for source, source_result in summary.sources.items():
                root_key = make_path_key(source_result.root)
                connection.execute(
                    """
                    INSERT INTO scan_root(source, root_path, root_path_key)
                    VALUES (?, ?, ?)
                    ON CONFLICT(source) DO UPDATE SET
                        root_path = excluded.root_path,
                        root_path_key = excluded.root_path_key
                    """,
                    (source.value, source_result.root, root_key),
                )

            self._batch_upsert_files(
                connection,
                run_id,
                summary.files,
                is_baseline=is_baseline or not had_success,
            )
            self._mark_missing(connection, run_id, summary, is_baseline=is_baseline)
            rd_result = summary.sources.get(SourceKind.RD)
            if rd_result is not None and not self._source_scan_unreliable(rd_result):
                # Next startup hydrate walks leftover rows. Do not scan every
                # RD path here: a kit rescan would pay ~0.5 s for 0 deletes.
                self._set_noncanonical_rd_clean(connection, clean=False)
            self._replace_overlays(connection, run_id, summary)
            self._replace_collision_snapshots(connection, run_id, summary)
            self._apply_void_annulled_flags(connection, summary.files)

            if summary.status is ScanRunStatus.SUCCESS:
                connection.execute(
                    """
                    UPDATE scan_root
                    SET last_successful_run_id = ?
                    WHERE source IN (
                        SELECT value FROM json_each(?)
                    )
                    """,
                    (
                        run_id,
                        json.dumps([source.value for source in summary.sources]),
                    ),
                )
            return run_id

    def purge_noncanonical_rd_files(
        self,
        rd_root: str | Path | None,
        *,
        skip_if_clean: bool = False,
    ) -> int:
        """Delete RD catalog rows that are not on an issued-transfer path.

        SQ, ROBOT, and canonical RD (including ``present=0`` missing files)
        stay. Used at startup so leftover rows from before the scan filter
        do not wait for the next walk.

        After a completed walk (0 or more deletes) the database is marked
        clean via ``schema_meta.noncanonical_rd_clean``. A reliable RD
        ``store_scan`` clears that flag; the next hydrate walks once.

        Args:
            rd_root: RD source root. Empty or ``None`` skips the purge
                (a missing root would otherwise match nothing as canonical).
            skip_if_clean: When True, return immediately if the last
                completed purge deleted nothing that still needs a walk.

        Returns:
            Number of ``file_entry`` rows deleted.
        """

        root = str(rd_root or "").strip()
        if not root:
            return 0
        with self._connection() as connection, connection:
            if skip_if_clean and self._noncanonical_rd_is_clean(connection):
                with perf_span("db.purge_noncanonical_rd", skipped=1):
                    return 0
            deleted = self._purge_noncanonical_rd_files(connection, root)
            self._set_noncanonical_rd_clean(connection, clean=True)
            return deleted

    @staticmethod
    def _noncanonical_rd_is_clean(connection: sqlite3.Connection) -> bool:
        row = connection.execute(
            "SELECT value FROM schema_meta WHERE key = ?",
            (_NONCANONICAL_RD_CLEAN_KEY,),
        ).fetchone()
        return bool(row) and str(row["value"]) == "1"

    @staticmethod
    def _set_noncanonical_rd_clean(
        connection: sqlite3.Connection,
        *,
        clean: bool,
    ) -> None:
        if clean:
            connection.execute(
                "INSERT OR REPLACE INTO schema_meta(key, value) VALUES (?, ?)",
                (_NONCANONICAL_RD_CLEAN_KEY, "1"),
            )
            return
        connection.execute(
            "DELETE FROM schema_meta WHERE key = ?",
            (_NONCANONICAL_RD_CLEAN_KEY,),
        )

    @staticmethod
    def _purge_noncanonical_rd_files(
        connection: sqlite3.Connection,
        rd_root: str,
    ) -> int:
        root = str(rd_root or "").strip()
        if not root:
            return 0
        with perf_span("db.purge_noncanonical_rd"):
            rows = connection.execute(
                "SELECT id, path FROM file_entry WHERE source = ?",
                (SourceKind.RD.value,),
            ).fetchall()
            ids = [
                int(row["id"])
                for row in rows
                if not has_canonical_rd_issued_path(str(row["path"]), root)
            ]
            if not ids:
                return 0
            CatalogDatabase._delete_file_entry_ids(connection, ids)
            return len(ids)

    @staticmethod
    def _delete_file_entry_ids(
        connection: sqlite3.Connection,
        ids: Sequence[int],
    ) -> None:
        for chunk in _id_chunks(ids):
            placeholders = ",".join("?" * len(chunk))
            connection.execute(
                f"DELETE FROM overlay_state WHERE file_id IN ({placeholders})",
                chunk,
            )
            connection.execute(
                f"DELETE FROM review_event WHERE file_id IN ({placeholders})",
                chunk,
            )
            connection.execute(
                f"""
                DELETE FROM mto_comparison
                WHERE rd_file_id IN ({placeholders})
                   OR robot_file_id IN ({placeholders})
                """,
                (*chunk, *chunk),
            )
            connection.execute(
                f"""
                DELETE FROM mto_pair_comparison
                WHERE left_file_id IN ({placeholders})
                   OR right_file_id IN ({placeholders})
                """,
                (*chunk, *chunk),
            )
            connection.execute(
                f"DELETE FROM file_entry WHERE id IN ({placeholders})",
                chunk,
            )

    @staticmethod
    def _replace_collision_snapshots(
        connection: sqlite3.Connection,
        run_id: int,
        summary: ScanSummary,
    ) -> None:
        for source, source_result in summary.sources.items():
            if CatalogDatabase._source_scan_unreliable(source_result):
                continue
            if not source_result.errors and not source_result.scoped_folders():
                connection.execute(
                    "DELETE FROM current_collision WHERE scope = 'source' AND source = ?",
                    (source.value,),
                )
                CatalogDatabase._insert_collisions(
                    connection,
                    run_id,
                    scope="source",
                    source=source,
                    collisions=source_result.collisions,
                )

            if source is not SourceKind.RD:
                continue
            for scope, overlay in (
                ("overlay_pdf", summary.rd_pdf_overlay),
                ("overlay_mto", summary.rd_mto_overlay),
            ):
                connection.execute(
                    "DELETE FROM current_collision WHERE scope = ? AND source = ?",
                    (scope, SourceKind.RD.value),
                )
                if overlay is not None:
                    CatalogDatabase._insert_collisions(
                        connection,
                        run_id,
                        scope=scope,
                        source=SourceKind.RD,
                        collisions=overlay.collisions,
                    )

    @staticmethod
    def _insert_collisions(
        connection: sqlite3.Connection,
        run_id: int,
        *,
        scope: str,
        source: SourceKind,
        collisions: Sequence[OverlayCollision],
    ) -> None:
        connection.executemany(
            """
            INSERT INTO current_collision(
                scope, source, kind, message, document_key,
                path_keys_json, scan_run_id
            ) VALUES (?, ?, ?, ?, ?, ?, ?)
            """,
            [
                (
                    scope,
                    source.value,
                    collision.kind.value,
                    collision.message,
                    collision.document_key,
                    json.dumps(collision.path_keys, ensure_ascii=False),
                    run_id,
                )
                for collision in collisions
            ],
        )

    @staticmethod
    def _source_scan_unreliable(source_result: SourceScanResult) -> bool:
        """Return whether this walk cannot safely detect missing files.

        A cancelled walk or an error on the source root (or scoped subtree
        root) means unseen rows must stay present. Nested folder errors are
        not enough to freeze the whole source.

        Args:
            source_result: One source's scan output.

        Returns:
            ``True`` when missing detection and overlay persist must be skipped.
        """

        if source_result.cancelled:
            return True
        if not source_result.errors:
            return False
        roots = [source_result.root]
        roots.extend(source_result.scoped_folders())
        root_keys = {make_path_key(root) for root in roots if root}
        return any(
            make_path_key(error.path) in root_keys for error in source_result.errors
        )

    def _batch_upsert_files(
        self,
        connection: sqlite3.Connection,
        run_id: int,
        files: Sequence[ParsedFile],
        *,
        is_baseline: bool,
    ) -> None:
        for file in files:
            existing = connection.execute(
                "SELECT id, size, mtime_ns, present, review_state "
                "FROM file_entry WHERE path_key = ? COLLATE NOCASE",
                (file.path_key,),
            ).fetchone()
            changed = bool(
                existing
                and (
                    existing["size"] != file.size
                    or existing["mtime_ns"] != file.mtime_ns
                    or not existing["present"]
                )
            )
            if existing and existing["review_state"] == ReviewState.IGNORED.value:
                review_state = ReviewState.IGNORED
            elif is_baseline:
                review_state = ReviewState.ACKNOWLEDGED
            elif existing is None or changed:
                review_state = ReviewState.PENDING
            else:
                review_state = ReviewState(existing["review_state"])

            values = self._file_values(file)
            connection.execute(
                """
                INSERT INTO file_entry(
                    path, path_key, source, file_kind, name, size, mtime_ns,
                    parse_status, contract, title_system, title, mark,
                    discipline_block, core_stem, revision, appendix, language,
                    extension, transfer_sequence, transfer_revision,
                    transfer_appendix, transfer_is_as_build, transfer_name,
                    parse_error, first_seen_run_id, last_seen_run_id, present,
                    review_state
                ) VALUES (
                    ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?,
                    ?, ?, ?, ?, ?, ?, ?, ?, 1, ?
                )
                ON CONFLICT(path_key) DO UPDATE SET
                    path = excluded.path,
                    source = excluded.source,
                    file_kind = excluded.file_kind,
                    name = excluded.name,
                    size = excluded.size,
                    mtime_ns = excluded.mtime_ns,
                    parse_status = excluded.parse_status,
                    contract = excluded.contract,
                    title_system = excluded.title_system,
                    title = excluded.title,
                    mark = excluded.mark,
                    discipline_block = excluded.discipline_block,
                    core_stem = excluded.core_stem,
                    revision = excluded.revision,
                    appendix = excluded.appendix,
                    language = excluded.language,
                    extension = excluded.extension,
                    transfer_sequence = excluded.transfer_sequence,
                    transfer_revision = excluded.transfer_revision,
                    transfer_appendix = excluded.transfer_appendix,
                    transfer_is_as_build = excluded.transfer_is_as_build,
                    transfer_name = excluded.transfer_name,
                    parse_error = excluded.parse_error,
                    last_seen_run_id = excluded.last_seen_run_id,
                    present = 1,
                    review_state = excluded.review_state
                """,
                (*values, run_id, run_id, review_state.value),
            )
            if not is_baseline and (existing is None or changed):
                file_id = (
                    int(existing["id"])
                    if existing
                    else int(
                        connection.execute(
                            "SELECT id FROM file_entry WHERE path_key = ? COLLATE NOCASE",
                            (file.path_key,),
                        ).fetchone()["id"]
                    )
                )
                connection.execute(
                    """
                    INSERT INTO review_event(file_id, action, comment, created_at)
                    VALUES (?, ?, NULL, ?)
                    """,
                    (file_id, "detected_changed" if existing else "detected_added", _utc_now()),
                )

    @staticmethod
    def _mark_missing(
        connection: sqlite3.Connection,
        run_id: int,
        summary: ScanSummary,
        *,
        is_baseline: bool,
    ) -> None:
        if is_baseline:
            return
        for source, source_result in summary.sources.items():
            if CatalogDatabase._source_scan_unreliable(source_result):
                continue
            error_roots = [error.path for error in source_result.errors if error.path]
            missing = connection.execute(
                """
                SELECT id, path, review_state FROM file_entry
                WHERE source = ? AND present = 1 AND last_seen_run_id <> ?
                """,
                (source.value, run_id),
            ).fetchall()
            scopes = source_result.scoped_folders()
            for row in missing:
                if scopes and not any(
                    path_is_under(row["path"], scope) for scope in scopes
                ):
                    continue
                if any(path_is_under(row["path"], err) for err in error_roots):
                    continue
                if not matches_agcc_filename(row["path"]):
                    connection.execute(
                        "UPDATE file_entry SET present = 0 WHERE id = ?",
                        (row["id"],),
                    )
                    continue
                review_state = (
                    ReviewState.IGNORED.value
                    if row["review_state"] == ReviewState.IGNORED.value
                    else ReviewState.PENDING.value
                )
                connection.execute(
                    "UPDATE file_entry SET present = 0, review_state = ? WHERE id = ?",
                    (review_state, row["id"]),
                )
                connection.execute(
                    """
                    INSERT INTO review_event(file_id, action, comment, created_at)
                    VALUES (?, 'detected_missing', NULL, ?)
                    """,
                    (row["id"], _utc_now()),
                )

    @staticmethod
    def _present_rd_parsed_files(
        connection: sqlite3.Connection,
    ) -> list[ParsedFile]:
        rows = connection.execute(
            """
            SELECT * FROM file_entry
            WHERE source = ? AND present = 1
            ORDER BY path_key
            """,
            (SourceKind.RD.value,),
        ).fetchall()
        return [CatalogDatabase._to_parsed_file(row) for row in rows]

    @staticmethod
    def _rd_root_for_overlay(
        connection: sqlite3.Connection,
        summary: ScanSummary | None = None,
    ) -> str:
        if summary is not None:
            rd_result = summary.sources.get(SourceKind.RD)
            if rd_result is not None and str(rd_result.root or "").strip():
                return str(rd_result.root)
        row = connection.execute(
            "SELECT root_path FROM scan_root WHERE source = ?",
            (SourceKind.RD.value,),
        ).fetchone()
        return str(row["root_path"]) if row else ""

    @staticmethod
    def _ignored_path_keys(connection: sqlite3.Connection) -> set[str]:
        rows = connection.execute(
            "SELECT path_key FROM file_entry WHERE review_state = ?",
            (ReviewState.IGNORED.value,),
        ).fetchall()
        return {str(row["path_key"]) for row in rows}

    @staticmethod
    def _replace_overlays(
        connection: sqlite3.Connection,
        run_id: int,
        summary: ScanSummary,
    ) -> None:
        rd_result = summary.sources.get(SourceKind.RD)
        if rd_result is None or CatalogDatabase._source_scan_unreliable(rd_result):
            return
        if rd_result.subtree or rd_result.subtrees or rd_result.errors:
            pdf_overlay, mto_overlay = build_rd_overlays(
                CatalogDatabase._present_rd_parsed_files(connection),
                ignored_path_keys=CatalogDatabase._ignored_path_keys(connection),
                rd_root=CatalogDatabase._rd_root_for_overlay(connection, summary),
            )
            summary.rd_pdf_overlay = pdf_overlay
            summary.rd_mto_overlay = mto_overlay
        CatalogDatabase._persist_overlays(
            connection,
            run_id,
            summary.rd_pdf_overlay,
            summary.rd_mto_overlay,
        )

    @staticmethod
    def _persist_overlays(
        connection: sqlite3.Connection,
        run_id: int,
        pdf_overlay: OverlayResult | None,
        mto_overlay: OverlayResult | None,
    ) -> None:
        for overlay in (pdf_overlay, mto_overlay):
            if overlay is None:
                continue
            connection.execute(
                "DELETE FROM overlay_state WHERE file_kind = ?",
                (overlay.file_kind.value,),
            )
            for document_key, file in overlay.current.items():
                row = connection.execute(
                    "SELECT id FROM file_entry WHERE path_key = ? COLLATE NOCASE",
                    (file.path_key,),
                ).fetchone()
                if row:
                    connection.execute(
                        """
                        INSERT INTO overlay_state(
                            document_key, file_kind, file_id, detected_current,
                            algorithm_version, scan_run_id
                        ) VALUES (?, ?, ?, 1, ?, ?)
                        """,
                        (
                            document_key,
                            overlay.file_kind.value,
                            row["id"],
                            OVERLAY_ALGORITHM_VERSION,
                            run_id,
                        ),
                    )

    def count_present_skipped(self, skip_dirs: Iterable[str]) -> int:
        """Count present RD/SQ files whose parent dirs match skip tokens.

        Args:
            skip_dirs: Current skip tokens.

        Returns:
            Number of present catalog rows that would be marked absent.
        """

        return len(self._present_skipped_ids(skip_dirs))

    def apply_skip_dirs(self, skip_dirs: Iterable[str]) -> int:
        """Mark skipped RD/SQ files absent and rebuild overlay from SQLite.

        Does not walk the filesystem, does not create a scan run, and does
        not change review state. Robot files are left untouched.

        Args:
            skip_dirs: Current skip tokens.

        Returns:
            Number of rows marked absent.
        """

        with perf_span("db.apply_skip_dirs"):
            tokens = tuple(skip_dirs)
            with self._connection() as connection, connection:
                ids = self._present_skipped_ids(tokens, connection=connection)
                for file_id in ids:
                    connection.execute(
                        "UPDATE file_entry SET present = 0 WHERE id = ?",
                        (file_id,),
                    )
                run_row = connection.execute(
                    "SELECT id FROM scan_run WHERE status = ? ORDER BY id DESC LIMIT 1",
                    (ScanRunStatus.SUCCESS.value,),
                ).fetchone()
                if run_row is None:
                    run_row = connection.execute(
                        "SELECT id FROM scan_run ORDER BY id DESC LIMIT 1"
                    ).fetchone()
                if run_row is not None:
                    pdf_overlay, mto_overlay = build_rd_overlays(
                        self._present_rd_parsed_files(connection),
                        ignored_path_keys=self._ignored_path_keys(connection),
                        rd_root=self._rd_root_for_overlay(connection),
                    )
                    self._persist_overlays(
                        connection,
                        int(run_row["id"]),
                        pdf_overlay,
                        mto_overlay,
                    )
                return len(ids)

    def _present_skipped_ids(
        self,
        skip_dirs: Iterable[str],
        *,
        connection: sqlite3.Connection | None = None,
    ) -> list[int]:
        tokens = tuple(skip_dirs)
        if connection is not None:
            return self._present_skipped_ids_on(connection, tokens)
        with self._connection() as owned:
            return self._present_skipped_ids_on(owned, tokens)

    @staticmethod
    def _present_skipped_ids_on(
        connection: sqlite3.Connection,
        skip_dirs: tuple[str, ...],
    ) -> list[int]:
        if not skip_dirs:
            return []
        rows = connection.execute(
            """
            SELECT id, path FROM file_entry
            WHERE present = 1 AND source IN (?, ?)
            """,
            (SourceKind.RD.value, SourceKind.SQ.value),
        ).fetchall()
        return [
            int(row["id"])
            for row in rows
            if path_has_skipped_dir(row["path"], skip_dirs)
        ]

    def list_files(
        self,
        *,
        source: SourceKind | None = None,
        present_only: bool = False,
        review_state: ReviewState | None = None,
    ) -> list[FileRecord]:
        """Read file rows for a future GUI model.

        Matching ``file_mtime_override`` rows replace ``data['mtime_ns']``
        with the catalog date; ``data['disk_mtime_ns']`` stays the scan
        value. A fingerprint miss leaves disk mtime and sets
        ``mtime_override_stale``.

        Args:
            source: Optional source filter.
            present_only: Return only currently present files.
            review_state: Optional review-state filter.

        Returns:
            Stable file records ordered by source and path.
        """

        with perf_span(
            "db.list_files",
            source=source.value if source is not None else "all",
            present_only=1 if present_only else 0,
        ):
            clauses: list[str] = []
            params: list[Any] = []
            if source is not None:
                clauses.append("source = ?")
                params.append(source.value)
            if present_only:
                clauses.append("present = 1")
            if review_state is not None:
                clauses.append("review_state = ?")
                params.append(review_state.value)
            where = f"WHERE {' AND '.join(clauses)}" if clauses else ""
            with self._connection() as connection:
                rows = connection.execute(
                    f"SELECT * FROM file_entry {where} ORDER BY source, path_key",
                    params,
                ).fetchall()
                overrides = CatalogDatabase._mtime_overrides_by_path(connection)
            return [
                self._to_file_record(
                    row,
                    override=overrides.get(str(row["path_key"]).casefold()),
                )
                for row in rows
            ]

    @staticmethod
    def _mtime_overrides_by_path(
        connection: sqlite3.Connection,
    ) -> dict[str, FileMtimeOverrideRow]:
        try:
            rows = connection.execute("SELECT * FROM file_mtime_override").fetchall()
        except sqlite3.OperationalError:
            return {}
        return {
            str(row["path_key"]).casefold(): _mtime_override_from_row(row)
            for row in rows
        }

    @staticmethod
    def _to_file_record(
        row: sqlite3.Row,
        *,
        override: FileMtimeOverrideRow | None = None,
    ) -> FileRecord:
        excluded = {
            "id",
            "path",
            "path_key",
            "source",
            "present",
            "review_state",
            "first_seen_run_id",
            "last_seen_run_id",
        }
        data = {key: row[key] for key in row.keys() if key not in excluded}
        apply_mtime_override_to_data(str(row["path_key"]), data, override)
        return FileRecord(
            id=row["id"],
            path=row["path"],
            path_key=row["path_key"],
            source=SourceKind(row["source"]),
            present=bool(row["present"]),
            review_state=ReviewState(row["review_state"]),
            first_seen_run_id=row["first_seen_run_id"],
            last_seen_run_id=row["last_seen_run_id"],
            data=data,
        )

    def get_ignored_path_keys(self) -> set[str]:
        """Return paths excluded from the next overlay calculation.

        Returns:
            Case-insensitive path keys whose current review state is ignored.
        """

        with self._connection() as connection:
            rows = connection.execute(
                "SELECT path_key FROM file_entry WHERE review_state = ?",
                (ReviewState.IGNORED.value,),
            ).fetchall()
        return {str(row["path_key"]) for row in rows}

    def record_review(
        self,
        file_id: int,
        action: ReviewState | str,
        *,
        comment: str | None = None,
    ) -> int:
        """Append review history and update current review state.

        Args:
            file_id: Existing file-entry identifier.
            action: ``pending``, ``acknowledged``, ``ignored``, or ``comment``.
            comment: Review note. It is mandatory for ``ignored``.

        Returns:
            New review-event identifier.

        Raises:
            ValueError: If action is invalid or ignore has no comment.
            KeyError: If ``file_id`` does not exist.
        """

        action_value = action.value if isinstance(action, ReviewState) else action
        allowed = {state.value for state in ReviewState} | {"comment"}
        if action_value not in allowed:
            raise ValueError(f"Unsupported review action: {action_value!r}")
        if action_value == ReviewState.IGNORED.value and not (comment or "").strip():
            raise ValueError("Ignoring a file requires a comment")

        with self._connection() as connection, connection:
            exists = connection.execute(
                "SELECT 1 FROM file_entry WHERE id = ?", (file_id,)
            ).fetchone()
            if not exists:
                raise KeyError(f"Unknown file id: {file_id}")
            cursor = connection.execute(
                """
                INSERT INTO review_event(file_id, action, comment, created_at)
                VALUES (?, ?, ?, ?)
                """,
                (file_id, action_value, comment, _utc_now()),
            )
            if action_value != "comment":
                connection.execute(
                    "UPDATE file_entry SET review_state = ? WHERE id = ?",
                    (action_value, file_id),
                )
            return int(cursor.lastrowid)

    def review_history(self, file_id: int) -> list[dict[str, Any]]:
        """Return chronological review and detection history for one file.

        Args:
            file_id: Existing file-entry identifier.

        Returns:
            Ordered review-event mappings.
        """

        with self._connection() as connection:
            rows = connection.execute(
                """
                SELECT id, action, comment, created_at
                FROM review_event WHERE file_id = ?
                ORDER BY id
                """,
                (file_id,),
            ).fetchall()
        return [dict(row) for row in rows]

    def current_overlay(
        self,
        file_kind: FileKind | None = None,
    ) -> list[dict[str, Any]]:
        """Read current overlay rows joined to their files.

        Args:
            file_kind: Optional PDF/MTO filter.

        Returns:
            Mapping rows suitable for later GUI models.
        """

        with perf_span("db.current_overlay"):
            where = "WHERE o.file_kind = ?" if file_kind else ""
            params = (file_kind.value,) if file_kind else ()
            with self._connection() as connection:
                rows = connection.execute(
                    f"""
                    SELECT o.*, f.path, f.path_key, f.review_state, f.present
                    FROM overlay_state o
                    JOIN file_entry f ON f.id = o.file_id
                    {where}
                    ORDER BY o.file_kind, o.document_key
                    """,
                    params,
                ).fetchall()
            return [dict(row) for row in rows]

    def list_current_collisions(self) -> list[dict[str, Any]]:
        """Return current successful collision snapshots with resolved paths.

        Returns:
            Collision mappings ordered by source, scope, and type. ``paths``
            follows ``path_keys`` order and contains only currently known paths.
        """

        with self._connection() as connection:
            rows = connection.execute(
                """
                SELECT id, scope, source, kind, message, document_key,
                       path_keys_json, scan_run_id
                FROM current_collision
                ORDER BY source, scope, kind, document_key, id
                """
            ).fetchall()
            result: list[dict[str, Any]] = []
            for row in rows:
                item = dict(row)
                path_keys = [
                    str(value)
                    for value in json.loads(item.pop("path_keys_json"))
                ]
                paths: list[str] = []
                for path_key in path_keys:
                    file_row = connection.execute(
                        """
                        SELECT path FROM file_entry
                        WHERE path_key = ? COLLATE NOCASE
                        """,
                        (path_key,),
                    ).fetchone()
                    if file_row is not None:
                        paths.append(str(file_row["path"]))
                item["path_keys"] = path_keys
                item["paths"] = paths
                result.append(item)
        return result

    @staticmethod
    def _to_parsed_file(row: sqlite3.Row) -> ParsedFile:
        transfer_name = row["transfer_name"]
        transfer_is_as_build = bool(row["transfer_is_as_build"])
        transfer = None
        if transfer_name or transfer_is_as_build:
            name = str(transfer_name or "")
            transfer = TransferMetadata(
                original_name=name,
                normalized_name=name,
                sequence=row["transfer_sequence"],
                revision=row["transfer_revision"],
                appendix=row["transfer_appendix"],
                title=row["title"],
                mark=row["mark"],
                title_system=row["title_system"],
                is_as_build=transfer_is_as_build,
                parse_status=ParseStatus.PARSED,
            )
        return ParsedFile(
            path=str(row["path"]),
            path_key=str(row["path_key"]),
            name=str(row["name"]),
            source=SourceKind(row["source"]),
            file_kind=FileKind(row["file_kind"]),
            size=int(row["size"]),
            mtime_ns=int(row["mtime_ns"]),
            parse_status=ParseStatus(row["parse_status"]),
            contract=row["contract"],
            title_system=row["title_system"],
            title=row["title"],
            mark=row["mark"],
            discipline_block=row["discipline_block"],
            core_stem=row["core_stem"],
            revision=row["revision"],
            appendix=row["appendix"],
            language=row["language"],
            extension=row["extension"],
            transfer=transfer,
            parse_error=row["parse_error"],
        )

    def reconstruct_current_rd_mto_overlay(self) -> OverlayResult:
        """Rebuild the current persisted RD MTO overlay for worker use.

        Returns:
            Overlay containing current RD MTO files without reading source files.
        """

        overlay = OverlayResult(file_kind=FileKind.MTO_XLSX)
        with self._connection() as connection:
            rows = connection.execute(
                """
                SELECT f.*, o.document_key AS overlay_document_key
                FROM overlay_state o
                JOIN file_entry f ON f.id = o.file_id
                WHERE o.file_kind = ? AND o.detected_current = 1
                ORDER BY o.document_key
                """,
                (FileKind.MTO_XLSX.value,),
            ).fetchall()
        for row in rows:
            overlay.current[str(row["overlay_document_key"])] = self._to_parsed_file(
                row
            )
        return overlay

    def list_present_robot_mto_files(self) -> list[ParsedFile]:
        """Rebuild present robot MTO records without opening XLSX files.

        Returns:
            ``ParsedFile`` robot records ordered by normalized path, including
            gracefully unparsed candidates needed for BLOCKED pairing.
        """

        with self._connection() as connection:
            rows = connection.execute(
                """
                SELECT * FROM file_entry
                WHERE source = ? AND file_kind = ? AND present = 1
                ORDER BY path_key
                """,
                (
                    SourceKind.ROBOT.value,
                    FileKind.MTO_XLSX.value,
                ),
            ).fetchall()
        return [self._to_parsed_file(row) for row in rows]

    def list_present_rd_mto_files(self) -> list[ParsedFile]:
        """Rebuild present RD MTO records without opening XLSX files.

        Returns:
            ``ParsedFile`` RD MTO rows ordered by normalized path.
        """

        with self._connection() as connection:
            rows = connection.execute(
                """
                SELECT * FROM file_entry
                WHERE source = ? AND file_kind = ? AND present = 1
                ORDER BY path_key
                """,
                (
                    SourceKind.RD.value,
                    FileKind.MTO_XLSX.value,
                ),
            ).fetchall()
        return [self._to_parsed_file(row) for row in rows]

    def record_mto_comparison(
        self,
        rd_file_id: int,
        robot_file_id: int | None,
        *,
        rd_fingerprint: str | None,
        robot_fingerprint: str | None,
        algorithm_version: int,
        status: str,
        stats: dict[str, Any] | None = None,
        diff: dict[str, Any] | None = None,
    ) -> int:
        """Upsert a future Phase 2 MTO comparison cache row.

        Args:
            rd_file_id: Current RD MTO file identifier.
            robot_file_id: Matched robot MTO identifier, if any.
            rd_fingerprint: RD semantic fingerprint.
            robot_fingerprint: Robot semantic fingerprint.
            algorithm_version: Phase 2 comparison algorithm version.
            status: Structured comparison status.
            stats: Optional comparison counters.
            diff: Optional structured comparison details.

        Returns:
            Comparison cache row identifier.
        """

        with self._connection() as connection, connection:
            connection.execute(
                """
                INSERT INTO mto_comparison(
                    rd_file_id, robot_file_id, rd_fingerprint,
                    robot_fingerprint, algorithm_version, status,
                    stats_json, diff_json, compared_at
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                ON CONFLICT DO UPDATE SET
                    rd_fingerprint = excluded.rd_fingerprint,
                    robot_fingerprint = excluded.robot_fingerprint,
                    status = excluded.status,
                    stats_json = excluded.stats_json,
                    diff_json = excluded.diff_json,
                    compared_at = excluded.compared_at
                """,
                (
                    rd_file_id,
                    robot_file_id,
                    rd_fingerprint,
                    robot_fingerprint,
                    algorithm_version,
                    status,
                    json.dumps(stats or {}, ensure_ascii=False),
                    json.dumps(diff or {}, ensure_ascii=False),
                    _utc_now(),
                ),
            )
            row = connection.execute(
                """
                SELECT id FROM mto_comparison
                WHERE rd_file_id = ?
                  AND robot_file_id IS ?
                  AND algorithm_version = ?
                """,
                (rd_file_id, robot_file_id, algorithm_version),
            ).fetchone()
            return int(row["id"])

    def get_mto_comparison(
        self,
        rd_file_id: int,
        robot_file_id: int | None,
        algorithm_version: int,
    ) -> dict[str, Any] | None:
        """Read a cached Phase 2 comparison result.

        Args:
            rd_file_id: Current RD MTO file identifier.
            robot_file_id: Matched robot MTO identifier, if any.
            algorithm_version: Phase 2 comparison algorithm version.

        Returns:
            Decoded cache mapping or ``None``.
        """

        with self._connection() as connection:
            row = connection.execute(
                """
                SELECT * FROM mto_comparison
                WHERE rd_file_id = ?
                  AND robot_file_id IS ?
                  AND algorithm_version = ?
                """,
                (rd_file_id, robot_file_id, algorithm_version),
            ).fetchone()
        if row is None:
            return None
        result = dict(row)
        result["stats"] = json.loads(result.pop("stats_json"))
        result["diff"] = json.loads(result.pop("diff_json"))
        return result

    def _file_id_for_path_key(self, path_key: str) -> int:
        with self._connection() as connection:
            row = connection.execute(
                "SELECT id FROM file_entry WHERE path_key = ? COLLATE NOCASE",
                (path_key,),
            ).fetchone()
        if row is None:
            raise KeyError(
                "MTO comparison requires files to be stored before orchestration: "
                f"{path_key}"
            )
        return int(row["id"])

    def mto_pair_cache_valid(self, pair: Any) -> bool:
        """Return True when a stored MTO comparison matches current stats.

        Args:
            pair: ``MtoPair`` from ``pair_current_mto``.

        Returns:
            True when cache row exists, algorithm version matches, RD stat
            signature matches, robot candidate signatures match, and cached
            readiness is non-empty.
        """

        from rd_catalog.mto_diff import MTO_COMPARE_ALGORITHM_VERSION

        try:
            rd_file_id = self._file_id_for_path_key(pair.rd_file.path_key)
            robot_file = (
                pair.robot_candidates[0]
                if len(pair.robot_candidates) == 1
                else None
            )
            robot_file_id = (
                self._file_id_for_path_key(robot_file.path_key)
                if robot_file is not None
                else None
            )
        except KeyError:
            return False
        rd_signature = list(pair.rd_file.stat_signature)
        robot_signatures = [
            list(candidate.stat_signature)
            for candidate in pair.robot_candidates
        ]
        cached = self.get_mto_comparison(
            rd_file_id,
            robot_file_id,
            MTO_COMPARE_ALGORITHM_VERSION,
        )
        return (
            cached is not None
            and cached["stats"].get("rd_stat_signature") == rd_signature
            and cached["stats"].get("robot_stat_signatures") == robot_signatures
            and cached["diff"].get("readiness")
        )

    @staticmethod
    def _mto_pair_in_scope(
        pair: Any,
        robot_subtree: str | Path | None,
        document_keys: set[tuple[str, str]] | None,
    ) -> bool:
        if robot_subtree is None and document_keys is None:
            return True
        if document_keys and pair.key in document_keys:
            return True
        if robot_subtree:
            return any(
                path_is_under(candidate.path, robot_subtree)
                for candidate in pair.robot_candidates
            )
        return False

    def compare_and_store_mto(
        self,
        rd_overlay: OverlayResult,
        robot_files: Sequence[ParsedFile],
        *,
        loader: RowLoader | None = None,
        is_cancelled: Callable[[], bool] | None = None,
        progress: Callable[[int, int], None] | None = None,
        robot_subtree: str | Path | None = None,
        document_keys: set[tuple[str, str]] | None = None,
    ) -> list[MtoComparisonResult]:
        """Compare current MTO pairs using stat-aware persistent caching.

        Files must already have been persisted with :meth:`store_scan`. A
        cached result is reused when the RD stat signature, the complete robot
        candidate signature list, and the comparison algorithm version match.
        Otherwise canonical fingerprints and the structured diff are rebuilt.

        Args:
            rd_overlay: Current RD MTO overlay.
            robot_files: Parsed robot-source MTO files from the same scan.
            loader: Optional test adapter matching ``mto_diff.RowLoader``.
            is_cancelled: Optional cooperative cancellation predicate checked
                between MTO pairs.
            progress: Optional callback receiving completed and total pair counts.
            robot_subtree: If set, only pairs with a robot file under this
                folder (or whose MTO key is in ``document_keys``) are stored.
            document_keys: Extra MTO keys to refresh after a scoped robot walk.

        Returns:
            Typed ``MtoComparisonResult`` objects in document-key order.

        Raises:
            KeyError: If a compared file was not stored in ``file_entry``.
        """

        from rd_catalog.mto_diff import (
            MTO_COMPARE_ALGORITHM_VERSION,
            compare_mto_pair,
            comparison_from_dict,
            comparison_to_dict,
            pair_current_mto,
        )

        pairing = pair_current_mto(rd_overlay, robot_files)
        scoped_pairs = [
            pair
            for pair in pairing.pairs
            if self._mto_pair_in_scope(pair, robot_subtree, document_keys)
        ]
        results: list[MtoComparisonResult] = []
        total = len(scoped_pairs)
        for pair in scoped_pairs:
            if is_cancelled and is_cancelled():
                break
            rd_file_id = self._file_id_for_path_key(pair.rd_file.path_key)
            robot_file = (
                pair.robot_candidates[0]
                if len(pair.robot_candidates) == 1
                else None
            )
            robot_file_id = (
                self._file_id_for_path_key(robot_file.path_key)
                if robot_file is not None
                else None
            )
            rd_signature = list(pair.rd_file.stat_signature)
            robot_signatures = [
                list(candidate.stat_signature)
                for candidate in pair.robot_candidates
            ]
            if self.mto_pair_cache_valid(pair):
                cached = self.get_mto_comparison(
                    rd_file_id,
                    robot_file_id,
                    MTO_COMPARE_ALGORITHM_VERSION,
                )
                results.append(
                    comparison_from_dict(
                        cached["diff"],
                        rd_file=pair.rd_file,
                        robot_file=robot_file,
                        cache_hit=True,
                    )
                )
                if progress:
                    progress(len(results), total)
                continue

            result = compare_mto_pair(pair, loader=loader)
            stats: dict[str, Any] = {
                **result.stats,
                "rd_stat_signature": rd_signature,
                "robot_stat_signatures": robot_signatures,
            }
            self.record_mto_comparison(
                rd_file_id,
                robot_file_id,
                rd_fingerprint=result.rd_fingerprint,
                robot_fingerprint=result.robot_fingerprint,
                algorithm_version=MTO_COMPARE_ALGORITHM_VERSION,
                status=result.readiness.value,
                stats=stats,
                diff=comparison_to_dict(result),
            )
            results.append(result)
            if progress:
                progress(len(results), total)
        return results

    def list_mto_comparisons(
        self,
        *,
        current_only: bool = True,
    ) -> list[dict[str, Any]]:
        """Return decoded MTO cache rows with both source paths for the GUI.

        Args:
            current_only: Restrict RD rows to the current MTO overlay.

        Returns:
            Comparison rows ordered by RD document identity.
        """

        current_join = (
            """
            JOIN overlay_state o
              ON o.file_id = c.rd_file_id
             AND o.file_kind = 'mto_xlsx'
             AND o.detected_current = 1
            """
            if current_only
            else ""
        )
        current_filter = (
            """
            WHERE c.id = (
                SELECT c2.id
                FROM mto_comparison c2
                WHERE c2.rd_file_id = c.rd_file_id
                ORDER BY c2.compared_at DESC, c2.id DESC
                LIMIT 1
            )
            """
            if current_only
            else ""
        )
        with self._connection() as connection:
            rows = connection.execute(
                f"""
                SELECT c.*, rd.path AS rd_path, rd.path_key AS rd_path_key,
                       rd.name AS rd_name, rd.revision AS rd_revision,
                       rd.appendix AS rd_appendix,
                       rd.transfer_revision AS rd_transfer_revision,
                       rd.transfer_appendix AS rd_transfer_appendix,
                       rd.transfer_is_as_build AS rd_is_as_build,
                       rd.review_state AS rd_review_state,
                       rd.present AS rd_present, rd.mtime_ns AS rd_mtime_ns,
                       robot.path AS robot_path,
                       robot.path_key AS robot_path_key,
                       robot.name AS robot_name,
                       robot.revision AS robot_revision,
                       robot.appendix AS robot_appendix,
                       robot.present AS robot_present,
                       robot.mtime_ns AS robot_mtime_ns,
                       rd.title_system, rd.title, rd.mark, rd.discipline_block
                FROM mto_comparison c
                JOIN file_entry rd ON rd.id = c.rd_file_id
                LEFT JOIN file_entry robot ON robot.id = c.robot_file_id
                {current_join}
                {current_filter}
                ORDER BY rd.title_system, rd.discipline_block, c.id DESC
                """
            ).fetchall()
        result: list[dict[str, Any]] = []
        for row in rows:
            item = dict(row)
            item["stats"] = json.loads(item.pop("stats_json"))
            item["diff"] = json.loads(item.pop("diff_json"))
            result.append(item)
        return result

    def get_mto_pair_comparison(
        self,
        left_file_id: int,
        right_file_id: int,
        algorithm_version: int,
    ) -> dict[str, Any] | None:
        """Read one symmetric MTO pair-comparison cache row.

        Pair ids are ordered with :func:`canonical_mto_pair_ids` so ``(A, B)``
        and ``(B, A)`` resolve to the same row.

        Args:
            left_file_id: One ``file_entry.id`` (any order).
            right_file_id: The other ``file_entry.id``.
            algorithm_version: Comparison algorithm version.

        Returns:
            Decoded cache mapping, or ``None`` when missing.
        """

        left_id, right_id = canonical_mto_pair_ids(left_file_id, right_file_id)
        with self._connection() as connection:
            row = connection.execute(
                """
                SELECT * FROM mto_pair_comparison
                WHERE left_file_id = ?
                  AND right_file_id = ?
                  AND algorithm_version = ?
                """,
                (left_id, right_id, int(algorithm_version)),
            ).fetchone()
        if row is None:
            return None
        return self._decode_mto_pair_row(row)

    def upsert_mto_pair_comparison(
        self,
        left_file_id: int,
        right_file_id: int,
        *,
        algorithm_version: int,
        content_status: str,
        left_fingerprint: str | None,
        right_fingerprint: str | None,
        left_signature: Any,
        right_signature: Any,
        stats: dict[str, Any] | None = None,
        diff: dict[str, Any] | None = None,
        error: str | None = None,
    ) -> int:
        """Insert or replace one symmetric MTO pair-comparison row.

        Ids, fingerprints, and signatures are stored in canonical file-id
        order. Existing ``mto_comparison`` rows are not touched.

        Args:
            left_file_id: One ``file_entry.id`` (any order).
            right_file_id: The other ``file_entry.id``.
            algorithm_version: Comparison algorithm version.
            content_status: ``content_equal``, ``content_diff``, or
                ``not_compared``.
            left_fingerprint: Semantic fingerprint of the first argument file.
            right_fingerprint: Semantic fingerprint of the second argument.
            left_signature: Current ``(path_key, size, mtime_ns)`` of the
                first argument file.
            right_signature: Current signature of the second argument file.
            stats: Optional counters (added / removed / changed / rows).
            diff: Optional structured diff payload.
            error: Loader or compare error text; ``None`` when the compare
                ran.

        Returns:
            Cache row identifier.

        Raises:
            sqlite3.IntegrityError: If either file id is missing from
                ``file_entry``.
        """

        (
            left_id,
            right_id,
            left_fp,
            right_fp,
            left_sig,
            right_sig,
        ) = self._ordered_pair_fields(
            left_file_id,
            right_file_id,
            left_fingerprint=left_fingerprint,
            right_fingerprint=right_fingerprint,
            left_signature=left_signature,
            right_signature=right_signature,
        )
        with self._connection() as connection, connection:
            connection.execute(
                """
                INSERT INTO mto_pair_comparison(
                    left_file_id, right_file_id, algorithm_version,
                    content_status, left_fingerprint, right_fingerprint,
                    left_signature, right_signature, stats_json, diff_json,
                    error, compared_at
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                ON CONFLICT(left_file_id, right_file_id, algorithm_version)
                DO UPDATE SET
                    content_status = excluded.content_status,
                    left_fingerprint = excluded.left_fingerprint,
                    right_fingerprint = excluded.right_fingerprint,
                    left_signature = excluded.left_signature,
                    right_signature = excluded.right_signature,
                    stats_json = excluded.stats_json,
                    diff_json = excluded.diff_json,
                    error = excluded.error,
                    compared_at = excluded.compared_at
                """,
                (
                    left_id,
                    right_id,
                    int(algorithm_version),
                    str(content_status),
                    left_fp,
                    right_fp,
                    _dump_signature(left_sig),
                    _dump_signature(right_sig),
                    json.dumps(stats or {}, ensure_ascii=False),
                    json.dumps(diff or {}, ensure_ascii=False),
                    error,
                    _utc_now(),
                ),
            )
            row = connection.execute(
                """
                SELECT id FROM mto_pair_comparison
                WHERE left_file_id = ?
                  AND right_file_id = ?
                  AND algorithm_version = ?
                """,
                (left_id, right_id, int(algorithm_version)),
            ).fetchone()
            return int(row["id"])

    def mto_pair_comparison_valid(
        self,
        *,
        left: FileRecord | ParsedFile,
        right: FileRecord | ParsedFile,
        algorithm_version: int | None = None,
    ) -> bool:
        """Return True when a stored material verdict matches both signatures.

        Cache is valid when both ``file_entry`` rows still exist, both stored
        signatures equal the current ``(path_key, size, mtime_ns)`` of
        ``left`` and ``right``, and ``content_status`` is ``content_equal``
        or ``content_diff``. A stored ``not_compared`` row is history, not
        cache: Excel locks and other read errors must be retried. Pair order
        does not matter.

        Args:
            left: One file of the pair.
            right: The other file.
            algorithm_version: Comparison algorithm; defaults to
                ``MTO_COMPARE_ALGORITHM_VERSION``.

        Returns:
            True when the cache row may be reused as a material verdict.
        """

        from rd_catalog.mto_diff import MTO_COMPARE_ALGORITHM_VERSION

        version = (
            MTO_COMPARE_ALGORITHM_VERSION
            if algorithm_version is None
            else int(algorithm_version)
        )
        left_identity = self._pair_file_identity(left)
        right_identity = self._pair_file_identity(right)
        if left_identity is None or right_identity is None:
            return False
        left_id, left_sig = left_identity
        right_id, right_sig = right_identity
        if not self._file_ids_exist(left_id, right_id):
            return False
        cached = self.get_mto_pair_comparison(left_id, right_id, version)
        if cached is None:
            return False
        status = str(cached.get("content_status") or "")
        if status not in {
            MtoContentStatus.EQUAL.value,
            MtoContentStatus.DIFF.value,
        }:
            return False
        ordered = self._ordered_pair_fields(
            left_id,
            right_id,
            left_fingerprint=None,
            right_fingerprint=None,
            left_signature=left_sig,
            right_signature=right_sig,
        )
        return _signatures_equal(
            cached["left_signature"], ordered[4]
        ) and _signatures_equal(cached["right_signature"], ordered[5])

    def list_mto_pair_comparisons(
        self,
        pairs: Sequence[tuple[int, int]],
        *,
        algorithm_version: int | None = None,
    ) -> dict[tuple[int, int], dict[str, Any]]:
        """Return stored pair verdicts for many id-pairs in one query.

        Intended for a GUI repaint of ~230 pairs. Incoming ``(A, B)`` and
        ``(B, A)`` collapse to one canonical key ``(min_id, max_id)``.
        Signature validity is not applied here; call
        :meth:`mto_pair_comparison_valid` or compare signatures in memory.

        Args:
            pairs: File-id pairs in any order. Duplicates are ignored.
            algorithm_version: Comparison algorithm; defaults to
                ``MTO_COMPARE_ALGORITHM_VERSION``.

        Returns:
            Canonical ``(left_file_id, right_file_id)`` → decoded row.
        """

        from rd_catalog.mto_diff import MTO_COMPARE_ALGORITHM_VERSION

        version = (
            MTO_COMPARE_ALGORITHM_VERSION
            if algorithm_version is None
            else int(algorithm_version)
        )
        canonical: list[tuple[int, int]] = []
        seen: set[tuple[int, int]] = set()
        for left, right in pairs:
            key = canonical_mto_pair_ids(int(left), int(right))
            if key in seen:
                continue
            seen.add(key)
            canonical.append(key)
        if not canonical:
            return {}
        result: dict[tuple[int, int], dict[str, Any]] = {}
        with self._connection() as connection:
            for offset in range(0, len(canonical), _MTO_PAIR_LIST_CHUNK):
                chunk = canonical[offset : offset + _MTO_PAIR_LIST_CHUNK]
                selects = " UNION ALL ".join(
                    ["SELECT ? AS left_file_id, ? AS right_file_id"] * len(chunk)
                )
                params: list[Any] = []
                for left_id, right_id in chunk:
                    params.extend([left_id, right_id])
                params.append(version)
                rows = connection.execute(
                    f"""
                    SELECT c.* FROM mto_pair_comparison c
                    INNER JOIN ({selects}) AS wanted
                      ON c.left_file_id = wanted.left_file_id
                     AND c.right_file_id = wanted.right_file_id
                    WHERE c.algorithm_version = ?
                    """,
                    params,
                ).fetchall()
                for row in rows:
                    item = self._decode_mto_pair_row(row)
                    result[
                        (int(item["left_file_id"]), int(item["right_file_id"]))
                    ] = item
        return result

    @staticmethod
    def _decode_mto_pair_row(row: sqlite3.Row) -> dict[str, Any]:
        item = dict(row)
        item["stats"] = json.loads(item.pop("stats_json") or "{}")
        item["diff"] = json.loads(item.pop("diff_json") or "{}")
        item["left_signature"] = _signature_mapping(item.get("left_signature"))
        item["right_signature"] = _signature_mapping(item.get("right_signature"))
        return item

    @staticmethod
    def _ordered_pair_fields(
        left_file_id: int,
        right_file_id: int,
        *,
        left_fingerprint: str | None,
        right_fingerprint: str | None,
        left_signature: Any,
        right_signature: Any,
    ) -> tuple[int, int, str | None, str | None, dict[str, Any], dict[str, Any]]:
        left_id = int(left_file_id)
        right_id = int(right_file_id)
        left_fp = left_fingerprint
        right_fp = right_fingerprint
        left_sig = _signature_mapping(left_signature)
        right_sig = _signature_mapping(right_signature)
        if left_id > right_id:
            return right_id, left_id, right_fp, left_fp, right_sig, left_sig
        return left_id, right_id, left_fp, right_fp, left_sig, right_sig

    def _pair_file_identity(
        self, file: FileRecord | ParsedFile
    ) -> tuple[int, dict[str, Any]] | None:
        signature = mto_file_stat_signature(file)
        if isinstance(file, FileRecord):
            return int(file.id), signature
        try:
            return self._file_id_for_path_key(file.path_key), signature
        except KeyError:
            return None

    def _file_ids_exist(self, left_file_id: int, right_file_id: int) -> bool:
        with self._connection() as connection:
            rows = connection.execute(
                "SELECT id FROM file_entry WHERE id IN (?, ?)",
                (int(left_file_id), int(right_file_id)),
            ).fetchall()
        found = {int(row["id"]) for row in rows}
        needed = {int(left_file_id), int(right_file_id)}
        return found == needed

    def list_robot_extra_mto(self) -> list[dict[str, Any]]:
        """Return present parsed robot MTO files lacking a current RD key.

        Returns:
            GUI-ready WARNING rows with ``robot_extra`` issue payloads.
        """

        with self._connection() as connection:
            rows = connection.execute(
                """
                SELECT robot.*
                FROM file_entry robot
                WHERE robot.source = ?
                  AND robot.file_kind = ?
                  AND robot.present = 1
                  AND robot.parse_status = ?
                  AND robot.title_system IS NOT NULL
                  AND robot.discipline_block IS NOT NULL
                  AND NOT EXISTS (
                      SELECT 1
                      FROM overlay_state o
                      JOIN file_entry rd ON rd.id = o.file_id
                      WHERE o.file_kind = ?
                        AND o.detected_current = 1
                        AND rd.title_system = robot.title_system COLLATE NOCASE
                        AND rd.discipline_block =
                            robot.discipline_block COLLATE NOCASE
                  )
                ORDER BY robot.title_system, robot.discipline_block, robot.path_key
                """,
                (
                    SourceKind.ROBOT.value,
                    FileKind.MTO_XLSX.value,
                    ParseStatus.PARSED.value,
                    FileKind.MTO_XLSX.value,
                ),
            ).fetchall()
        return [
            {
                "rd_file_id": None,
                "robot_file_id": int(row["id"]),
                "status": "warning",
                "title_system": row["title_system"],
                "title": row["title"],
                "mark": row["mark"],
                "discipline_block": row["discipline_block"],
                "rd_path": None,
                "rd_present": False,
                "robot_path": row["path"],
                "robot_path_key": row["path_key"],
                "robot_name": row["name"],
                "robot_revision": row["revision"],
                "robot_appendix": row["appendix"],
                "robot_present": True,
                "robot_mtime_ns": row["mtime_ns"],
                "diff": {
                    "content_status": "not_compared",
                    "issues": [MtoIssueKind.ROBOT_EXTRA.value],
                    "error": "MTO робота не имеет текущей пары в РД",
                },
                "stats": {},
            }
            for row in rows
        ]

    def last_scan_info(self, *, successful_only: bool = False) -> dict[str, Any] | None:
        """Return the latest persisted scan metadata for the GUI.

        Args:
            successful_only: Restrict the lookup to complete successful scans.

        Returns:
            A scan-run mapping or ``None`` when no scan has been stored.
        """

        where = "WHERE status = ?" if successful_only else ""
        params = (ScanRunStatus.SUCCESS.value,) if successful_only else ()
        with self._connection() as connection:
            row = connection.execute(
                f"""
                SELECT id, started_at, completed_at, status, is_baseline,
                       error_json
                FROM scan_run
                {where}
                ORDER BY id DESC
                LIMIT 1
                """,
                params,
            ).fetchone()
        if row is None:
            return None
        result = dict(row)
        result["is_baseline"] = bool(result["is_baseline"])
        result["errors"] = json.loads(result.pop("error_json"))
        return result

    @staticmethod
    def _kit_filter_sql(
        title: str | None,
        mark: str | None,
    ) -> tuple[str, list[Any]]:
        clauses: list[str] = []
        params: list[Any] = []
        if title is not None:
            clauses.append("title = ? COLLATE NOCASE")
            params.append(title)
        if mark is not None:
            clauses.append("mark = ? COLLATE NOCASE")
            params.append(mark)
        where = f"WHERE {' AND '.join(clauses)}" if clauses else ""
        return where, params

    @staticmethod
    def _event_from_row(row: sqlite3.Row) -> KitEvent:
        raw = str(row["raw"] or "")
        if raw.strip():
            return parse_history_line(raw)
        transmittals = tuple(json.loads(row["transmittals_json"] or "[]"))
        return KitEvent(
            raw=raw,
            date=row["event_date"],
            stage=str(row["stage"] or "other"),
            stage_label=str(row["stage_label"] or ""),
            revision=row["revision"],
            appendix=row["appendix"],
            transmittals=transmittals,
            parsed=bool(row["parsed"]),
        )

    @staticmethod
    def _last_event_of(events: Sequence[KitEvent]) -> KitEvent | None:
        dated = [event for event in events if event.date]
        if dated:
            return dated[-1]
        return events[-1] if events else None

    @staticmethod
    def _google_kit_from_parts(
        row: sqlite3.Row,
        events: Sequence[KitEvent],
    ) -> GoogleKit:
        event_tuple = tuple(events)
        return GoogleKit(
            title=str(row["title"]),
            mark=str(row["mark"]),
            mark_raw=str(row["mark_raw"] or ""),
            title_system=str(row["title_system"] or ""),
            sheet_revision=row["sheet_revision"],
            sheet_appendix=row["sheet_appendix"],
            sheet_revision_text=str(row["sheet_revision_text"] or ""),
            status_sheet=str(row["status_sheet"] or ""),
            comment_raw=str(row["comment_raw"] or ""),
            events=event_tuple,
            last_event=CatalogDatabase._last_event_of(event_tuple),
            row_index=int(row["row_index"] or 0),
        )

    @staticmethod
    def _issuance_from_row(row: sqlite3.Row) -> IssuanceKit:
        return IssuanceKit(
            title=str(row["title"]),
            mark=str(row["mark"]),
            mark_raw=str(row["mark_raw"] or ""),
            title_system=str(row["title_system"] or ""),
            revision=row["revision"],
            appendix=row["appendix"],
            revision_text=str(row["revision_text"] or ""),
            status=str(row["status"] or ""),
            send_date=str(row["send_date"] or ""),
            send_date_sortable=str(row["send_date_sortable"] or ""),
            send_transmittal=str(row["send_transmittal"] or ""),
            incoming_control_date=str(row["incoming_control_date"] or ""),
            incoming_control_date_sortable=str(
                row["incoming_control_date_sortable"] or ""
            ),
            confirm_transmittal=str(row["confirm_transmittal"] or ""),
            note_raw=str(row["note_raw"] or ""),
            row_index=int(row["row_index"] or 0),
        )

    @staticmethod
    def _package_from_row(row: sqlite3.Row) -> KitPackageRow:
        return KitPackageRow(
            id=int(row["id"]),
            title=str(row["title"]),
            mark=str(row["mark"]),
            source=str(row["source"]),
            sequence=row["sequence"],
            transfer_name=row["transfer_name"],
            package_path=str(row["package_path"] or ""),
            revision_text=str(row["revision_text"] or ""),
            max_mtime_ns=row["max_mtime_ns"],
            pdf_count=int(row["pdf_count"] or 0),
            editable_count=int(row["editable_count"] or 0),
            mto_revision_text=str(row["mto_revision_text"] or ""),
            overlay_current_count=int(row["overlay_current_count"] or 0),
            is_grey=bool(row["is_grey"]),
            is_current=bool(row["is_current"]),
            is_as_build=bool(row["is_as_build"]) if "is_as_build" in row.keys() else False,
        )

    @staticmethod
    def _cycle_from_row(row: sqlite3.Row) -> KitCycleRow:
        return KitCycleRow(
            id=int(row["id"]),
            title=str(row["title"]),
            mark=str(row["mark"]),
            revision_text=str(row["revision_text"] or ""),
            match_reason=str(row["match_reason"] or ""),
            send_id=row["send_id"],
            package_id=row["package_id"],
            tdo_event_id=row["tdo_event_id"],
            code_event_id=row["code_event_id"],
        )

    @staticmethod
    def _pipeline_from_row(row: sqlite3.Row) -> KitPipelineRow:
        return KitPipelineRow(
            title=str(row["title"]),
            mark=str(row["mark"]),
            status=str(row["status"] or ""),
            code=row["code"],
            code_origin=row["code_origin"],
            working_revision_text=str(row["working_revision_text"] or ""),
            official_revision_text=str(row["official_revision_text"] or ""),
            suspicious=bool(row["suspicious"]),
            algorithm_version=int(row["algorithm_version"] or 1),
            code_stale=bool(row["code_stale"]),
            code_revision_text=str(row["code_revision_text"] or ""),
            code_date=str(row["code_date"] or ""),
            tdo_date=(
                str(row["tdo_date"] or "") if "tdo_date" in row.keys() else ""
            ),
            review_as_build=bool(row["review_as_build"]),
            working_as_build=bool(
                row["working_as_build"]
            )
            if "working_as_build" in row.keys()
            else False,
            working_transfer_names=_json_str_tuple(
                row["working_transfer_names_json"]
                if "working_transfer_names_json" in row.keys()
                else "[]"
            ),
            working_sequences=_json_int_tuple(
                row["working_sequences_json"]
                if "working_sequences_json" in row.keys()
                else "[]"
            ),
            annulled_transfer_names=_json_str_tuple(
                row["annulled_transfer_names_json"]
                if "annulled_transfer_names_json" in row.keys()
                else "[]"
            ),
            annulled_sequences=_json_int_tuple(
                row["annulled_sequences_json"]
                if "annulled_sequences_json" in row.keys()
                else "[]"
            ),
        )

    @staticmethod
    def _revision_from_row(row: sqlite3.Row) -> KitRevisionRow:
        return KitRevisionRow(
            id=int(row["id"]),
            title=str(row["title"]),
            mark=str(row["mark"]),
            revision_text=str(row["revision_text"] or ""),
            pipeline_status=str(row["pipeline_status"] or ""),
            letters=str(row["letters"] or ""),
            is_as_build=bool(row["is_as_build"]),
            is_current=bool(row["is_current"]),
            is_current_ifc=bool(row["is_current_ifc"]),
            has_mto=bool(row["has_mto"]),
            problem_kinds_json=str(row["problem_kinds_json"] or "[]"),
            package_ids_json=str(row["package_ids_json"] or "[]"),
            algorithm_version=int(row["algorithm_version"] or 1),
        )

    @staticmethod
    def _liquidity_from_row(row: sqlite3.Row) -> LiquidityReviewRow:
        return LiquidityReviewRow(
            id=int(row["id"]),
            title=str(row["title"]),
            mark=str(row["mark"]),
            transfer_name=str(row["transfer_name"] or ""),
            revision_text=str(row["revision_text"] or ""),
            decision=str(row["decision"]),
            sequence=row["sequence"],
            comment=row["comment"],
            evidence_mtime_ns=row["evidence_mtime_ns"],
            decided_at=str(row["decided_at"] or ""),
        )

    @staticmethod
    def _working_flag_from_row(row: sqlite3.Row) -> KitWorkingFlagRow:
        return KitWorkingFlagRow(
            id=int(row["id"]),
            title=str(row["title"]),
            mark=str(row["mark"]),
            revision_text=str(row["revision_text"] or ""),
            transfer_name=str(row["transfer_name"] or ""),
            sequence=(
                int(row["sequence"])
                if "sequence" in row.keys() and row["sequence"] is not None
                else None
            ),
            decided_at=str(row["decided_at"] or ""),
        )

    @staticmethod
    def _annulled_flag_from_row(row: sqlite3.Row) -> KitAnnulledFlagRow:
        return KitAnnulledFlagRow(
            id=int(row["id"]),
            title=str(row["title"]),
            mark=str(row["mark"]),
            revision_text=str(row["revision_text"] or ""),
            transfer_name=str(row["transfer_name"] or ""),
            sequence=(
                int(row["sequence"])
                if "sequence" in row.keys() and row["sequence"] is not None
                else None
            ),
            decided_at=str(row["decided_at"] or ""),
        )

    @staticmethod
    def _issuance_review_from_row(row: sqlite3.Row) -> IssuanceReviewRow:
        return IssuanceReviewRow(
            id=int(row["id"]),
            title=str(row["title"]),
            mark=str(row["mark"]),
            kind=str(row["kind"]),
            source=str(row["source"]),
            decision=str(row["decision"] or ""),
            revision_text=str(row["revision_text"] or ""),
            send_date=str(row["send_date"] or ""),
            send_date_sortable=str(row["send_date_sortable"] or ""),
            send_transmittal=str(row["send_transmittal"] or ""),
            incoming_control_date=str(row["incoming_control_date"] or ""),
            incoming_control_date_sortable=str(
                row["incoming_control_date_sortable"] or ""
            ),
            confirm_transmittal=str(row["confirm_transmittal"] or ""),
            sheet_status=str(row["sheet_status"] or ""),
            note=str(row["note"] or ""),
            comment=row["comment"],
            identity_fingerprint=str(row["identity_fingerprint"]),
            evidence_fingerprint=str(row["evidence_fingerprint"] or ""),
            match_state=str(row["match_state"] or ""),
            source_path=str(row["source_path"] or ""),
            path_key=str(row["path_key"] or ""),
            decided_at=str(row["decided_at"] or ""),
        )

    def replace_google_snapshot(
        self,
        kits: Sequence[GoogleKit],
        sends: Sequence[IssuanceKit],
        *,
        loaded_at: str,
        source: str,
        warning: str | None = None,
    ) -> None:
        """Replace Google kit, event, and issuance-send rows in one transaction.

        Does not delete ``kit_liquidity_review``, ``issuance_review``,
        ``kit_working_flag``, ``kit_annulled_flag``, ``file_mtime_override``,
        ``kit_package``,
        ``kit_pipeline``, ``kit_cycle``, ``file_entry``,
        ``an_mto_file``, or ``rd_dump_mto_file``.
        Duplicate
        ``(title, mark)`` kits are collapsed with last-wins using
        :func:`kit_identity_key`. All issuance sends are stored.

        Args:
            kits: Parsed KSB ИД kits (events are inserted in order).
            sends: Every parsed «Выдача РД ПД» row, not latest-wins.
            loaded_at: ISO timestamp of this snapshot load.
            source: Load origin (``network``, ``cache``, test label, …).
            warning: Optional load warning stored on ``google_load``.
        """

        unique_kits: dict[tuple[str, str], GoogleKit] = {}
        for kit in kits:
            unique_kits[kit_identity_key(kit.title, kit.mark)] = kit
        ordered_kits = tuple(
            sorted(unique_kits.values(), key=lambda kit: (kit.title, kit.mark))
        )
        with self._connection() as connection, connection:
            connection.execute("DELETE FROM google_event")
            connection.execute("DELETE FROM google_kit")
            connection.execute("DELETE FROM issuance_send")
            connection.execute("DELETE FROM google_load")
            for kit in ordered_kits:
                cursor = connection.execute(
                    """
                    INSERT INTO google_kit(
                        title, mark, mark_raw, title_system, sheet_revision,
                        sheet_appendix, sheet_revision_text, status_sheet,
                        comment_raw, row_index
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    (
                        kit.title,
                        kit.mark,
                        kit.mark_raw,
                        kit.title_system,
                        kit.sheet_revision,
                        kit.sheet_appendix,
                        kit.sheet_revision_text,
                        kit.status_sheet,
                        kit.comment_raw,
                        kit.row_index,
                    ),
                )
                kit_id = int(cursor.lastrowid)
                for seq, event in enumerate(kit.events):
                    connection.execute(
                        """
                        INSERT INTO google_event(
                            kit_id, seq, raw, event_date, stage, stage_label,
                            revision, appendix, transmittals_json, parsed
                        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                        """,
                        (
                            kit_id,
                            seq,
                            event.raw,
                            event.date,
                            event.stage,
                            event.stage_label,
                            event.revision,
                            event.appendix,
                            json.dumps(list(event.transmittals), ensure_ascii=False),
                            int(bool(event.parsed)),
                        ),
                    )
            for send in sends:
                connection.execute(
                    """
                    INSERT INTO issuance_send(
                        title, mark, mark_raw, title_system, revision, appendix,
                        revision_text, status, send_date, send_date_sortable,
                        send_transmittal, incoming_control_date,
                        incoming_control_date_sortable, confirm_transmittal,
                        note_raw, row_index
                    ) VALUES (
                        ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?
                    )
                    """,
                    (
                        send.title,
                        send.mark,
                        send.mark_raw,
                        send.title_system,
                        send.revision,
                        send.appendix,
                        send.revision_text,
                        send.status,
                        send.send_date,
                        send.send_date_sortable,
                        send.send_transmittal,
                        send.incoming_control_date,
                        send.incoming_control_date_sortable,
                        send.confirm_transmittal,
                        send.note_raw,
                        send.row_index,
                    ),
                )
            connection.execute(
                """
                INSERT INTO google_load(id, loaded_at, source, warning)
                VALUES (1, ?, ?, ?)
                """,
                (loaded_at, source, warning),
            )

    def google_load_info(self) -> dict[str, Any] | None:
        """Return Google snapshot load metadata, if any.

        Returns:
            Mapping with ``loaded_at``, ``source``, and ``warning``, or ``None``.
        """

        with self._connection() as connection:
            row = connection.execute(
                "SELECT loaded_at, source, warning FROM google_load WHERE id = 1"
            ).fetchone()
        return dict(row) if row is not None else None

    def list_google_kits(self) -> list[GoogleKit]:
        """Return persisted KSB ИД kits with events in sheet order.

        Returns:
            Kits ordered by title and mark.
        """

        with self._connection() as connection:
            kit_rows = connection.execute(
                "SELECT * FROM google_kit ORDER BY title, mark, id"
            ).fetchall()
            event_rows = connection.execute(
                "SELECT * FROM google_event ORDER BY kit_id, seq"
            ).fetchall()
        events_by_kit: dict[int, list[KitEvent]] = {}
        for row in event_rows:
            events_by_kit.setdefault(int(row["kit_id"]), []).append(
                self._event_from_row(row)
            )
        return [
            self._google_kit_from_parts(
                row,
                events_by_kit.get(int(row["id"]), ()),
            )
            for row in kit_rows
        ]

    def list_google_events_with_ids(
        self, title: str, mark: str
    ) -> list[tuple[int, KitEvent]]:
        """Return column-F events with SQLite ids for one kit identity.

        Args:
            title: Four-digit title.
            mark: Latin AGCC mark.

        Returns:
            ``(google_event.id, event)`` pairs ordered by ``seq``; empty when
            the kit is absent.
        """

        with self._connection() as connection:
            kit_row = connection.execute(
                """
                SELECT id FROM google_kit
                WHERE title = ? COLLATE NOCASE AND mark = ? COLLATE NOCASE
                """,
                (title, mark),
            ).fetchone()
            if kit_row is None:
                return []
            rows = connection.execute(
                """
                SELECT * FROM google_event
                WHERE kit_id = ?
                ORDER BY seq
                """,
                (int(kit_row["id"]),),
            ).fetchall()
        return [(int(row["id"]), self._event_from_row(row)) for row in rows]

    def list_google_events(self, title: str, mark: str) -> list[KitEvent]:
        """Return column-F events for one kit identity.

        Args:
            title: Four-digit title.
            mark: Latin AGCC mark.

        Returns:
            Events ordered by ``seq``; empty when the kit is absent.
        """

        return [event for _event_id, event in self.list_google_events_with_ids(title, mark)]

    def list_google_events_by_kit(
        self,
    ) -> dict[tuple[str, str], list[tuple[int, KitEvent]]]:
        """Return all column-F events with ids, grouped by kit identity.

        Returns:
            ``kit_identity_key`` → ``(google_event.id, event)`` pairs in
            sheet ``seq`` order.
        """

        with self._connection() as connection:
            rows = connection.execute(
                """
                SELECT e.*, k.title AS kit_title, k.mark AS kit_mark
                FROM google_event AS e
                JOIN google_kit AS k ON k.id = e.kit_id
                ORDER BY k.title, k.mark, e.seq
                """
            ).fetchall()
        grouped: dict[tuple[str, str], list[tuple[int, KitEvent]]] = {}
        for row in rows:
            key = kit_identity_key(str(row["kit_title"]), str(row["kit_mark"]))
            grouped.setdefault(key, []).append(
                (int(row["id"]), self._event_from_row(row))
            )
        return grouped

    def list_issuance_sends_with_ids(
        self,
        title: str | None = None,
        mark: str | None = None,
    ) -> list[tuple[int, IssuanceKit]]:
        """Return «Выдача РД ПД» rows with SQLite ids (every send).

        Args:
            title: Optional title filter (case-insensitive).
            mark: Optional mark filter (case-insensitive).

        Returns:
            ``(issuance_send.id, send)`` pairs ordered by sortable date, then
            sheet row.
        """

        where, params = self._kit_filter_sql(title, mark)
        with self._connection() as connection:
            rows = connection.execute(
                f"""
                SELECT * FROM issuance_send
                {where}
                ORDER BY send_date_sortable, row_index, id
                """,
                params,
            ).fetchall()
        return [(int(row["id"]), self._issuance_from_row(row)) for row in rows]

    def list_issuance_sends(
        self,
        title: str | None = None,
        mark: str | None = None,
    ) -> list[IssuanceKit]:
        """Return persisted «Выдача РД ПД» rows (every send, not latest-wins).

        Args:
            title: Optional title filter (case-insensitive).
            mark: Optional mark filter (case-insensitive).

        Returns:
            Sends ordered by sortable date, then sheet row.
        """

        return [
            send for _send_id, send in self.list_issuance_sends_with_ids(title, mark)
        ]

    @staticmethod
    def _delete_kit_identity_rows(
        connection: sqlite3.Connection,
        table: str,
        identities: Collection[tuple[str, str]],
    ) -> None:
        """Delete rows whose ``(title, mark)`` is in ``identities``.

        Args:
            connection: Open SQLite connection (caller owns the transaction).
            table: One of ``kit_cycle``, ``kit_package``, ``kit_pipeline``,
                ``kit_revision_cell``.
            identities: Kit keys to delete. Empty is a no-op.

        Raises:
            ValueError: ``table`` is not a known kit derived table.
        """

        allowed = {
            "kit_cycle",
            "kit_package",
            "kit_pipeline",
            "kit_revision_cell",
        }
        if table not in allowed:
            raise ValueError(f"unknown kit table: {table}")
        if not identities:
            return
        connection.executemany(
            f"DELETE FROM {table} "
            "WHERE title = ? COLLATE NOCASE AND mark = ? COLLATE NOCASE",
            list(identities),
        )

    def replace_kit_derived(
        self,
        packages: Sequence[KitPackageRow],
        cycles: Sequence[KitCycleRow],
        pipelines: Sequence[KitPipelineRow],
        *,
        identities: Collection[tuple[str, str]] | None = None,
    ) -> None:
        """Replace ``kit_package``, ``kit_cycle``, and ``kit_pipeline`` only.

        Does not touch ``kit_liquidity_review``, ``issuance_review``,
        ``kit_working_flag``, ``kit_annulled_flag``, ``file_mtime_override``,
        Google snapshot
        tables, ``file_entry``, ``kit_revision_cell``, ``an_mto_file``,
        or ``rd_dump_mto_file``.
        ``KitCycleRow.package_id``
        values are remapped from matching :attr:`KitPackageRow.id` in
        ``packages``. Duplicate pipeline identities keep the last row via
        :func:`kit_identity_key`.

        Args:
            packages: Issued folders and grey stubs.
            cycles: Official send cycles.
            pipelines: Current derived status per kit.
            identities: When set, delete and reinsert only these kit keys.
                ``None`` replaces every derived row. Caller must pass rows
                only for those keys.
        """

        unique_pipelines: dict[tuple[str, str], KitPipelineRow] = {}
        for pipeline in pipelines:
            unique_pipelines[kit_identity_key(pipeline.title, pipeline.mark)] = (
                pipeline
            )
        with self._connection() as connection, connection:
            if identities is None:
                connection.execute("DELETE FROM kit_cycle")
                connection.execute("DELETE FROM kit_package")
                connection.execute("DELETE FROM kit_pipeline")
            else:
                scoped = {
                    kit_identity_key(title, mark) for title, mark in identities
                }
                self._delete_kit_identity_rows(connection, "kit_cycle", scoped)
                self._delete_kit_identity_rows(connection, "kit_package", scoped)
                self._delete_kit_identity_rows(connection, "kit_pipeline", scoped)
            id_map: dict[int, int] = {}
            for package in packages:
                cursor = connection.execute(
                    """
                    INSERT INTO kit_package(
                        title, mark, source, sequence, transfer_name,
                        package_path, revision_text, max_mtime_ns, pdf_count,
                        editable_count, mto_revision_text, overlay_current_count,
                        is_grey, is_current, is_as_build
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    (
                        package.title,
                        package.mark,
                        package.source,
                        package.sequence,
                        package.transfer_name,
                        package.package_path,
                        package.revision_text,
                        package.max_mtime_ns,
                        package.pdf_count,
                        package.editable_count,
                        package.mto_revision_text,
                        package.overlay_current_count,
                        int(bool(package.is_grey)),
                        int(bool(package.is_current)),
                        int(bool(package.is_as_build)),
                    ),
                )
                new_id = int(cursor.lastrowid)
                if package.id is not None:
                    id_map[package.id] = new_id
            for cycle in cycles:
                package_id = (
                    id_map.get(cycle.package_id)
                    if cycle.package_id is not None
                    else None
                )
                connection.execute(
                    """
                    INSERT INTO kit_cycle(
                        title, mark, send_id, package_id, tdo_event_id,
                        code_event_id, revision_text, match_reason
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    (
                        cycle.title,
                        cycle.mark,
                        cycle.send_id,
                        package_id,
                        cycle.tdo_event_id,
                        cycle.code_event_id,
                        cycle.revision_text,
                        cycle.match_reason,
                    ),
                )
            for pipeline in unique_pipelines.values():
                connection.execute(
                    """
                    INSERT INTO kit_pipeline(
                        title, mark, status, code, code_origin,
                        working_revision_text, official_revision_text,
                        suspicious, algorithm_version, code_stale,
                        code_revision_text, code_date, tdo_date,
                        review_as_build, working_as_build,
                        working_transfer_names_json, working_sequences_json,
                        annulled_transfer_names_json, annulled_sequences_json
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    (
                        pipeline.title,
                        pipeline.mark,
                        pipeline.status,
                        pipeline.code,
                        pipeline.code_origin,
                        pipeline.working_revision_text,
                        pipeline.official_revision_text,
                        int(bool(pipeline.suspicious)),
                        pipeline.algorithm_version,
                        int(bool(pipeline.code_stale)),
                        pipeline.code_revision_text,
                        pipeline.code_date,
                        pipeline.tdo_date,
                        int(bool(pipeline.review_as_build)),
                        int(bool(pipeline.working_as_build)),
                        json.dumps(list(pipeline.working_transfer_names), ensure_ascii=False),
                        json.dumps(list(pipeline.working_sequences)),
                        json.dumps(
                            list(
                                getattr(pipeline, "annulled_transfer_names", ())
                                or ()
                            ),
                            ensure_ascii=False,
                        ),
                        json.dumps(
                            list(getattr(pipeline, "annulled_sequences", ()) or ())
                        ),
                    ),
                )

    def list_kit_packages(
        self,
        title: str | None = None,
        mark: str | None = None,
    ) -> list[KitPackageRow]:
        """Return persisted package rows.

        Args:
            title: Optional title filter (case-insensitive).
            mark: Optional mark filter (case-insensitive).

        Returns:
            Packages ordered by title, mark, sequence, and id.
        """

        where, params = self._kit_filter_sql(title, mark)
        with self._connection() as connection:
            rows = connection.execute(
                f"""
                SELECT * FROM kit_package
                {where}
                ORDER BY title, mark, sequence, id
                """,
                params,
            ).fetchall()
        return [self._package_from_row(row) for row in rows]

    def list_kit_cycles(
        self,
        title: str | None = None,
        mark: str | None = None,
    ) -> list[KitCycleRow]:
        """Return persisted official-send cycle rows.

        Args:
            title: Optional title filter (case-insensitive).
            mark: Optional mark filter (case-insensitive).

        Returns:
            Cycles ordered by title, mark, and id.
        """

        where, params = self._kit_filter_sql(title, mark)
        with self._connection() as connection:
            rows = connection.execute(
                f"""
                SELECT * FROM kit_cycle
                {where}
                ORDER BY title, mark, id
                """,
                params,
            ).fetchall()
        return [self._cycle_from_row(row) for row in rows]

    def list_kit_pipelines(
        self,
        title: str | None = None,
        mark: str | None = None,
    ) -> list[KitPipelineRow]:
        """Return derived pipeline status rows.

        Args:
            title: Optional title filter (case-insensitive).
            mark: Optional mark filter (case-insensitive).

        Returns:
            Pipeline rows ordered by title and mark.
        """

        where, params = self._kit_filter_sql(title, mark)
        with self._connection() as connection:
            rows = connection.execute(
                f"""
                SELECT * FROM kit_pipeline
                {where}
                ORDER BY title, mark
                """,
                params,
            ).fetchall()
        return [self._pipeline_from_row(row) for row in rows]

    def upsert_liquidity_review(
        self,
        title: str,
        mark: str,
        transfer_name: str,
        revision_text: str,
        *,
        decision: str,
        sequence: int | None = None,
        comment: str | None = None,
        evidence_mtime_ns: int | None = None,
        decided_at: str | None = None,
    ) -> int:
        """Insert or update a user liquidity decision.

        Identity is ``(title, mark, transfer_name, revision_text)`` with
        case-insensitive comparison. Never deleted by Google or derived
        rebuilds.

        Args:
            title: Four-digit title.
            mark: Latin AGCC mark.
            transfer_name: Issued folder name (or empty for unnamed).
            revision_text: Filename revision of the package.
            decision: ``confirmed_ok`` or ``confirmed_illiquid``.
            sequence: Optional folder ``NN``.
            comment: Optional reviewer note.
            evidence_mtime_ns: Package ``max_mtime_ns`` at decision time.
            decided_at: ISO UTC timestamp; default now.

        Returns:
            Liquidity-review row identifier.

        Raises:
            ValueError: If ``decision`` is not an allowed value.
        """

        if decision not in LIQUIDITY_DECISIONS:
            raise ValueError(f"Unsupported liquidity decision: {decision!r}")
        stamped = decided_at or _utc_now()
        with self._connection() as connection, connection:
            existing = connection.execute(
                """
                SELECT id FROM kit_liquidity_review
                WHERE title = ? COLLATE NOCASE
                  AND mark = ? COLLATE NOCASE
                  AND transfer_name = ? COLLATE NOCASE
                  AND revision_text = ? COLLATE NOCASE
                """,
                (title, mark, transfer_name, revision_text),
            ).fetchone()
            if existing is None:
                cursor = connection.execute(
                    """
                    INSERT INTO kit_liquidity_review(
                        title, mark, transfer_name, revision_text, sequence,
                        decision, comment, evidence_mtime_ns, decided_at
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    (
                        title,
                        mark,
                        transfer_name,
                        revision_text,
                        sequence,
                        decision,
                        comment,
                        evidence_mtime_ns,
                        stamped,
                    ),
                )
                return int(cursor.lastrowid)
            connection.execute(
                """
                UPDATE kit_liquidity_review
                SET sequence = ?, decision = ?, comment = ?,
                    evidence_mtime_ns = ?, decided_at = ?,
                    title = ?, mark = ?, transfer_name = ?, revision_text = ?
                WHERE id = ?
                """,
                (
                    sequence,
                    decision,
                    comment,
                    evidence_mtime_ns,
                    stamped,
                    title,
                    mark,
                    transfer_name,
                    revision_text,
                    int(existing["id"]),
                ),
            )
            return int(existing["id"])

    def list_liquidity_reviews(
        self,
        title: str | None = None,
        mark: str | None = None,
    ) -> list[LiquidityReviewRow]:
        """Return user liquidity decisions.

        Args:
            title: Optional title filter (case-insensitive).
            mark: Optional mark filter (case-insensitive).

        Returns:
            Reviews ordered by title, mark, and id.
        """

        where, params = self._kit_filter_sql(title, mark)
        with self._connection() as connection:
            rows = connection.execute(
                f"""
                SELECT * FROM kit_liquidity_review
                {where}
                ORDER BY title, mark, id
                """,
                params,
            ).fetchall()
        return [self._liquidity_from_row(row) for row in rows]

    def upsert_working_flag(
        self,
        title: str,
        mark: str,
        revision_text: str,
        *,
        transfer_name: str = "",
        sequence: int | None = None,
        decided_at: str | None = None,
    ) -> int:
        """Insert or update a manual working-folder mark.

        Identity is ``(title, mark, transfer_name)`` case-insensitive.
        Never deleted by scan, Google ingest, or derived rebuilds.

        Args:
            title: Four-digit title.
            mark: Latin AGCC mark.
            revision_text: Filename revision of that folder (display).
            transfer_name: Issued folder name of the tree node.
            sequence: Transfer ``NN``; parsed from ``transfer_name`` when omitted.
            decided_at: ISO UTC timestamp; default now.

        Returns:
            Row identifier.

        Raises:
            ValueError: If title, mark, revision_text, or transfer_name is empty.
        """

        parsed_title = str(title or "").strip()
        parsed_mark = str(mark or "").strip()
        parsed_rev = str(revision_text or "").strip()
        folder = str(transfer_name or "").strip()
        if not parsed_title or not parsed_mark or not parsed_rev or not folder:
            raise ValueError(
                "Working flag needs title, mark, revision_text, and transfer_name"
            )
        stamped = decided_at or _utc_now()
        parsed_sequence = _working_flag_sequence(folder, sequence)
        with self._connection() as connection, connection:
            existing = connection.execute(
                """
                SELECT id FROM kit_working_flag
                WHERE title = ? COLLATE NOCASE
                  AND mark = ? COLLATE NOCASE
                  AND transfer_name = ? COLLATE NOCASE
                """,
                (parsed_title, parsed_mark, folder),
            ).fetchone()
            connection.execute(
                """
                DELETE FROM kit_annulled_flag
                WHERE title = ? COLLATE NOCASE
                  AND mark = ? COLLATE NOCASE
                  AND transfer_name = ? COLLATE NOCASE
                """,
                (parsed_title, parsed_mark, folder),
            )
            if existing is None:
                cursor = connection.execute(
                    """
                    INSERT INTO kit_working_flag(
                        title, mark, revision_text, transfer_name, sequence, decided_at
                    ) VALUES (?, ?, ?, ?, ?, ?)
                    """,
                    (
                        parsed_title,
                        parsed_mark,
                        parsed_rev,
                        folder,
                        parsed_sequence,
                        stamped,
                    ),
                )
                return int(cursor.lastrowid)
            connection.execute(
                """
                UPDATE kit_working_flag
                SET revision_text = ?, sequence = ?, decided_at = ?
                WHERE id = ?
                """,
                (parsed_rev, parsed_sequence, stamped, int(existing["id"])),
            )
            return int(existing["id"])

    def delete_working_flag(
        self,
        title: str,
        mark: str,
        revision_text: str | None = None,
        *,
        transfer_name: str = "",
    ) -> bool:
        """Remove a manual working-folder mark.

        Args:
            title: Four-digit title.
            mark: Latin AGCC mark.
            revision_text: Unused when ``transfer_name`` is set; legacy
                whole-revision delete when the folder is omitted.
            transfer_name: Issued folder previously marked working.

        Returns:
            ``True`` when a row was deleted.
        """

        folder = str(transfer_name or "").strip()
        with self._connection() as connection, connection:
            if folder:
                cursor = connection.execute(
                    """
                    DELETE FROM kit_working_flag
                    WHERE title = ? COLLATE NOCASE
                      AND mark = ? COLLATE NOCASE
                      AND transfer_name = ? COLLATE NOCASE
                    """,
                    (title, mark, folder),
                )
            else:
                cursor = connection.execute(
                    """
                    DELETE FROM kit_working_flag
                    WHERE title = ? COLLATE NOCASE
                      AND mark = ? COLLATE NOCASE
                      AND revision_text = ? COLLATE NOCASE
                    """,
                    (title, mark, revision_text or ""),
                )
            return cursor.rowcount > 0

    def list_working_flags(
        self,
        title: str | None = None,
        mark: str | None = None,
    ) -> list[KitWorkingFlagRow]:
        """Return manual working-revision marks.

        Args:
            title: Optional title filter (case-insensitive).
            mark: Optional mark filter (case-insensitive).

        Returns:
            Flags ordered by title, mark, and id.
        """

        where, params = self._kit_filter_sql(title, mark)
        with self._connection() as connection:
            rows = connection.execute(
                f"""
                SELECT * FROM kit_working_flag
                {where}
                ORDER BY title, mark, id
                """,
                params,
            ).fetchall()
        return [self._working_flag_from_row(row) for row in rows]

    def apply_void_annulled_flags_from_records(
        self, records: Sequence[FileRecord]
    ) -> int:
        """Upsert ``kit_annulled_flag`` for present RD folders named Void.

        Existing rows are left untouched (``decided_at`` stays). Working
        flags on the same folder are deleted. Folders without a filename
        revision are skipped.

        Args:
            records: Catalog files (typically the pipeline contour).

        Returns:
            Number of Void packages that received an insert.
        """

        grouped: dict[tuple[str, str, str], tuple[int | None, str]] = {}
        for record in records:
            if record.source is not SourceKind.RD or not record.present:
                continue
            name = str(record.data.get("transfer_name") or "").strip()
            if not transfer_name_is_void(name):
                continue
            title = str(record.data.get("title") or "").strip()
            mark = str(record.data.get("mark") or "").strip()
            if not title or not mark:
                continue
            revision = format_revision(
                str(record.data.get("revision") or "") or None,
                str(record.data.get("appendix") or "") or None,
            )
            sequence = _working_flag_sequence(
                name, record.data.get("transfer_sequence")
            )
            key = (title, mark, name)
            prev = grouped.get(key)
            if prev is None:
                grouped[key] = (sequence, revision)
            else:
                grouped[key] = (prev[0] if prev[0] is not None else sequence, _better_revision_text(prev[1], revision))
        inserted = 0
        for (title, mark, name), (sequence, revision) in grouped.items():
            if not revision:
                continue
            self.upsert_annulled_flag(
                title,
                mark,
                revision,
                transfer_name=name,
                sequence=sequence,
                touch_existing=False,
            )
            inserted += 1
        return inserted

    @staticmethod
    def _apply_void_annulled_flags(
        connection: sqlite3.Connection,
        files: Sequence[ParsedFile],
    ) -> None:
        grouped: dict[tuple[str, str, str], tuple[int | None, str]] = {}
        for file in files:
            if file.source is not SourceKind.RD:
                continue
            transfer = file.transfer
            name = str(transfer.original_name if transfer else "").strip()
            if not transfer_name_is_void(name):
                continue
            title = str(file.title or "").strip()
            mark = str(file.mark or "").strip()
            if not title or not mark:
                continue
            revision = format_revision(file.revision, file.appendix)
            sequence = transfer.sequence if transfer is not None else None
            key = (title, mark, name)
            prev = grouped.get(key)
            if prev is None:
                grouped[key] = (sequence, revision)
            else:
                grouped[key] = (
                    prev[0] if prev[0] is not None else sequence,
                    _better_revision_text(prev[1], revision),
                )
        stamped = _utc_now()
        for (title, mark, name), (sequence, revision) in grouped.items():
            if not revision:
                continue
            CatalogDatabase._upsert_annulled_flag_on(
                connection,
                title,
                mark,
                revision,
                name,
                sequence,
                stamped,
                touch_existing=False,
            )

    def upsert_annulled_flag(
        self,
        title: str,
        mark: str,
        revision_text: str,
        *,
        transfer_name: str = "",
        sequence: int | None = None,
        decided_at: str | None = None,
        touch_existing: bool = True,
    ) -> int:
        """Insert or update a manual annulled-folder mark.

        Identity is ``(title, mark, transfer_name)`` case-insensitive.
        Deletes the working flag on the same folder. Never deleted by
        scan, Google ingest, or derived rebuilds.

        Args:
            title: Four-digit title.
            mark: Latin AGCC mark.
            revision_text: Filename revision of that folder (display).
            transfer_name: Issued folder name of the tree node.
            sequence: Transfer ``NN``; parsed from ``transfer_name`` when omitted.
            decided_at: ISO UTC timestamp; default now.
            touch_existing: When False, an existing row is left unchanged
                (scan/rebuild Void sync must not rewrite ``decided_at``).

        Returns:
            Row identifier.

        Raises:
            ValueError: If title, mark, revision_text, or transfer_name is empty.
        """

        parsed_title = str(title or "").strip()
        parsed_mark = str(mark or "").strip()
        parsed_rev = str(revision_text or "").strip()
        folder = str(transfer_name or "").strip()
        if not parsed_title or not parsed_mark or not parsed_rev or not folder:
            raise ValueError(
                "Annulled flag needs title, mark, revision_text, and transfer_name"
            )
        stamped = decided_at or _utc_now()
        parsed_sequence = _working_flag_sequence(folder, sequence)
        with self._connection() as connection, connection:
            return CatalogDatabase._upsert_annulled_flag_on(
                connection,
                parsed_title,
                parsed_mark,
                parsed_rev,
                folder,
                parsed_sequence,
                stamped,
                touch_existing=touch_existing,
            )

    @staticmethod
    def _upsert_annulled_flag_on(
        connection: sqlite3.Connection,
        title: str,
        mark: str,
        revision_text: str,
        transfer_name: str,
        sequence: int | None,
        decided_at: str,
        *,
        touch_existing: bool,
    ) -> int:
        existing = connection.execute(
            """
            SELECT id FROM kit_annulled_flag
            WHERE title = ? COLLATE NOCASE
              AND mark = ? COLLATE NOCASE
              AND transfer_name = ? COLLATE NOCASE
            """,
            (title, mark, transfer_name),
        ).fetchone()
        connection.execute(
            """
            DELETE FROM kit_working_flag
            WHERE title = ? COLLATE NOCASE
              AND mark = ? COLLATE NOCASE
              AND transfer_name = ? COLLATE NOCASE
            """,
            (title, mark, transfer_name),
        )
        if existing is None:
            cursor = connection.execute(
                """
                INSERT INTO kit_annulled_flag(
                    title, mark, revision_text, transfer_name,
                    sequence, decided_at
                ) VALUES (?, ?, ?, ?, ?, ?)
                """,
                (
                    title,
                    mark,
                    revision_text,
                    transfer_name,
                    sequence,
                    decided_at,
                ),
            )
            return int(cursor.lastrowid)
        if not touch_existing:
            return int(existing["id"])
        connection.execute(
            """
            UPDATE kit_annulled_flag
            SET revision_text = ?, sequence = ?, decided_at = ?
            WHERE id = ?
            """,
            (revision_text, sequence, decided_at, int(existing["id"])),
        )
        return int(existing["id"])

    def delete_annulled_flag(
        self,
        title: str,
        mark: str,
        *,
        transfer_name: str = "",
    ) -> bool:
        """Remove a manual annulled-folder mark.

        Args:
            title: Four-digit title.
            mark: Latin AGCC mark.
            transfer_name: Issued folder previously marked annulled.

        Returns:
            ``True`` when a row was deleted.
        """

        folder = str(transfer_name or "").strip()
        if not folder:
            return False
        with self._connection() as connection, connection:
            cursor = connection.execute(
                """
                DELETE FROM kit_annulled_flag
                WHERE title = ? COLLATE NOCASE
                  AND mark = ? COLLATE NOCASE
                  AND transfer_name = ? COLLATE NOCASE
                """,
                (title, mark, folder),
            )
            return cursor.rowcount > 0

    def list_annulled_flags(
        self,
        title: str | None = None,
        mark: str | None = None,
    ) -> list[KitAnnulledFlagRow]:
        """Return manual annulled-folder marks.

        Args:
            title: Optional title filter (case-insensitive).
            mark: Optional mark filter (case-insensitive).

        Returns:
            Flags ordered by title, mark, and id.
        """

        where, params = self._kit_filter_sql(title, mark)
        with self._connection() as connection:
            rows = connection.execute(
                f"""
                SELECT * FROM kit_annulled_flag
                {where}
                ORDER BY title, mark, id
                """,
                params,
            ).fetchall()
        return [self._annulled_flag_from_row(row) for row in rows]

    def upsert_file_mtime_override(
        self,
        path_key: str,
        override_date: str,
        *,
        reason: str = "code_a",
        decided_at: str | None = None,
    ) -> FileMtimeOverrideRow:
        """Store a catalog date for one catalog file; scan must not clear it.

        Fingerprint is ``path_key|size|disk_mtime_ns`` of the current
        ``file_entry`` row. A later scan of the same bytes keeps the
        override; a different size or disk mtime makes it stale.

        Args:
            path_key: File identity (``make_path_key``).
            override_date: ``DD.MM.YYYY`` (F letter, typed date, or folder mean).
            reason: ``code_a`` / ``code_b`` / ``code_c`` / ``manual`` /
                ``folder_mean``.
            decided_at: ISO UTC timestamp; default now.

        Returns:
            Persisted override row.

        Raises:
            KeyError: No ``file_entry`` for ``path_key``.
            ValueError: Not a catalog document file, or the date/reason is invalid.
        """

        key = str(path_key or "").strip()
        if not key:
            raise ValueError("Override needs a path_key")
        token = str(reason or "").strip() or "code_a"
        if token not in FILE_MTIME_OVERRIDE_REASONS:
            raise ValueError(f"Unsupported mtime override reason: {token!r}")
        override_mtime_ns = override_mtime_ns_from_date(override_date)
        stamped = decided_at or _utc_now()
        with self._connection() as connection, connection:
            file_row = connection.execute(
                """
                SELECT path_key, size, mtime_ns, file_kind
                FROM file_entry
                WHERE path_key = ? COLLATE NOCASE
                """,
                (key,),
            ).fetchone()
            if file_row is None:
                raise KeyError(f"Unknown file path_key: {key}")
            kind = str(file_row["file_kind"] or "")
            if kind not in _FILE_MTIME_OVERRIDE_KINDS:
                raise ValueError("Date override is only for PDF, MTO xlsx, or editable")
            stored_key = str(file_row["path_key"])
            size = int(file_row["size"] or 0)
            disk_mtime_ns = int(file_row["mtime_ns"] or 0)
            fingerprint = file_mtime_fingerprint(stored_key, size, disk_mtime_ns)
            existing = connection.execute(
                """
                SELECT id FROM file_mtime_override
                WHERE path_key = ? COLLATE NOCASE
                """,
                (stored_key,),
            ).fetchone()
            if existing is None:
                cursor = connection.execute(
                    """
                    INSERT INTO file_mtime_override(
                        path_key, size, disk_mtime_ns, fingerprint,
                        override_date, override_mtime_ns, reason, decided_at
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    (
                        stored_key,
                        size,
                        disk_mtime_ns,
                        fingerprint,
                        str(override_date).strip(),
                        override_mtime_ns,
                        token,
                        stamped,
                    ),
                )
                row_id = int(cursor.lastrowid)
            else:
                row_id = int(existing["id"])
                connection.execute(
                    """
                    UPDATE file_mtime_override
                    SET size = ?, disk_mtime_ns = ?, fingerprint = ?,
                        override_date = ?, override_mtime_ns = ?,
                        reason = ?, decided_at = ?
                    WHERE id = ?
                    """,
                    (
                        size,
                        disk_mtime_ns,
                        fingerprint,
                        str(override_date).strip(),
                        override_mtime_ns,
                        token,
                        stamped,
                        row_id,
                    ),
                )
        stored = self.get_file_mtime_override(stored_key)
        if stored is None:
            raise RuntimeError("Failed to persist file_mtime_override")
        return stored

    def delete_file_mtime_override(self, path_key: str) -> bool:
        """Remove a catalog-date override.

        Args:
            path_key: File identity.

        Returns:
            ``True`` when a row was deleted.
        """

        key = str(path_key or "").strip()
        if not key:
            return False
        with self._connection() as connection, connection:
            cursor = connection.execute(
                "DELETE FROM file_mtime_override WHERE path_key = ? COLLATE NOCASE",
                (key,),
            )
            return cursor.rowcount > 0

    def get_file_mtime_override(self, path_key: str) -> FileMtimeOverrideRow | None:
        """Return the override for one path, or ``None``.

        Args:
            path_key: File identity.
        """

        key = str(path_key or "").strip()
        if not key:
            return None
        with self._connection() as connection:
            row = connection.execute(
                """
                SELECT * FROM file_mtime_override
                WHERE path_key = ? COLLATE NOCASE
                """,
                (key,),
            ).fetchone()
        if row is None:
            return None
        return _mtime_override_from_row(row)

    def list_file_mtime_overrides(self) -> list[FileMtimeOverrideRow]:
        """Return every catalog-date override."""

        with self._connection() as connection:
            rows = connection.execute(
                "SELECT * FROM file_mtime_override ORDER BY path_key, id"
            ).fetchall()
        return [_mtime_override_from_row(row) for row in rows]

    def upsert_issuance_review(
        self,
        title: str,
        mark: str,
        kind: str,
        source: str,
        identity_fingerprint: str,
        *,
        decision: str = "",
        revision_text: str = "",
        send_date: str = "",
        send_date_sortable: str = "",
        send_transmittal: str = "",
        incoming_control_date: str = "",
        incoming_control_date_sortable: str = "",
        confirm_transmittal: str = "",
        sheet_status: str = "",
        note: str = "",
        comment: str | None = None,
        evidence_fingerprint: str = "",
        match_state: str = "",
        source_path: str = "",
        path_key: str = "",
        decided_at: str | None = None,
        review_id: int | None = None,
    ) -> int:
        """Insert or update a user issuance decision.

        Identity is ``(kind, identity_fingerprint)``. Never deleted by Google
        or derived rebuilds. Pass ``review_id`` to update that row even when
        the identity fingerprint changes (used by rematch).

        Args:
            title: Four-digit title.
            mark: Latin AGCC mark.
            kind: ``send``, ``orphan``, or ``manual``.
            source: ``issuance``, ``google_f``, ``rd``, ``robot``,
                ``auto_mto``, or ``manual``.
            identity_fingerprint: Stable SHA-256 identity hex.
            decision: Allowed issuance decision, including empty.
            revision_text: Filename or sheet revision text.
            send_date: Display send date.
            send_date_sortable: Sortable send date.
            send_transmittal: Send TRM token.
            incoming_control_date: Display incoming-control date.
            incoming_control_date_sortable: Sortable incoming-control date.
            confirm_transmittal: Confirmation TRM token.
            sheet_status: Sheet status cell.
            note: Sheet note / remark.
            comment: Optional reviewer comment.
            evidence_fingerprint: SHA-256 evidence hex at decision time.
            match_state: ``matched``, ``unmatched``, ``ambiguous``, or empty.
            source_path: Optional file or sheet path (display only).
            path_key: Optional path identity for robot / Auto MTO.
            decided_at: ISO UTC timestamp; default now.
            review_id: Existing row id to update in place.

        Returns:
            Issuance-review row identifier.

        Raises:
            ValueError: If ``kind``, ``source``, ``decision``, or
                ``match_state`` is not an allowed value.
        """

        if kind not in ISSUANCE_REVIEW_KINDS:
            raise ValueError(f"Unsupported issuance review kind: {kind!r}")
        if source not in ISSUANCE_REVIEW_SOURCES:
            raise ValueError(f"Unsupported issuance review source: {source!r}")
        if decision not in ISSUANCE_REVIEW_DECISIONS:
            raise ValueError(f"Unsupported issuance review decision: {decision!r}")
        if match_state not in ISSUANCE_REVIEW_MATCH_STATES:
            raise ValueError(
                f"Unsupported issuance review match_state: {match_state!r}"
            )
        stamped = decided_at or _utc_now()
        values = (
            title,
            mark,
            kind,
            source,
            decision,
            revision_text,
            send_date,
            send_date_sortable,
            send_transmittal,
            incoming_control_date,
            incoming_control_date_sortable,
            confirm_transmittal,
            sheet_status,
            note,
            comment,
            identity_fingerprint,
            evidence_fingerprint,
            match_state,
            source_path,
            path_key,
            stamped,
        )
        with self._connection() as connection, connection:
            existing = None
            if review_id is not None:
                existing = connection.execute(
                    "SELECT id FROM issuance_review WHERE id = ?",
                    (int(review_id),),
                ).fetchone()
            if existing is None:
                existing = connection.execute(
                    """
                    SELECT id FROM issuance_review
                    WHERE kind = ? AND identity_fingerprint = ?
                    """,
                    (kind, identity_fingerprint),
                ).fetchone()
            if existing is None:
                cursor = connection.execute(
                    """
                    INSERT INTO issuance_review(
                        title, mark, kind, source, decision, revision_text,
                        send_date, send_date_sortable, send_transmittal,
                        incoming_control_date, incoming_control_date_sortable,
                        confirm_transmittal, sheet_status, note, comment,
                        identity_fingerprint, evidence_fingerprint, match_state,
                        source_path, path_key, decided_at
                    ) VALUES (
                        ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?,
                        ?, ?, ?
                    )
                    """,
                    values,
                )
                return int(cursor.lastrowid)
            connection.execute(
                """
                UPDATE issuance_review
                SET title = ?, mark = ?, kind = ?, source = ?, decision = ?,
                    revision_text = ?, send_date = ?, send_date_sortable = ?,
                    send_transmittal = ?, incoming_control_date = ?,
                    incoming_control_date_sortable = ?, confirm_transmittal = ?,
                    sheet_status = ?, note = ?, comment = ?,
                    identity_fingerprint = ?, evidence_fingerprint = ?,
                    match_state = ?, source_path = ?, path_key = ?,
                    decided_at = ?
                WHERE id = ?
                """,
                (*values, int(existing["id"])),
            )
            return int(existing["id"])

    def list_issuance_reviews(
        self,
        title: str | None = None,
        mark: str | None = None,
    ) -> list[IssuanceReviewRow]:
        """Return user issuance decisions.

        Args:
            title: Optional title filter (case-insensitive).
            mark: Optional mark filter (case-insensitive).

        Returns:
            Reviews ordered by title, mark, and id.
        """

        where, params = self._kit_filter_sql(title, mark)
        with self._connection() as connection:
            rows = connection.execute(
                f"""
                SELECT * FROM issuance_review
                {where}
                ORDER BY title, mark, id
                """,
                params,
            ).fetchall()
        return [self._issuance_review_from_row(row) for row in rows]

    def get_issuance_review(self, review_id: int) -> IssuanceReviewRow | None:
        """Return one issuance review by SQLite id.

        Args:
            review_id: ``issuance_review.id``.

        Returns:
            The row, or ``None`` when absent.
        """

        with self._connection() as connection:
            row = connection.execute(
                "SELECT * FROM issuance_review WHERE id = ?",
                (int(review_id),),
            ).fetchone()
        if row is None:
            return None
        return self._issuance_review_from_row(row)

    def delete_issuance_review(self, review_id: int) -> None:
        """Delete one issuance review by SQLite id.

        Args:
            review_id: ``issuance_review.id``.
        """

        with self._connection() as connection, connection:
            connection.execute(
                "DELETE FROM issuance_review WHERE id = ?",
                (int(review_id),),
            )

    def replace_revision_cells(
        self,
        cells: Sequence[KitRevisionRow],
        *,
        identities: Collection[tuple[str, str]] | None = None,
    ) -> None:
        """Replace ``kit_revision_cell`` rows in one transaction.

        Does not touch ``kit_liquidity_review``, Google snapshot tables,
        ``file_entry``, or ``kit_package``. Duplicate
        ``(title, mark, revision_text)`` identities keep the last row via
        :func:`kit_identity_key` plus case-insensitive revision text.

        Args:
            cells: Derived heatmap cells for the current rebuild.
            identities: When set, delete and reinsert only these kit keys.
                ``None`` replaces every heatmap row.
        """

        unique: dict[tuple[str, str, str], KitRevisionRow] = {}
        for cell in cells:
            unique[
                (
                    *kit_identity_key(cell.title, cell.mark),
                    (cell.revision_text or "").casefold(),
                )
            ] = cell
        with self._connection() as connection, connection:
            if identities is None:
                connection.execute("DELETE FROM kit_revision_cell")
            else:
                scoped = {
                    kit_identity_key(title, mark) for title, mark in identities
                }
                self._delete_kit_identity_rows(
                    connection, "kit_revision_cell", scoped
                )
            for cell in unique.values():
                connection.execute(
                    """
                    INSERT INTO kit_revision_cell(
                        title, mark, revision_text, pipeline_status, letters,
                        is_as_build, is_current, is_current_ifc, has_mto,
                        problem_kinds_json, package_ids_json, algorithm_version
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    (
                        cell.title,
                        cell.mark,
                        cell.revision_text,
                        cell.pipeline_status,
                        cell.letters,
                        int(bool(cell.is_as_build)),
                        int(bool(cell.is_current)),
                        int(bool(cell.is_current_ifc)),
                        int(bool(cell.has_mto)),
                        cell.problem_kinds_json or "[]",
                        cell.package_ids_json or "[]",
                        cell.algorithm_version,
                    ),
                )

    def list_revision_cells(
        self,
        title: str | None = None,
        mark: str | None = None,
    ) -> list[KitRevisionRow]:
        """Return persisted heatmap cells.

        Args:
            title: Optional title filter (case-insensitive).
            mark: Optional mark filter (case-insensitive).

        Returns:
            Cells ordered by title, mark, revision text, and id.
        """

        where, params = self._kit_filter_sql(title, mark)
        with self._connection() as connection:
            rows = connection.execute(
                f"""
                SELECT * FROM kit_revision_cell
                {where}
                ORDER BY title, mark, revision_text, id
                """,
                params,
            ).fetchall()
        return [self._revision_from_row(row) for row in rows]

    @staticmethod
    def _an_mto_file_from_row(row: sqlite3.Row) -> AnMtoFile:
        return AnMtoFile(
            path=str(row["path"] or ""),
            path_key=str(row["path_key"] or ""),
            title=str(row["title"] or ""),
            mark=str(row["mark"] or ""),
            revision_text=str(row["revision_text"] or ""),
            core_stem=str(row["core_stem"] or ""),
            discipline_block=str(row["discipline_block"] or ""),
            name=str(row["name"] or ""),
            parent_dir=str(row["parent_dir"] or ""),
            mtime_ns=int(row["mtime_ns"] or 0),
            size=int(row["size"] or 0),
        )

    def replace_an_snapshot(
        self,
        files: Sequence[AnMtoFile],
        *,
        scanned_at: str,
    ) -> None:
        """Replace all ``an_mto_file`` rows in one transaction.

        Does not touch ``file_entry``, overlay, ``issuance_review``,
        ``kit_liquidity_review``, ``kit_working_flag``, or
        ``kit_annulled_flag``.

        Args:
            files: Parsed AN MTO workbooks from the last walk.
            scanned_at: ISO timestamp stored on every inserted row.
        """

        with self._connection() as connection, connection:
            connection.execute("DELETE FROM an_mto_file")
            connection.executemany(
                """
                INSERT INTO an_mto_file(
                    path, path_key, title, mark, revision_text, core_stem,
                    discipline_block, name, parent_dir, mtime_ns, size,
                    scanned_at
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                [
                    (
                        file.path,
                        file.path_key,
                        file.title,
                        file.mark,
                        file.revision_text,
                        file.core_stem,
                        file.discipline_block,
                        file.name,
                        file.parent_dir,
                        int(file.mtime_ns),
                        int(file.size),
                        scanned_at,
                    )
                    for file in files
                ],
            )

    def list_an_mto_files(self) -> tuple[AnMtoFile, ...]:
        """Return every stored AN MTO file.

        Returns:
            Rows ordered by title, mark, revision text, and path.
        """

        with self._connection() as connection:
            rows = connection.execute(
                """
                SELECT path, path_key, title, mark, revision_text, core_stem,
                       discipline_block, name, parent_dir, mtime_ns, size
                FROM an_mto_file
                ORDER BY title, mark, revision_text, path
                """
            ).fetchall()
        return tuple(self._an_mto_file_from_row(row) for row in rows)

    def list_an_files_by_kit(self) -> dict[tuple[str, str], tuple[AnMtoFile, ...]]:
        """Group stored AN MTO files by kit identity.

        Returns:
            Mapping of :func:`kit_identity_key` to files in list order.
        """

        grouped: dict[tuple[str, str], list[AnMtoFile]] = {}
        for file in self.list_an_mto_files():
            key = kit_identity_key(file.title, file.mark)
            grouped.setdefault(key, []).append(file)
        return {key: tuple(items) for key, items in grouped.items()}

    def replace_rd_dump_snapshot(
        self,
        files: Sequence[AnMtoFile],
        *,
        scanned_at: str,
    ) -> None:
        """Replace all ``rd_dump_mto_file`` rows in one transaction.

        Does not touch ``file_entry``, overlay, ``an_mto_file``,
        ``issuance_review``, ``kit_liquidity_review``, ``kit_working_flag``,
        or ``kit_annulled_flag``.

        Args:
            files: Parsed RD-dump MTO / OD files from the last walk.
            scanned_at: ISO timestamp stored on every inserted row.
        """

        with self._connection() as connection, connection:
            connection.execute("DELETE FROM rd_dump_mto_file")
            connection.executemany(
                """
                INSERT INTO rd_dump_mto_file(
                    path, path_key, title, mark, revision_text, core_stem,
                    discipline_block, name, parent_dir, mtime_ns, size,
                    scanned_at
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                [
                    (
                        file.path,
                        file.path_key,
                        file.title,
                        file.mark,
                        file.revision_text,
                        file.core_stem,
                        file.discipline_block,
                        file.name,
                        file.parent_dir,
                        int(file.mtime_ns),
                        int(file.size),
                        scanned_at,
                    )
                    for file in files
                ],
            )

    def list_rd_dump_mto_files(self) -> tuple[AnMtoFile, ...]:
        """Return every stored RD-dump MTO / OD file.

        Returns:
            Rows ordered by title, mark, revision text, and path.
        """

        with self._connection() as connection:
            rows = connection.execute(
                """
                SELECT path, path_key, title, mark, revision_text, core_stem,
                       discipline_block, name, parent_dir, mtime_ns, size
                FROM rd_dump_mto_file
                ORDER BY title, mark, revision_text, path
                """
            ).fetchall()
        return tuple(self._an_mto_file_from_row(row) for row in rows)

    def list_rd_dump_files_by_kit(
        self,
    ) -> dict[tuple[str, str], tuple[AnMtoFile, ...]]:
        """Group stored RD-dump MTO files by kit identity.

        Returns:
            Mapping of :func:`kit_identity_key` to files in list order.
        """

        grouped: dict[tuple[str, str], list[AnMtoFile]] = {}
        for file in self.list_rd_dump_mto_files():
            key = kit_identity_key(file.title, file.mark)
            grouped.setdefault(key, []).append(file)
        return {key: tuple(items) for key, items in grouped.items()}
