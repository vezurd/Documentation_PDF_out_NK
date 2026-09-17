"""Typed domain models for RD catalog scanning and persistence."""

from __future__ import annotations

from dataclasses import dataclass, field
from enum import StrEnum
from pathlib import Path
from typing import Any


class SourceKind(StrEnum):
    """Catalog input source."""

    RD = "rd"
    SQ = "sq"
    ROBOT = "robot"


class FileKind(StrEnum):
    """Supported catalog file kind."""

    PDF = "pdf"
    MTO_XLSX = "mto_xlsx"
    SOURCE_EDITABLE = "source_editable"


class ParseStatus(StrEnum):
    """Result of parsing a discovered path."""

    PARSED = "parsed"
    UNPARSED_FILE = "unparsed_file"
    UNPARSED_FOLDER = "unparsed_folder"
    STAT_ERROR = "stat_error"


class OverlayStatus(StrEnum):
    """A file's objective state in an overlay calculation."""

    DETECTED_CURRENT = "detected_current"
    SUPERSEDED = "superseded"
    AMBIGUOUS = "ambiguous"
    IGNORED = "ignored"


class ReviewState(StrEnum):
    """User review state, kept separate from objective overlay state."""

    PENDING = "pending"
    ACKNOWLEDGED = "acknowledged"
    IGNORED = "ignored"


class CollisionKind(StrEnum):
    """Non-fatal catalog collision category."""

    UNPARSED_FOLDER = "unparsed_folder"
    UNPARSED_FILE = "unparsed_file"
    DUP_SAME_REVISION = "dup_same_revision"
    FOLDER_VS_FILENAME_REVISION = "folder_vs_filename_revision"
    TRANSFER_ORDER_CONFLICT = "transfer_order_conflict"
    TRANSFER_MTIME_CONFLICT = "transfer_mtime_conflict"
    RD_VS_SQ = "rd_vs_sq"
    FILE_MISSING = "file_missing"
    SOURCE_ERROR = "source_error"


_COLLISION_KIND_LABELS: dict[str, str] = {
    CollisionKind.TRANSFER_ORDER_CONFLICT.value: "Порядок передач",
    CollisionKind.TRANSFER_MTIME_CONFLICT.value: (
        "Дата vs номер передачи"
    ),
    CollisionKind.DUP_SAME_REVISION.value: "Дубль в передаче",
    CollisionKind.UNPARSED_FOLDER.value: "Папка не разобрана",
    CollisionKind.UNPARSED_FILE.value: "Файл не разобран",
    CollisionKind.FOLDER_VS_FILENAME_REVISION.value: (
        "Ревизия папки ≠ файла"
    ),
    CollisionKind.RD_VS_SQ.value: "РД vs SQ",
    CollisionKind.FILE_MISSING.value: "Файл отсутствует",
    CollisionKind.SOURCE_ERROR.value: "Ошибка источника",
    "liquidity": "Ликвидность",
}


def collision_kind_label(kind: str) -> str:
    """Return a short Russian label for a stored collision kind.

    Args:
        kind: Stored ``CollisionKind`` value or a derived problem token.

    Returns:
        Localized label, or the raw token when unknown.
    """

    return _COLLISION_KIND_LABELS.get(kind, kind)


class ScanRunStatus(StrEnum):
    """Overall scan completion state."""

    SUCCESS = "success"
    PARTIAL = "partial"
    CANCELLED = "cancelled"


class MtoReadinessStatus(StrEnum):
    """Robot MTO readiness shown by the catalog."""

    READY = "ready"
    WARNING = "warning"
    BLOCKED = "blocked"


class MtoContentStatus(StrEnum):
    """Outcome of semantic MTO content comparison."""

    EQUAL = "content_equal"
    DIFF = "content_diff"
    NOT_COMPARED = "not_compared"
    COMPARE_ERROR = "compare_error"


class MtoIssueKind(StrEnum):
    """Structured reason affecting MTO readiness."""

    ROBOT_MISSING = "robot_missing"
    ROBOT_DUPLICATE = "robot_duplicate"
    ROBOT_EXTRA = "robot_extra"
    RD_UNPARSED = "rd_unparsed"
    ROBOT_UNPARSED = "robot_unparsed"
    LOAD_ERROR = "load_error"
    COMPARE_ERROR = "compare_error"
    CONTENT_DIFF = "content_diff"
    REVISION_MISMATCH = "revision_mismatch"
    MTIME_SUSPICION = "mtime_suspicion"


@dataclass(frozen=True, slots=True)
class TransferMetadata:
    """Parsed metadata from a transfer folder.

    Attributes:
        original_name: Folder name exactly as found on disk.
        normalized_name: Folder name after Unicode dash normalization.
        sequence: Transfer sequence number.
        revision: Declared transfer revision, if present.
        appendix: Declared ``AN`` token without the ``AN`` prefix.
        title: Four-digit title number, if present.
        mark: Mark/system component, if present.
        title_system: Combined title and mark, if present.
        is_as_build: Whether the folder identifies an as-built transfer.
        parse_status: Folder parsing result.
        error: Human-readable parse problem.
    """

    original_name: str
    normalized_name: str
    sequence: int | None = None
    revision: str | None = None
    appendix: str | None = None
    title: str | None = None
    mark: str | None = None
    title_system: str | None = None
    is_as_build: bool = False
    parse_status: ParseStatus = ParseStatus.PARSED
    error: str | None = None


@dataclass(frozen=True, slots=True)
class ParsedFile:
    """One statted and parsed catalog file."""

    path: str
    path_key: str
    name: str
    source: SourceKind
    file_kind: FileKind
    size: int
    mtime_ns: int
    parse_status: ParseStatus
    contract: str | None = None
    title_system: str | None = None
    title: str | None = None
    mark: str | None = None
    discipline_block: str | None = None
    core_stem: str | None = None
    revision: str | None = None
    appendix: str | None = None
    language: str | None = None
    extension: str | None = None
    transfer: TransferMetadata | None = None
    parse_error: str | None = None

    @property
    def document_key(self) -> str | tuple[str, str] | None:
        """Return the prescribed identity key for this file.

        Returns:
            Normalized PDF/source-editable core stem, the MTO tuple key, or
            ``None``.
        """

        if self.file_kind in (FileKind.PDF, FileKind.SOURCE_EDITABLE):
            return self.core_stem.casefold() if self.core_stem else None
        if self.title_system and self.discipline_block:
            return (self.title_system.casefold(), self.discipline_block.casefold())
        return None

    @property
    def stat_signature(self) -> tuple[str, int, int]:
        """Return the incremental scan signature."""

        return (self.path_key, self.size, self.mtime_ns)


@dataclass(frozen=True, slots=True)
class OverlayEntry:
    """One file classified by the delta overlay."""

    file: ParsedFile
    document_key: str
    status: OverlayStatus


@dataclass(frozen=True, slots=True)
class OverlayCollision:
    """A non-fatal ambiguity or ordering problem."""

    kind: CollisionKind
    message: str
    path_keys: tuple[str, ...] = ()
    document_key: str | None = None


@dataclass(slots=True)
class OverlayResult:
    """Delta-overlay output for one file kind."""

    file_kind: FileKind
    current: dict[str, ParsedFile] = field(default_factory=dict)
    entries: list[OverlayEntry] = field(default_factory=list)
    collisions: list[OverlayCollision] = field(default_factory=list)


@dataclass(frozen=True, slots=True)
class SourceError:
    """Failure confined to one scan source or path."""

    source: SourceKind
    path: str
    message: str


@dataclass(frozen=True, slots=True)
class ScanProgress:
    """Progress callback payload."""

    source: SourceKind
    path: str
    files_seen: int
    message: str = ""


@dataclass(slots=True)
class SourceScanResult:
    """Files and diagnostics produced by scanning one source."""

    source: SourceKind
    root: str
    files: list[ParsedFile] = field(default_factory=list)
    collisions: list[OverlayCollision] = field(default_factory=list)
    errors: list[SourceError] = field(default_factory=list)
    candidates_seen: int = 0
    excluded_files: int = 0
    cancelled: bool = False
    subtree: str | None = None
    subtrees: tuple[str, ...] = ()

    def scoped_folders(self) -> tuple[str, ...]:
        """Return folders that bound missing-file detection.

        A scoped walk stores one or more subtree roots. Older callers may
        set only ``subtree``; a full-source walk leaves both empty.

        Returns:
            Unique folder prefixes, or an empty tuple for a full source walk.
        """

        if self.subtrees:
            return self.subtrees
        if self.subtree:
            return (self.subtree,)
        return ()


@dataclass(slots=True)
class ScanSummary:
    """Combined source scan result suitable for persistence and later GUI use."""

    sources: dict[SourceKind, SourceScanResult] = field(default_factory=dict)
    status: ScanRunStatus = ScanRunStatus.SUCCESS
    rd_pdf_overlay: OverlayResult | None = None
    rd_mto_overlay: OverlayResult | None = None

    @property
    def files(self) -> list[ParsedFile]:
        """Return all discovered files across sources."""

        return [
            file
            for source_result in self.sources.values()
            for file in source_result.files
        ]

    @property
    def errors(self) -> list[SourceError]:
        """Return all source errors without discarding partial results."""

        return [
            error
            for source_result in self.sources.values()
            for error in source_result.errors
        ]


@dataclass(frozen=True, slots=True)
class FileRecord:
    """Persisted file row exposed to later GUI phases."""

    id: int
    path: str
    path_key: str
    source: SourceKind
    present: bool
    review_state: ReviewState
    first_seen_run_id: int
    last_seen_run_id: int
    data: dict[str, Any] = field(default_factory=dict)


def make_path_key(path: str | Path) -> str:
    """Build a case-insensitive normalized path identity.

    Args:
        path: Local or UNC path.

    Returns:
        Windows-compatible normalized, case-folded path text.
    """

    import os

    return os.path.normcase(os.path.normpath(str(path))).casefold()
