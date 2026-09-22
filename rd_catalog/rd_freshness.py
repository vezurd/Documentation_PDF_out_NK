"""Quiet RD file-stamp census and Google snapshot fingerprint.

The census lists canonical RD names and compares size and mtime with the
last catalog snapshot. It does not open workbooks or PDFs. A changed or
missing file yields the mark folder to rescan, never the RD root.
"""

from __future__ import annotations

import hashlib
import json
import os
from collections.abc import Callable, Iterable, Sequence
from dataclasses import dataclass

from rd_catalog.kits import GoogleKit, IssuanceKit
from rd_catalog.models import FileRecord, SourceKind, make_path_key
from rd_catalog.parse import (
    has_canonical_rd_issued_path,
    is_package_media_folder,
    kit_rd_mark_folder_from_path,
    matches_agcc_filename,
)
from rd_catalog.path_actions import path_is_under
from rd_catalog.scan import _candidate_kind, _is_excel_lock_name, _is_skipped_dir, collapse_nested_folders

AUTO_REFRESH_PERIOD_S = 75 * 60
AUTO_REFRESH_TICK_MS = 60_000
AUTO_REFRESH_MAX_FOLDERS = 4

CancelCallback = Callable[[], bool]


@dataclass(frozen=True, slots=True)
class RdFileStamp:
    """Size and mtime of one catalog file, without its contents."""

    path: str
    path_key: str
    size: int
    mtime_ns: int


@dataclass(frozen=True, slots=True)
class RdCensusResult:
    """Finished or cancelled listing of canonical RD file stamps."""

    stamps: tuple[RdFileStamp, ...]
    completed: bool
    errors: tuple[str, ...] = ()


def rd_read_error_is_missing(error: str) -> bool:
    """Return whether an Auto MTO read error means the RD file is absent.

    Args:
        error: ``AutoMtoCompareResult.error`` text.

    Returns:
        True for a missing path (``Errno 2`` and the Windows equivalents).
    """

    text = str(error or "").casefold()
    if not text:
        return False
    markers = (
        "no such file",
        "errno 2",
        "filenotfounderror",
        "winerror 2",
        "cannot find the file",
        "cannot find the path",
        "не удается найти",
        "система не может найти",
    )
    return any(marker in text for marker in markers)


def google_snapshot_fingerprint(
    kits: Sequence[GoogleKit],
    sends: Sequence[IssuanceKit],
) -> str:
    """Return a stable hash of the kits sheet and issuance sends.

    Args:
        kits: Parsed КСБ ИД rows.
        sends: All parsed «Выдача РД ПД» sends, not only the latest.

    Returns:
        SHA-256 hex of the fields that change a kit's Google picture.
    """

    kit_part = sorted(
        (
            str(kit.title or "").casefold(),
            str(kit.mark or "").casefold(),
            str(kit.sheet_revision_text or ""),
            str(kit.status_sheet or ""),
            str(kit.comment_raw or ""),
        )
        for kit in kits
    )
    send_part = sorted(
        (
            str(send.title or "").casefold(),
            str(send.mark or "").casefold(),
            str(send.revision_text or ""),
            str(send.status or ""),
            str(send.send_date or ""),
            str(send.send_transmittal or ""),
            str(send.incoming_control_date or ""),
            str(send.confirm_transmittal or ""),
            str(send.note_raw or ""),
        )
        for send in sends
    )
    blob = json.dumps(
        {"kits": kit_part, "sends": send_part},
        ensure_ascii=False,
        separators=(",", ":"),
    )
    return hashlib.sha256(blob.encode("utf-8")).hexdigest()


def stamps_from_records(records: Sequence[FileRecord]) -> tuple[RdFileStamp, ...]:
    """Return disk size and mtime for present RD rows.

    Uses ``disk_mtime_ns`` when a catalog date override replaced
    ``mtime_ns``. Absent or non-RD rows are skipped.

    Args:
        records: Catalog file rows already loaded in the window.

    Returns:
        One stamp per present RD file.
    """

    stamps: list[RdFileStamp] = []
    for record in records:
        if record.source is not SourceKind.RD or not record.present:
            continue
        path = str(record.path or "").strip()
        if not path:
            continue
        data = record.data or {}
        size = int(data.get("size") or 0)
        disk_mtime = data.get("disk_mtime_ns")
        if disk_mtime in (None, ""):
            disk_mtime = data.get("mtime_ns")
        stamps.append(
            RdFileStamp(
                path=path,
                path_key=str(record.path_key or make_path_key(path)),
                size=size,
                mtime_ns=int(disk_mtime or 0),
            )
        )
    return tuple(stamps)


def rescan_folder_for_file(path: str, rd_root: str) -> str:
    """Return the mark folder to rescan for one RD file, else its title folder.

    Never returns the RD root. A missing mark segment falls back to
    ``rd_root / <титул>`` when the first folder is four digits.

    Args:
        path: File path from the catalog or from the census.
        rd_root: Configured RD root.

    Returns:
        Folder under ``rd_root``, or empty when the file is not under it.
    """

    raw = str(path or "").strip()
    root = str(rd_root or "").strip()
    if not raw or not root or not path_is_under(raw, root):
        return ""
    if path_is_under(root, raw):
        return ""
    relative = os.path.relpath(raw, root)
    parts = [part for part in relative.split(os.sep) if part and part != "."]
    title = parts[0] if parts else ""
    folder = kit_rd_mark_folder_from_path(raw, root, title=title)
    if folder and path_is_under(folder, root) and not path_is_under(root, folder):
        return folder
    if len(title) == 4 and title.isdigit():
        candidate = os.path.normpath(os.path.join(root, title))
        if path_is_under(candidate, root) and not path_is_under(root, candidate):
            return candidate
    return ""


def dirty_mark_folders(
    baseline: Sequence[RdFileStamp],
    seen: Sequence[RdFileStamp],
    rd_root: str,
) -> tuple[str, ...]:
    """Return mark folders whose file size, date, or presence changed.

    Args:
        baseline: Present RD stamps from the last catalog snapshot.
        seen: Stamps just listed on disk.
        rd_root: RD root. The root itself is never returned.

    Returns:
        Unique folders, parents kept and nested children dropped.
    """

    base = {item.path_key: (item.size, item.mtime_ns) for item in baseline}
    seen_map = {item.path_key: item for item in seen}
    folders: list[str] = []
    seen_keys: set[str] = set()

    def add(path: str) -> None:
        folder = rescan_folder_for_file(path, rd_root)
        if not folder:
            return
        key = make_path_key(folder)
        if key in seen_keys:
            return
        seen_keys.add(key)
        folders.append(folder)

    for key, stamp in seen_map.items():
        previous = base.get(key)
        if previous != (stamp.size, stamp.mtime_ns):
            add(stamp.path)
    for key, stamp in ((item.path_key, item) for item in baseline):
        if key not in seen_map:
            add(stamp.path)
    return collapse_nested_folders(folders)


def take_rescan_batch(
    folders: Sequence[str],
    *,
    limit: int = AUTO_REFRESH_MAX_FOLDERS,
) -> tuple[tuple[str, ...], int]:
    """Return the first ``limit`` folders and how many were left out.

    Args:
        folders: Dirty mark folders, stable order.
        limit: Maximum folders to rescan in one cycle.

    Returns:
        Batch, and the count of folders postponed to the next census.
    """

    unique = collapse_nested_folders(folders)
    cap = max(0, int(limit))
    batch = unique[:cap]
    return batch, max(0, len(unique) - len(batch))


def collect_rd_file_stamps(
    rd_root: str | os.PathLike[str],
    *,
    skip_dirs: Iterable[str],
    sq_root: str | os.PathLike[str] | None = None,
    is_cancelled: CancelCallback | None = None,
) -> RdCensusResult:
    """List canonical RD files with size and mtime from the directory entry.

    Same name and skip-dir filters as the RD scan. File bodies are not read.
    A cancelled walk returns ``completed=False`` and must not be diffed:
    a partial listing would look like mass deletions.

    Args:
        rd_root: RD source root.
        skip_dirs: Directory-name tokens pruned during the walk.
        sq_root: SQ tree excluded from the RD walk.
        is_cancelled: Cooperative cancel check, called per directory.

    Returns:
        Stamps, completion flag, and directory-read errors.
    """

    root_text = os.path.normpath(str(rd_root))
    if not root_text or not os.path.isdir(root_text):
        return RdCensusResult((), False, ("корень РД недоступен",))
    sq_key = make_path_key(sq_root) if sq_root else ""
    stamps: list[RdFileStamp] = []
    errors: list[str] = []
    cancelled = False

    def cancelled_now() -> bool:
        return bool(is_cancelled and is_cancelled())

    def walk(current: str) -> None:
        nonlocal cancelled
        if cancelled_now():
            cancelled = True
            return
        try:
            entries = list(os.scandir(current))
        except OSError as exc:
            errors.append(f"{current}: {exc}")
            return
        directories: list[os.DirEntry[str]] = []
        files: list[os.DirEntry[str]] = []
        for entry in entries:
            try:
                is_dir = entry.is_dir(follow_symlinks=False)
            except OSError:
                continue
            if is_dir:
                directories.append(entry)
            else:
                files.append(entry)
        for entry in directories:
            name = entry.name
            if _is_skipped_dir(name, skip_dirs):
                continue
            child = entry.path
            if sq_key and make_path_key(child) == sq_key:
                continue
            walk(child)
            if cancelled:
                return
        names = [entry.name for entry in files]
        is_media_dir = is_package_media_folder(os.path.basename(current))
        has_pdf = any(name.casefold().endswith(".pdf") for name in names)
        in_issued_package = has_canonical_rd_issued_path(
            os.path.join(current, "_"), root_text
        )
        allow_source = is_media_dir or has_pdf or in_issued_package
        for entry in files:
            if cancelled_now():
                cancelled = True
                return
            name = entry.name
            if _is_excel_lock_name(name):
                continue
            if _candidate_kind(name, sq_only=False, allow_source=allow_source) is None:
                continue
            if not matches_agcc_filename(name):
                continue
            path = entry.path
            if not has_canonical_rd_issued_path(path, root_text):
                continue
            try:
                stat = entry.stat(follow_symlinks=False)
            except OSError as exc:
                errors.append(f"{path}: {exc}")
                continue
            stamps.append(
                RdFileStamp(
                    path=path,
                    path_key=make_path_key(path),
                    size=int(stat.st_size),
                    mtime_ns=int(stat.st_mtime_ns),
                )
            )

    walk(root_text)
    return RdCensusResult(
        tuple(stamps),
        completed=not cancelled,
        errors=tuple(errors),
    )
