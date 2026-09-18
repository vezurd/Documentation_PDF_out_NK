"""Read-only, cancellable filesystem scanners for RD catalog sources."""

from __future__ import annotations

import os
from collections.abc import Callable, Iterable, Sequence
from pathlib import Path

from rd_catalog.config import CatalogConfig
from rd_catalog.models import (
    CollisionKind,
    FileKind,
    OverlayCollision,
    ParseStatus,
    ScanProgress,
    ScanRunStatus,
    ScanSummary,
    SourceError,
    SourceKind,
    SourceScanResult,
    TransferMetadata,
    make_path_key,
)
from rd_catalog.overlay import build_rd_overlays
from rd_catalog.parse import (
    _file_kind_from_name,
    find_transfer_in_path,
    has_canonical_rd_issued_path,
    is_package_media_folder,
    matches_agcc_filename,
    parse_catalog_file,
)
from rd_catalog.path_actions import path_is_under
from rd_catalog.skip_dirs import is_skipped_dir_name


ProgressCallback = Callable[[ScanProgress], None]
CancelCallback = Callable[[], bool]
ScanSubtree = str | Path | Sequence[str | Path]


def normalize_scan_subtrees(
    subtree: ScanSubtree | None,
) -> tuple[str, ...]:
    """Return unique normalized folder paths from a scan subtree spec.

    ``str`` and ``Path`` stay a single folder. A sequence walks each
    folder in one source result. Empty strings are dropped.

    Args:
        subtree: One folder, several folders, or ``None``.

    Returns:
        Normalized paths in first-seen order.
    """

    if subtree is None:
        return ()
    if isinstance(subtree, (str, Path)):
        items: Sequence[str | Path] = (subtree,)
    else:
        items = subtree
    out: list[str] = []
    seen: set[str] = set()
    for item in items:
        text = os.path.normpath(str(item)).strip()
        if not text:
            continue
        key = make_path_key(text)
        if key in seen:
            continue
        seen.add(key)
        out.append(text)
    return tuple(out)


def collapse_nested_folders(folders: Sequence[str]) -> tuple[str, ...]:
    """Keep parent folders and drop children already covered by a parent.

    Args:
        folders: Candidate walk roots.

    Returns:
        Folders sorted shorter-first, without nested duplicates.
    """

    unique = normalize_scan_subtrees(folders)
    kept: list[str] = []
    for folder in sorted(
        unique, key=lambda item: (len(make_path_key(item)), item.casefold())
    ):
        if any(path_is_under(folder, parent) for parent in kept):
            continue
        kept.append(folder)
    return tuple(kept)

_ROBOT_EXCLUDED_PARTS = (
    "old",
    "temp",
    "tmp",
    "backup",
    "archive",
    "архив",
    "старые",
    "test",
    "__результат_",
    "проверка",
)
_ROBOT_VALID_SUBFOLDERS = frozenset(
    {
        "POS",
        "SPP",
        "SOS",
        "SKUD",
        "SOT",
        "SOO",
        "KSB",
        "KSB1",
        "KSB2",
        "SOS1",
        "SKUD1",
        "SOT1",
        "KBI",
        "SAGD",
        "KSB3",
        "KSB4",
        "POS1",
        "SOO1",
        "POS2",
        "SOT2",
        "SOO2",
        "POS3",
        "SOT3",
        "SOO3",
        "POS4",
        "SOT4",
        "SOO4",
        "SOT5",
        "POS5",
        "SOO5",
        "SOT.1",
    }
)


def _is_cancelled(callback: CancelCallback | None) -> bool:
    return bool(callback and callback())


def _emit_progress(
    callback: ProgressCallback | None,
    source: SourceKind,
    path: str,
    files_seen: int,
    message: str = "",
) -> None:
    if callback:
        callback(
            ScanProgress(
                source=source,
                path=path,
                files_seen=files_seen,
                message=message,
            )
        )


def _is_excel_lock_name(name: str) -> bool:
    return Path(name).name.startswith("~$")


def _is_skipped_dir(name: str, skip_dirs: Iterable[str]) -> bool:
    return is_skipped_dir_name(name, skip_dirs)


def _find_transfer(root: str, source_root: str) -> TransferMetadata:
    return find_transfer_in_path(root, source_root)


def _candidate_kind(
    name: str,
    *,
    sq_only: bool = False,
    allow_source: bool = False,
) -> FileKind | None:
    kind = _file_kind_from_name(name)
    if kind is None or kind is FileKind.PDF:
        return kind
    if kind is FileKind.MTO_XLSX:
        return kind
    if sq_only:
        return None
    if kind is FileKind.SOURCE_EDITABLE and allow_source:
        return kind
    return None


def scan_document_source(
    root: str | Path,
    source: SourceKind,
    *,
    skip_dirs: Iterable[str],
    sq_root: str | Path | None = None,
    progress: ProgressCallback | None = None,
    is_cancelled: CancelCallback | None = None,
    subtree: ScanSubtree | None = None,
) -> SourceScanResult:
    """Scan RD or SQ using a pruned, read-only ``os.walk``.

    RD collects PDF, MTO XLSX, and source-editables, but stores only files
    on a canonical issued path
    (``title/mark/gate/NN_/[PDF|DWG|BBB|ПДФ]``). Source-editables are
    taken from a media folder, from a folder that also contains a PDF, or
    from the issued ``NN_…`` package root itself (BOE/BOM/BOQ xlsx next
    to MTO without a PDF sibling). Names that match AGCC but sit outside
    that layout are skipped. SQ collects PDF and MTO XLSX in any folder
    under the SQ root. Names that do not match
    ``AgccFilenamePatterns`` (``parse_strict`` then ``parse_loose``) are
    skipped.

    Args:
        root: Source root. Transfer detection always uses this path even
            when ``subtree`` is set.
        source: ``SourceKind.RD`` or ``SourceKind.SQ``.
        skip_dirs: Directory names/substrings pruned in-place.
        sq_root: SQ subtree excluded from an RD walk.
        progress: Optional callback invoked after each candidate.
        is_cancelled: Optional cooperative cancellation predicate.
        subtree: Optional folder or folders under ``root``. When set, only
            those folders are walked; missing-file detection in
            ``store_scan`` is scoped to their union. Nested folders
            collapse to the parent. ``result.root`` stays the configured
            source root.

    Returns:
        Source result retaining files and partial errors.

    Raises:
        ValueError: If called for the robot source.
    """

    if source is SourceKind.ROBOT:
        raise ValueError("Use scan_robot_source for the robot source")

    root_text = str(root)
    result = SourceScanResult(source=source, root=root_text)
    walk_roots: list[str]
    if subtree is not None:
        requested = collapse_nested_folders(normalize_scan_subtrees(subtree))
        if not requested:
            return result
        valid: list[str] = []
        for folder in requested:
            if not path_is_under(folder, root_text):
                result.errors.append(
                    SourceError(
                        source,
                        folder,
                        f"{source.value} subtree is outside source root",
                    )
                )
                continue
            valid.append(folder)
        if not valid:
            return result
        result.subtree = valid[0]
        result.subtrees = tuple(valid)
        walk_roots = valid
    else:
        if not os.path.isdir(root_text):
            result.errors.append(
                SourceError(
                    source,
                    root_text,
                    "source root is unavailable or not a directory",
                )
            )
            return result
        walk_roots = [root_text]

    sq_key = make_path_key(sq_root) if sq_root is not None else None

    def on_walk_error(error: OSError) -> None:
        result.errors.append(
            SourceError(source, error.filename or root_text, str(error))
        )

    for walk_text in walk_roots:
        if result.cancelled:
            break
        if not os.path.isdir(walk_text):
            continue
        for current_root, dirs, names in os.walk(
            walk_text, onerror=on_walk_error
        ):
            if _is_cancelled(is_cancelled):
                result.cancelled = True
                break

            kept_dirs: list[str] = []
            for directory in dirs:
                child = os.path.join(current_root, directory)
                if _is_skipped_dir(directory, skip_dirs):
                    continue
                if (
                    source is SourceKind.RD
                    and sq_key
                    and make_path_key(child) == sq_key
                ):
                    continue
                kept_dirs.append(directory)
            dirs[:] = kept_dirs

            transfer = _find_transfer(current_root, root_text)
            sq = source is SourceKind.SQ
            is_media_dir = is_package_media_folder(Path(current_root).name)
            has_pdf = any(n.casefold().endswith(".pdf") for n in names)
            in_issued_package = has_canonical_rd_issued_path(
                os.path.join(current_root, "_"), root_text
            )
            allow_source = (not sq) and (
                is_media_dir or has_pdf or in_issued_package
            )
            for name in names:
                if _is_cancelled(is_cancelled):
                    result.cancelled = True
                    break
                if _is_excel_lock_name(name):
                    result.excluded_files += 1
                    continue
                kind = _candidate_kind(
                    name, sq_only=sq, allow_source=allow_source
                )
                if kind is None:
                    continue
                if not matches_agcc_filename(name):
                    result.excluded_files += 1
                    continue

                path = os.path.join(current_root, name)
                if source is SourceKind.RD and not has_canonical_rd_issued_path(
                    path, root_text
                ):
                    result.excluded_files += 1
                    continue
                result.candidates_seen += 1
                try:
                    stat = os.stat(path, follow_symlinks=False)
                    parsed = parse_catalog_file(
                        path,
                        source,
                        size=stat.st_size,
                        mtime_ns=stat.st_mtime_ns,
                        transfer=transfer,
                    )
                except OSError as exc:
                    result.errors.append(SourceError(source, path, str(exc)))
                    result.collisions.append(
                        OverlayCollision(
                            kind=CollisionKind.SOURCE_ERROR,
                            message=str(exc),
                            path_keys=(make_path_key(path),),
                        )
                    )
                else:
                    result.files.append(parsed)
                    if parsed.parse_status is not ParseStatus.PARSED:
                        result.collisions.append(
                            OverlayCollision(
                                kind=CollisionKind.UNPARSED_FILE,
                                message=parsed.parse_error or "unparsed file",
                                path_keys=(parsed.path_key,),
                            )
                        )
                    if (
                        transfer.parse_status is not ParseStatus.PARSED
                        and parsed.file_kind is not FileKind.SOURCE_EDITABLE
                    ):
                        result.collisions.append(
                            OverlayCollision(
                                kind=CollisionKind.UNPARSED_FOLDER,
                                message=transfer.error or "unparsed transfer folder",
                                path_keys=(parsed.path_key,),
                            )
                        )
                _emit_progress(progress, source, path, result.candidates_seen)
            if result.cancelled:
                break
    return result


def _robot_structure_valid(
    relative_file: Path,
    *,
    flat_structure: bool,
) -> bool:
    if flat_structure:
        return True
    parent_parts = relative_file.parent.parts
    if len(parent_parts) not in (1, 2):
        return False
    if len(parent_parts[0]) != 4:
        return False
    return len(parent_parts) == 1 or parent_parts[1].upper() in _ROBOT_VALID_SUBFOLDERS


def scan_robot_source(
    root: str | Path,
    *,
    flat_structure: bool,
    progress: ProgressCallback | None = None,
    is_cancelled: CancelCallback | None = None,
    subtree: str | Path | None = None,
) -> SourceScanResult:
    """Scan robot MTO files with legacy ``filter_mto_files`` semantics.

    This narrow adapter keeps the old name, excluded-folder, and folder-layout
    rules while returning plain domain records and avoiding document factories.

    Args:
        root: Robot MTO root.
        flat_structure: Skip title/subfolder structure validation when true.
        progress: Optional callback invoked after each XLSX candidate.
        is_cancelled: Optional cooperative cancellation predicate.
        subtree: Optional folder under ``root``. When set, only that folder is
            walked; missing-file detection in ``store_scan`` is scoped to it.

    Returns:
        Robot source result with the count of filtered-out XLSX files.
        ``result.root`` stays the configured robot root even for a subtree walk.
    """

    root_path = Path(root)
    root_text = str(root_path)
    result = SourceScanResult(source=SourceKind.ROBOT, root=root_text)
    walk_text = root_text
    if subtree is not None:
        subtree_text = os.path.normpath(str(subtree))
        if not path_is_under(subtree_text, root_text):
            result.errors.append(
                SourceError(
                    SourceKind.ROBOT,
                    subtree_text,
                    "robot subtree is outside robot_root",
                )
            )
            return result
        result.subtree = subtree_text
        walk_text = subtree_text
        if not os.path.isdir(walk_text):
            return result
    elif not os.path.isdir(root_text):
        result.errors.append(
            SourceError(
                SourceKind.ROBOT,
                root_text,
                "source root is unavailable or not a directory",
            )
        )
        return result

    def on_walk_error(error: OSError) -> None:
        result.errors.append(
            SourceError(SourceKind.ROBOT, error.filename or root_text, str(error))
        )

    for current_root, dirs, names in os.walk(walk_text, onerror=on_walk_error):
        if _is_cancelled(is_cancelled):
            result.cancelled = True
            break
        dirs[:] = [
            directory
            for directory in dirs
            if not any(
                excluded in directory.casefold()
                for excluded in _ROBOT_EXCLUDED_PARTS
            )
        ]
        for name in names:
            if _is_cancelled(is_cancelled):
                result.cancelled = True
                break
            if _is_excel_lock_name(name):
                result.excluded_files += 1
                continue
            if not name.casefold().endswith(".xlsx"):
                continue
            path = Path(current_root, name)
            result.candidates_seen += 1
            folded_name = name.casefold()
            try:
                relative = path.relative_to(root_path)
            except ValueError:
                relative = Path(name)
            if (
                "mto" not in folded_name
                and "мто" not in folded_name
                or not _robot_structure_valid(
                    relative, flat_structure=flat_structure
                )
                or not matches_agcc_filename(name)
            ):
                result.excluded_files += 1
                _emit_progress(
                    progress,
                    SourceKind.ROBOT,
                    str(path),
                    result.candidates_seen,
                    "excluded by robot MTO filter",
                )
                continue
            try:
                stat = path.stat()
                parsed = parse_catalog_file(
                    path,
                    SourceKind.ROBOT,
                    size=stat.st_size,
                    mtime_ns=stat.st_mtime_ns,
                )
            except OSError as exc:
                result.errors.append(
                    SourceError(SourceKind.ROBOT, str(path), str(exc))
                )
            else:
                result.files.append(parsed)
                if parsed.parse_status is not ParseStatus.PARSED:
                    result.collisions.append(
                        OverlayCollision(
                            kind=CollisionKind.UNPARSED_FILE,
                            message=parsed.parse_error or "unparsed robot MTO",
                            path_keys=(parsed.path_key,),
                        )
                    )
            _emit_progress(
                progress, SourceKind.ROBOT, str(path), result.candidates_seen
            )
        if result.cancelled:
            break
    return result


def scan_catalog(
    config: CatalogConfig,
    *,
    sources: Iterable[SourceKind] | None = None,
    ignored_path_keys: Iterable[str] = (),
    progress: ProgressCallback | None = None,
    is_cancelled: CancelCallback | None = None,
    robot_subtree: str | Path | None = None,
    rd_subtree: ScanSubtree | None = None,
    sq_subtree: str | Path | None = None,
) -> ScanSummary:
    """Scan configured sources and calculate RD overlays.

    Args:
        config: Resolved catalog configuration.
        sources: Optional source subset; all sources are scanned by default.
        ignored_path_keys: User-ignored paths excluded from overlays.
        progress: Optional per-candidate progress callback.
        is_cancelled: Optional cooperative cancellation predicate.
        robot_subtree: Optional folder under ``robot_root``; only that folder
            is walked for the robot source.
        rd_subtree: Optional folder or folders under ``rd_root``; only
            those folders are walked for RD. After an SQ→RD move this is
            the gate (``Для передачи`` / ``На_отправку``), so sibling
            transfers of the same kit are restatted. A kit rescan with
            mixed-title files may pass the kit folder plus foreign
            title/mark folders. Overlay persistence rebuilds from all
            present RD files so other kits stay current.
        sq_subtree: Optional folder under ``sq_root``; only that folder is
            walked for SQ.

    Returns:
        Combined summary. One source failure produces ``PARTIAL`` while keeping
        successful source results.
    """

    requested = set(SourceKind) if sources is None else set(sources)
    summary = ScanSummary()
    if SourceKind.RD in requested:
        summary.sources[SourceKind.RD] = scan_document_source(
            config.rd_root,
            SourceKind.RD,
            skip_dirs=config.skip_dirs,
            sq_root=config.sq_root,
            progress=progress,
            is_cancelled=is_cancelled,
            subtree=rd_subtree,
        )
    if SourceKind.SQ in requested:
        summary.sources[SourceKind.SQ] = scan_document_source(
            config.sq_root,
            SourceKind.SQ,
            skip_dirs=config.skip_dirs,
            progress=progress,
            is_cancelled=is_cancelled,
            subtree=sq_subtree,
        )
    if SourceKind.ROBOT in requested:
        summary.sources[SourceKind.ROBOT] = scan_robot_source(
            config.robot_root,
            flat_structure=config.robot_flat_structure,
            progress=progress,
            is_cancelled=is_cancelled,
            subtree=robot_subtree,
        )

    if any(result.cancelled for result in summary.sources.values()):
        summary.status = ScanRunStatus.CANCELLED
    elif any(result.errors for result in summary.sources.values()):
        summary.status = ScanRunStatus.PARTIAL
    else:
        summary.status = ScanRunStatus.SUCCESS

    summary.rd_pdf_overlay, summary.rd_mto_overlay = build_rd_overlays(
        summary.files,
        ignored_path_keys=ignored_path_keys,
        rd_root=config.rd_root,
    )
    return summary
