"""Deterministic newest-to-oldest delta overlay for RD files."""

from __future__ import annotations

from collections import defaultdict
from collections.abc import Iterable
from pathlib import Path

from rd_catalog.models import (
    CollisionKind,
    FileKind,
    OverlayCollision,
    OverlayEntry,
    OverlayResult,
    OverlayStatus,
    ParsedFile,
    ParseStatus,
    SourceKind,
)
from rd_catalog.parse import record_has_canonical_layout


OVERLAY_ALGORITHM_VERSION = 3

# Highest-wins order: missing < V (cancelled) < S (superseded) < 0 < 1 < …
# Letter tokens must stay below every numeric revision so a cancelled file
# cannot hide a working ``0-AN02`` in kit / folder "current revision".
_RANK_MISSING = -3
_RANK_LETTER = {"V": -2, "S": -1}


def document_key_text(file: ParsedFile) -> str | None:
    """Serialize the prescribed PDF or MTO identity key.

    Args:
        file: Parsed catalog file.

    Returns:
        Stable key text or ``None`` for an unkeyed file.
    """

    key = file.document_key
    if isinstance(key, tuple):
        return f"mto:{key[0]}|{key[1]}"
    if isinstance(key, str):
        return f"pdf:{key}"
    return None


def revision_rank(value: str | None, appendix: str | None) -> tuple[int, int, str]:
    """Return a comparable rank for a revision / AN pair.

    Numeric tokens rank as their integer value. Letter tokens ``V`` / ``S``
    rank below every numeric revision (a cancelled ``V`` must not outrank
    ``0-AN02``). ``S`` still ranks above ``V``. Missing and unknown tokens
    rank below both letters.

    Args:
        value: Revision token such as ``01``, ``0``, ``V``.
        appendix: Optional AN digits without the ``AN`` prefix.

    Returns:
        ``(revision_int, appendix_int, casefolded_raw)``. Missing rank is
        ``(_RANK_MISSING, -1, "")``.
    """

    return _revision_rank(value, appendix)


def _revision_rank(value: str | None, appendix: str | None) -> tuple[int, int, str]:
    if value is None:
        return (_RANK_MISSING, -1, "")
    if value.isdigit():
        revision_int = int(value)
    else:
        revision_int = _RANK_LETTER.get(value.upper(), _RANK_MISSING)
    appendix_rank = int(appendix) if appendix and appendix.isdigit() else -1
    return (revision_int, appendix_rank, value.casefold())


def _newest_sort_key(file: ParsedFile) -> tuple[int, tuple[int, int, str], int, str]:
    transfer = file.transfer
    sequence = transfer.sequence if transfer and transfer.sequence is not None else -1
    return (
        sequence,
        _revision_rank(file.revision, file.appendix),
        file.mtime_ns,
        file.path_key,
    )


def _add_order_conflicts(
    key: str,
    files: list[ParsedFile],
    result: OverlayResult,
) -> None:
    ordered = sorted(files, key=_newest_sort_key, reverse=True)
    for index, newer in enumerate(ordered):
        new_transfer = newer.transfer
        if not new_transfer or new_transfer.sequence is None:
            continue
        for older in ordered[index + 1 :]:
            old_transfer = older.transfer
            if not old_transfer or old_transfer.sequence is None:
                continue
            if new_transfer.sequence <= old_transfer.sequence:
                continue
            newer_rank = _revision_rank(newer.revision, newer.appendix)
            older_rank = _revision_rank(older.revision, older.appendix)
            if newer_rank[0] >= 0 and older_rank[0] >= 0 and newer_rank < older_rank:
                result.collisions.append(
                    OverlayCollision(
                        kind=CollisionKind.TRANSFER_ORDER_CONFLICT,
                        message=(
                            f"Передача {new_transfer.sequence:02d} старше по ревизии "
                            f"файла ({_file_rev_text(newer)}), чем передача "
                            f"{old_transfer.sequence:02d} ({_file_rev_text(older)}). "
                            f"Текущей остаётся {new_transfer.sequence:02d}; "
                            "проверьте, ушла ли заказчику более высокая ревизия."
                        ),
                        path_keys=(newer.path_key, older.path_key),
                        document_key=key,
                    )
                )
            if newer.mtime_ns < older.mtime_ns:
                result.collisions.append(
                    OverlayCollision(
                        kind=CollisionKind.TRANSFER_MTIME_CONFLICT,
                        message=(
                            f"Передача {new_transfer.sequence:02d} новее по номеру, "
                            f"но файл старше по дате, чем в передаче "
                            f"{old_transfer.sequence:02d}."
                        ),
                        path_keys=(newer.path_key, older.path_key),
                        document_key=key,
                    )
                )


def _file_rev_text(file: ParsedFile) -> str:
    if not file.revision:
        return "—"
    return (
        f"{file.revision}-AN{file.appendix}" if file.appendix else str(file.revision)
    )


def build_overlay(
    files: Iterable[ParsedFile],
    file_kind: FileKind,
    *,
    ignored_path_keys: Iterable[str] = (),
) -> OverlayResult:
    """Build one RD delta overlay, newest transfer first.

    Newest means highest transfer sequence, then highest filename revision,
    then later file ``mtime``. Folder ``рев.*`` text is not a ranking signal.
        Duplicate copies of the same revision inside one transfer stay as a
        visible ``DUP_SAME_REVISION`` collision; the later ``mtime`` remains
        selectable so a leftover nested copy cannot hide the issued file.
        A later sequence stays ``detected_current`` even when an older sequence
    has a higher filename revision or a newer ``mtime``; those cases are
    non-fatal review collisions (``TRANSFER_ORDER_CONFLICT`` /
    ``TRANSFER_MTIME_CONFLICT``).

    Args:
        files: Parsed files from any sources.
        file_kind: Either PDF or MTO XLSX; kinds are always overlaid separately.
        ignored_path_keys: User-ignored path identities excluded from selection.

    Returns:
        Current files, per-file states, and non-fatal collisions.
    """

    result = OverlayResult(file_kind=file_kind)
    ignored = {path_key.casefold() for path_key in ignored_path_keys}
    grouped: dict[str, list[ParsedFile]] = defaultdict(list)

    for file in files:
        if file.source is not SourceKind.RD or file.file_kind is not file_kind:
            continue
        key = document_key_text(file)
        if file.parse_status is not ParseStatus.PARSED or key is None:
            result.collisions.append(
                OverlayCollision(
                    kind=CollisionKind.UNPARSED_FILE,
                    message=file.parse_error or "File has no document key",
                    path_keys=(file.path_key,),
                    document_key=key,
                )
            )
            continue
        if not file.transfer or file.transfer.parse_status is not ParseStatus.PARSED:
            result.collisions.append(
                OverlayCollision(
                    kind=CollisionKind.UNPARSED_FOLDER,
                    message=(
                        file.transfer.error
                        if file.transfer
                        else "Containing transfer folder is unknown"
                    )
                    or "Transfer folder is unparsed",
                    path_keys=(file.path_key,),
                    document_key=key,
                )
            )
            continue
        grouped[key].append(file)

    for key in sorted(grouped):
        candidates = sorted(grouped[key], key=_newest_sort_key, reverse=True)
        _add_order_conflicts(key, candidates, result)
        selectable: list[ParsedFile] = []
        ambiguous_keys: set[str] = set()
        same_transfer: dict[
            tuple[int | None, str, str | None, str | None], list[ParsedFile]
        ] = defaultdict(list)

        for file in candidates:
            if file.path_key.casefold() in ignored:
                result.entries.append(
                    OverlayEntry(file, key, OverlayStatus.IGNORED)
                )
                continue
            transfer = file.transfer
            same_transfer[
                (
                    transfer.sequence if transfer else None,
                    (transfer.normalized_name or "").casefold() if transfer else "",
                    file.revision,
                    file.appendix,
                )
            ].append(file)

        for duplicate_group in same_transfer.values():
            if len(duplicate_group) > 1:
                ranked = sorted(
                    duplicate_group,
                    key=lambda item: (int(item.mtime_ns), item.path_key),
                    reverse=True,
                )
                duplicate_paths = tuple(file.path_key for file in ranked)
                ambiguous_keys.update(file.path_key for file in ranked[1:])
                result.collisions.append(
                    OverlayCollision(
                        kind=CollisionKind.DUP_SAME_REVISION,
                        message=(
                            "Multiple files have the same key in one "
                            "transfer/revision; later mtime stays current"
                        ),
                        path_keys=duplicate_paths,
                        document_key=key,
                    )
                )

        for file in candidates:
            if file.path_key.casefold() in ignored:
                continue
            if file.path_key in ambiguous_keys:
                result.entries.append(
                    OverlayEntry(file, key, OverlayStatus.AMBIGUOUS)
                )
            else:
                selectable.append(file)

        if selectable:
            current = selectable[0]
            result.current[key] = current
            result.entries.append(
                OverlayEntry(current, key, OverlayStatus.DETECTED_CURRENT)
            )
            result.entries.extend(
                OverlayEntry(file, key, OverlayStatus.SUPERSEDED)
                for file in selectable[1:]
            )

    result.entries.sort(
        key=lambda entry: (
            entry.document_key,
            -_newest_sort_key(entry.file)[0],
            entry.file.path_key,
        )
    )
    return result


def build_rd_overlays(
    files: Iterable[ParsedFile],
    *,
    ignored_path_keys: Iterable[str] = (),
    rd_root: str | Path | None = None,
) -> tuple[OverlayResult, OverlayResult]:
    """Build independent PDF and MTO overlays.

    Args:
        files: Parsed files from all sources.
        ignored_path_keys: User-ignored path identities.
        rd_root: RD source root. When set, non-canonical RD files are
            excluded via :func:`record_has_canonical_layout`. Empty or
            ``None`` keeps the unfiltered input (unit-test fixtures).

    Returns:
        ``(pdf_overlay, mto_overlay)``.
    """

    return build_rd_overlays_as_of(
        files,
        max_sequence=None,
        ignored_path_keys=ignored_path_keys,
        rd_root=rd_root,
    )


def build_rd_overlays_as_of(
    files: Iterable[ParsedFile],
    *,
    max_sequence: int | None,
    ignored_path_keys: Iterable[str] = (),
    rd_root: str | Path | None = None,
) -> tuple[OverlayResult, OverlayResult]:
    """Build PDF and MTO overlays using only transfers up to ``max_sequence``.

    Ranking is the existing newest-to-oldest overlay (sequence, then filename
    revision, then mtime). This helper only filters the input; it does not
    duplicate ``_newest_sort_key``.

    Args:
        files: Parsed files from all sources.
        max_sequence: Inclusive transfer-sequence bound. ``None`` keeps every
            file (same as :func:`build_rd_overlays`). Files whose transfer
            sequence is ``None`` are kept — they already sort as ``-1`` and
            cannot outrank a sequenced file.
        ignored_path_keys: User-ignored path identities.
        rd_root: RD source root. When set, non-canonical RD files are
            excluded via :func:`record_has_canonical_layout`.

    Returns:
        ``(pdf_overlay, mto_overlay)``, same shape as :func:`build_rd_overlays`.
    """

    materialized = list(files)
    root = str(rd_root or "").strip()
    if root:
        materialized = [
            file
            for file in materialized
            if record_has_canonical_layout(file, root)
        ]
    if max_sequence is not None:
        materialized = [
            file
            for file in materialized
            if not _sequence_exceeds_bound(file, max_sequence)
        ]
    return (
        build_overlay(
            materialized, FileKind.PDF, ignored_path_keys=ignored_path_keys
        ),
        build_overlay(
            materialized, FileKind.MTO_XLSX, ignored_path_keys=ignored_path_keys
        ),
    )


def _sequence_exceeds_bound(file: ParsedFile, max_sequence: int) -> bool:
    transfer = file.transfer
    if transfer is None or transfer.sequence is None:
        return False
    return transfer.sequence > max_sequence
