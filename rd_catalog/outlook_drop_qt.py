"""Qt helper: pull ``.msg`` files from Explorer or Outlook MIME.

Does not call Outlook ``Selection``. FileContents is read from QMimeData
(first virtual file). Empty Qt streams fall back to the live OLE
``IDataObject`` (IStorage/IStream) captured during ``RegisterDragDrop``,
then to MAPI IDs from ``RenPrivateMessages``.
"""

from __future__ import annotations

from dataclasses import dataclass, replace
from pathlib import Path

from PySide6.QtCore import QMimeData, QUrl

from rd_catalog.outlook_drop import (
    OutlookDropError,
    collect_msg_paths,
    is_msg_name,
    parse_file_group_descriptors,
    parse_ren_private_message,
)
from rd_catalog.outlook_ole import (
    hook_status,
    last_drop_data_object,
    read_file_contents,
    read_ole_format,
)

_DESCRIPTOR_HINTS = ("FileGroupDescriptorW", "FileGroupDescriptor")
_CONTENTS_HINTS = ("FileContents",)
_PRIVATE_MESSAGE_HINTS = ("RenPrivateMessages", "RenPrivateLatestMessages")
_DUMP_MIME_LIMIT = 24
_HEX_LIMIT = 64


@dataclass(frozen=True, slots=True)
class DroppedMsg:
    """One dropped message: path, bytes, or exact Outlook MAPI identity."""

    name: str
    path: str = ""
    payload: bytes = b""
    outlook_entry_id: str = ""
    outlook_store_id: str = ""
    mime_formats: tuple[str, ...] = ()
    mime_sizes: tuple[tuple[str, int], ...] = ()
    debug_note: str = ""
    error: str = ""

    @property
    def usable(self) -> bool:
        return not self.error and (
            bool(self.path)
            or bool(self.payload)
            or bool(self.outlook_entry_id and self.outlook_store_id)
        )


def mime_has_approval_mail(mime: QMimeData | None) -> bool:
    """Return True when the MIME payload may contain approval letters."""

    if mime is None:
        return False
    # Do not request virtual-file bytes while Outlook is still dragging.
    # Rendering OLE data can block and this function is called on DragMove.
    if _mime_has_format(mime, _DESCRIPTOR_HINTS):
        return True
    if mime.hasUrls():
        for url in mime.urls():
            if not url.isLocalFile():
                continue
            path = Path(url.toLocalFile())
            if path.is_dir() or is_msg_name(path.name):
                return True
    return False


def extract_dropped_messages(
    mime: QMimeData | None,
    *,
    recursive: bool = False,
) -> tuple[DroppedMsg, ...]:
    """Collect ``.msg`` paths and Outlook virtual files from a drop.

    Args:
        mime: Qt MIME from ``QDropEvent`` / clipboard.
        recursive: Walk subfolders when the drop contains a directory.

    Returns:
        Dropped messages in order. Outlook extras without FileContents get
        ``error`` set instead of raising.
    """

    if mime is None:
        return ()
    formats = tuple(mime.formats())
    sizes = _mime_format_sizes(mime, formats)
    found: list[DroppedMsg] = []
    descriptor = _mime_bytes(mime, _DESCRIPTOR_HINTS)
    if descriptor:
        try:
            infos = parse_file_group_descriptors(descriptor)
        except OutlookDropError as exc:
            return (
                DroppedMsg(
                    name="",
                    error=str(exc),
                    mime_formats=formats,
                    mime_sizes=sizes,
                ),
            )
        contents = _mime_bytes(mime, _CONTENTS_HINTS, index=0)
        ole_note = ""
        if not contents:
            contents, ole_note = _ole_file_contents()
        msg_infos = [item for item in infos if is_msg_name(item.name)]
        if not msg_infos:
            return (
                DroppedMsg(
                    name=infos[0].name if infos else "",
                    error="Outlook передал файлы, но среди них нет .msg.",
                    mime_formats=formats,
                    mime_sizes=sizes,
                    debug_note=ole_note,
                ),
            )
        first = msg_infos[0]
        size_note = (
            f"дескриптор {first.size} байт" if first.size else "дескриптор без размера"
        )
        if contents:
            found.append(
                DroppedMsg(
                    name=first.name,
                    payload=contents,
                    debug_note=_join_notes(ole_note, size_note),
                )
            )
        else:
            found.append(
                _from_private_messages(
                    mime,
                    first.name,
                    extra_note=_join_notes(ole_note, size_note),
                )
            )
        for extra in msg_infos[1:]:
            found.append(
                DroppedMsg(
                    name=extra.name,
                    error=(
                        "Outlook передал несколько писем; прочитано только первое. "
                        "Сохраните остальные как .msg."
                    ),
                )
            )
        return tuple(
            replace(item, mime_formats=formats, mime_sizes=sizes)
            for item in found
        )
    if mime.hasUrls():
        sources = [
            Path(url.toLocalFile())
            for url in mime.urls()
            if url.isLocalFile()
        ]
        for path in collect_msg_paths(sources, recursive=recursive):
            found.append(DroppedMsg(name=path.name, path=str(path)))
    return tuple(
        replace(item, mime_formats=formats, mime_sizes=sizes) for item in found
    )


def urls_from_paths(paths: list[str | Path]) -> list[QUrl]:
    """Build ``file://`` URLs for tests."""

    return [QUrl.fromLocalFile(str(Path(path))) for path in paths]


def _from_private_messages(
    mime: QMimeData,
    name: str,
    *,
    extra_note: str,
) -> DroppedMsg:
    notes = [extra_note] if extra_note else []
    private_payload = _mime_bytes(mime, _PRIVATE_MESSAGE_HINTS)
    if not private_payload:
        ole_private, ole_note = _ole_private_messages()
        if ole_note:
            notes.append(ole_note)
        private_payload = ole_private
    else:
        notes.append(f"RenPrivateMessages Qt: {len(private_payload)} байт")
    if not private_payload:
        notes.append(hook_status() or "нет IDataObject")
        return DroppedMsg(
            name=name,
            error=(
                "Outlook не передал содержимое письма "
                "(FileContents/RenPrivateMessages). Сохраните .msg "
                "и перетащите файл из Проводника."
            ),
            debug_note=_join_notes(*notes),
        )
    try:
        private = parse_ren_private_message(private_payload)
    except OutlookDropError as exc:
        notes.append(f"разбор RenPrivate: {exc}")
        notes.append(f"RenPrivate hex: {private_payload[:_HEX_LIMIT].hex()}")
        return DroppedMsg(
            name=name,
            error=(
                "Outlook не передал содержимое письма "
                "(FileContents/RenPrivateMessages). Сохраните .msg "
                "и перетащите файл из Проводника."
            ),
            debug_note=_join_notes(*notes),
        )
    return DroppedMsg(
        name=name,
        outlook_entry_id=private.entry_id,
        outlook_store_id=private.store_id,
        debug_note=_join_notes(*notes, f"тема RenPrivate: {private.subject}"),
    )


def _ole_file_contents() -> tuple[bytes, str]:
    data_object = last_drop_data_object()
    if not data_object:
        status = hook_status()
        return b"", status or "OLE FileContents: нет IDataObject"
    payload, note = read_file_contents(data_object, index=0)
    return payload, f"OLE {note}" if note else ""


def _ole_private_messages() -> tuple[bytes, str]:
    data_object = last_drop_data_object()
    if not data_object:
        return b"", ""
    last_note = ""
    for name in _PRIVATE_MESSAGE_HINTS:
        payload, note = read_ole_format(data_object, name, index=-1)
        if payload:
            return payload, f"OLE {note}"
        if note:
            last_note = note
    return b"", last_note


def _mime_bytes(
    mime: QMimeData,
    hints: tuple[str, ...],
    *,
    index: int | None = None,
) -> bytes:
    """Read the first non-empty matching Windows MIME payload."""

    formats = list(mime.formats())
    candidates: list[str] = []
    for hint in hints:
        if index is not None:
            candidates.append(
                'application/x-qt-windows-mime;'
                f'value="{hint}";index={index}'
            )
        for fmt in formats:
            folded = fmt.casefold()
            if hint.casefold() not in folded:
                continue
            if index is not None and ";index=" in folded:
                if f";index={index}" not in folded:
                    continue
            candidates.append(fmt)
        candidates.append(hint)
    seen: set[str] = set()
    for candidate in candidates:
        key = candidate.casefold()
        if key in seen:
            continue
        seen.add(key)
        payload = bytes(mime.data(candidate))
        if payload:
            return payload
    return b""


def _mime_format_sizes(
    mime: QMimeData,
    formats: tuple[str, ...],
) -> tuple[tuple[str, int], ...]:
    """Byte sizes of advertised formats (for the Outlook dump)."""

    sizes: list[tuple[str, int]] = []
    for fmt in formats[:_DUMP_MIME_LIMIT]:
        sizes.append((fmt, len(bytes(mime.data(fmt)))))
    return tuple(sizes)


def _mime_has_format(mime: QMimeData, hints: tuple[str, ...]) -> bool:
    """Check advertised MIME names without rendering their OLE payload."""

    formats = tuple(mime.formats())
    return any(
        hint.casefold() in fmt.casefold()
        for hint in hints
        for fmt in formats
    )


def _join_notes(*parts: str) -> str:
    return "; ".join(part for part in parts if part)
