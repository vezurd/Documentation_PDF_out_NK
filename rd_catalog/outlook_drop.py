"""Parse Outlook/Explorer drag payloads into ``.msg`` names and paths.

Qt-free. FileGroupDescriptorW layout matches the EisenhowerMatrix
``OutlookDataObjectReader`` (FILEDESCRIPTORW, pack 4, 260 WCHAR). Does not
use ``Outlook.Application`` / ``ActiveExplorer().Selection``.
"""

from __future__ import annotations

import os
import struct
from collections.abc import Iterable, Sequence
from dataclasses import dataclass
from pathlib import Path

_FILEDESCRIPTORW_SIZE = 592
_FILENAME_OFFSET = 72
_FILENAME_BYTES = 520  # 260 UTF-16 code units
_MAX_ITEMS = 10_000


class OutlookDropError(ValueError):
    """Malformed FileGroupDescriptorW or an unsupported drop payload."""


@dataclass(frozen=True, slots=True)
class OutlookPrivateMessage:
    """Identifiers of the first item in Outlook ``RenPrivateMessages``."""

    store_id: str
    entry_id: str
    subject: str


def is_msg_name(name: str) -> bool:
    """Return True when ``name`` looks like a saved Outlook message."""

    return Path(name).suffix.casefold() == ".msg"


_FD_FILESIZE = 0x00000040
_FILESIZE_HIGH_OFFSET = 64
_FILESIZE_LOW_OFFSET = 68


@dataclass(frozen=True, slots=True)
class FileDescriptorInfo:
    """One ``FILEDESCRIPTORW``: name and claimed size."""

    name: str
    size: int


def parse_file_group_descriptor_w(payload: bytes) -> tuple[str, ...]:
    """Parse ``FileGroupDescriptorW`` into file names.

    Args:
        payload: Raw CFSTR_FILEDESCRIPTORW bytes (count + descriptors).

    Returns:
        File names in drop order, including non-``.msg`` entries.

    Raises:
        OutlookDropError: Truncated or oversized descriptor stream.
    """

    return tuple(item.name for item in parse_file_group_descriptors(payload))


def parse_file_group_descriptors(payload: bytes) -> tuple[FileDescriptorInfo, ...]:
    """Parse ``FileGroupDescriptorW`` into names and claimed file sizes.

    Args:
        payload: Raw CFSTR_FILEDESCRIPTORW bytes (count + descriptors).

    Returns:
        Descriptors in drop order.

    Raises:
        OutlookDropError: Truncated or oversized descriptor stream.
    """

    data = bytes(payload or b"")
    if len(data) < 4:
        raise OutlookDropError("Повреждён FileGroupDescriptorW.")
    count = struct.unpack_from("<I", data, 0)[0]
    if count > _MAX_ITEMS:
        raise OutlookDropError("Слишком много перетаскиваемых элементов.")
    needed = 4 + count * _FILEDESCRIPTORW_SIZE
    if len(data) < needed:
        raise OutlookDropError("FileGroupDescriptorW имеет неверную длину.")
    items: list[FileDescriptorInfo] = []
    offset = 4
    for _ in range(count):
        raw_name = data[
            offset + _FILENAME_OFFSET : offset + _FILENAME_OFFSET + _FILENAME_BYTES
        ]
        flags = struct.unpack_from("<I", data, offset)[0]
        size = 0
        if flags & _FD_FILESIZE:
            high = struct.unpack_from("<I", data, offset + _FILESIZE_HIGH_OFFSET)[0]
            low = struct.unpack_from("<I", data, offset + _FILESIZE_LOW_OFFSET)[0]
            size = (high << 32) | low
        items.append(FileDescriptorInfo(name=_decode_wchar_z(raw_name), size=size))
        offset += _FILEDESCRIPTORW_SIZE
    return tuple(items)


def parse_ren_private_message(payload: bytes) -> OutlookPrivateMessage:
    """Read the first dragged Outlook item without using current selection.

    Outlook sometimes exposes an empty ``FileContents`` stream but provides
    the exact dragged item's MAPI identifiers in ``RenPrivateMessages``.

    Args:
        payload: Raw Outlook ``RenPrivateMessages`` bytes.

    Returns:
        Store ID, item Entry ID, and subject of the first dragged item.

    Raises:
        OutlookDropError: If the private stream is missing or malformed.
    """

    data = bytes(payload or b"")
    offset = 0

    def read(size: int, label: str) -> bytes:
        nonlocal offset
        if size < 0 or offset + size > len(data):
            raise OutlookDropError(f"Повреждён RenPrivateMessages: {label}.")
        value = data[offset : offset + size]
        offset += size
        return value

    def read_u32(label: str) -> int:
        return struct.unpack("<I", read(4, label))[0]

    folder_id_size = read_u32("длина FolderID")
    read(folder_id_size, "FolderID")
    store_id_size = read_u32("длина StoreID")
    store_id = read(store_id_size, "StoreID")
    read(12, "служебный заголовок")
    item_count = read_u32("число писем")
    if item_count < 1 or item_count > _MAX_ITEMS:
        raise OutlookDropError("RenPrivateMessages не содержит писем.")

    read(4, "SideEffects")
    class_size = read(1, "длина MessageClass")[0]
    read(class_size, "MessageClass")
    subject_chars = read(1, "длина темы")[0]
    subject = read(subject_chars * 2, "тема").decode(
        "utf-16le", errors="replace"
    )
    entry_id_size = read_u32("длина EntryID")
    entry_id = read(entry_id_size, "EntryID")
    if not store_id or not entry_id:
        raise OutlookDropError("RenPrivateMessages не содержит MAPI ID.")
    return OutlookPrivateMessage(
        store_id=store_id.hex().upper(),
        entry_id=entry_id.hex().upper(),
        subject=subject,
    )


def build_file_group_descriptor_w(names: Sequence[str]) -> bytes:
    """Serialize names as FileGroupDescriptorW (tests and fakes).

    Args:
        names: File names, typically ``letter.msg``.

    Returns:
        Bytes Outlook would put on the clipboard.
    """

    payload = bytearray(struct.pack("<I", len(names)))
    for name in names:
        descriptor = bytearray(_FILEDESCRIPTORW_SIZE)
        encoded = name.encode("utf-16le") + b"\x00\x00"
        if len(encoded) > _FILENAME_BYTES:
            raise OutlookDropError(f"Имя файла слишком длинное: {name!r}")
        descriptor[_FILENAME_OFFSET : _FILENAME_OFFSET + len(encoded)] = encoded
        payload.extend(descriptor)
    return bytes(payload)


def collect_msg_paths(
    sources: Iterable[str | Path],
    *,
    recursive: bool = False,
) -> tuple[Path, ...]:
    """Expand dropped files/folders into existing ``.msg`` paths.

    Args:
        sources: Files or directories from Explorer ``CF_HDROP``.
        recursive: When True, walk directory trees (UNC trees included).
            Default is only the chosen folder itself, not subfolders.

    Returns:
        Existing ``.msg`` files, folders expanded with a case-insensitive
        suffix match, sorted by path.
    """

    found: list[Path] = []
    seen: set[str] = set()
    for raw in sources:
        path = Path(raw)
        candidates: list[Path] = []
        if path.is_file() and is_msg_name(path.name):
            candidates.append(path)
        elif path.is_dir():
            if recursive:
                candidates.extend(_msg_files_under(path))
            else:
                candidates.extend(
                    item
                    for item in sorted(path.iterdir())
                    if item.is_file() and is_msg_name(item.name)
                )
        for item in candidates:
            key = str(item.resolve()) if item.exists() else str(item)
            if key in seen:
                continue
            seen.add(key)
            found.append(item)
    return tuple(found)


def _msg_files_under(root: Path) -> list[Path]:
    """Return ``.msg`` files under ``root``, skipping unreadable directories."""

    found: list[Path] = []

    def _ignore_walk_error(_exc: OSError) -> None:
        return

    try:
        walker = os.walk(root, followlinks=False, onerror=_ignore_walk_error)
        for dirpath, dirnames, filenames in walker:
            dirnames.sort(key=str.casefold)
            for name in sorted(filenames, key=str.casefold):
                if is_msg_name(name):
                    found.append(Path(dirpath) / name)
    except OSError:
        return found
    return found


def _decode_wchar_z(raw: bytes) -> str:
    text = raw.decode("utf-16le", errors="replace")
    return text.split("\x00", 1)[0]
