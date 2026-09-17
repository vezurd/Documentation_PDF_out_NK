"""Local checks for FileGroupDescriptorW and Explorer ``.msg`` collection."""

from __future__ import annotations

import struct
import sys
import tempfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from rd_catalog.outlook_drop import (
    OutlookDropError,
    build_file_group_descriptor_w,
    collect_msg_paths,
    is_msg_name,
    parse_file_group_descriptor_w,
    parse_file_group_descriptors,
    parse_ren_private_message,
)


def _private_message_payload() -> bytes:
    folder_id = b"\x01\x02"
    store_id = b"\x10\x20\x30"
    entry_id = b"\xAA\xBB\xCC\xDD"
    message_class = b"IPM.Note"
    subject = "Письмо".encode("utf-16le")
    return b"".join(
        (
            struct.pack("<I", len(folder_id)),
            folder_id,
            struct.pack("<I", len(store_id)),
            store_id,
            b"\x00" * 12,
            struct.pack("<I", 1),
            struct.pack("<I", 0),
            bytes((len(message_class),)),
            message_class,
            bytes((len(subject) // 2,)),
            subject,
            struct.pack("<I", len(entry_id)),
            entry_id,
        )
    )


def main() -> None:
    """Round-trip descriptor bytes and expand a folder of ``.msg`` files."""

    assert is_msg_name("letter.MSG")
    assert not is_msg_name("letter.pdf")
    names = ("Notification.msg", "RE ответ.msg")
    payload = build_file_group_descriptor_w(names)
    assert parse_file_group_descriptor_w(payload) == names
    infos = parse_file_group_descriptors(payload)
    assert [item.name for item in infos] == list(names)
    assert all(item.size == 0 for item in infos)
    try:
        parse_file_group_descriptor_w(b"\x01")
        raise AssertionError("short payload must fail")
    except OutlookDropError:
        pass
    try:
        parse_file_group_descriptor_w(b"\x02\x00\x00\x00" + b"\x00" * 10)
        raise AssertionError("truncated descriptors must fail")
    except OutlookDropError:
        pass

    private = parse_ren_private_message(_private_message_payload())
    assert private.store_id == "102030"
    assert private.entry_id == "AABBCCDD"
    assert private.subject == "Письмо"
    try:
        parse_ren_private_message(b"\x01")
        raise AssertionError("short private payload must fail")
    except OutlookDropError:
        pass

    with tempfile.TemporaryDirectory() as raw:
        root = Path(raw)
        (root / "keep.msg").write_bytes(b"msg")
        (root / "skip.txt").write_text("no", encoding="utf-8")
        nested = root / "nested"
        nested.mkdir()
        (nested / "inner.msg").write_bytes(b"inner")
        lone = root / "single.MSG"
        lone.write_bytes(b"one")
        collected = collect_msg_paths([root, lone, root / "missing.msg"])
        stems = sorted(path.name.casefold() for path in collected)
        assert stems == ["keep.msg", "single.msg"]
        assert collect_msg_paths([nested])[0].name == "inner.msg"
        recursive = collect_msg_paths([root], recursive=True)
        recursive_names = sorted(path.name.casefold() for path in recursive)
        assert recursive_names == ["inner.msg", "keep.msg", "single.msg"]
        assert collect_msg_paths([root], recursive=False) == collect_msg_paths(
            [root]
        )

    print("outlook_drop ok")


if __name__ == "__main__":
    main()
