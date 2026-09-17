"""Dump text for an Outlook drop row (no Qt)."""

from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from rd_catalog.approval_mail import parse_approval_mail_text
from rd_catalog.approval_mail_dump import (
    DUMP_HEADER,
    TEXT_PREVIEW_HEADER,
    build_mail_text_preview,
    build_outlook_dump,
)
from rd_catalog.outlook_drop_qt import DroppedMsg


def main() -> None:
    """Format a drop dump with MIME names and parser fields."""

    mail = parse_approval_mail_text(
        subject="Notification of Transmittal Number 'AGCC-BCC-TRM-000001'",
        body=(
            "Transmittal Date\t02.12.2024 16:11:16\n"
            "1. AGCC.287-1600-SOT.OD-0001 \\ 04 \\ IFC \\ x \\ A - "
            "Замечания отсутствуют\n"
        ),
    )
    dump = build_outlook_dump(
        DroppedMsg(
            name="letter.msg",
            payload=b"fake",
            mime_formats=("FileGroupDescriptorW", "FileContents"),
            mime_sizes=(("FileContents", 4),),
            debug_note="OLE FileContents ISTORAGE 4 байт",
        ),
        mail,
    )
    assert dump.startswith(DUMP_HEADER)
    assert "Outlook FileContents" in dump
    assert "FileGroupDescriptorW" in dump
    assert "FileContents байт: 4" in dump
    assert "FileContents · 4 байт" in dump
    assert "OLE FileContents ISTORAGE" in dump
    assert "1600" in dump
    assert dump.endswith("\n")
    preview = build_mail_text_preview(
        DroppedMsg(name="letter.msg", error="нет тела"),
        mail,
    )
    assert preview.startswith(TEXT_PREVIEW_HEADER)
    assert "Тема: Notification of Transmittal" in preview
    assert "Титул–марка: 1600-SOT" in preview
    assert "— текст —" in preview
    print("approval_mail_dump ok")


if __name__ == "__main__":
    main()
