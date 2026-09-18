"""Offscreen GUI checks for the approval-mail tab (no CatalogWindow)."""

from __future__ import annotations

import os
import struct
import sys
import tempfile
from pathlib import Path
from unittest.mock import patch

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from PySide6.QtCore import QByteArray, QMimeData, QUrl
from PySide6.QtWidgets import QApplication

from rd_catalog.approval_mail import parse_approval_mail_text
from rd_catalog.approval_mail_preview import (
    build_mail_previews,
    comment_lookup_from_kits,
)
from rd_catalog.approval_mail_tab import (
    ApprovalMailTab,
    _COL_F_AFTER,
    _COL_F_BEFORE,
    _COL_MTO,
    _COL_REV,
    _COL_STAGE,
    _COL_TRM,
    _COL_WRITE,
    _HEADERS,
    _ROLE_DIFF_SPANS,
)
from rd_catalog.sheet_de_sync_dialog import SHEET_DE_SYNC_BUTTON
from rd_catalog.approval_mail_log import ingest_log_path
from rd_catalog.google_f_write import GoogleWriteResult
from rd_catalog.kits import parse_google_matrix
from rd_catalog.outlook_drop import build_file_group_descriptor_w
from rd_catalog.outlook_drop_qt import (
    DroppedMsg,
    extract_dropped_messages,
    mime_has_approval_mail,
)

_HEADER = ["титул", "ИД/ТО/МДЗ/РД", "Наименование", "Тип/ Стадия", "Статус", "Комментарий"]
_SOT_F = "02.12.2024 код А на рев. 04 AGCC-BCC-TRM-000582"
_NOTIFICATION = """
Transmittal Date	09.09.2026 12:00:00
1. AGCC.287-1600-SOT.OD-0001 \\ 04 \\ IFC - x \\ x \\ \\ Offsite facilities \\ A - Замечания отсутствуют
"""


class _ObservedMimeData(QMimeData):
    """Record payload reads made while only detecting a drag."""

    def __init__(self) -> None:
        super().__init__()
        self.data_calls = 0

    def data(self, mime_type: str) -> QByteArray:
        self.data_calls += 1
        return super().data(mime_type)


def _private_message_payload() -> bytes:
    store_id = b"\x10\x20\x30"
    entry_id = b"\xAA\xBB\xCC\xDD"
    message_class = b"IPM.Note"
    subject = "one".encode("utf-16le")
    return b"".join(
        (
            struct.pack("<I", 2),
            b"\x01\x02",
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


def _app() -> QApplication:
    existing = QApplication.instance()
    if existing is not None:
        return existing
    return QApplication([])


def _writable_mail():
    return parse_approval_mail_text(
        subject="Notification of Transmittal Number 'AGCC-BCC-TRM-000999'",
        body=_NOTIFICATION,
    )


def main() -> None:
    """Drop MIME, F columns, write without dialog, duplicates, delete."""

    app = _app()
    kits, _stats = parse_google_matrix(
        [_HEADER, ["1600", "SOT", "x", "Рев. 04", "ok", _SOT_F]]
    )
    lookup = comment_lookup_from_kits(kits)
    mail = _writable_mail()
    assert mail.error == ""
    rows = build_mail_previews([mail], comment_lookup=lookup)
    assert rows[0].writable

    tab = ApprovalMailTab()
    assert list(_HEADERS)[ _COL_REV ] == "Рев."
    assert _HEADERS[_COL_TRM] == "TRM"
    assert _HEADERS[_COL_MTO] == "MTO"
    assert _HEADERS[_COL_WRITE] == "Запись"
    assert tab._table.columnWidth(_COL_STAGE) == 150
    assert tab._table.columnWidth(_COL_TRM) == 150
    assert tab._de_sync_button.text() == SHEET_DE_SYNC_BUTTON
    assert tab._open_log_button.isEnabled() is False
    de_sync_clicks: list[int] = []
    tab.de_sync_requested.connect(lambda: de_sync_clicks.append(1))
    tab._de_sync_button.click()
    assert de_sync_clicks == [1]
    tab.set_lookups(comment_lookup=lookup, kit_from_transmittal=None)
    captured: list[object] = []
    tab.write_requested.connect(captured.append)

    tab.ingest_dropped(
        [
            DroppedMsg(
                name="bad.pdf",
                error="Outlook передал файлы, но среди них нет .msg.",
                mime_formats=("FileGroupDescriptorW", "FileContents"),
            )
        ]
    )
    app.processEvents()
    assert tab._table.rowCount() == 1
    dump = tab.dump_text_for_row(0)
    assert "Дамп письма Outlook" in dump
    assert "FileGroupDescriptorW" in dump
    assert "bad.pdf" in dump
    preview = tab.text_preview_for_row(0)
    assert "Текст письма" in preview
    assert "Outlook передал файлы" in preview
    app.processEvents()
    assert tab._table.rowCount() == 1
    assert tab.preview_rows()[0].writable is False
    assert tab._write_button.isEnabled() is False

    with tempfile.TemporaryDirectory() as raw:
        dummy = Path(raw) / "letter.msg"
        dummy.write_bytes(b"not-a-msg")
        mime = QMimeData()
        mime.setUrls([QUrl.fromLocalFile(str(dummy))])
        assert mime_has_approval_mail(mime)
        tab.ingest_mime(mime)
        app.processEvents()
        assert tab._table.rowCount() == 2
        assert any(row.error for row in tab.preview_rows())

        tab._clear()
        item = DroppedMsg(name="letter.msg", path=str(dummy), payload=b"not-a-msg")
        tab.ingest_dropped([item, item])
        app.processEvents()
        assert tab._table.rowCount() == 1
        assert "дубл" in tab._status.text()

        same_bytes = DroppedMsg(name="a.msg", payload=b"same-bytes")
        twin_bytes = DroppedMsg(name="b.msg", payload=b"same-bytes")
        tab._clear()
        tab.ingest_dropped([same_bytes, twin_bytes])
        app.processEvents()
        assert tab._table.rowCount() == 1
        assert "дубл" in tab._status.text()

    descriptor = build_file_group_descriptor_w(("one.msg", "two.msg"))
    outlook = QMimeData()
    outlook.setData("FileGroupDescriptorW", QByteArray(descriptor))
    outlook.setData("FileContents", QByteArray(b"fake-msg-bytes"))
    extracted = extract_dropped_messages(outlook)
    assert extracted[0].name == "one.msg"
    assert extracted[0].payload == b"fake-msg-bytes"
    assert "несколько писем" in extracted[1].error

    indexed_outlook = _ObservedMimeData()
    indexed_outlook.setData("FileGroupDescriptorW", QByteArray(descriptor))
    indexed_outlook.setData("FileContents", QByteArray())
    indexed_outlook.setData(
        'application/x-qt-windows-mime;value="FileContents";index=0',
        QByteArray(b"indexed-msg-bytes"),
    )
    assert mime_has_approval_mail(indexed_outlook)
    assert indexed_outlook.data_calls == 0
    indexed = extract_dropped_messages(indexed_outlook)
    assert indexed[0].payload == b"indexed-msg-bytes"
    assert any("FileGroupDescriptorW" in fmt for fmt in indexed[0].mime_formats)

    private_outlook = QMimeData()
    private_outlook.setData("FileGroupDescriptorW", QByteArray(descriptor))
    private_outlook.setData("FileContents", QByteArray())
    private_outlook.setData(
        "RenPrivateMessages", QByteArray(_private_message_payload())
    )
    private = extract_dropped_messages(private_outlook)
    assert private[0].outlook_store_id == "102030"
    assert private[0].outlook_entry_id == "AABBCCDD"
    assert not private[0].error
    assert private[0].mime_sizes

    broken_private = QMimeData()
    broken_private.setData("FileGroupDescriptorW", QByteArray(descriptor))
    broken_private.setData("FileContents", QByteArray())
    broken_private.setData("RenPrivateMessages", QByteArray(b"\x01\x00"))
    broken = extract_dropped_messages(broken_private)
    assert broken[0].error
    assert "RenPrivate hex:" in broken[0].debug_note or "разбор RenPrivate" in broken[0].debug_note

    tab._clear()
    with patch.object(tab, "_parse_dropped", return_value=mail) as parse_dropped:
        tab.ingest_dropped(
            [
                DroppedMsg(name="one.msg", path="letter-one"),
                DroppedMsg(name="two.msg", path="letter-two"),
            ]
        )
        app.processEvents()
        assert tab._table.rowCount() == 1
        assert "дубл" in tab._status.text()
        assert "09.09.2026" in tab._table.item(0, _COL_F_AFTER).text()
        assert "02.12.2024" in tab._table.item(0, _COL_F_BEFORE).text()
        assert "на рев. 04" in tab._table.item(0, _COL_F_AFTER).text()
        rev_combo = tab._table.cellWidget(0, _COL_REV)
        assert rev_combo is not None
        rev_combo.setEditText("05")
        tab._on_part_edited(0, "od_revision", rev_combo)
        app.processEvents()
        assert "на рев. 05" in tab._table.item(0, _COL_F_AFTER).text()
        mto_combo = tab._table.cellWidget(0, _COL_MTO)
        assert mto_combo is not None
        after_spans = tab._table.item(0, _COL_F_AFTER).data(_ROLE_DIFF_SPANS) or []
        assert after_spans
        after_text = tab._table.item(0, _COL_F_AFTER).text()
        marked = "".join(after_text[int(a) : int(b)] for a, b in after_spans)
        assert "09.09.2026" in marked
        assert "02.12.2024" not in marked
        before_spans = tab._table.item(0, _COL_F_BEFORE).data(_ROLE_DIFF_SPANS) or []
        assert not before_spans
        assert tab._write_button.isEnabled() is True
        parse_count = parse_dropped.call_count
        assert parse_count >= 1
        tab._start_write()
        app.processEvents()
        assert captured
        jobs = captured[-1]
        assert jobs[0].title == "1600"
        assert jobs[0].mark == "SOT"
        tab.apply_write_results(
            [
                GoogleWriteResult(
                    title="1600",
                    mark="SOT",
                    row_index=4,
                    sheet_title="Лист1",
                    comment_before=_SOT_F,
                    comment_after=rows[0].patch.comment_after,
                    update_de=True,
                )
            ]
        )
        app.processEvents()
        assert parse_dropped.call_count == parse_count
        assert tab._table.rowCount() == 0
        tab.set_lookups(
            comment_lookup=comment_lookup_from_kits(kits),
            kit_from_transmittal=tab._kit_lookup,
        )
        app.processEvents()
        assert parse_dropped.call_count == parse_count
        tab.set_lookups(comment_lookup=lookup, kit_from_transmittal=object())
        app.processEvents()
        assert parse_dropped.call_count == parse_count

        tab.ingest_dropped([DroppedMsg(name="again.msg", path="letter-again")])
        tab.apply_write_results(
            [
                GoogleWriteResult(
                    title="1600",
                    mark="SOT",
                    row_index=4,
                    sheet_title="Лист1",
                    comment_before=_SOT_F,
                    comment_after=rows[0].patch.comment_after,
                    update_de=True,
                )
            ]
        )
        assert tab.preview_rows()[0].write_status != "written"
        tab._pending_indexes = (0,)
        tab.apply_write_results(
            [
                GoogleWriteResult(
                    title="1600",
                    mark="SOT",
                    row_index=4,
                    sheet_title="Лист1",
                    comment_before=_SOT_F,
                    comment_after=rows[0].patch.comment_after,
                    update_de=True,
                    error="нет строки",
                )
            ]
        )
        assert tab._table.rowCount() == 1
        assert tab.preview_rows()[0].write_status == "failed"
        tab._pending_indexes = (0,)
        tab.apply_write_results(
            [
                GoogleWriteResult(
                    title="1600",
                    mark="SOT",
                    row_index=4,
                    sheet_title="Лист1",
                    comment_before=_SOT_F,
                    comment_after=rows[0].patch.comment_after,
                    update_de=True,
                )
            ]
        )
        app.processEvents()
        assert tab._table.rowCount() == 0

    already_lookup = comment_lookup_from_kits(
        parse_google_matrix(
            [
                _HEADER,
                [
                    "1600",
                    "SOT",
                    "x",
                    "Рев. 04",
                    "РД Согласовано",
                    rows[0].patch.comment_after,
                ],
            ]
        )[0]
    )
    tab.set_lookups(comment_lookup=already_lookup, kit_from_transmittal=None)
    captured.clear()
    with patch.object(tab, "_parse_dropped", return_value=mail):
        tab.ingest_dropped([DroppedMsg(name="old.msg", path="already-in-sheet")])
        app.processEvents()
        assert tab.preview_rows()[0].already_complete
        assert tab.preview_rows()[0].already_recorded
        assert tab.preview_rows()[0].writable is False
        assert tab._write_button.isEnabled() is False
        assert tab._clear_written_button.isEnabled() is True
        assert "уже в F" in tab._table.item(0, _COL_WRITE).text()
        tab._clear_written()
        app.processEvents()
        assert tab._table.rowCount() == 0
        assert tab._clear_written_button.isEnabled() is False

    tab._clear()
    real_parse = tab._parse_dropped

    def _parse_keep_errors(dropped: DroppedMsg):
        if dropped.path == "already-in-sheet":
            return mail
        return real_parse(dropped)

    with patch.object(tab, "_parse_dropped", side_effect=_parse_keep_errors):
        tab.ingest_dropped(
            [
                DroppedMsg(name="bad.pdf", error="Outlook не передал .msg."),
                DroppedMsg(name="old.msg", path="already-in-sheet"),
            ]
        )
        app.processEvents()
        assert tab._table.rowCount() == 2
        tab._clear_written()
        app.processEvents()
        assert tab._table.rowCount() == 1
        assert tab.preview_rows()[0].error
        assert tab._clear_written_button.isEnabled() is False

    with tempfile.TemporaryDirectory() as raw:
        root = Path(raw)
        nested = root / "17_Даты TRM" / "inner"
        nested.mkdir(parents=True)
        (root / "top.msg").write_bytes(b"not-a-msg")
        (nested / "deep.msg").write_bytes(b"not-a-msg")
        tab._clear()
        tab._subfolders_box.setChecked(False)
        tab.ingest_paths([root])
        app.processEvents()
        assert tab._table.rowCount() == 1
        tab._clear()
        tab._subfolders_box.setChecked(True)
        tab.ingest_paths([root])
        app.processEvents()
        assert tab._table.rowCount() == 2

    with tempfile.TemporaryDirectory() as raw:
        runtime = Path(raw)
        logged = ApprovalMailTab(runtime_dir=runtime)
        logged.set_lookups(comment_lookup=lookup, kit_from_transmittal=None)
        with patch.object(logged, "_parse_dropped", return_value=mail):
            logged.ingest_dropped(
                [DroppedMsg(name="logged.msg", path="letter-log")]
            )
            app.processEvents()
        log_path = ingest_log_path(runtime)
        text = log_path.read_text(encoding="utf-8")
        assert "Текст письма" in text
        assert "logged.msg" in text or "1600" in text
        assert logged._open_log_button.isEnabled() is True
        opened: list[str] = []
        with patch(
            "rd_catalog.approval_mail_tab.open_path",
            side_effect=lambda path: opened.append(path) or (True, path),
        ):
            logged._open_ingest_log()
        assert opened == [str(log_path)]
        empty = ApprovalMailTab(runtime_dir=runtime / "empty")
        empty._open_log_button.setEnabled(True)
        shown: list[str] = []
        with patch(
            "rd_catalog.approval_mail_tab.QMessageBox.information",
            side_effect=lambda *args, **kwargs: shown.append(str(args)),
        ):
            empty._open_ingest_log()
        assert shown
        assert "ещё пуст" in shown[0]
        empty.close()
        logged.close()

    tab.close()
    print("approval_mail_tab ok")


if __name__ == "__main__":
    main()
