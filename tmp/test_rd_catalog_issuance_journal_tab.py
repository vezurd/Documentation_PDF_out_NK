"""Offscreen GUI checks for the issuance-journal decision combo."""

from __future__ import annotations

import os
import sys
from pathlib import Path
from unittest.mock import MagicMock, patch

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from PySide6.QtWidgets import QApplication, QComboBox, QWidget
from shiboken6 import delete, isValid

from rd_catalog.issuance_journal_tab import IssuanceJournalTab, _COL_DECISION
from rd_catalog.issuance_review import IssuanceJournalRow, send_identity_fingerprint
from rd_catalog.kits import IssuanceKit


def _app() -> QApplication:
    existing = QApplication.instance()
    if existing is not None:
        return existing
    return QApplication([])


def _sheet_row(**overrides: object) -> IssuanceJournalRow:
    payload = dict(
        title="2868",
        mark="SKUD",
        kind="send",
        source="issuance",
        revision_text="01-AN02",
        send_date="27.03.2026",
        send_date_sortable="2026-03-27",
        send_transmittal="AGCC-287-BCC-TRM-0001",
        incoming_control_date="",
        incoming_control_date_sortable="",
        confirm_transmittal="",
        sheet_status="На рассмотрении",
        note="",
        decision="active",
        comment=None,
        match_state="matched",
        identity_fingerprint="id",
        evidence_fingerprint="ev",
        source_path="",
        path_key="",
        in_f=True,
        in_rd=True,
        in_robot=True,
        in_auto_mto=True,
        review_id=None,
        issuance_send_id=1,
    )
    payload.update(overrides)
    return IssuanceJournalRow(**payload)


def _deleted_combo() -> QComboBox:
    parent = QWidget()
    combo = QComboBox(parent)
    combo.addItem("Активна", "active")
    combo.addItem("Аннулирована", "annulled")
    combo.setCurrentIndex(1)
    delete(parent)
    assert not isValid(combo)
    return combo


def _live_combo(parent: QWidget, *, current: str = "annulled") -> QComboBox:
    combo = QComboBox(parent)
    combo.addItem("Активна", "active")
    combo.addItem("Аннулирована", "annulled")
    pos = combo.findData(current)
    combo.setCurrentIndex(pos if pos >= 0 else 0)
    return combo


def test_restore_combo_after_cpp_deleted(app: QApplication) -> None:
    tab = IssuanceJournalTab()
    combo = _deleted_combo()
    tab._restore_combo(combo, "active")
    tab.close()


def test_apply_combo_skips_deleted_editor(app: QApplication) -> None:
    tab = IssuanceJournalTab()
    tab.set_database(MagicMock())
    tab.set_rows((_sheet_row(),), is_banned=lambda _t, _m: False)
    index = tab.table().model().index(0, _COL_DECISION)
    combo = _deleted_combo()
    with patch(
        "rd_catalog.issuance_journal_tab.apply_journal_decision"
    ) as persist:
        tab.apply_combo_decision(combo, index)
        persist.assert_not_called()
    tab.close()


def test_exclude_cancel_after_editor_destroyed_does_not_crash(
    app: QApplication,
) -> None:
    tab = IssuanceJournalTab()
    tab.set_database(MagicMock())
    tab.set_rows((_sheet_row(),), is_banned=lambda _t, _m: False)
    index = tab.table().model().index(0, _COL_DECISION)
    host = QWidget()
    combo = _live_combo(host)
    combo.setCurrentIndex(combo.findData("annulled"))

    def _cancel_after_delete(*_args: object, **_kwargs: object):
        delete(host)
        return ("отозвали", False)

    with (
        patch(
            "rd_catalog.issuance_journal_tab.QInputDialog.getMultiLineText",
            side_effect=_cancel_after_delete,
        ),
        patch(
            "rd_catalog.issuance_journal_tab.apply_journal_decision"
        ) as persist,
    ):
        tab.apply_combo_decision(combo, index)
        persist.assert_not_called()
    tab.close()


def test_exclude_comment_persists_after_editor_destroyed(
    app: QApplication,
) -> None:
    tab = IssuanceJournalTab()
    tab.set_database(MagicMock())
    row = _sheet_row()
    tab.set_rows((row,), is_banned=lambda _t, _m: False)
    tab.reviews_changed.connect(lambda _keys: tab._rebuild_table())
    index = tab.table().model().index(0, _COL_DECISION)
    host = QWidget()
    combo = _live_combo(host)
    combo.setCurrentIndex(combo.findData("annulled"))

    def _accept_after_delete(*_args: object, **_kwargs: object):
        delete(host)
        return ("отозвали", True)

    with (
        patch(
            "rd_catalog.issuance_journal_tab.QInputDialog.getMultiLineText",
            side_effect=_accept_after_delete,
        ),
        patch(
            "rd_catalog.issuance_journal_tab.apply_journal_decision",
            return_value=1,
        ) as persist,
    ):
        tab.apply_combo_decision(combo, index)
        persist.assert_called_once()
        assert persist.call_args.kwargs["decision"] == "annulled"
        assert persist.call_args.kwargs["comment"] == "отозвали"
    tab.close()


def _issuance_from_row(row: IssuanceJournalRow) -> IssuanceKit:
    return IssuanceKit(
        title=row.title,
        mark=row.mark,
        mark_raw=row.mark,
        title_system=f"{row.title}-{row.mark}",
        revision=None,
        appendix=None,
        revision_text=row.revision_text,
        status=row.sheet_status,
        send_date=row.send_date,
        send_date_sortable=row.send_date_sortable,
        send_transmittal=row.send_transmittal,
        incoming_control_date=row.incoming_control_date,
        incoming_control_date_sortable=row.incoming_control_date_sortable,
        confirm_transmittal=row.confirm_transmittal,
        note_raw=row.note,
        row_index=row.issuance_send_id or 0,
    )


def test_focus_kit_filters_and_selects_effective_row(app: QApplication) -> None:
    older = _sheet_row(
        title="2869",
        mark="SOS",
        revision_text="01-AN01",
        send_date="22.12.2025",
        send_date_sortable="2025-12-22",
        send_transmittal="TRM-1",
        issuance_send_id=1,
        identity_fingerprint="old",
    )
    newer = _sheet_row(
        title="2869",
        mark="SOS",
        revision_text="01-AN02",
        send_date="15.09.2026",
        send_date_sortable="2026-09-15",
        send_transmittal="TRM-2",
        issuance_send_id=2,
        identity_fingerprint="new",
    )
    other = _sheet_row(title="2868", mark="SKUD", issuance_send_id=3)
    issuance = _issuance_from_row(newer)
    newer = _sheet_row(
        title="2869",
        mark="SOS",
        revision_text="01-AN02",
        send_date="15.09.2026",
        send_date_sortable="2026-09-15",
        send_transmittal="TRM-2",
        issuance_send_id=2,
        identity_fingerprint=send_identity_fingerprint(issuance),
    )
    tab = IssuanceJournalTab()
    tab.set_rows((older, newer, other), is_banned=lambda _t, _m: False)
    tab._excluded.setChecked(True)
    tab._orphans.setChecked(True)
    tab._unmatched.setChecked(True)
    tab.focus_kit("2869", "SOS", issuance=issuance)
    assert tab._filter.text() == "2869-SOS"
    assert not tab._excluded.isChecked()
    assert not tab._orphans.isChecked()
    assert not tab._unmatched.isChecked()
    selected = tab.selected_row()
    assert selected is not None
    assert selected.revision_text == "01-AN02"
    assert tab.table().currentColumn() == _COL_DECISION
    hidden_other = False
    visible_revs = []
    for index in range(tab.table().rowCount()):
        payload = tab.row_at(index)
        if payload is None:
            continue
        if payload.title == "2868":
            assert tab.table().isRowHidden(index)
            hidden_other = True
            continue
        if payload.title == "2869":
            assert not tab.table().isRowHidden(index)
            visible_revs.append(payload.revision_text)
    assert hidden_other
    assert "01-AN01" in visible_revs
    assert "01-AN02" in visible_revs
    tab.close()


def main() -> None:
    app = _app()
    test_restore_combo_after_cpp_deleted(app)
    test_apply_combo_skips_deleted_editor(app)
    test_exclude_cancel_after_editor_destroyed_does_not_crash(app)
    test_exclude_comment_persists_after_editor_destroyed(app)
    test_focus_kit_filters_and_selects_effective_row(app)
    print("RD catalog issuance journal tab: OK")


if __name__ == "__main__":
    main()
