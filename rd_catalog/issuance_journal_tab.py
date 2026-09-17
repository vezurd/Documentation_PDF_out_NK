"""Issuance journal tab: sheet sends, orphans, and manual reviews.

Qt monitor only. Rows come from ``list_issuance_journal``. The widget
does not parse Google CSV.
"""

from __future__ import annotations

from collections.abc import Callable, Sequence

from shiboken6 import isValid
from PySide6.QtCore import QModelIndex, QPoint, QSettings, Qt, Signal, Slot
from PySide6.QtWidgets import (
    QAbstractItemView,
    QCheckBox,
    QComboBox,
    QDialog,
    QDialogButtonBox,
    QFormLayout,
    QHBoxLayout,
    QHeaderView,
    QInputDialog,
    QLabel,
    QLineEdit,
    QMenu,
    QMessageBox,
    QPushButton,
    QStyledItemDelegate,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from rd_catalog.context_menu_qt import exec_tracked_menu
from rd_catalog.context_menu_usage import MENU_ISSUANCE_JOURNAL
from rd_catalog.db import CatalogDatabase
from rd_catalog.issuance_review import (
    IssuanceJournalRow,
    add_manual_journal_row,
    apply_journal_decision,
    pick_journal_row_for_issuance,
)
from rd_catalog.kits import IssuanceKit, kit_identity_key

_ROLE_ROW = Qt.ItemDataRole.UserRole
_HEADERS = (
    "Титул",
    "Марка",
    "Источник",
    "Рев.",
    "Дата отпр.",
    "TRM",
    "Статус листа",
    "Наш статус",
    "Комментарий",
    "F",
    "РД",
    "Робот",
    "Авто МТО",
    "Вх.контр.",
    "TRM подтв.",
    "Примечание",
    "Сопоставление",
)
_COL_REV = 3
_COL_SEND_DATE = 4
_COL_TRM = 5
_COL_SHEET_STATUS = 6
_COL_DECISION = 7
_COL_COMMENT = 8
_COL_INCOMING = 13
_COL_CONFIRM = 14
_COL_NOTE = 15
_FIELD_COLUMNS = {
    _COL_REV,
    _COL_SEND_DATE,
    _COL_TRM,
    _COL_SHEET_STATUS,
    _COL_INCOMING,
    _COL_CONFIRM,
    _COL_NOTE,
}
_DECISION_LABELS = {
    "": "не задан",
    "active": "Активна",
    "legalized": "Легализована",
    "annulled": "Аннулирована",
    "erroneous": "Ошибочна",
    "duplicate": "Дубликат",
}
_SOURCE_LABELS = {
    "issuance": "Выдача",
    "google_f": "F",
    "rd": "РД",
    "robot": "Робот",
    "auto_mto": "Авто МТО",
    "manual": "вручную",
}
_MATCH_LABELS = {
    "": "",
    "matched": "сопоставлено",
    "unmatched": "не сопоставлено",
    "ambiguous": "неоднозначно",
}
_SHEET_DECISIONS = ("active", "annulled", "erroneous", "duplicate")
_ORPHAN_DECISIONS = ("", "legalized", "annulled", "erroneous")
_EXCLUDE_DECISIONS = frozenset({"annulled", "erroneous", "duplicate"})
_ADD_DECISIONS = ("", "legalized", "annulled", "erroneous")


def _is_sheet_row(row: IssuanceJournalRow) -> bool:
    return row.kind == "send" or row.source == "issuance"


def _fields_editable(row: IssuanceJournalRow) -> bool:
    return row.kind == "manual" or row.decision == "legalized"


def _decision_choices(row: IssuanceJournalRow) -> tuple[tuple[str, str], ...]:
    tokens = _SHEET_DECISIONS if _is_sheet_row(row) else _ORPHAN_DECISIONS
    choices = [(token, _DECISION_LABELS[token]) for token in tokens]
    if row.decision not in {token for token, _label in choices}:
        choices.insert(0, (row.decision, _DECISION_LABELS.get(row.decision, row.decision)))
    return tuple(choices)


def _presence(flag: bool) -> str:
    return "да" if flag else ""


def _qt_alive(widget: QWidget) -> bool:
    """Return whether a PySide wrapper still owns a C++ object."""

    try:
        return bool(isValid(widget))
    except RuntimeError:
        return False


def _settings_bool(settings: QSettings, key: str, default: bool) -> bool:
    value = settings.value(key, default)
    if isinstance(value, bool):
        return value
    if isinstance(value, (int, float)):
        return bool(value)
    text = str(value).strip().casefold()
    if text in {"1", "true", "yes", "on"}:
        return True
    if text in {"0", "false", "no", "off", ""}:
        return False
    return default


class _DecisionComboDelegate(QStyledItemDelegate):
    """In-cell decision combo for «Наш статус»."""

    def __init__(self, tab: IssuanceJournalTab) -> None:
        super().__init__(tab)
        self._tab = tab

    def createEditor(
        self,
        parent: QWidget,
        _option: object,
        index: QModelIndex,
    ) -> QComboBox:
        combo = QComboBox(parent)
        payload = self._tab.row_at(index.row())
        if payload is None:
            return combo
        for value, label in _decision_choices(payload):
            combo.addItem(label, value)
        combo.activated.connect(
            lambda _idx, editor=combo, model_index=QModelIndex(index): (
                self._tab.apply_combo_decision(editor, model_index)
            )
        )
        return combo

    def setEditorData(self, editor: QWidget, index: QModelIndex) -> None:
        if not isinstance(editor, QComboBox):
            super().setEditorData(editor, index)
            return
        payload = self._tab.row_at(index.row())
        if payload is None:
            return
        editor.blockSignals(True)
        pos = editor.findData(payload.decision)
        editor.setCurrentIndex(pos if pos >= 0 else 0)
        editor.blockSignals(False)

    def setModelData(
        self,
        _editor: QWidget,
        _model: object,
        _index: QModelIndex,
    ) -> None:
        return


class _JournalRowDialog(QDialog):
    """Add or edit issuance-journal fields."""

    def __init__(
        self,
        parent: QWidget | None,
        *,
        add_mode: bool,
        row: IssuanceJournalRow | None = None,
    ) -> None:
        super().__init__(parent)
        self._add_mode = add_mode
        self.setWindowTitle("Добавить строку" if add_mode else "Изменить строку")
        self.setMinimumWidth(420)
        layout = QVBoxLayout(self)
        form = QFormLayout()
        self._title = QLineEdit(self)
        self._mark = QLineEdit(self)
        self._revision = QLineEdit(self)
        self._send_date = QLineEdit(self)
        self._send_trm = QLineEdit(self)
        self._incoming = QLineEdit(self)
        self._confirm = QLineEdit(self)
        self._sheet_status = QLineEdit(self)
        self._note = QLineEdit(self)
        self._comment = QLineEdit(self)
        self._decision = QComboBox(self)
        tokens = _ADD_DECISIONS if add_mode else _ORPHAN_DECISIONS
        if row is not None and not add_mode and _is_sheet_row(row):
            tokens = _SHEET_DECISIONS
        for token in tokens:
            self._decision.addItem(_DECISION_LABELS[token], token)
        if row is not None:
            self._title.setText(row.title)
            self._mark.setText(row.mark)
            self._revision.setText(row.revision_text)
            self._send_date.setText(row.send_date)
            self._send_trm.setText(row.send_transmittal)
            self._incoming.setText(row.incoming_control_date)
            self._confirm.setText(row.confirm_transmittal)
            self._sheet_status.setText(row.sheet_status)
            self._note.setText(row.note)
            self._comment.setText(row.comment or "")
            pos = self._decision.findData(row.decision)
            if pos >= 0:
                self._decision.setCurrentIndex(pos)
        if not add_mode:
            self._title.setReadOnly(True)
            self._mark.setReadOnly(True)
        form.addRow("Титул:", self._title)
        form.addRow("Марка:", self._mark)
        form.addRow("Рев.:", self._revision)
        form.addRow("Дата отпр.:", self._send_date)
        form.addRow("TRM:", self._send_trm)
        form.addRow("Вх.контр.:", self._incoming)
        form.addRow("TRM подтв.:", self._confirm)
        form.addRow("Статус листа:", self._sheet_status)
        form.addRow("Примечание:", self._note)
        form.addRow("Комментарий:", self._comment)
        form.addRow("Наш статус:", self._decision)
        layout.addLayout(form)
        hint = QLabel(
            "нет даты и нет TRM — сопоставление с F только по ревизии",
            self,
        )
        hint.setWordWrap(True)
        layout.addWidget(hint)
        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok
            | QDialogButtonBox.StandardButton.Cancel,
            self,
        )
        buttons.accepted.connect(self._on_accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)
        self._revision.textChanged.connect(self._sync_legalized_enabled)
        self._decision.currentIndexChanged.connect(self._sync_legalized_enabled)
        self._sync_legalized_enabled()

    def reset_fields(self) -> None:
        """Clear add-mode fields back to an empty draft."""

        self.setWindowTitle("Добавить строку")
        self._title.clear()
        self._mark.clear()
        self._revision.clear()
        self._send_date.clear()
        self._send_trm.clear()
        self._incoming.clear()
        self._confirm.clear()
        self._sheet_status.clear()
        self._note.clear()
        self._comment.clear()
        self._title.setReadOnly(False)
        self._mark.setReadOnly(False)
        empty = self._decision.findData("")
        self._decision.setCurrentIndex(empty if empty >= 0 else 0)
        self._sync_legalized_enabled()

    def apply_prefill(
        self,
        *,
        title: str = "",
        mark: str = "",
        revision_text: str = "",
        send_date: str = "",
        send_transmittal: str = "",
        incoming_control_date: str = "",
        confirm_transmittal: str = "",
        sheet_status: str = "",
        note: str = "",
        comment: str = "",
        decision: str = "",
        lock_identity: bool = False,
        window_title: str = "",
    ) -> None:
        """Fill add-mode fields from a kit (or other caller).

        Args:
            title: Four-digit title.
            mark: Latin AGCC mark.
            revision_text: Filename revision to legalize.
            send_date: Optional send date ``DD.MM.YYYY``.
            send_transmittal: Optional send TRM.
            incoming_control_date: Optional incoming-control date.
            confirm_transmittal: Optional confirmation TRM.
            sheet_status: Optional sheet status.
            note: Optional note.
            comment: Optional reviewer comment.
            decision: Decision token (``legalized`` for an RD handoff).
            lock_identity: Make title and mark read-only.
            window_title: Override the dialog caption.
        """

        self.reset_fields()
        if window_title:
            self.setWindowTitle(window_title)
        self._title.setText(title)
        self._mark.setText(mark)
        self._revision.setText(revision_text)
        self._send_date.setText(send_date)
        self._send_trm.setText(send_transmittal)
        self._incoming.setText(incoming_control_date)
        self._confirm.setText(confirm_transmittal)
        self._sheet_status.setText(sheet_status)
        self._note.setText(note)
        self._comment.setText(comment)
        self._title.setReadOnly(lock_identity)
        self._mark.setReadOnly(lock_identity)
        pos = self._decision.findData(decision)
        if pos >= 0:
            self._decision.setCurrentIndex(pos)
        self._sync_legalized_enabled()

    def _legalized_index(self) -> int:
        return self._decision.findData("legalized")

    @Slot()
    def _sync_legalized_enabled(self) -> None:
        index = self._legalized_index()
        if index < 0:
            return
        item = self._decision.model().item(index)
        if item is None:
            return
        item.setEnabled(bool(self._revision.text().strip()))
        if (
            self._decision.currentData() == "legalized"
            and not self._revision.text().strip()
        ):
            self._decision.setCurrentIndex(self._decision.findData("") if self._add_mode else 0)

    def values(self) -> dict[str, str | None]:
        """Return dialog fields as persist kwargs.

        Returns:
            Title, mark, issuance fields, decision, and comment.
        """

        comment = self._comment.text().strip()
        return {
            "title": self._title.text().strip(),
            "mark": self._mark.text().strip(),
            "revision_text": self._revision.text().strip(),
            "send_date": self._send_date.text().strip(),
            "send_transmittal": self._send_trm.text().strip(),
            "incoming_control_date": self._incoming.text().strip(),
            "confirm_transmittal": self._confirm.text().strip(),
            "sheet_status": self._sheet_status.text().strip(),
            "note": self._note.text().strip(),
            "decision": str(self._decision.currentData() or ""),
            "comment": comment or None,
        }

    @Slot()
    def _on_accept(self) -> None:
        values = self.values()
        if self._add_mode and (not values["title"] or not values["mark"]):
            QMessageBox.warning(self, "Добавить строку", "Нужны титул и марка.")
            return
        decision = str(values["decision"] or "")
        revision = str(values["revision_text"] or "")
        comment = values["comment"]
        if decision == "legalized" and not revision:
            QMessageBox.warning(
                self,
                self.windowTitle(),
                "Укажите ревизию, чтобы легализовать строку.",
            )
            return
        if decision in _EXCLUDE_DECISIONS and not (comment or "").strip():
            QMessageBox.warning(
                self,
                self.windowTitle(),
                "Для исключения обязателен комментарий.",
            )
            return
        self.accept()


class IssuanceJournalTab(QWidget):
    """Filterable issuance journal with in-cell status decisions."""

    kit_activated = Signal(str, str)
    reviews_changed = Signal(object)
    prepare_context_menu = Signal(QMenu)

    def __init__(
        self,
        parent: QWidget | None = None,
        *,
        database: CatalogDatabase | None = None,
    ) -> None:
        super().__init__(parent)
        self._database = database
        self._rows: tuple[IssuanceJournalRow, ...] = ()
        self._is_banned: Callable[[str, str], bool] = lambda _t, _m: False
        self._allowed_kits: set[tuple[str, str]] | None = None
        self._build()

    def set_database(self, database: CatalogDatabase) -> None:
        """Bind the catalog database used to persist reviews.

        Args:
            database: Initialized catalog database.
        """

        self._database = database

    def table(self) -> QTableWidget:
        return self._table

    def selected_row(self) -> IssuanceJournalRow | None:
        """Return the journal payload for the current table selection.

        Returns:
            The ``IssuanceJournalRow`` on the selected row, or ``None``.
        """

        return self.row_at(self._table.currentRow())

    def row_at(self, row_index: int) -> IssuanceJournalRow | None:
        """Return the journal payload stored on a table row.

        Args:
            row_index: Table row index.

        Returns:
            Stored ``IssuanceJournalRow``, or ``None``.
        """

        if row_index < 0:
            return None
        item = self._table.item(row_index, 0)
        payload = item.data(_ROLE_ROW) if item else None
        return payload if isinstance(payload, IssuanceJournalRow) else None

    def _build(self) -> None:
        layout = QVBoxLayout(self)
        filters = QHBoxLayout()
        filters.addWidget(QLabel("Фильтр:"))
        self._filter = QLineEdit(self)
        self._filter.setPlaceholderText("Титул, марка, ревизия, TRM…")
        self._filter.textChanged.connect(self._apply_row_visibility)
        filters.addWidget(self._filter, 1)
        self._excluded = QCheckBox("Только исключённые", self)
        self._orphans = QCheckBox("Только сироты", self)
        self._unmatched = QCheckBox("Не сопоставлено", self)
        for box in (self._excluded, self._orphans, self._unmatched):
            box.toggled.connect(self._apply_row_visibility)
            filters.addWidget(box)
        layout.addLayout(filters)

        actions = QHBoxLayout()
        self._add_button = QPushButton("Добавить строку…", self)
        self._add_button.clicked.connect(self._on_add_row)
        actions.addWidget(self._add_button)
        actions.addStretch(1)
        layout.addLayout(actions)

        self._add_dialog = _JournalRowDialog(self, add_mode=True)
        self._decision = self._add_dialog._decision

        self._table = QTableWidget(0, len(_HEADERS), self)
        self._table.setHorizontalHeaderLabels(list(_HEADERS))
        self._table.setSelectionBehavior(
            QAbstractItemView.SelectionBehavior.SelectRows
        )
        self._table.setSelectionMode(
            QAbstractItemView.SelectionMode.SingleSelection
        )
        self._table.setEditTriggers(
            QAbstractItemView.EditTrigger.DoubleClicked
            | QAbstractItemView.EditTrigger.SelectedClicked
            | QAbstractItemView.EditTrigger.EditKeyPressed
        )
        self._table.setAlternatingRowColors(True)
        self._table.setSortingEnabled(True)
        self._table.verticalHeader().setVisible(False)
        self._table.setWordWrap(False)
        header = self._table.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        header.setSectionsMovable(True)
        header.setFirstSectionMovable(True)
        header.setStretchLastSection(False)
        header.setMinimumSectionSize(40)
        self._decision_delegate = _DecisionComboDelegate(self)
        self._table.setItemDelegateForColumn(
            _COL_DECISION, self._decision_delegate
        )
        self._table.cellDoubleClicked.connect(self._on_cell_activated)
        self._table.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self._table.customContextMenuRequested.connect(self._show_context_menu)
        layout.addWidget(self._table, 1)

    def set_rows(
        self,
        rows: Sequence[IssuanceJournalRow],
        *,
        is_banned: Callable[[str, str], bool],
        allowed_kits: set[tuple[str, str]] | None = None,
    ) -> None:
        """Replace the journal snapshot and rebuild the table.

        Args:
            rows: ``IssuanceJournalRow`` values from ``list_issuance_journal``.
            is_banned: Hide banned ``(title, mark)`` pairs.
            allowed_kits: If set, only these identities are shown.
                ``None`` keeps every non-banned kit (Комплекты universe).
        """

        self._rows = tuple(rows)
        self._is_banned = is_banned
        self._allowed_kits = allowed_kits
        self._rebuild_table()

    def restore_filters(self, settings: QSettings) -> None:
        """Load filter widgets from QSettings without applying twice.

        Args:
            settings: Catalog window QSettings.
        """

        boxes = (self._excluded, self._orphans, self._unmatched)
        self._filter.blockSignals(True)
        for box in boxes:
            box.blockSignals(True)
        try:
            self._filter.setText(
                str(settings.value("window/issuance_journal_filter") or "")
            )
            self._excluded.setChecked(
                _settings_bool(settings, "window/issuance_journal_excluded", False)
            )
            self._orphans.setChecked(
                _settings_bool(settings, "window/issuance_journal_orphans", False)
            )
            self._unmatched.setChecked(
                _settings_bool(
                    settings, "window/issuance_journal_unmatched", False
                )
            )
        finally:
            self._filter.blockSignals(False)
            for box in boxes:
                box.blockSignals(False)

    def save_filters(self, settings: QSettings) -> None:
        """Persist filter widgets into QSettings.

        Args:
            settings: Catalog window QSettings.
        """

        settings.setValue("window/issuance_journal_filter", self._filter.text())
        settings.setValue(
            "window/issuance_journal_excluded", self._excluded.isChecked()
        )
        settings.setValue(
            "window/issuance_journal_orphans", self._orphans.isChecked()
        )
        settings.setValue(
            "window/issuance_journal_unmatched", self._unmatched.isChecked()
        )

    def set_kit_filter(self, title: str, mark: str) -> None:
        """Show only this title–mark and clear restrictive checkboxes.

        Jump from Комплекты must keep the effective «Выдача · рев.» row
        visible (it is included, usually a sheet send).

        Args:
            title: Four-digit title.
            mark: Latin AGCC mark.
        """

        boxes = (self._excluded, self._orphans, self._unmatched)
        self._filter.blockSignals(True)
        for box in boxes:
            box.blockSignals(True)
        try:
            self._excluded.setChecked(False)
            self._orphans.setChecked(False)
            self._unmatched.setChecked(False)
            self._filter.setText(f"{title}-{mark}")
        finally:
            self._filter.blockSignals(False)
            for box in boxes:
                box.blockSignals(False)
        self._apply_row_visibility()

    def focus_kit(
        self,
        title: str,
        mark: str,
        *,
        issuance: IssuanceKit | None = None,
    ) -> None:
        """Filter to ``title`` / ``mark`` and select the painted issuance row.

        Args:
            title: Four-digit title.
            mark: Latin AGCC mark.
            issuance: Effective send from Комплекты «Выдача · рев.».
        """

        self.set_kit_filter(title, mark)
        target = pick_journal_row_for_issuance(
            self._rows,
            title=title,
            mark=mark,
            issuance=issuance,
        )
        key = kit_identity_key(title, mark)
        table = self._table
        fallback_visible: int | None = None
        fallback_any: int | None = None
        for index in range(table.rowCount()):
            payload = self.row_at(index)
            if payload is None:
                continue
            if kit_identity_key(payload.title, payload.mark) != key:
                continue
            hidden = table.isRowHidden(index)
            if fallback_any is None:
                fallback_any = index
            if fallback_visible is None and not hidden:
                fallback_visible = index
            if target is None:
                continue
            if (
                payload == target
                or payload.identity_fingerprint == target.identity_fingerprint
            ):
                self._focus_table_row(index)
                return
        chosen = (
            fallback_visible if fallback_visible is not None else fallback_any
        )
        if chosen is not None:
            self._focus_table_row(chosen)

    def _focus_table_row(self, index: int) -> None:
        """Select ``index`` and put keyboard focus on «Наш статус»."""

        table = self._table
        table.selectRow(index)
        item = table.item(index, _COL_DECISION)
        if item is None:
            item = table.item(index, 0)
        if item is not None:
            table.setCurrentItem(item)
            table.scrollToItem(item)
        table.setFocus(Qt.FocusReason.OtherFocusReason)

    def apply_combo_decision(self, combo: QComboBox, index: QModelIndex) -> None:
        """Persist a combo change after closing the cell editor.

        Modal comment/warning dialogs (and the journal rebuild after a
        successful write) destroy the in-cell ``QComboBox``. The decision
        is copied first; the editor is dismissed before any dialog so a
        cancelled exclude does not touch a deleted C++ object.

        Args:
            combo: Decision combo that just changed.
            index: Table model index of the decision cell.
        """

        payload = self.row_at(index.row())
        if payload is None or self._database is None:
            return
        if not _qt_alive(combo):
            return
        new_decision = str(
            combo.currentData() if combo.currentData() is not None else ""
        )
        self._dismiss_decision_editor(combo)
        if new_decision == payload.decision:
            return
        comment = payload.comment
        if new_decision in _EXCLUDE_DECISIONS and not (comment or "").strip():
            text, accepted = QInputDialog.getMultiLineText(
                self,
                "Комментарий",
                "Для исключения обязателен комментарий:",
                "",
            )
            if not accepted or not text.strip():
                return
            comment = text.strip()
        if new_decision == "legalized" and not (payload.revision_text or "").strip():
            QMessageBox.warning(
                self,
                "Выдача · Журнал",
                "Укажите ревизию, чтобы легализовать строку.",
            )
            return
        self._persist_decision(payload, new_decision, comment)

    def _dismiss_decision_editor(self, combo: QComboBox) -> None:
        """Close the in-cell combo before a modal dialog or table rebuild."""

        if not _qt_alive(combo):
            return
        try:
            combo.blockSignals(True)
            self._decision_delegate.closeEditor.emit(
                combo,
                QStyledItemDelegate.EndEditHint.NoHint,
            )
        except RuntimeError:
            return

    def _restore_combo(self, combo: QComboBox, decision: str) -> None:
        """Reset a still-alive editor; no-op when Qt already deleted it."""

        if not _qt_alive(combo):
            return
        try:
            combo.blockSignals(True)
            pos = combo.findData(decision)
            combo.setCurrentIndex(pos if pos >= 0 else 0)
            combo.blockSignals(False)
        except RuntimeError:
            return

    def _persist_decision(
        self,
        row: IssuanceJournalRow,
        decision: str,
        comment: str | None,
        **field_overrides: str | None,
    ) -> bool:
        if self._database is None:
            return False
        try:
            apply_journal_decision(
                self._database,
                row,
                decision=decision,
                comment=comment,
                **field_overrides,
            )
        except ValueError as exc:
            QMessageBox.warning(self, "Выдача · Журнал", str(exc))
            return False
        self.reviews_changed.emit([(row.title, row.mark)])
        return True

    def _kit_is_hidden(self, title: str, mark: str) -> bool:
        if self._is_banned(title, mark):
            return True
        if self._allowed_kits is None:
            return False
        return kit_identity_key(title, mark) not in self._allowed_kits

    def _rebuild_table(self) -> None:
        table = self._table
        table.setSortingEnabled(False)
        visible_rows = [
            row
            for row in self._rows
            if not self._kit_is_hidden(row.title, row.mark)
        ]
        table.setRowCount(len(visible_rows))
        for row_index, row in enumerate(visible_rows):
            self._fill_row(row_index, row)
        table.setSortingEnabled(True)
        self._apply_row_visibility()

    def _fill_row(self, row_index: int, row: IssuanceJournalRow) -> None:
        values = (
            row.title,
            row.mark,
            _SOURCE_LABELS.get(row.source, row.source),
            row.revision_text,
            row.send_date,
            row.send_transmittal,
            row.sheet_status,
            _DECISION_LABELS.get(row.decision, row.decision),
            row.comment or "",
            _presence(row.in_f),
            _presence(row.in_rd),
            _presence(row.in_robot),
            _presence(row.in_auto_mto),
            row.incoming_control_date,
            row.confirm_transmittal,
            row.note,
            _MATCH_LABELS.get(row.match_state, row.match_state),
        )
        for column, text in enumerate(values):
            item = QTableWidgetItem(text)
            item.setData(_ROLE_ROW, row)
            if column == _COL_DECISION:
                item.setFlags(item.flags() | Qt.ItemFlag.ItemIsEditable)
            else:
                item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEditable)
            self._table.setItem(row_index, column, item)

    def _row_matches_text(self, row: IssuanceJournalRow, needle: str) -> bool:
        if not needle:
            return True
        chunks = [
            row.title,
            row.mark,
            f"{row.title}-{row.mark}",
            row.revision_text,
            row.send_date,
            row.send_transmittal,
            row.sheet_status,
            row.note,
            row.comment or "",
            row.source,
            _SOURCE_LABELS.get(row.source, ""),
            row.decision,
            _DECISION_LABELS.get(row.decision, ""),
            row.match_state,
            _MATCH_LABELS.get(row.match_state, ""),
            row.confirm_transmittal,
            row.incoming_control_date,
        ]
        return needle in " ".join(chunks).casefold()

    def _row_matches_filters(self, row: IssuanceJournalRow) -> bool:
        if self._excluded.isChecked() and row.decision not in _EXCLUDE_DECISIONS:
            return False
        if self._orphans.isChecked() and row.kind not in {"orphan", "manual"}:
            return False
        if self._unmatched.isChecked() and not (
            row.issuance_send_id is None
            or row.match_state in {"unmatched", "ambiguous"}
        ):
            return False
        return True

    @Slot()
    def _apply_row_visibility(self) -> None:
        needle = self._filter.text().strip().casefold()
        table = self._table
        for index in range(table.rowCount()):
            payload = self.row_at(index)
            if payload is None or self._kit_is_hidden(payload.title, payload.mark):
                table.setRowHidden(index, True)
                continue
            visible = self._row_matches_filters(payload) and self._row_matches_text(
                payload, needle
            )
            table.setRowHidden(index, not visible)

    @Slot(int, int)
    def _on_cell_activated(self, row_index: int, column: int) -> None:
        payload = self.row_at(row_index)
        if payload is None:
            return
        if column == _COL_DECISION:
            return
        if column == _COL_COMMENT:
            self._edit_comment(payload)
            return
        if column in _FIELD_COLUMNS and _fields_editable(payload):
            self._edit_fields(payload)
            return
        if _is_sheet_row(payload):
            return
        if _fields_editable(payload):
            self._edit_fields(payload)

    def _edit_comment(self, row: IssuanceJournalRow) -> None:
        text, accepted = QInputDialog.getMultiLineText(
            self,
            "Комментарий",
            "Комментарий:",
            row.comment or "",
        )
        if not accepted:
            return
        comment = text.strip() or None
        if (comment or "") == (row.comment or ""):
            return
        if row.decision in _EXCLUDE_DECISIONS and not comment:
            QMessageBox.warning(
                self,
                "Выдача · Журнал",
                "Для исключения обязателен комментарий.",
            )
            return
        decision = row.decision
        if _is_sheet_row(row) and not decision:
            decision = "active"
        self._persist_decision(row, decision, comment)

    def _edit_fields(self, row: IssuanceJournalRow) -> None:
        dialog = _JournalRowDialog(self, add_mode=False, row=row)
        if dialog.exec() != QDialog.DialogCode.Accepted:
            return
        values = dialog.values()
        self._persist_decision(
            row,
            str(values["decision"] or row.decision),
            values["comment"],
            revision_text=str(values["revision_text"] or ""),
            send_date=str(values["send_date"] or ""),
            send_transmittal=str(values["send_transmittal"] or ""),
            incoming_control_date=str(values["incoming_control_date"] or ""),
            confirm_transmittal=str(values["confirm_transmittal"] or ""),
            sheet_status=str(values["sheet_status"] or ""),
            note=str(values["note"] or ""),
        )

    def open_add_row_dialog(
        self,
        *,
        title: str = "",
        mark: str = "",
        revision_text: str = "",
        send_date: str = "",
        send_transmittal: str = "",
        incoming_control_date: str = "",
        confirm_transmittal: str = "",
        sheet_status: str = "",
        note: str = "",
        comment: str = "",
        decision: str = "",
        lock_identity: bool = False,
        window_title: str = "",
    ) -> bool:
        """Show the add-row dialog, optionally pre-filled from a kit.

        Args:
            title: Four-digit title.
            mark: Latin AGCC mark.
            revision_text: Filename revision to legalize.
            send_date: Optional send date ``DD.MM.YYYY``.
            send_transmittal: Optional send TRM.
            incoming_control_date: Optional incoming-control date.
            confirm_transmittal: Optional confirmation TRM.
            sheet_status: Optional sheet status.
            note: Optional note.
            comment: Optional reviewer comment.
            decision: Decision token.
            lock_identity: Make title and mark read-only.
            window_title: Override the dialog caption.

        Returns:
            ``True`` when a row was persisted.
        """

        dialog = self._add_dialog
        if any((title, mark, revision_text, send_date, note, decision, window_title)):
            dialog.apply_prefill(
                title=title,
                mark=mark,
                revision_text=revision_text,
                send_date=send_date,
                send_transmittal=send_transmittal,
                incoming_control_date=incoming_control_date,
                confirm_transmittal=confirm_transmittal,
                sheet_status=sheet_status,
                note=note,
                comment=comment,
                decision=decision,
                lock_identity=lock_identity,
                window_title=window_title,
            )
        else:
            dialog.reset_fields()
        accepted = dialog.exec() == QDialog.DialogCode.Accepted
        if not accepted:
            dialog.reset_fields()
            return False
        values = dialog.values()
        dialog.reset_fields()
        if self._database is None:
            return False
        try:
            add_manual_journal_row(
                self._database,
                str(values["title"] or ""),
                str(values["mark"] or ""),
                revision_text=str(values["revision_text"] or ""),
                send_date=str(values["send_date"] or ""),
                send_transmittal=str(values["send_transmittal"] or ""),
                incoming_control_date=str(values["incoming_control_date"] or ""),
                confirm_transmittal=str(values["confirm_transmittal"] or ""),
                sheet_status=str(values["sheet_status"] or ""),
                note=str(values["note"] or ""),
                decision=str(values["decision"] or ""),
                comment=values["comment"],
            )
        except ValueError as exc:
            QMessageBox.warning(self, "Добавить строку", str(exc))
            return False
        self.reviews_changed.emit(
            [(str(values["title"] or ""), str(values["mark"] or ""))]
        )
        return True

    @Slot()
    def _on_add_row(self) -> None:
        self.open_add_row_dialog()

    def _show_context_menu(self, position: QPoint) -> None:
        table = self._table
        row_index = table.rowAt(position.y())
        if row_index < 0:
            return
        table.selectRow(row_index)
        payload = self.row_at(row_index)
        if payload is None:
            return
        menu = QMenu(self)
        show_kits = menu.addAction('Показать в «Комплекты»')
        self.prepare_context_menu.emit(menu)
        chosen = exec_tracked_menu(
            menu, MENU_ISSUANCE_JOURNAL, table.viewport().mapToGlobal(position)
        )
        if chosen == show_kits:
            self.kit_activated.emit(payload.title, payload.mark)
