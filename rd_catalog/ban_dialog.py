"""Dialog to view, add, and unban hidden title+mark pairs."""

from __future__ import annotations

from PySide6.QtCore import Qt, Signal, Slot
from PySide6.QtWidgets import (
    QAbstractItemView,
    QDialog,
    QDialogButtonBox,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from rd_catalog.ban_filter import BanFilterStore, BannedTitleMark, parse_title_mark

_ROLE_PAIR = Qt.ItemDataRole.UserRole


class BannedTitlesDialog(QDialog):
    """Manage the persisted ban list of title+mark pairs."""

    changed = Signal()

    def __init__(
        self,
        store: BanFilterStore,
        parent: QWidget | None = None,
    ) -> None:
        """Build the dialog bound to an existing store.

        Args:
            store: Ban-filter persistence object.
            parent: Optional Qt parent.
        """

        super().__init__(parent)
        self._store = store
        self.setWindowTitle("Забаненные титулы")
        self.setMinimumSize(640, 420)
        self.resize(760, 480)
        self._build()
        self._reload_table()

    def _build(self) -> None:
        layout = QVBoxLayout(self)
        hint = QLabel(
            "Скрытые пары титул–марка не показываются в комплектах, "
            "готовности MTO и дереве документов.",
            self,
        )
        hint.setWordWrap(True)
        layout.addWidget(hint)

        self._table = QTableWidget(0, 4, self)
        self._table.setHorizontalHeaderLabels(
            ["Титул", "Марка", "Комментарий", "Добавлен"]
        )
        self._table.setSelectionBehavior(
            QAbstractItemView.SelectionBehavior.SelectRows
        )
        self._table.setSelectionMode(
            QAbstractItemView.SelectionMode.ExtendedSelection
        )
        self._table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self._table.setAlternatingRowColors(True)
        self._table.verticalHeader().setVisible(False)
        header = self._table.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        header.setStretchLastSection(True)
        header.setMinimumSectionSize(40)
        self._table.setColumnWidth(0, 90)
        self._table.setColumnWidth(1, 110)
        self._table.setColumnWidth(2, 280)
        layout.addWidget(self._table, 1)

        form = QHBoxLayout()
        form.addWidget(QLabel("Титул:"))
        self._title_edit = QLineEdit(self)
        self._title_edit.setPlaceholderText("5850")
        self._title_edit.setMaximumWidth(110)
        form.addWidget(self._title_edit)
        form.addWidget(QLabel("Марка:"))
        self._mark_edit = QLineEdit(self)
        self._mark_edit.setPlaceholderText("SKUD")
        self._mark_edit.setMaximumWidth(140)
        form.addWidget(self._mark_edit)
        form.addWidget(QLabel("Комментарий:"))
        self._comment_edit = QLineEdit(self)
        self._comment_edit.setPlaceholderText("аннулирован, старый…")
        form.addWidget(self._comment_edit, 1)
        add_button = QPushButton("Добавить", self)
        add_button.clicked.connect(self._on_add)
        form.addWidget(add_button)
        layout.addLayout(form)

        buttons = QHBoxLayout()
        self._unban_button = QPushButton("Разбанить", self)
        self._unban_button.clicked.connect(self._on_unban)
        buttons.addWidget(self._unban_button)
        buttons.addStretch(1)
        box = QDialogButtonBox(QDialogButtonBox.StandardButton.Close, self)
        box.rejected.connect(self.reject)
        box.accepted.connect(self.accept)
        buttons.addWidget(box)
        layout.addLayout(buttons)

        self._title_edit.returnPressed.connect(self._on_add)
        self._mark_edit.returnPressed.connect(self._on_add)
        self._comment_edit.returnPressed.connect(self._on_add)
        self._table.itemSelectionChanged.connect(self._update_unban_enabled)
        self._update_unban_enabled()

    def _reload_table(self) -> None:
        self._table.setSortingEnabled(False)
        self._table.setRowCount(0)
        for pair in self._store.pairs():
            added = pair.added_at.replace("T", " ")[:19] or "—"
            values = [pair.title, pair.mark, pair.comment or "—", added]
            row_index = self._table.rowCount()
            self._table.insertRow(row_index)
            for column, value in enumerate(values):
                item = QTableWidgetItem(value)
                item.setData(_ROLE_PAIR, pair)
                self._table.setItem(row_index, column, item)
        self._table.setSortingEnabled(True)
        self._update_unban_enabled()

    def _selected_pairs(self) -> list[BannedTitleMark]:
        seen: set[tuple[str, str]] = set()
        result: list[BannedTitleMark] = []
        for item in self._table.selectedItems():
            pair = item.data(_ROLE_PAIR)
            if not isinstance(pair, BannedTitleMark):
                continue
            if pair.identity in seen:
                continue
            seen.add(pair.identity)
            result.append(pair)
        return result

    @Slot()
    def _update_unban_enabled(self) -> None:
        self._unban_button.setEnabled(bool(self._selected_pairs()))

    @Slot()
    def _on_add(self) -> None:
        try:
            title, mark = parse_title_mark(
                self._title_edit.text(), self._mark_edit.text()
            )
        except ValueError as exc:
            QMessageBox.warning(self, "Бан-фильтр", str(exc))
            return
        comment = self._comment_edit.text().strip()
        try:
            pair, added = self._store.add(title, mark, comment)
        except OSError as exc:
            QMessageBox.warning(
                self, "Бан-фильтр", f"Не удалось сохранить список:\n{exc}"
            )
            return
        if not added:
            QMessageBox.information(
                self,
                "Бан-фильтр",
                f"{pair.label} уже в списке скрытых.",
            )
            return
        self._title_edit.clear()
        self._mark_edit.clear()
        self._comment_edit.clear()
        self._reload_table()
        self.changed.emit()

    @Slot()
    def _on_unban(self) -> None:
        selected = self._selected_pairs()
        if not selected:
            return
        labels = ", ".join(pair.label for pair in selected)
        reply = QMessageBox.question(
            self,
            "Разбанить",
            f"Вернуть в списки: {labels}?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.Yes,
        )
        if reply != QMessageBox.StandardButton.Yes:
            return
        try:
            for pair in selected:
                self._store.remove(pair.title, pair.mark)
        except OSError as exc:
            QMessageBox.warning(
                self, "Бан-фильтр", f"Не удалось сохранить список:\n{exc}"
            )
            return
        self._reload_table()
        self.changed.emit()
