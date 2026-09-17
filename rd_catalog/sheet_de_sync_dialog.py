"""Dialog: bulk-write KSB ИД D/E from the last column-F event.

Qt monitor. Does not call Sheets. The parent window starts
:class:`~rd_catalog.google_f_write_thread.GoogleFWriteThread` with
:meth:`selected_jobs` after **Записать выбранное**.
"""

from __future__ import annotations

from collections.abc import Sequence

from PySide6.QtCore import Qt
from PySide6.QtGui import QBrush, QColor
from PySide6.QtWidgets import (
    QAbstractItemView,
    QDialog,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QMessageBox,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from rd_catalog.google_f_write import JournalWriteJob
from rd_catalog.monitor_views import REV_MATCH_FILL, ROBOT_ORPHAN_FILL
from rd_catalog.sheet_de_sync import SheetDeSyncRow

SHEET_DE_SYNC_BUTTON = "D/E из журнала F…"
SHEET_DE_SYNC_TITLE = "D/E отстают от журнала F"
SHEET_DE_SYNC_INTRO = (
    "Последнее событие F совпадает с ревизией файлов РД, а D/E ещё нет. "
    "«как есть» — эту ячейку не пишем. D показан без префикса «Рев.» "
    "(в таблицу уйдёт «Рев. …»). F не переписывается. "
    "Строку в КСБ ИД не создаём, «Выдача РД ПД» не трогаем."
)
_ROLE_ROW = Qt.ItemDataRole.UserRole
_HEADERS = (
    "",
    "Титул",
    "Марка",
    "РД",
    "D сейчас",
    "D будет",
    "E сейчас",
    "E будет",
    "Пишем",
    "Последняя строка F",
)
_COL_CHECK = 0
_COL_D_NOW = 4
_COL_D_NEXT = 5
_COL_E_NOW = 6
_COL_E_NEXT = 7
_WRONG_BG = QColor(ROBOT_ORPHAN_FILL)
_RIGHT_BG = QColor(REV_MATCH_FILL)


class SheetDeSyncDialog(QDialog):
    """Checkbox list of kits whose D/E lag the last F event."""

    def __init__(
        self,
        rows: Sequence[SheetDeSyncRow],
        parent: QWidget | None = None,
    ) -> None:
        """Fill the table. Every row starts checked.

        Args:
            rows: Domain rows from :func:`list_sheet_de_sync_rows`.
            parent: Optional Qt parent (the catalog window).
        """

        super().__init__(parent)
        self._rows = tuple(rows)
        self.setWindowTitle(SHEET_DE_SYNC_TITLE)
        self.setMinimumSize(980, 480)
        self.resize(1180, 560)
        self._build()
        self._fill()

    def selected_jobs(self) -> tuple[JournalWriteJob, ...]:
        """Return jobs for currently checked rows, table order."""

        jobs: list[JournalWriteJob] = []
        for table_row in range(self._table.rowCount()):
            check = self._table.item(table_row, _COL_CHECK)
            payload = self._table.item(table_row, 1)
            if check is None or payload is None:
                continue
            if check.checkState() != Qt.CheckState.Checked:
                continue
            item = payload.data(_ROLE_ROW)
            if isinstance(item, SheetDeSyncRow):
                jobs.append(item.job)
        return tuple(jobs)

    def _build(self) -> None:
        layout = QVBoxLayout(self)
        intro = QLabel(SHEET_DE_SYNC_INTRO, self)
        intro.setWordWrap(True)
        layout.addWidget(intro)
        self._status = QLabel(self)
        layout.addWidget(self._status)

        self._table = QTableWidget(0, len(_HEADERS), self)
        self._table.setHorizontalHeaderLabels(list(_HEADERS))
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
        header.setMinimumSectionSize(28)
        self._table.setColumnWidth(_COL_CHECK, 36)
        self._table.setColumnWidth(1, 70)
        self._table.setColumnWidth(2, 80)
        self._table.setColumnWidth(3, 90)
        self._table.setColumnWidth(4, 90)
        self._table.setColumnWidth(5, 100)
        self._table.setColumnWidth(6, 140)
        self._table.setColumnWidth(7, 160)
        self._table.setColumnWidth(8, 70)
        self._table.itemChanged.connect(self._update_write_enabled)
        layout.addWidget(self._table, 1)

        buttons = QHBoxLayout()
        select_all = QPushButton("Выделить все", self)
        select_all.clicked.connect(lambda: self._set_all_checked(True))
        clear_all = QPushButton("Снять все", self)
        clear_all.clicked.connect(lambda: self._set_all_checked(False))
        self._write_button = QPushButton("Записать выбранное", self)
        self._write_button.clicked.connect(self._accept_write)
        close_button = QPushButton("Закрыть", self)
        close_button.clicked.connect(self.reject)
        buttons.addWidget(select_all)
        buttons.addWidget(clear_all)
        buttons.addStretch(1)
        buttons.addWidget(self._write_button)
        buttons.addWidget(close_button)
        layout.addLayout(buttons)

    def _fill(self) -> None:
        self._table.blockSignals(True)
        self._table.setRowCount(len(self._rows))
        for index, row in enumerate(self._rows):
            check = QTableWidgetItem()
            check.setFlags(
                Qt.ItemFlag.ItemIsUserCheckable
                | Qt.ItemFlag.ItemIsEnabled
                | Qt.ItemFlag.ItemIsSelectable
            )
            check.setCheckState(Qt.CheckState.Checked)
            self._table.setItem(index, _COL_CHECK, check)
            values = (
                row.title,
                row.mark,
                row.rd_revision_text or "—",
                row.d_now,
                row.d_next,
                row.e_now,
                row.e_next,
                row.write_label(),
                row.last_f_line,
            )
            tooltips = (
                row.last_f_line,
                row.last_f_line,
                f"Официальная рев. РД: {row.rd_revision_text or '—'}",
                f"Сейчас в D: {row.d_now}",
                (
                    f"В D запишется: {row.d_cell}"
                    if row.write_d and row.d_cell
                    else "D не меняем"
                ),
                f"Сейчас в E: {row.e_now}",
                (
                    f"В E запишется: {row.e_cell}"
                    if row.write_e and row.e_cell
                    else "E не меняем"
                ),
                f"Пишем столбцы: {row.write_label()}",
                row.last_f_line,
            )
            for column, text in enumerate(values, start=1):
                item = QTableWidgetItem(text)
                if column == 1:
                    item.setData(_ROLE_ROW, row)
                item.setToolTip(tooltips[column - 1])
                self._table.setItem(index, column, item)
            self._paint_change(index, row)
        self._table.blockSignals(False)
        self._status.setText(f"Комплектов: {len(self._rows)}")
        self._update_write_enabled()

    def _paint_change(self, table_row: int, row: SheetDeSyncRow) -> None:
        """Tint only cells that will change: wrong now, proposed next."""

        if row.write_d:
            self._set_fill(table_row, _COL_D_NOW, _WRONG_BG)
            self._set_fill(table_row, _COL_D_NEXT, _RIGHT_BG)
        if row.write_e:
            self._set_fill(table_row, _COL_E_NOW, _WRONG_BG)
            self._set_fill(table_row, _COL_E_NEXT, _RIGHT_BG)

    def _set_fill(self, table_row: int, column: int, color: QColor) -> None:
        item = self._table.item(table_row, column)
        if item is not None:
            item.setBackground(QBrush(color))

    def _set_all_checked(self, checked: bool) -> None:
        state = Qt.CheckState.Checked if checked else Qt.CheckState.Unchecked
        self._table.blockSignals(True)
        for table_row in range(self._table.rowCount()):
            item = self._table.item(table_row, _COL_CHECK)
            if item is not None:
                item.setCheckState(state)
        self._table.blockSignals(False)
        self._update_write_enabled()

    def _checked_count(self) -> int:
        return sum(
            1
            for table_row in range(self._table.rowCount())
            if (item := self._table.item(table_row, _COL_CHECK)) is not None
            and item.checkState() == Qt.CheckState.Checked
        )

    def _update_write_enabled(self) -> None:
        count = self._checked_count()
        self._write_button.setEnabled(count > 0)
        total = len(self._rows)
        self._status.setText(f"Выбрано: {count} / {total}")

    def _accept_write(self) -> None:
        if not self.selected_jobs():
            QMessageBox.information(
                self,
                SHEET_DE_SYNC_TITLE,
                "Нет выбранных строк для записи в КСБ ИД.",
            )
            return
        self.accept()
