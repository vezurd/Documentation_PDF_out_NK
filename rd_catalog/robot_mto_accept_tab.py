"""Qt tab listing stored ``robot_mto_accept`` rows.

Join and live/stale status stay in ``robot_mto_accept``. This widget
does not walk UNC or rebuild the pipeline.
"""

from __future__ import annotations

from collections.abc import Callable, Sequence
from pathlib import Path

from PySide6.QtCore import QPoint, QSettings, Qt, Signal, Slot
from PySide6.QtGui import QBrush, QColor, QFont
from PySide6.QtWidgets import (
    QAbstractItemView,
    QCheckBox,
    QDialog,
    QDialogButtonBox,
    QFormLayout,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QMenu,
    QTableWidget,
    QTableWidgetItem,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from rd_catalog.context_menu_qt import exec_tracked_menu
from rd_catalog.context_menu_usage import MENU_ROBOT_MTO_ACCEPT
from rd_catalog.models import FileRecord
from rd_catalog.path_actions import open_path
from rd_catalog.robot_mto_accept import (
    ACCEPT_TAB_HEADERS,
    ROBOT_MTO_ACCEPT_FOREGROUND,
    RobotMtoAcceptView,
    record_revision_text,
)

_ROLE_ROW = Qt.ItemDataRole.UserRole
_ROLE_SORT = Qt.ItemDataRole.UserRole + 1
_STALE_FILL = "#F8E3B0"
_MISSING_FILL = "#E8EAED"
_FILTER_TEXT_KEY = "window/robot_mto_accept_filter"
_FILTER_STALE_KEY = "window/robot_mto_accept_stale_only"


class RobotMtoAcceptDialog(QDialog):
    """Confirm that the robot MTO is accepted vs the official RD MTO."""

    def __init__(
        self,
        *,
        title: str,
        mark: str,
        robot: FileRecord,
        rd: FileRecord,
        parent: QWidget | None = None,
    ) -> None:
        super().__init__(parent)
        self.setWindowTitle("Подтвердить MTO робота")
        layout = QVBoxLayout(self)
        intro = QLabel(
            "Зелёная заливка и синий текст в «Робот МТО · рев.»: "
            "этот файл робота — подтверждённая замена MTO официальной "
            "папки РД с корректировками. Это не копия по дате и не "
            "сверка байтов. Имя файла робота не меняется."
        )
        intro.setWordWrap(True)
        layout.addWidget(intro)
        form = QFormLayout()
        form.addRow("Комплект:", QLabel(f"{title}-{mark}"))
        form.addRow(
            "MTO робота:",
            QLabel(
                f"{record_revision_text(robot) or '—'}  {Path(robot.path).name}"
            ),
        )
        form.addRow("Путь робота:", QLabel(robot.path))
        form.addRow(
            "Эталон MTO РД:",
            QLabel(f"{record_revision_text(rd) or '—'}  {Path(rd.path).name}"),
        )
        form.addRow("Путь РД:", QLabel(rd.path))
        layout.addLayout(form)
        self._comment_edit = QTextEdit(self)
        self._comment_edit.setPlaceholderText("Комментарий (необязательно)")
        self._comment_edit.setFixedHeight(64)
        layout.addWidget(self._comment_edit)
        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok
            | QDialogButtonBox.StandardButton.Cancel
        )
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def comment(self) -> str:
        """Return the optional operator note."""

        return self._comment_edit.toPlainText().strip()


class RobotMtoAcceptTab(QWidget):
    """Listing of stored robot-MTO accepts with live/stale status."""

    kit_activated = Signal(str, str)
    unmark_requested = Signal(str, str)
    prepare_context_menu = Signal(QMenu)

    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self._rows: tuple[RobotMtoAcceptView, ...] = ()
        self._is_banned: Callable[[str, str], bool] = lambda _t, _m: False
        layout = QVBoxLayout(self)
        layout.setContentsMargins(8, 8, 8, 8)
        hint = QLabel(
            "Подтверждённые замены MTO робота относительно эталона "
            "официальной папки РД. «актуально» — зелёная ячейка и синий "
            "текст в Комплектах; «устарело» — сменился эталон РД; "
            "«нет файла» — другой файл робота."
        )
        hint.setWordWrap(True)
        layout.addWidget(hint)
        filters = QHBoxLayout()
        self._filter = QLineEdit(self)
        self._filter.setPlaceholderText("Фильтр: титул, марка, статус, файл…")
        self._filter.textChanged.connect(self._apply_row_visibility)
        self._stale_only = QCheckBox("Только устаревшие", self)
        self._stale_only.toggled.connect(self._apply_row_visibility)
        filters.addWidget(self._filter, 1)
        filters.addWidget(self._stale_only)
        layout.addLayout(filters)
        self._table = QTableWidget(0, len(ACCEPT_TAB_HEADERS), self)
        self._table.setHorizontalHeaderLabels(list(ACCEPT_TAB_HEADERS))
        self._table.setSelectionBehavior(
            QAbstractItemView.SelectionBehavior.SelectRows
        )
        self._table.setSelectionMode(
            QAbstractItemView.SelectionMode.SingleSelection
        )
        self._table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
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
        self._table.cellDoubleClicked.connect(self._on_cell_activated)
        self._table.setContextMenuPolicy(
            Qt.ContextMenuPolicy.CustomContextMenu
        )
        self._table.customContextMenuRequested.connect(self._show_context_menu)
        layout.addWidget(self._table, 1)

    def table(self) -> QTableWidget:
        """Return the listing table for header persistence."""

        return self._table

    def set_rows(
        self,
        rows: Sequence[RobotMtoAcceptView],
        *,
        is_banned: Callable[[str, str], bool],
    ) -> None:
        """Replace the snapshot and rebuild the table.

        Args:
            rows: Joined accept views.
            is_banned: Hide banned ``(title, mark)`` pairs.
        """

        self._rows = tuple(rows)
        self._is_banned = is_banned
        self._rebuild_table()

    def restore_filters(self, settings: QSettings) -> None:
        """Load filter widgets from QSettings."""

        self._filter.blockSignals(True)
        self._stale_only.blockSignals(True)
        try:
            self._filter.setText(str(settings.value(_FILTER_TEXT_KEY) or ""))
            self._stale_only.setChecked(
                _settings_bool(settings, _FILTER_STALE_KEY, False)
            )
        finally:
            self._filter.blockSignals(False)
            self._stale_only.blockSignals(False)
        self._apply_row_visibility()

    def save_filters(self, settings: QSettings) -> None:
        """Persist filter widgets."""

        settings.setValue(_FILTER_TEXT_KEY, self._filter.text())
        settings.setValue(_FILTER_STALE_KEY, self._stale_only.isChecked())

    def current_view(self) -> RobotMtoAcceptView | None:
        """Return the selected row payload, if any."""

        return self._row_at(self._table.currentRow())

    def _rebuild_table(self) -> None:
        table = self._table
        table.setSortingEnabled(False)
        table.setRowCount(len(self._rows))
        for row_index, view in enumerate(self._rows):
            values = (
                view.title,
                view.mark,
                view.robot_revision_text,
                view.rd_revision_text,
                view.status,
                view.decided_at,
                view.robot_name,
                view.rd_name,
                view.robot_path,
            )
            for column, text in enumerate(values):
                item = QTableWidgetItem(str(text or "—"))
                item.setData(_ROLE_ROW, view)
                item.setData(_ROLE_SORT, str(text or "").casefold())
                if column == 4:
                    self._paint_status(item, view)
                table.setItem(row_index, column, item)
        table.setSortingEnabled(True)
        self._apply_row_visibility()

    def _paint_status(
        self, item: QTableWidgetItem, view: RobotMtoAcceptView
    ) -> None:
        status = view.status.casefold()
        if status == "актуально":
            item.setForeground(QBrush(QColor(ROBOT_MTO_ACCEPT_FOREGROUND)))
            font = QFont(item.font())
            font.setBold(True)
            item.setFont(font)
            return
        if status == "устарело":
            item.setBackground(QBrush(QColor(_STALE_FILL)))
            return
        if status == "нет файла":
            item.setBackground(QBrush(QColor(_MISSING_FILL)))

    def _row_at(self, row_index: int) -> RobotMtoAcceptView | None:
        if row_index < 0:
            return None
        item = self._table.item(row_index, 0)
        if item is None:
            return None
        payload = item.data(_ROLE_ROW)
        return payload if isinstance(payload, RobotMtoAcceptView) else None

    @Slot()
    def _apply_row_visibility(self) -> None:
        needle = self._filter.text().strip().casefold()
        stale_only = self._stale_only.isChecked()
        for row_index in range(self._table.rowCount()):
            view = self._row_at(row_index)
            if view is None:
                self._table.setRowHidden(row_index, True)
                continue
            banned = self._is_banned(view.title, view.mark)
            haystack = view.haystack
            visible = not banned
            if visible and needle:
                visible = needle in haystack
            if visible and stale_only:
                visible = view.status.casefold() in {"устарело", "нет файла"}
            self._table.setRowHidden(row_index, not visible)

    @Slot(int, int)
    def _on_cell_activated(self, row_index: int, _column: int) -> None:
        view = self._row_at(row_index)
        if view is None:
            return
        self.kit_activated.emit(view.title, view.mark)

    def _show_context_menu(self, position: QPoint) -> None:
        table = self._table
        row_index = table.rowAt(position.y())
        if row_index < 0:
            return
        table.selectRow(row_index)
        view = self._row_at(row_index)
        if view is None:
            return
        menu = QMenu(self)
        open_folder = menu.addAction("Открыть папку робота")
        parent = str(Path(view.robot_path).parent) if view.robot_path else ""
        open_folder.setEnabled(bool(parent))
        show_kits = menu.addAction('Показать в «Комплекты»')
        unmark = menu.addAction("Снять подтверждение")
        self.prepare_context_menu.emit(menu)
        chosen = exec_tracked_menu(
            menu,
            MENU_ROBOT_MTO_ACCEPT,
            table.viewport().mapToGlobal(position),
        )
        if chosen == open_folder and parent:
            open_path(parent)
        elif chosen == show_kits:
            self.kit_activated.emit(view.title, view.mark)
        elif chosen == unmark:
            self.unmark_requested.emit(view.title, view.mark)


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
