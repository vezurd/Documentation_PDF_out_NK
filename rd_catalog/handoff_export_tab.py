"""Official-kit dump tab: preview table and copy destination controls.

Qt monitor only. Row building and copy stay in ``handoff_export``.
"""

from __future__ import annotations

from PySide6.QtCore import QPoint, QSettings, Qt, Signal, Slot
from PySide6.QtGui import QBrush, QColor, QFont
from PySide6.QtWidgets import (
    QAbstractItemView,
    QApplication,
    QButtonGroup,
    QCheckBox,
    QComboBox,
    QFileDialog,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QMenu,
    QPushButton,
    QRadioButton,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from rd_catalog.context_menu_qt import exec_tracked_menu
from rd_catalog.context_menu_usage import MENU_HANDOFF_EXPORT
from rd_catalog.handoff_export import (
    HANDOFF_HEADERS,
    HANDOFF_LAYOUT_FLAT,
    HANDOFF_LAYOUT_TITLE_MARK,
    HANDOFF_PROBLEM_FILL,
    HandoffDocCell,
    HandoffExportRow,
    load_handoff_destinations,
    remember_handoff_destination,
)
from rd_catalog.path_actions import open_path

_ROLE_ROW = Qt.ItemDataRole.UserRole
_ROLE_SORT = Qt.ItemDataRole.UserRole + 1
_COL_MTO = HANDOFF_HEADERS.index("MTO")
_COL_BOE = HANDOFF_HEADERS.index("BOE")
_COL_BOM = HANDOFF_HEADERS.index("BOM")
_COL_BOQ = HANDOFF_HEADERS.index("BOQ")
_COL_NOTES = HANDOFF_HEADERS.index("Примечание")
_COL_PACKAGE = HANDOFF_HEADERS.index("Пакет")
_COUNTS_EMPTY = "Комплекты: 0 · проблемы: 0 · файлов: 0"
_SETTINGS_INCLUDE_TDO = "window/handoff_export_include_tdo"
_SETTINGS_LAYOUT = "window/handoff_export_layout"


class _SortItem(QTableWidgetItem):
    """Table item that sorts by ``_ROLE_SORT`` when both sides have it."""

    def __lt__(self, other: QTableWidgetItem) -> bool:
        left = self.data(_ROLE_SORT)
        right = other.data(_ROLE_SORT) if other is not None else None
        if left is not None and right is not None:
            return bool(left < right)
        return super().__lt__(other)


class HandoffExportTab(QWidget):
    """Preview official-kit dump rows and request a copy to a dest folder."""

    kit_activated = Signal(str, str)
    copy_requested = Signal()
    filters_changed = Signal()

    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self._runtime_dir = ""
        self._rows: tuple[HandoffExportRow, ...] = ()
        self._copy_enabled = True
        self._build()

    def table(self) -> QTableWidget:
        """Return the preview table.

        Returns:
            The dump ``QTableWidget``.
        """

        return self._table

    def include_tdo(self) -> bool:
        """Return whether sent/passed TDO kits without code A are included.

        Returns:
            True when the TDO checkbox is on.
        """

        return self._include_tdo.isChecked()

    def layout_mode(self) -> str:
        """Return the destination layout key.

        Returns:
            ``title_mark`` or ``flat``.
        """

        if self._layout_flat.isChecked():
            return HANDOFF_LAYOUT_FLAT
        return HANDOFF_LAYOUT_TITLE_MARK

    def dest_root(self) -> str:
        """Return the typed or remembered dump folder.

        Returns:
            Stripped destination path, possibly empty.
        """

        return self._dest_combo.currentText().strip()

    def visible_rows(self) -> tuple[HandoffExportRow, ...]:
        """Return all loaded preview rows.

        Returns:
            The last ``set_rows`` payload (no extra text filter).
        """

        return self._rows

    def is_copy_running(self) -> bool:
        """Return whether this tab owns a running copy thread.

        Returns:
            Always False: the catalog window owns ``HandoffExportThread``.
        """

        return False

    def set_copy_enabled(self, enabled: bool) -> None:
        """Enable or disable the dump button.

        Args:
            enabled: False while a catalog worker is busy.
        """

        self._copy_enabled = bool(enabled)
        self._copy_button.setEnabled(self._copy_enabled)

    def configure(self, runtime_dir: str) -> None:
        """Load remembered dump roots into the destination combo.

        Args:
            runtime_dir: Catalog runtime folder (local JSON, never UNC).
        """

        self._runtime_dir = str(runtime_dir or "")
        store = load_handoff_destinations(self._runtime_dir)
        current = self.dest_root()
        self._dest_combo.blockSignals(True)
        try:
            self._dest_combo.clear()
            for path in store.paths:
                self._dest_combo.addItem(path)
            chosen = store.last or current
            if chosen:
                if self._dest_combo.findText(chosen) < 0:
                    self._dest_combo.insertItem(0, chosen)
                self._dest_combo.setCurrentText(chosen)
        finally:
            self._dest_combo.blockSignals(False)

    def set_rows(self, rows: tuple[HandoffExportRow, ...]) -> None:
        """Replace the preview table.

        Args:
            rows: Domain preview rows (already filtered by ban / eligibility).
        """

        self._rows = tuple(rows)
        self._rebuild_table()

    def restore_filters(self, settings: QSettings) -> None:
        """Load include-TDO and layout from QSettings without applying twice.

        Args:
            settings: Catalog window QSettings.
        """

        self._include_tdo.blockSignals(True)
        self._layout_title_mark.blockSignals(True)
        self._layout_flat.blockSignals(True)
        try:
            self._include_tdo.setChecked(
                _settings_bool(settings, _SETTINGS_INCLUDE_TDO, False)
            )
            layout = str(settings.value(_SETTINGS_LAYOUT) or "").strip()
            if layout == HANDOFF_LAYOUT_FLAT:
                self._layout_flat.setChecked(True)
            else:
                self._layout_title_mark.setChecked(True)
        finally:
            self._include_tdo.blockSignals(False)
            self._layout_title_mark.blockSignals(False)
            self._layout_flat.blockSignals(False)

    def save_filters(self, settings: QSettings) -> None:
        """Persist include-TDO and layout into QSettings.

        Args:
            settings: Catalog window QSettings.
        """

        settings.setValue(_SETTINGS_INCLUDE_TDO, self.include_tdo())
        settings.setValue(_SETTINGS_LAYOUT, self.layout_mode())

    def _build(self) -> None:
        layout = QVBoxLayout(self)

        dest_row = QHBoxLayout()
        dest_row.addWidget(QLabel("Папка:"))
        self._dest_combo = QComboBox(self)
        self._dest_combo.setEditable(True)
        self._dest_combo.setInsertPolicy(QComboBox.InsertPolicy.NoInsert)
        self._dest_combo.setMinimumContentsLength(28)
        dest_row.addWidget(self._dest_combo, 1)
        self._browse_button = QPushButton("Обзор…", self)
        self._browse_button.clicked.connect(self._on_browse)
        dest_row.addWidget(self._browse_button)
        layout.addLayout(dest_row)

        options = QHBoxLayout()
        self._layout_title_mark = QRadioButton("По титул / марка", self)
        self._layout_flat = QRadioButton("Плоская", self)
        self._layout_title_mark.setChecked(True)
        group = QButtonGroup(self)
        group.addButton(self._layout_title_mark)
        group.addButton(self._layout_flat)
        options.addWidget(self._layout_title_mark)
        options.addWidget(self._layout_flat)
        self._include_tdo = QCheckBox("Включая отправленные на ТДО", self)
        self._include_tdo.toggled.connect(self._on_include_tdo_toggled)
        options.addWidget(self._include_tdo)
        self._copy_button = QPushButton("Выгрузить", self)
        self._copy_button.clicked.connect(self.copy_requested.emit)
        options.addWidget(self._copy_button)
        options.addStretch(1)
        layout.addLayout(options)

        self._coverage = QLabel(_COUNTS_EMPTY, self)
        coverage_font = QFont(self._coverage.font())
        coverage_font.setBold(True)
        self._coverage.setFont(coverage_font)
        layout.addWidget(self._coverage)

        self._table = QTableWidget(0, len(HANDOFF_HEADERS), self)
        self._table.setHorizontalHeaderLabels(list(HANDOFF_HEADERS))
        self._table.setSelectionBehavior(
            QAbstractItemView.SelectionBehavior.SelectRows
        )
        self._table.setSelectionMode(
            QAbstractItemView.SelectionMode.SingleSelection
        )
        self._table.setEditTriggers(
            QAbstractItemView.EditTrigger.NoEditTriggers
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
        self._table.cellDoubleClicked.connect(self._on_cell_activated)
        self._table.setContextMenuPolicy(
            Qt.ContextMenuPolicy.CustomContextMenu
        )
        self._table.customContextMenuRequested.connect(
            self._show_context_menu
        )
        layout.addWidget(self._table, 1)

    def _rebuild_table(self) -> None:
        table = self._table
        table.setSortingEnabled(False)
        table.setRowCount(len(self._rows))
        problems = 0
        files = 0
        for index, row in enumerate(self._rows):
            if row.has_problem:
                problems += 1
            files += len(row.copy_files)
            values = (
                (row.title, row.title),
                (row.mark, row.mark.casefold()),
                (row.official_revision_text, row.official_revision_text),
                (row.status_label, row.status_label),
                ("да" if row.code_a else "нет", int(row.code_a)),
                (row.mto.text, row.mto.text),
                (row.boe.text, row.boe.text),
                (row.bom.text, row.bom.text),
                (row.boq.text, row.boq.text),
                (row.notes, row.notes),
                (row.package_label, row.package_label),
            )
            for column, (text, sort_key) in enumerate(values):
                item = _SortItem(str(text))
                item.setData(_ROLE_ROW, row)
                item.setData(_ROLE_SORT, sort_key)
                table.setItem(index, column, item)
            self._paint_row(index, row)
        table.setSortingEnabled(True)
        self._coverage.setText(
            f"Комплекты: {len(self._rows)} · проблемы: {problems} · файлов: {files}"
        )
        if self._rows:
            table.selectRow(0)

    def _paint_row(self, index: int, row: HandoffExportRow) -> None:
        table = self._table
        for column, cell in (
            (_COL_MTO, row.mto),
            (_COL_BOE, row.boe),
            (_COL_BOM, row.bom),
            (_COL_BOQ, row.boq),
        ):
            item = table.item(index, column)
            if item is not None:
                _paint_doc_cell(item, cell)
        notes_item = table.item(index, _COL_NOTES)
        if notes_item is not None and row.has_problem:
            _paint_fill(notes_item, HANDOFF_PROBLEM_FILL)
            notes_item.setToolTip(row.notes)
        package_item = table.item(index, _COL_PACKAGE)
        if package_item is not None and row.package_path:
            package_item.setToolTip(row.package_path)

    def _row_at(self, row_index: int) -> HandoffExportRow | None:
        if row_index < 0:
            return None
        item = self._table.item(row_index, 0)
        payload = item.data(_ROLE_ROW) if item else None
        return payload if isinstance(payload, HandoffExportRow) else None

    @Slot()
    def _on_include_tdo_toggled(self) -> None:
        self.filters_changed.emit()

    @Slot()
    def _on_browse(self) -> None:
        folder = QFileDialog.getExistingDirectory(
            self,
            "Папка выгрузки",
            self.dest_root(),
            QFileDialog.Option.ShowDirsOnly,
        )
        if not folder:
            return
        if self._runtime_dir:
            remember_handoff_destination(self._runtime_dir, folder)
            self.configure(self._runtime_dir)
        elif self._dest_combo.findText(folder) < 0:
            self._dest_combo.insertItem(0, folder)
        self._dest_combo.setCurrentText(folder)

    @Slot(int, int)
    def _on_cell_activated(self, row: int, _column: int) -> None:
        payload = self._row_at(row)
        if payload is None:
            return
        self.kit_activated.emit(payload.title, payload.mark)

    def _show_context_menu(self, position: QPoint) -> None:
        table = self._table
        row_index = table.rowAt(position.y())
        if row_index < 0:
            return
        table.selectRow(row_index)
        column = table.columnAt(position.x())
        if column >= 0:
            table.setCurrentCell(row_index, column)
        payload = self._row_at(row_index)
        if payload is None:
            return
        menu = QMenu(self)
        show_kits = menu.addAction("Показать в Комплекты")
        open_pkg = menu.addAction("Открыть папку пакета")
        copy_path = menu.addAction("Копировать путь")
        open_pkg.setEnabled(bool(payload.package_path))
        copy_path.setEnabled(bool(payload.package_path))
        chosen = exec_tracked_menu(
            menu, MENU_HANDOFF_EXPORT, table.viewport().mapToGlobal(position)
        )
        if chosen == show_kits:
            self.kit_activated.emit(payload.title, payload.mark)
        elif chosen == open_pkg:
            open_path(payload.package_path)
        elif chosen == copy_path:
            _copy_text(payload.package_path)


def _paint_doc_cell(item: QTableWidgetItem, cell: HandoffDocCell) -> None:
    if cell.fill_hex:
        _paint_fill(item, cell.fill_hex)
    if cell.path:
        item.setToolTip(cell.path)


def _paint_fill(item: QTableWidgetItem, hex_color: str) -> None:
    color = QColor(hex_color)
    if not color.isValid():
        return
    item.setBackground(QBrush(color))
    item.setForeground(
        QBrush(
            QColor("#ffffff")
            if color.lightness() < 140
            else QColor("#202124")
        )
    )


def _copy_text(text: str) -> None:
    if text:
        QApplication.clipboard().setText(text)


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
