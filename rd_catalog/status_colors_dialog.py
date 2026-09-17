"""Dialog to edit pipeline and F-stage status colors."""

from __future__ import annotations

from collections.abc import Mapping
from pathlib import Path

from PySide6.QtCore import Qt, Signal, Slot
from PySide6.QtGui import QBrush, QColor
from PySide6.QtWidgets import (
    QAbstractItemView,
    QColorDialog,
    QDialog,
    QDialogButtonBox,
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

from rd_catalog.status_colors import (
    UNKNOWN_STATUS_COLOR,
    color_for,
    default_palette,
    load_status_colors,
    save_status_colors,
    status_color_label,
)

_ROLE_KEY = Qt.ItemDataRole.UserRole


class StatusColorsDialog(QDialog):
    """Edit and persist the local status-color palette."""

    changed = Signal()

    def __init__(
        self,
        path: str | Path,
        parent: QWidget | None = None,
        *,
        palette: Mapping[str, str] | None = None,
    ) -> None:
        """Build the dialog bound to a runtime JSON path.

        Args:
            path: ``status_colors.json`` under the catalog runtime directory.
            parent: Optional Qt parent.
            palette: Optional in-memory palette; loaded from ``path`` if omitted.
        """

        super().__init__(parent)
        self._path = Path(path)
        self._draft = dict(palette if palette is not None else load_status_colors(path))
        self.setWindowTitle("Цвета статусов")
        self.setMinimumSize(560, 420)
        self.resize(640, 520)
        self._build()
        self._reload_table()

    def _build(self) -> None:
        layout = QVBoxLayout(self)
        hint = QLabel(
            "Цвета столбца «Статус рассмотрения», бейджа карточки и матрицы «Ревизии MTO». "
            "Файл — локальный status_colors.json, не сеть и не Google.",
            self,
        )
        hint.setWordWrap(True)
        layout.addWidget(hint)

        self._table = QTableWidget(0, 3, self)
        self._table.setHorizontalHeaderLabels(["Ключ", "Подпись", "Цвет"])
        self._table.setSelectionBehavior(
            QAbstractItemView.SelectionBehavior.SelectRows
        )
        self._table.setSelectionMode(
            QAbstractItemView.SelectionMode.SingleSelection
        )
        self._table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self._table.setAlternatingRowColors(True)
        self._table.verticalHeader().setVisible(False)
        header = self._table.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        header.setStretchLastSection(True)
        header.setMinimumSectionSize(40)
        self._table.setColumnWidth(0, 140)
        self._table.setColumnWidth(1, 260)
        self._table.cellDoubleClicked.connect(self._on_cell_double_clicked)
        layout.addWidget(self._table, 1)

        row = QHBoxLayout()
        pick_button = QPushButton("Выбрать цвет…", self)
        pick_button.clicked.connect(self._on_pick_color)
        reset_button = QPushButton("Сбросить", self)
        reset_button.setToolTip("Вернуть заводскую палитру (пока не сохранено).")
        reset_button.clicked.connect(self._on_reset)
        row.addWidget(pick_button)
        row.addWidget(reset_button)
        row.addStretch(1)
        layout.addLayout(row)

        box = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Save
            | QDialogButtonBox.StandardButton.Cancel,
            self,
        )
        box.accepted.connect(self._on_save)
        box.rejected.connect(self.reject)
        layout.addWidget(box)

    def _ordered_keys(self) -> list[str]:
        keys = list(default_palette())
        extras = [key for key in self._draft if key not in keys]
        extras.sort()
        keys.extend(extras)
        return keys

    def _reload_table(self, select_key: str | None = None) -> None:
        current = self._selected_key() if select_key is None else select_key
        self._table.setSortingEnabled(False)
        self._table.setRowCount(0)
        for key in self._ordered_keys():
            hex_color = color_for(self._draft, key)
            values = [key, status_color_label(key), hex_color]
            row_index = self._table.rowCount()
            self._table.insertRow(row_index)
            for column, value in enumerate(values):
                item = QTableWidgetItem(value)
                item.setData(_ROLE_KEY, key)
                if column == 2:
                    color = QColor(hex_color)
                    item.setBackground(QBrush(color))
                    item.setForeground(
                        QBrush(
                            QColor("#ffffff")
                            if color.lightness() < 140
                            else QColor("#202124")
                        )
                    )
                self._table.setItem(row_index, column, item)
        self._table.setSortingEnabled(True)
        if current:
            for row_index in range(self._table.rowCount()):
                item = self._table.item(row_index, 0)
                if item is not None and item.data(_ROLE_KEY) == current:
                    self._table.selectRow(row_index)
                    break

    def _selected_key(self) -> str | None:
        row = self._table.currentRow()
        item = self._table.item(row, 0) if row >= 0 else None
        key = item.data(_ROLE_KEY) if item else None
        return str(key) if key else None

    def _pick_for_key(self, key: str) -> None:
        current = QColor(color_for(self._draft, key) or UNKNOWN_STATUS_COLOR)
        chosen = QColorDialog.getColor(current, self, "Цвет статуса")
        if not chosen.isValid():
            return
        self._draft[key] = chosen.name().casefold()
        self._reload_table(select_key=key)

    @Slot(int, int)
    def _on_cell_double_clicked(self, _row: int, column: int) -> None:
        key = self._selected_key()
        if key is None or column != 2:
            return
        self._pick_for_key(key)

    @Slot()
    def _on_pick_color(self) -> None:
        key = self._selected_key()
        if key is None:
            return
        self._pick_for_key(key)

    @Slot()
    def _on_reset(self) -> None:
        self._draft = default_palette()
        self._reload_table()

    @Slot()
    def _on_save(self) -> None:
        try:
            save_status_colors(self._path, self._draft)
        except OSError as exc:
            QMessageBox.warning(
                self,
                "Цвета статусов",
                f"Не удалось сохранить палитру:\n{exc}",
            )
            return
        self.changed.emit()
        self.accept()
