"""Excel column settings tab for ``ds_compare_center``."""

from __future__ import annotations

from typing import Any, Callable

from PySide6.QtCore import Qt
from PySide6.QtGui import QColor
from PySide6.QtWidgets import (
    QAbstractItemView,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QMessageBox,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
    QColorDialog,
)

from RFQ.ds_compare.ds_compare_config import (
    load_ds_compare_config,
    normalize_ds_vs_mto_output,
    save_ds_compare_config,
)
from RFQ.ds_compare.ds_vs_mto_excel_columns import (
    MAX_COLUMN_WIDTH,
    MIN_COLUMN_WIDTH,
    SECTION_LABELS,
    default_columns_settings,
    merge_column_settings,
    normalize_header_color,
)

_COL_OUTPUT = 0
_COL_NAME = 1
_COL_SECTION = 2
_COL_LABEL = 3
_COL_WIDTH = 4
_COL_COLOR = 5


class ColumnsPanel(QWidget):
    """Editor for DS vs MTO Excel column visibility, labels, widths, and header colors."""

    def __init__(
        self,
        parent: QWidget | None = None,
        *,
        on_saved: Callable[[], None] | None = None,
    ) -> None:
        super().__init__(parent)
        self._on_saved = on_saved

        hint = QLabel(
            "Настройки применяются к финальному Excel (_xw). "
            "Колонка «Код» — внутреннее имя поля в RowStd.",
            self,
        )
        hint.setWordWrap(True)

        self._table = QTableWidget(self)
        self._table.setColumnCount(6)
        self._table.setHorizontalHeaderLabels(
            [
                "Вывод",
                "Код",
                "Блок",
                "Название для Excel",
                "Ширина",
                "Цвет заголовка",
            ]
        )
        self._table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self._table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self._table.verticalHeader().setVisible(False)
        header = self._table.horizontalHeader()
        header.setSectionResizeMode(_COL_OUTPUT, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(_COL_NAME, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(_COL_SECTION, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(_COL_LABEL, QHeaderView.ResizeMode.Stretch)
        header.setSectionResizeMode(_COL_WIDTH, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(_COL_COLOR, QHeaderView.ResizeMode.ResizeToContents)

        btn_row = QHBoxLayout()
        btn_reload = QPushButton("Перезагрузить", self)
        btn_reload.clicked.connect(self.reload_from_disk)
        btn_reset = QPushButton("Сбросить к умолчанию", self)
        btn_reset.clicked.connect(self.reset_to_defaults)
        btn_color = QPushButton("Цвет выбранной строки…", self)
        btn_color.clicked.connect(self._pick_color_for_selected_row)
        btn_save = QPushButton("Сохранить", self)
        btn_save.clicked.connect(self.save_to_disk)
        btn_row.addWidget(btn_reload)
        btn_row.addWidget(btn_reset)
        btn_row.addWidget(btn_color)
        btn_row.addStretch(1)
        btn_row.addWidget(btn_save)

        layout = QVBoxLayout(self)
        layout.addWidget(hint)
        layout.addWidget(self._table, stretch=1)
        layout.addLayout(btn_row)

        self.reload_from_disk()

    def reload_from_disk(self) -> None:
        cfg = load_ds_compare_config()
        out = normalize_ds_vs_mto_output(cfg.get("ds_vs_mto_output"))
        self._populate_table(out.get("columns") or default_columns_settings())

    def reset_to_defaults(self) -> None:
        self._populate_table(default_columns_settings())

    def _populate_table(self, rows: list[dict[str, Any]]) -> None:
        self._table.setRowCount(len(rows))
        for row_idx, item in enumerate(rows):
            output_item = QTableWidgetItem()
            output_item.setFlags(
                Qt.ItemFlag.ItemIsUserCheckable | Qt.ItemFlag.ItemIsEnabled
            )
            output_item.setCheckState(
                Qt.CheckState.Checked if item.get("output", True) else Qt.CheckState.Unchecked
            )
            self._table.setItem(row_idx, _COL_OUTPUT, output_item)

            name_item = QTableWidgetItem(str(item.get("col_name", "")))
            name_item.setFlags(Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsSelectable)
            self._table.setItem(row_idx, _COL_NAME, name_item)

            section = str(item.get("section", "ds"))
            section_item = QTableWidgetItem(SECTION_LABELS.get(section, section))
            section_item.setData(Qt.ItemDataRole.UserRole, section)
            section_item.setFlags(Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsSelectable)
            self._table.setItem(row_idx, _COL_SECTION, section_item)

            label_item = QTableWidgetItem(str(item.get("header_label", "")))
            label_item.setFlags(
                Qt.ItemFlag.ItemIsEnabled
                | Qt.ItemFlag.ItemIsSelectable
                | Qt.ItemFlag.ItemIsEditable
            )
            self._table.setItem(row_idx, _COL_LABEL, label_item)

            width_item = QTableWidgetItem(str(int(item.get("width", 15))))
            width_item.setFlags(
                Qt.ItemFlag.ItemIsEnabled
                | Qt.ItemFlag.ItemIsSelectable
                | Qt.ItemFlag.ItemIsEditable
            )
            self._table.setItem(row_idx, _COL_WIDTH, width_item)

            color = normalize_header_color(
                item.get("header_color"),
                fallback="#dbbcdb",
            )
            color_item = QTableWidgetItem(color)
            color_item.setFlags(
                Qt.ItemFlag.ItemIsEnabled
                | Qt.ItemFlag.ItemIsSelectable
                | Qt.ItemFlag.ItemIsEditable
            )
            color_item.setBackground(QColor(color))
            self._table.setItem(row_idx, _COL_COLOR, color_item)

    def _pick_color_for_selected_row(self) -> None:
        row = self._table.currentRow()
        if row < 0:
            QMessageBox.information(self, "Цвет", "Выберите строку в таблице.")
            return
        color_item = self._table.item(row, _COL_COLOR)
        if color_item is None:
            return
        initial = QColor(color_item.text().strip() or "#dbbcdb")
        picked = QColorDialog.getColor(initial, self, "Цвет заголовка колонки")
        if not picked.isValid():
            return
        hex_color = picked.name().lower()
        color_item.setText(hex_color)
        color_item.setBackground(picked)

    def _collect_columns(self) -> list[dict[str, Any]] | None:
        rows: list[dict[str, Any]] = []
        for row_idx in range(self._table.rowCount()):
            name_item = self._table.item(row_idx, _COL_NAME)
            section_item = self._table.item(row_idx, _COL_SECTION)
            label_item = self._table.item(row_idx, _COL_LABEL)
            width_item = self._table.item(row_idx, _COL_WIDTH)
            color_item = self._table.item(row_idx, _COL_COLOR)
            output_item = self._table.item(row_idx, _COL_OUTPUT)
            if not name_item or not section_item:
                continue

            try:
                width = int(str(width_item.text() if width_item else "15").strip())
            except ValueError:
                QMessageBox.warning(
                    self,
                    "Ширина",
                    f"Неверная ширина в строке {row_idx + 1}: "
                    f"{width_item.text() if width_item else ''!r}",
                )
                return None
            width = max(MIN_COLUMN_WIDTH, min(MAX_COLUMN_WIDTH, width))

            section = str(section_item.data(Qt.ItemDataRole.UserRole) or "ds")
            fallback = normalize_header_color(None, fallback="#dbbcdb")
            header_color = normalize_header_color(
                color_item.text() if color_item else "",
                fallback=fallback,
            )
            label = str(label_item.text() if label_item else "").strip()
            if not label:
                QMessageBox.warning(
                    self,
                    "Название",
                    f"Пустое название колонки в строке {row_idx + 1}.",
                )
                return None

            rows.append(
                {
                    "col_name": name_item.text().strip(),
                    "output": output_item.checkState() == Qt.CheckState.Checked
                    if output_item
                    else True,
                    "width": width,
                    "header_label": label,
                    "header_color": header_color,
                    "section": section,
                }
            )
        return merge_column_settings(rows)

    def _collect_config(self) -> dict[str, Any]:
        cfg = load_ds_compare_config()
        columns = self._collect_columns()
        if columns is None:
            raise ValueError("invalid columns")
        prev = normalize_ds_vs_mto_output(cfg.get("ds_vs_mto_output"))
        cfg["ds_vs_mto_output"] = {
            "show_internal_column_names": bool(
                prev.get("show_internal_column_names", False)
            ),
            "expand_aggregated_replacement_codes": bool(
                prev.get("expand_aggregated_replacement_codes", True)
            ),
            "excel_export_mode": prev.get("excel_export_mode", "both"),
            "columns": columns,
        }
        return cfg

    def save_to_disk(self) -> bool:
        try:
            cfg = self._collect_config()
        except ValueError:
            return False
        if save_ds_compare_config(cfg):
            QMessageBox.information(
                self,
                "Сохранено",
                "Настройки столбцов записаны в ds_compare_config.json.",
            )
            if self._on_saved:
                self._on_saved()
            return True
        QMessageBox.critical(self, "Ошибка", "Не удалось сохранить файл конфигурации.")
        return False
