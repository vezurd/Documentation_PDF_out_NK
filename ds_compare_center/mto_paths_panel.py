"""MTO folder presets tab for ``ds_compare_center``."""

from __future__ import annotations

from copy import deepcopy
from typing import Any, Callable

from PySide6.QtWidgets import (
    QComboBox,
    QFileDialog,
    QFormLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QListWidget,
    QMessageBox,
    QPushButton,
    QScrollArea,
    QVBoxLayout,
    QWidget,
)

from RFQ.ds_compare.ds_compare_config import (
    MtoPathPresetDict,
    load_ds_compare_config,
    normalize_mto_path_entries,
    save_ds_compare_config,
)


class MtoPathsPanel(QWidget):
    """Editor for ``mto_paths`` and ``mto_path_selected_index`` in ds_compare_config."""

    def __init__(
        self,
        parent: QWidget | None = None,
        *,
        on_saved: Callable[[], None] | None = None,
    ) -> None:
        super().__init__(parent)
        self._on_saved = on_saved
        self._mto_items: list[MtoPathPresetDict] = []

        scroll = QScrollArea(self)
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QScrollArea.Shape.NoFrame)

        inner = QWidget()
        layout = QVBoxLayout(inner)
        layout.addWidget(self._build_mto_paths_group())
        layout.addStretch(1)
        scroll.setWidget(inner)

        btn_row = QHBoxLayout()
        btn_row.addStretch(1)
        btn_reload = QPushButton("Перезагрузить", self)
        btn_reload.clicked.connect(self.reload_from_disk)
        btn_save = QPushButton("Сохранить", self)
        btn_save.clicked.connect(self.save_to_disk)
        btn_row.addWidget(btn_reload)
        btn_row.addWidget(btn_save)

        root = QVBoxLayout(self)
        root.addWidget(scroll, stretch=1)
        root.addLayout(btn_row)

        self.reload_from_disk()

    def _build_mto_paths_group(self) -> QGroupBox:
        box = QGroupBox("Папки МТО (пресеты)", self)
        v = QVBoxLayout(box)

        hint = QLabel(
            "Список корневых папок МТО для сравнения с ДС. "
            "На вкладке «Запуск» выбирается активный пресет.",
            box,
        )
        hint.setWordWrap(True)
        v.addWidget(hint)

        self._mto_list = QListWidget(box)
        self._mto_list.currentRowChanged.connect(self._load_mto_row_to_form)
        v.addWidget(self._mto_list)

        form = QFormLayout()
        self._mto_label = QLineEdit(box)
        self._mto_path = QLineEdit(box)
        form.addRow("Краткое имя:", self._mto_label)
        form.addRow("Путь:", self._mto_path)
        v.addLayout(form)

        row = QHBoxLayout()
        btn_add = QPushButton("Добавить", box)
        btn_add.clicked.connect(self._mto_add)
        btn_remove = QPushButton("Удалить", box)
        btn_remove.clicked.connect(self._mto_remove)
        btn_apply = QPushButton("Обновить строку", box)
        btn_apply.clicked.connect(self._mto_apply_row)
        btn_browse = QPushButton("Обзор…", box)
        btn_browse.clicked.connect(self._mto_browse)
        row.addWidget(btn_add)
        row.addWidget(btn_remove)
        row.addWidget(btn_apply)
        row.addWidget(btn_browse)
        row.addStretch(1)
        v.addLayout(row)

        sel_row = QHBoxLayout()
        sel_row.addWidget(QLabel("Пресет по умолчанию:", box))
        self._mto_default_combo = QComboBox(box)
        sel_row.addWidget(self._mto_default_combo, stretch=1)
        v.addLayout(sel_row)
        return box

    def reload_from_disk(self) -> None:
        cfg = load_ds_compare_config()
        self._mto_items = deepcopy(normalize_mto_path_entries(cfg.get("mto_paths")))
        sel_idx = int(cfg.get("mto_path_selected_index", 0))
        if self._mto_items:
            sel_idx = max(0, min(sel_idx, len(self._mto_items) - 1))
        else:
            sel_idx = 0

        self._mto_list.blockSignals(True)
        self._mto_list.clear()
        for i, it in enumerate(self._mto_items):
            self._mto_list.addItem(f"{i + 1}. {it['label']}  →  {it['path']}")
        self._mto_list.blockSignals(False)
        if self._mto_items:
            self._mto_list.setCurrentRow(sel_idx)

        self._refresh_mto_default_combo(sel_idx)

    def _refresh_mto_default_combo(self, selected_index: int) -> None:
        self._mto_default_combo.blockSignals(True)
        self._mto_default_combo.clear()
        for i, it in enumerate(self._mto_items):
            self._mto_default_combo.addItem(f"{i + 1}. {it['label']}", i)
        if self._mto_items:
            idx = max(0, min(selected_index, len(self._mto_items) - 1))
            self._mto_default_combo.setCurrentIndex(idx)
        self._mto_default_combo.blockSignals(False)

    def _load_mto_row_to_form(self, row: int) -> None:
        if row < 0 or row >= len(self._mto_items):
            self._mto_label.clear()
            self._mto_path.clear()
            return
        it = self._mto_items[row]
        self._mto_label.setText(it["label"])
        self._mto_path.setText(it["path"])

    def _mto_add(self) -> None:
        self._mto_apply_row(silent=True)
        self._mto_items.append({"label": "Новая папка", "path": ""})
        self._mto_list.addItem(f"{len(self._mto_items)}. Новая папка  →  ")
        self._mto_list.setCurrentRow(len(self._mto_items) - 1)
        self._refresh_mto_default_combo(self._mto_list.currentRow())

    def _mto_remove(self) -> None:
        row = self._mto_list.currentRow()
        if row < 0:
            QMessageBox.warning(self, "Удаление", "Выберите строку в списке.")
            return
        del self._mto_items[row]
        self._mto_list.takeItem(row)
        for i in range(self._mto_list.count()):
            it = self._mto_items[i]
            self._mto_list.item(i).setText(f"{i + 1}. {it['label']}  →  {it['path']}")
        if self._mto_items:
            self._mto_list.setCurrentRow(min(row, len(self._mto_items) - 1))
        self._refresh_mto_default_combo(self._mto_list.currentRow())

    def _mto_apply_row(self, *, silent: bool = False) -> None:
        row = self._mto_list.currentRow()
        if row < 0:
            if not silent:
                QMessageBox.warning(self, "Применить", "Выберите строку в списке.")
            return
        path = self._mto_path.text().strip()
        if not path:
            if not silent:
                QMessageBox.warning(self, "Применить", "Путь к папке не может быть пустым.")
            return
        label = self._mto_label.text().strip() or path
        self._mto_items[row] = {"label": label, "path": path}
        self._mto_list.item(row).setText(f"{row + 1}. {label}  →  {path}")
        self._refresh_mto_default_combo(row)

    def _mto_browse(self) -> None:
        path = QFileDialog.getExistingDirectory(self, "Папка МТО")
        if path:
            self._mto_path.setText(path)

    def _collect_config(self) -> dict[str, Any]:
        self._mto_apply_row(silent=True)
        cfg = load_ds_compare_config()
        normalized = normalize_mto_path_entries(self._mto_items)
        cfg["mto_paths"] = normalized
        idx = self._mto_default_combo.currentData()
        if idx is None:
            idx = 0
        idx = int(idx)
        if normalized:
            idx = max(0, min(idx, len(normalized) - 1))
        else:
            idx = 0
        cfg["mto_path_selected_index"] = idx
        return cfg

    def save_to_disk(self) -> bool:
        cfg = self._collect_config()
        if save_ds_compare_config(cfg):
            QMessageBox.information(
                self,
                "Сохранено",
                "Пресеты МТО записаны в ds_compare_config.json.",
            )
            if self._on_saved:
                self._on_saved()
            return True
        QMessageBox.critical(self, "Ошибка", "Не удалось сохранить файл конфигурации.")
        return False
