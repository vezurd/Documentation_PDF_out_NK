"""Collapsible panel: prealign overlay layer visibility (QSettings + tree)."""

from __future__ import annotations

from PySide6.QtCore import QSettings, Qt, Signal
from PySide6.QtGui import QFont
from PySide6.QtWidgets import (
    QAbstractItemView,
    QGroupBox,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QPushButton,
    QToolButton,
    QTreeWidget,
    QTreeWidgetItem,
    QVBoxLayout,
    QWidget,
)

from .prealign_overlay_layers import LAYER_SPEC_BY_ID, PREALIGN_LAYER_SPECS

_SETTINGS_VISIBILITY_PREFIX = "prealign_overlay_visibility"
_SETTINGS_COLLAPSED = "prealign_overlay_panel_collapsed"
_SETTINGS_CANVAS_PREFIX = "canvas_layer_visibility"

# Слои рабочего поля (не PNG prealign): id, тип, подпись, подсказка 4-я колонка
CANVAS_LAYER_ROWS: list[tuple[str, str, str, str]] = [
    ("stamp_fields", "rect", "Поля штампа (прямоугольники)", "FieldRectItem"),
    ("template_grid", "line", "Каркас шаблона (линии wireframe)", "GridLineGraphicsItem"),
    ("pdf_grid", "line", "Каркас PDF (find_tables, merged)", "DetectedGridLineItem"),
    ("adapt_overlay", "mix", "Оверлей «Тест-адаптация»", "рамки, подписи, matched cells"),
    ("consolidation_overlay", "line", "Оверлей консолидации сетки", "линии кластеров"),
    ("origin_markers", "marker", "Маркеры начала координат", "кресты углов + линия к полю"),
]

CANVAS_ROWS_BY_ID: dict[str, tuple[str, str, str, str]] = {r[0]: r for r in CANVAS_LAYER_ROWS}

# Как нейтральный список (см. «Список полей»): без синего hover/selection системной темы
_TREE_STYLE_SHEET = """
QTreeWidget {
    show-decoration-selected: 0;
}
QTreeWidget::item {
    min-height: 22px;
    padding-top: 3px;
    padding-bottom: 3px;
}
QTreeWidget::item:hover:!selected {
    background-color: #ebebeb;
}
QTreeWidget::item:selected {
    background-color: #d8d8d8;
    color: #1a1a1a;
}
QTreeWidget::item:selected:hover {
    background-color: #cfcfcf;
    color: #1a1a1a;
}
"""


class PrealignOverlayPanel(QWidget):
    """Checklist of overlay layers; persists visibility in QSettings."""

    visibility_changed = Signal(str, bool)
    overlay_refresh_requested = Signal()
    detach_toggle_requested = Signal()
    canvas_layer_visibility_changed = Signal(str, bool)

    def __init__(self, settings: QSettings, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self._settings = settings
        self._tree = QTreeWidget(self)
        self._tree.setHeaderLabels(["Вкл.", "Тип", "Слой (полное имя)", "PNG-группа"])
        hdr = self._tree.header()
        hdr.setStretchLastSection(False)
        # 0 и 2 — тянутся пользователем; 1 и 3 — по содержимому: при полностью Interactive
        # узкий сплиттер сжимал «Тип» и «PNG-группа» почти до нуля, текст пропадал из ячеек.
        hdr.setMinimumSectionSize(32)
        hdr.setSectionResizeMode(0, QHeaderView.ResizeMode.Interactive)
        hdr.setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        hdr.setSectionResizeMode(2, QHeaderView.ResizeMode.Interactive)
        hdr.setSectionResizeMode(3, QHeaderView.ResizeMode.ResizeToContents)
        self._tree.setColumnWidth(0, 246)
        self._tree.setColumnWidth(2, 880)
        self._tree.setRootIsDecorated(True)
        self._tree.setIndentation(20)
        self._tree.setAlternatingRowColors(False)
        self._tree.setUniformRowHeights(True)
        self._tree.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self._tree.itemChanged.connect(self._on_item_changed)
        self._tree.itemDoubleClicked.connect(self._on_prealign_tree_double_clicked)
        self._tree.setStyleSheet(_TREE_STYLE_SHEET)

        bar = QHBoxLayout()
        bar.setContentsMargins(0, 0, 0, 2)
        self._lbl_bar_title = QLabel("Границы prealign / adapt (оверлей)", self)
        bf = self._lbl_bar_title.font()
        bf.setBold(True)
        self._lbl_bar_title.setFont(bf)
        bar.addWidget(self._lbl_bar_title)
        bar.addStretch(1)
        self._btn_detach = QToolButton(self)
        self._btn_detach.setText("⬆")
        self._btn_detach.setToolTip("Открыть в отдельном окне")
        self._btn_detach.setFixedWidth(26)
        self._btn_detach.clicked.connect(self.detach_toggle_requested.emit)
        bar.addWidget(self._btn_detach)

        grp = QGroupBox("Слои на canvas", self)
        lay = QVBoxLayout(grp)
        hint = QLabel(
            "Те же слои, что в PNG дампа adapt_debug. Координаты — displayed space PDF. "
            "Двойной клик по колонкам «Тип», «Слой», «PNG-группа» — режим редактирования для копирования текста."
        )
        hint.setWordWrap(True)
        lay.addWidget(hint)
        btn_row = QHBoxLayout()
        self._btn_all = QPushButton("Все", self)
        self._btn_none = QPushButton("Ни одного", self)
        btn_row.addWidget(self._btn_all)
        btn_row.addWidget(self._btn_none)
        btn_row.addStretch(1)
        lay.addLayout(btn_row)
        lay.addWidget(self._tree, stretch=1)

        grp_canvas = QGroupBox("Рабочее поле (canvas)", self)
        lay_c = QVBoxLayout(grp_canvas)
        hint_c = QLabel(
            "Пошаговое отключение того, что рисуется поверх PDF. "
            "Дублирует/дополняет чекбоксы в режиме Wireframe для каркасов. "
            "Двойной клик по колонкам «Тип», «Слой», «Примечание» — копирование текста."
        )
        hint_c.setWordWrap(True)
        lay_c.addWidget(hint_c)
        btn_c = QHBoxLayout()
        self._btn_canvas_all = QPushButton("Все", self)
        self._btn_canvas_none = QPushButton("Ни одного", self)
        btn_c.addWidget(self._btn_canvas_all)
        btn_c.addWidget(self._btn_canvas_none)
        btn_c.addStretch(1)
        lay_c.addLayout(btn_c)

        self._tree_canvas = QTreeWidget(self)
        self._tree_canvas.setHeaderLabels(["Вкл.", "Тип", "Слой", "Примечание"])
        h2 = self._tree_canvas.header()
        h2.setStretchLastSection(False)
        h2.setMinimumSectionSize(32)
        h2.setSectionResizeMode(0, QHeaderView.ResizeMode.Interactive)
        h2.setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        h2.setSectionResizeMode(2, QHeaderView.ResizeMode.Interactive)
        h2.setSectionResizeMode(3, QHeaderView.ResizeMode.Interactive)
        self._tree_canvas.setColumnWidth(0, 246)
        self._tree_canvas.setColumnWidth(2, 260)
        self._tree_canvas.setColumnWidth(3, 200)
        self._tree_canvas.setRootIsDecorated(False)
        self._tree_canvas.setAlternatingRowColors(False)
        self._tree_canvas.setUniformRowHeights(True)
        self._tree_canvas.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self._tree_canvas.setStyleSheet(_TREE_STYLE_SHEET)
        self._tree_canvas.itemChanged.connect(self._on_canvas_item_changed)
        self._tree_canvas.itemDoubleClicked.connect(self._on_canvas_tree_double_clicked)
        lay_c.addWidget(self._tree_canvas, stretch=1)

        outer = QVBoxLayout(self)
        outer.setContentsMargins(2, 2, 2, 2)
        outer.setSpacing(4)
        outer.addLayout(bar)
        outer.addWidget(grp, stretch=2)
        outer.addWidget(grp_canvas, stretch=1)

        self._btn_all.clicked.connect(self._select_all)
        self._btn_none.clicked.connect(self._select_none)
        self._btn_canvas_all.clicked.connect(self._select_all_canvas)
        self._btn_canvas_none.clicked.connect(self._select_none_canvas)

        self._root = QTreeWidgetItem(["", "", "Слои (порядок пайплайна)", ""])
        self._root.setFlags(
            Qt.ItemFlag.ItemIsEnabled
            | Qt.ItemFlag.ItemIsSelectable
            | Qt.ItemFlag.ItemIsEditable
        )
        self._tree.addTopLevelItem(self._root)
        self._items_by_layer: dict[str, QTreeWidgetItem] = {}
        self._populate_tree()
        self._load_visibility_from_settings()

        self._canvas_items_by_layer: dict[str, QTreeWidgetItem] = {}
        self._populate_canvas_tree()
        self._load_canvas_visibility_from_settings()

        self._tree.resizeColumnToContents(1)
        self._tree.resizeColumnToContents(3)
        self._tree_canvas.resizeColumnToContents(1)

    def _populate_tree(self) -> None:
        for spec in sorted(PREALIGN_LAYER_SPECS, key=lambda s: s.pipeline_order):
            parent = self._resolve_parent_item(spec)
            it = QTreeWidgetItem(parent)
            it.setFlags(
                Qt.ItemFlag.ItemIsEnabled
                | Qt.ItemFlag.ItemIsSelectable
                | Qt.ItemFlag.ItemIsUserCheckable
                | Qt.ItemFlag.ItemIsEditable,
            )
            it.setCheckState(0, Qt.CheckState.Checked)
            it.setText(1, spec.geometry)
            it.setText(2, spec.label_full)
            it.setText(3, spec.png_group)
            it.setData(0, Qt.ItemDataRole.UserRole, spec.layer_id)
            tip = f"layer_id={spec.layer_id}\nзависит от: {', '.join(spec.depends_on) or '—'}"
            it.setToolTip(2, tip)
            self._items_by_layer[spec.layer_id] = it

    def _resolve_parent_item(self, spec) -> QTreeWidgetItem:
        for dep in spec.depends_on:
            if dep in self._items_by_layer:
                return self._items_by_layer[dep]
        return self._root

    def _load_visibility_from_settings(self) -> None:
        self._tree.blockSignals(True)
        for spec in PREALIGN_LAYER_SPECS:
            it = self._items_by_layer.get(spec.layer_id)
            if it is None:
                continue
            key = f"{_SETTINGS_VISIBILITY_PREFIX}/{spec.layer_id}"
            default_on = spec.layer_id != "extent_probe_merged_grid"
            vis = self._settings.value(key, default_on)
            checked = vis in (True, "true", 1, "1")
            it.setCheckState(0, Qt.CheckState.Checked if checked else Qt.CheckState.Unchecked)
        self._tree.expandAll()
        self._tree.blockSignals(False)

    def _save_layer(self, layer_id: str, visible: bool) -> None:
        self._settings.setValue(f"{_SETTINGS_VISIBILITY_PREFIX}/{layer_id}", visible)

    def _on_prealign_tree_double_clicked(self, item: QTreeWidgetItem, column: int) -> None:
        """Двойной клик по колонкам 1–3 — режим редактирования для копирования текста (данные не сохраняются)."""
        if column == 0:
            return
        self._tree.editItem(item, column)

    def _on_canvas_tree_double_clicked(self, item: QTreeWidgetItem, column: int) -> None:
        if column == 0:
            return
        self._tree_canvas.editItem(item, column)

    @staticmethod
    def _canonical_prealign_text(item: QTreeWidgetItem, column: int) -> str | None:
        lid = item.data(0, Qt.ItemDataRole.UserRole)
        if not lid:
            if column == 2:
                return "Слои (порядок пайплайна)"
            if column in (1, 3):
                return ""
            return None
        spec = LAYER_SPEC_BY_ID.get(str(lid))
        if spec is None:
            return None
        if column == 1:
            return str(spec.geometry)
        if column == 2:
            return spec.label_full
        if column == 3:
            return spec.png_group
        return None

    @staticmethod
    def _canonical_canvas_text(item: QTreeWidgetItem, column: int) -> str | None:
        lid = item.data(0, Qt.ItemDataRole.UserRole)
        if not lid:
            return None
        row = CANVAS_ROWS_BY_ID.get(str(lid))
        if row is None:
            return None
        _lid, geom, label, note = row
        if column == 1:
            return geom
        if column == 2:
            return label
        if column == 3:
            return note
        return None

    def _restore_prealign_cell_if_needed(self, item: QTreeWidgetItem, column: int) -> None:
        canon = self._canonical_prealign_text(item, column)
        if canon is None:
            return
        if item.text(column) == canon:
            return
        self._tree.blockSignals(True)
        item.setText(column, canon)
        self._tree.blockSignals(False)

    def _restore_canvas_cell_if_needed(self, item: QTreeWidgetItem, column: int) -> None:
        canon = self._canonical_canvas_text(item, column)
        if canon is None:
            return
        if item.text(column) == canon:
            return
        self._tree_canvas.blockSignals(True)
        item.setText(column, canon)
        self._tree_canvas.blockSignals(False)

    def _on_item_changed(self, item: QTreeWidgetItem, column: int) -> None:
        if column in (1, 2, 3):
            self._restore_prealign_cell_if_needed(item, column)
            return
        if column != 0:
            return
        lid = item.data(0, Qt.ItemDataRole.UserRole)
        if not lid:
            return
        vis = item.checkState(0) == Qt.CheckState.Checked
        self._save_layer(str(lid), vis)
        self.visibility_changed.emit(str(lid), vis)
        self.overlay_refresh_requested.emit()

    def is_layer_visible(self, layer_id: str) -> bool:
        it = self._items_by_layer.get(layer_id)
        if it is None:
            return True
        return it.checkState(0) == Qt.CheckState.Checked

    def visibility_map(self) -> dict[str, bool]:
        return {lid: self.is_layer_visible(lid) for lid in self._items_by_layer}

    def _select_all(self) -> None:
        self._tree.blockSignals(True)
        for spec in PREALIGN_LAYER_SPECS:
            it = self._items_by_layer.get(spec.layer_id)
            if it:
                it.setCheckState(0, Qt.CheckState.Checked)
                self._save_layer(spec.layer_id, True)
        self._tree.blockSignals(False)
        self.overlay_refresh_requested.emit()

    def _select_none(self) -> None:
        self._tree.blockSignals(True)
        for spec in PREALIGN_LAYER_SPECS:
            it = self._items_by_layer.get(spec.layer_id)
            if it:
                it.setCheckState(0, Qt.CheckState.Unchecked)
                self._save_layer(spec.layer_id, False)
        self._tree.blockSignals(False)
        self.overlay_refresh_requested.emit()

    def _populate_canvas_tree(self) -> None:
        for lid, geom, label, note in CANVAS_LAYER_ROWS:
            it = QTreeWidgetItem()
            it.setFlags(
                Qt.ItemFlag.ItemIsEnabled
                | Qt.ItemFlag.ItemIsSelectable
                | Qt.ItemFlag.ItemIsUserCheckable
                | Qt.ItemFlag.ItemIsEditable,
            )
            it.setCheckState(0, Qt.CheckState.Checked)
            it.setText(1, geom)
            it.setText(2, label)
            it.setText(3, note)
            it.setData(0, Qt.ItemDataRole.UserRole, lid)
            it.setToolTip(2, f"layer_id={lid}")
            self._tree_canvas.addTopLevelItem(it)
            self._canvas_items_by_layer[lid] = it

    def _load_canvas_visibility_from_settings(self) -> None:
        self._tree_canvas.blockSignals(True)
        for lid, *_rest in CANVAS_LAYER_ROWS:
            it = self._canvas_items_by_layer.get(lid)
            if it is None:
                continue
            key = f"{_SETTINGS_CANVAS_PREFIX}/{lid}"
            vis = self._settings.value(key, True)
            checked = vis in (True, "true", 1, "1")
            it.setCheckState(0, Qt.CheckState.Checked if checked else Qt.CheckState.Unchecked)
        self._tree_canvas.blockSignals(False)

    def _save_canvas_layer(self, layer_id: str, visible: bool) -> None:
        self._settings.setValue(f"{_SETTINGS_CANVAS_PREFIX}/{layer_id}", visible)

    def _on_canvas_item_changed(self, item: QTreeWidgetItem, column: int) -> None:
        if column in (1, 2, 3):
            self._restore_canvas_cell_if_needed(item, column)
            return
        if column != 0:
            return
        lid = item.data(0, Qt.ItemDataRole.UserRole)
        if not lid:
            return
        lid_s = str(lid)
        vis = item.checkState(0) == Qt.CheckState.Checked
        self._save_canvas_layer(lid_s, vis)
        self.canvas_layer_visibility_changed.emit(lid_s, vis)

    def is_canvas_layer_visible(self, layer_id: str) -> bool:
        it = self._canvas_items_by_layer.get(layer_id)
        if it is None:
            return True
        return it.checkState(0) == Qt.CheckState.Checked

    def sync_canvas_checkbox(self, layer_id: str, visible: bool) -> None:
        """Синхронизация из Wireframe-панели без повторной эмиссии сигнала."""
        it = self._canvas_items_by_layer.get(layer_id)
        if it is None:
            return
        self._tree_canvas.blockSignals(True)
        it.setCheckState(0, Qt.CheckState.Checked if visible else Qt.CheckState.Unchecked)
        self._save_canvas_layer(layer_id, visible)
        self._tree_canvas.blockSignals(False)

    def _select_all_canvas(self) -> None:
        self._tree_canvas.blockSignals(True)
        for lid, *_rest in CANVAS_LAYER_ROWS:
            it = self._canvas_items_by_layer.get(lid)
            if it:
                it.setCheckState(0, Qt.CheckState.Checked)
                self._save_canvas_layer(lid, True)
        self._tree_canvas.blockSignals(False)
        for lid, *_rest in CANVAS_LAYER_ROWS:
            self.canvas_layer_visibility_changed.emit(lid, True)

    def _select_none_canvas(self) -> None:
        self._tree_canvas.blockSignals(True)
        for lid, *_rest in CANVAS_LAYER_ROWS:
            it = self._canvas_items_by_layer.get(lid)
            if it:
                it.setCheckState(0, Qt.CheckState.Unchecked)
                self._save_canvas_layer(lid, False)
        self._tree_canvas.blockSignals(False)
        for lid, *_rest in CANVAS_LAYER_ROWS:
            self.canvas_layer_visibility_changed.emit(lid, False)
