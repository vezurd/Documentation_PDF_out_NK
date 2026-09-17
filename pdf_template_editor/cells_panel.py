"""Список ячеек (QTreeWidget) — боковая панель с навигацией Canvas <-> Список."""

from __future__ import annotations

from typing import TYPE_CHECKING

from PySide6.QtCore import Qt, Signal
from PySide6.QtGui import QColor, QBrush
from PySide6.QtWidgets import (
    QAbstractItemView,
    QCheckBox,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QTreeWidget,
    QTreeWidgetItem,
    QVBoxLayout,
    QWidget,
)

from pdf_parsing_v2_engine.models import FieldDef, FieldResult

if TYPE_CHECKING:
    from pdf_parsing_v2_engine.models import FieldCatalog
    from pdf_parsing_v2_engine.grid_matcher import AdaptResult
    from pdf_parsing_v2_engine.models import V2PageResult
    from .cell_items import FieldRectItem
    from .graphics_view import StampGraphicsView


_DETAIL_TOOLTIP_HEADERS = frozenset({"Значение", "Предупр."})
_DETAIL_MULTILINE_HEADERS = frozenset({"Значение", "Предупр."})

# Нейтральный hover/selection (без системного синего), как панель «Границы»
_FIELD_LIST_TREE_QSS = """
QTreeWidget {
    show-decoration-selected: 0;
}
QTreeWidget::item {
    min-height: 20px;
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


def _format_field_warnings_for_ui(fr: FieldResult) -> tuple[str, str]:
    """Short table text and full tooltip for ``FieldResult.parse_warnings``."""
    if not fr.parse_warnings:
        return "", ""
    tier_head = f"clean_tier={fr.clean_tier}\n" if fr.clean_tier else ""
    lines: list[str] = []
    for w in fr.parse_warnings:
        tier = f" [{w.tier}]" if w.tier else ""
        lines.append(f"{w.code}{tier}: {w.message}")
    full = tier_head + "\n".join(lines)
    short = " ; ".join(lines)
    if len(short) > 100:
        short = short[:97] + "…"
    return short, full


class NumericSortTreeWidgetItem(QTreeWidgetItem):
    """Сортировка: столбец «#» по числу (1, 2, … 10), не лексикографически (1, 10, 2)."""

    def __lt__(self, other: QTreeWidgetItem) -> bool:
        tree = self.treeWidget()
        col = tree.sortColumn() if tree is not None else 0
        if col == 0:
            try:
                return int(self.text(0)) < int(other.text(0))
            except ValueError:
                return self.text(0) < other.text(0)
        return self.text(col) < other.text(col)


class CellsPanel(QWidget):
    """Side panel: list of all FieldRectItem objects on the scene."""

    # FieldRectItem (canvas) or FieldDef (hidden document-property row)
    cell_selected = Signal(object)
    cell_double_clicked = Signal(object)

    def __init__(
        self,
        gfx_view: StampGraphicsView,
        catalog: "FieldCatalog | None" = None,
        parent=None,
    ):
        super().__init__(parent)
        self._gfx_view = gfx_view
        self._catalog = catalog
        self._item_map: dict[int, tuple[QTreeWidgetItem, FieldRectItem]] = {}
        # id(FieldDef) -> (tree row, FieldDef) for PDF document-property rows (no canvas item)
        self._docprop_map: dict[int, tuple[QTreeWidgetItem, FieldDef]] = {}
        self._warn_item_ids: set[int] = set()

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)

        hdr = QLabel("Список полей")
        hf = hdr.font()
        hf.setBold(True)
        hdr.setFont(hf)
        layout.addWidget(hdr)

        btn_row = QHBoxLayout()
        self._chk_filter = QCheckBox("Только неназначенные")
        self._chk_filter.stateChanged.connect(self._apply_filter)
        btn_row.addWidget(self._chk_filter)
        btn_row.addStretch()
        layout.addLayout(btn_row)

        self._tree = QTreeWidget()
        self._tree.setHeaderLabels(
            ["#", "ID (кат.)", "Label", "Score", "Значение", "Предупр."]
        )
        self._tree.setColumnWidth(0, 40)
        self._tree.setColumnWidth(1, 150)
        self._tree.setColumnWidth(2, 160)
        self._tree.setColumnWidth(3, 72)
        self._tree.setColumnWidth(4, 130)
        self._tree.setColumnWidth(5, 180)
        from PySide6.QtWidgets import QHeaderView
        header = self._tree.header()
        header.setSectionResizeMode(3, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(5, QHeaderView.ResizeMode.Stretch)
        self._tree.setSelectionMode(QTreeWidget.SelectionMode.ExtendedSelection)
        self._tree.setRootIsDecorated(False)
        self._tree.setAlternatingRowColors(False)
        self._tree.setStyleSheet(_FIELD_LIST_TREE_QSS)
        self._tree.setSortingEnabled(True)
        self._tree.itemClicked.connect(self._on_item_clicked)
        self._tree.itemDoubleClicked.connect(self._on_item_double_clicked)
        layout.addWidget(self._tree)

    def set_catalog(self, catalog: "FieldCatalog | None") -> None:
        self._catalog = catalog
        for _, (tw, cell_item) in self._item_map.items():
            self._set_row_from_item(tw, cell_item)

    def rebuild_list(self, document_property_defs: list[FieldDef] | None = None) -> None:
        """Rebuild the tree from the current scene items.

        Optional *document_property_defs*: hidden PDF property fields (no FieldRectItem),
        shown after canvas fields with a distinct style.
        """
        from .cell_items import FieldRectItem

        self._tree.clear()
        self._item_map.clear()
        self._docprop_map.clear()
        items = sorted(
            (it for it in self._gfx_view.scene().items() if isinstance(it, FieldRectItem)),
            key=lambda it: it.cell_index,
        )
        for cell_item in items:
            tw = NumericSortTreeWidgetItem()
            tw.setText(0, str(cell_item.cell_index))
            tw.setData(0, Qt.ItemDataRole.UserRole, id(cell_item))
            tw.setData(0, Qt.ItemDataRole.UserRole + 1, "scene")
            self._set_row_from_item(tw, cell_item)
            self._tree.addTopLevelItem(tw)
            self._item_map[id(cell_item)] = (tw, cell_item)

        if document_property_defs:
            base = len(items)
            for j, fd in enumerate(document_property_defs):
                tw = NumericSortTreeWidgetItem()
                tw.setText(0, str(base + j + 1))
                tw.setData(0, Qt.ItemDataRole.UserRole, id(fd))
                tw.setData(0, Qt.ItemDataRole.UserRole + 1, "docprop")
                self._set_row_for_document_property(tw, fd)
                self._tree.addTopLevelItem(tw)
                self._docprop_map[id(fd)] = (tw, fd)

        self._apply_filter()

    def update_item_row(self, cell_item: FieldRectItem) -> None:
        key = id(cell_item)
        if key in self._item_map:
            tw, _ = self._item_map[key]
            self._set_row_from_item(tw, cell_item)

    def get_item_details_rows(self, cell_item: FieldRectItem) -> list[tuple[str, str, bool]]:
        """Return the current visible row as labeled fields for the details dialog."""
        entry = self._item_map.get(id(cell_item))
        if not entry:
            return []
        tw, _ = entry
        header = self._tree.headerItem()
        rows: list[tuple[str, str, bool]] = []
        for col in range(self._tree.columnCount()):
            label = (header.text(col) or "").strip() or f"Колонка {col + 1}"
            value = tw.text(col)
            if label in _DETAIL_TOOLTIP_HEADERS:
                value = tw.toolTip(col) or value
            rows.append((label, value, label in _DETAIL_MULTILINE_HEADERS))
        return rows

    def get_field_def_details_rows(self, fd: FieldDef) -> list[tuple[str, str, bool]]:
        """Details dialog rows for a document-property ``FieldDef`` (no canvas item)."""
        entry = self._docprop_map.get(id(fd))
        if not entry:
            return []
        tw, _ = entry
        header = self._tree.headerItem()
        rows: list[tuple[str, str, bool]] = []
        for col in range(self._tree.columnCount()):
            label = (header.text(col) or "").strip() or f"Колонка {col + 1}"
            value = tw.text(col)
            if label in _DETAIL_TOOLTIP_HEADERS:
                value = tw.toolTip(col) or value
            rows.append((label, value, label in _DETAIL_MULTILINE_HEADERS))
        return rows

    def set_warning_items(self, item_ids: set[int]) -> None:
        self._warn_item_ids = set(item_ids)
        for _, (tw, cell_item) in self._item_map.items():
            self._set_row_from_item(tw, cell_item)

    def update_test_values(self, result: V2PageResult) -> None:
        for _, (tw, cell_item) in self._item_map.items():
            tw.setText(4, "")
            tw.setText(5, "")
            tw.setToolTip(4, "")
            tw.setToolTip(5, "")
            if not cell_item.field_def:
                continue
            fid = cell_item.field_def.id
            if fid not in result.fields:
                continue
            fr = result.fields[fid]
            val = (fr.cleaned_value or "")[:80]
            tw.setText(4, val)
            tw.setToolTip(4, fr.cleaned_value or "")
            warn_text, warn_tip = _format_field_warnings_for_ui(fr)
            tw.setText(5, warn_text)
            tw.setToolTip(5, warn_tip)
        for _, (tw, fd) in self._docprop_map.items():
            tw.setText(4, "")
            tw.setText(5, "")
            tw.setToolTip(4, "")
            tw.setToolTip(5, "")
            fid = fd.id
            if fid not in result.fields:
                continue
            fr = result.fields[fid]
            val = (fr.cleaned_value or "")[:80]
            tw.setText(4, val)
            tw.setToolTip(4, fr.cleaned_value or "")
            warn_text, warn_tip = _format_field_warnings_for_ui(fr)
            tw.setText(5, warn_text)
            tw.setToolTip(5, warn_tip)

    def update_adapt_results(self, adapted: dict[int, AdaptResult]) -> None:
        """Update Score and Iter columns from adapt results.

        *adapted* is keyed by ``id(cell_item)`` (python object id).
        """
        for _, (tw, cell_item) in self._item_map.items():
            if not cell_item.field_def:
                tw.setText(3, "")
                continue
            ar = adapted.get(id(cell_item))
            if ar is None:
                tw.setText(3, "")
                continue

            if ar.status == "anchor":
                tw.setText(3, f"{ar.confidence:.2f} A")
            elif ar.status == "derived":
                tw.setText(3, "~ D")
            elif ar.status == "excluded":
                tw.setText(3, "—")
            else:
                tw.setText(3, f"{ar.confidence:.2f}")

            score_color = QColor(100, 180, 100) if ar.confidence > 0.7 else (
                QColor(200, 180, 60) if ar.confidence >= 0.4 else QColor(200, 80, 80)
            )
            if ar.status == "derived":
                score_color = QColor(160, 160, 160)
            tw.setForeground(3, QBrush(score_color))

    def select_field_id(self, field_id: str) -> bool:
        """Select tree row by catalog field id (canvas or [PDF] document-property)."""
        fid = (field_id or "").strip()
        if not fid:
            return False
        for i in range(self._tree.topLevelItemCount()):
            tw = self._tree.topLevelItem(i)
            if (tw.text(1) or "").strip() != fid:
                continue
            self._tree.clearSelection()
            tw.setSelected(True)
            self._tree.setCurrentItem(tw)
            self._tree.scrollToItem(tw, QAbstractItemView.ScrollHint.PositionAtCenter)
            return True
        return False

    def highlight_items(self, items: list[FieldRectItem]) -> None:
        ids = {id(it) for it in items}
        for i in range(self._tree.topLevelItemCount()):
            tw = self._tree.topLevelItem(i)
            if tw.data(0, Qt.ItemDataRole.UserRole + 1) == "docprop":
                tw.setSelected(False)
                continue
            obj_id = tw.data(0, Qt.ItemDataRole.UserRole)
            tw.setSelected(obj_id in ids)

        if not items:
            return
        for it in items:
            key = id(it)
            entry = self._item_map.get(key)
            if not entry:
                continue
            tw, _ = entry
            if tw.isHidden():
                continue
            self._tree.setCurrentItem(tw)
            self._tree.scrollToItem(tw, QAbstractItemView.ScrollHint.PositionAtCenter)
            break

    # ---- internal ----

    def _set_row_from_item(self, tw: QTreeWidgetItem, cell_item: FieldRectItem) -> None:
        if cell_item.is_assigned and cell_item.field_def:
            fd = cell_item.field_def
            tw.setText(1, fd.id)
            tw.setText(2, (fd.label or "").strip() or "—")
            tw.setToolTip(1, "")
            cat_synced = False
            if self._catalog:
                entry = self._catalog.get_entry(fd.id)
                if entry:
                    from pdf_parsing_v2_engine.models import normalize_field_type

                    lbl_fd = (fd.label or "").strip()
                    lbl_cat = (entry.label or "").strip()
                    ft_fd = normalize_field_type(fd.field_type)
                    ft_cat = normalize_field_type(entry.default_field_type or "data")
                    cat_synced = (lbl_fd == lbl_cat) and (ft_fd == ft_cat)
            if cat_synced:
                tw.setToolTip(1, "Поле соответствует строке каталога (id/label/type)")
                tw.setForeground(1, QBrush(QColor(40, 135, 70)))
            else:
                tw.setForeground(1, QBrush(QColor(90, 90, 90)))
            if cell_item.field_def.outside_stamp:
                color = QColor(199, 93, 0)
            else:
                color = QColor(40, 135, 70) if cat_synced else QColor(60, 120, 220)
        else:
            tw.setText(1, "")
            tw.setText(2, "—")
            tw.setToolTip(1, "")
            tw.setForeground(1, QBrush(QColor(120, 120, 120)))
            color = QColor(160, 160, 160)
        if id(cell_item) in self._warn_item_ids:
            tw.setText(2, f"{tw.text(2)} ⚠")
        tw.setForeground(0, QBrush(color))

    def _set_row_for_document_property(self, tw: QTreeWidgetItem, fd: FieldDef) -> None:
        tw.setText(1, fd.id)
        lbl = (fd.label or "").strip() or "—"
        tw.setText(2, f"[PDF] {lbl}")
        tw.setToolTip(1, "Служебное поле свойств PDF (не на холсте)")
        tw.setForeground(1, QBrush(QColor(120, 80, 160)))
        tw.setForeground(0, QBrush(QColor(120, 80, 160)))
        tw.setForeground(2, QBrush(QColor(120, 80, 160)))

    def _apply_filter(self, _state=None) -> None:
        only_unassigned = self._chk_filter.isChecked()
        for i in range(self._tree.topLevelItemCount()):
            tw = self._tree.topLevelItem(i)
            obj_id = tw.data(0, Qt.ItemDataRole.UserRole)
            entry = self._item_map.get(obj_id)
            if entry and only_unassigned and entry[1].is_assigned:
                tw.setHidden(True)
            else:
                tw.setHidden(False)

    def _on_item_clicked(self, tw: QTreeWidgetItem, col: int) -> None:
        obj_id = tw.data(0, Qt.ItemDataRole.UserRole)
        if tw.data(0, Qt.ItemDataRole.UserRole + 1) == "docprop":
            entry = self._docprop_map.get(obj_id)
            if entry:
                self.cell_selected.emit(entry[1])
            return
        entry = self._item_map.get(obj_id)
        if entry:
            self.cell_selected.emit(entry[1])

    def _on_item_double_clicked(self, tw: QTreeWidgetItem, col: int) -> None:
        obj_id = tw.data(0, Qt.ItemDataRole.UserRole)
        if tw.data(0, Qt.ItemDataRole.UserRole + 1) == "docprop":
            entry = self._docprop_map.get(obj_id)
            if entry:
                self.cell_double_clicked.emit(entry[1])
            return
        entry = self._item_map.get(obj_id)
        if entry:
            self.cell_double_clicked.emit(entry[1])
