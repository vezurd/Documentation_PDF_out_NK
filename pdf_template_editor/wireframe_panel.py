"""Wireframe Mode panels: GridLinesPanel + FieldBindingsPanel.

GridLinesPanel — table of all TemplateGridLine objects with Generate/Auto-bind
buttons, line editing (add / delete / edit properties), and start_line/end_line
intersection reference selectors with "pick from canvas" buttons.

FieldBindingsPanel — 4 combo-boxes for the currently selected field's
bound_top / bound_bottom / bound_left / bound_right line references.
"""

from __future__ import annotations

from dataclasses import replace
from typing import TYPE_CHECKING, Sequence

from PySide6.QtCore import Qt, Signal
from PySide6.QtGui import QColor, QBrush, QFont
from PySide6.QtWidgets import (
    QAbstractItemView,
    QCheckBox,
    QComboBox,
    QFormLayout,
    QGroupBox,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QSplitter,
    QTableWidget,
    QTableWidgetItem,
    QToolButton,
    QVBoxLayout,
    QWidget,
)

if TYPE_CHECKING:
    from pdf_parsing_v2_engine.models import FieldDef, TemplateGridLine

_COL_IDX = 0
_COL_ID = 1
_COL_ORIENT = 2
_COL_POS = 3
_COL_START = 4
_COL_END = 5
_COL_BOUNDARY = 6
_NUM_COLS = 7

_HEADER_LABELS = ["#", "ID", "Orient.", "Pos mm", "Start", "End", "Boundary"]


class _NumSortItem(QTableWidgetItem):
    def __lt__(self, other: QTableWidgetItem) -> bool:
        try:
            return float(self.text()) < float(other.text())
        except ValueError:
            return self.text() < other.text()


# ---------------------------------------------------------------------------
# GridLinesPanel
# ---------------------------------------------------------------------------
class GridLinesPanel(QWidget):
    """Table of TemplateGridLine objects + generate / auto-bind / add / delete.

    Also contains start_line/end_line intersection reference editors with
    "pick from canvas" buttons.
    """

    lines_changed = Signal()
    line_selected = Signal(str)  # grid line id
    pick_line_requested = Signal(str)  # "start" or "end" — enters pick mode on canvas
    hide_fields_toggled = Signal(bool)  # True = hide field rectangles on canvas
    show_template_grid_toggled = Signal(bool)
    show_detected_grid_toggled = Signal(bool)
    interactive_add_line_requested = Signal(str)  # "h" or "v" — place line on canvas (two clicks)

    def __init__(self, parent=None):
        super().__init__(parent)
        self._lines: list[TemplateGridLine] = []
        self._block_signals = False
        self._build_ui()

    def _build_ui(self) -> None:
        root = QVBoxLayout(self)
        root.setContentsMargins(2, 2, 2, 2)
        root.setSpacing(4)

        top_row = QHBoxLayout()
        top_row.setSpacing(6)
        lbl = QLabel("Линии сетки (wireframe)")
        lbl.setStyleSheet("font-weight: bold;")
        top_row.addWidget(lbl)
        top_row.addStretch()
        self._chk_hide_fields = QCheckBox("Скрыть поля")
        self._chk_hide_fields.setToolTip("Скрыть прямоугольники полей, оставив только линии сетки")
        self._chk_hide_fields.toggled.connect(self.hide_fields_toggled.emit)
        top_row.addWidget(self._chk_hide_fields)
        root.addLayout(top_row)

        vis_row = QHBoxLayout()
        vis_row.setSpacing(6)
        self._chk_show_template_grid = QCheckBox("Каркас шаблона")
        self._chk_show_template_grid.setChecked(True)
        self._chk_show_template_grid.setToolTip("Показать/скрыть линии сетки шаблона на canvas")
        self._chk_show_template_grid.toggled.connect(self.show_template_grid_toggled.emit)
        vis_row.addWidget(self._chk_show_template_grid)
        self._chk_show_detected_grid = QCheckBox("Каркас PDF")
        self._chk_show_detected_grid.setChecked(True)
        self._chk_show_detected_grid.setEnabled(False)
        self._chk_show_detected_grid.setToolTip(
            "Показать/скрыть detected grid lines из find_tables (после удаления дублей)"
        )
        self._chk_show_detected_grid.toggled.connect(self.show_detected_grid_toggled.emit)
        vis_row.addWidget(self._chk_show_detected_grid)
        vis_row.addStretch()
        root.addLayout(vis_row)

        btn_row = QHBoxLayout()
        btn_row.setSpacing(4)
        self._btn_generate = QPushButton("Генерировать")
        self._btn_generate.setToolTip(
            "Автоматически создать линии из рёбер полей текущего шаблона"
        )
        self._btn_autobind = QPushButton("Привязать все поля")
        self._btn_autobind.setToolTip(
            "Привязать каждое поле к ближайшим линиям и подогнать размеры (bound_top/bottom/left/right)"
        )
        btn_row.addWidget(self._btn_generate)
        btn_row.addWidget(self._btn_autobind)
        root.addLayout(btn_row)

        btn_row2 = QHBoxLayout()
        btn_row2.setSpacing(4)
        self._btn_add_h = QPushButton("+ H-линия")
        self._btn_add_v = QPushButton("+ V-линия")
        self._btn_delete = QPushButton("Удалить")
        self._btn_delete.setEnabled(False)
        btn_row2.addWidget(self._btn_add_h)
        btn_row2.addWidget(self._btn_add_v)
        btn_row2.addWidget(self._btn_delete)
        root.addLayout(btn_row2)

        self._chk_disable_auto_ref_snap = QCheckBox("Отключить автопривязку пересечений")
        self._chk_disable_auto_ref_snap.setChecked(True)
        self._chk_disable_auto_ref_snap.setToolTip(
            "Если включено: при перетаскивании линий и правках не вызываются "
            "snap пересечений (start/end к перпендикулярным линиям). "
            "Удобно при создании каркаса; отключите чекбокс, чтобы снова подтягивать упирания."
        )
        root.addWidget(self._chk_disable_auto_ref_snap)

        self._table = QTableWidget(0, _NUM_COLS)
        self._table.setHorizontalHeaderLabels(_HEADER_LABELS)
        self._table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self._table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self._table.verticalHeader().setVisible(False)
        hh = self._table.horizontalHeader()
        hh.setStretchLastSection(True)
        hh.setSectionResizeMode(_COL_IDX, QHeaderView.ResizeMode.ResizeToContents)
        hh.setSectionResizeMode(_COL_ORIENT, QHeaderView.ResizeMode.ResizeToContents)
        self._table.setAlternatingRowColors(True)
        root.addWidget(self._table, stretch=1)

        self._lbl_count = QLabel("Линий: 0")
        root.addWidget(self._lbl_count)

        # --- Line intersection references (start_line / end_line) ---
        self._refs_group = QGroupBox("Пересечения (упирания)")
        refs_layout = QVBoxLayout(self._refs_group)
        refs_layout.setContentsMargins(4, 4, 4, 4)
        refs_layout.setSpacing(3)

        self._refs_lbl = QLabel("(выберите линию)")
        refs_layout.addWidget(self._refs_lbl)

        start_row = QHBoxLayout()
        start_row.setSpacing(3)
        start_row.addWidget(QLabel("Начало:"))
        self._combo_start_line = QComboBox()
        self._combo_start_line.setSizePolicy(
            self._combo_start_line.sizePolicy().horizontalPolicy(),
            self._combo_start_line.sizePolicy().verticalPolicy(),
        )
        start_row.addWidget(self._combo_start_line, stretch=1)
        self._btn_pick_start = QToolButton()
        self._btn_pick_start.setText("⊕")
        self._btn_pick_start.setToolTip("Выбрать линию на canvas (кликните по линии)")
        self._btn_pick_start.setFixedSize(24, 24)
        start_row.addWidget(self._btn_pick_start)
        refs_layout.addLayout(start_row)

        end_row = QHBoxLayout()
        end_row.setSpacing(3)
        end_row.addWidget(QLabel("Конец:"))
        self._combo_end_line = QComboBox()
        end_row.addWidget(self._combo_end_line, stretch=1)
        self._btn_pick_end = QToolButton()
        self._btn_pick_end.setText("⊕")
        self._btn_pick_end.setToolTip("Выбрать линию на canvas (кликните по линии)")
        self._btn_pick_end.setFixedSize(24, 24)
        end_row.addWidget(self._btn_pick_end)
        refs_layout.addLayout(end_row)

        self._btn_auto_refs = QPushButton("Авто-привязка пересечений")
        self._btn_auto_refs.setToolTip(
            "Автоматически найти пересекающиеся перпендикулярные линии для начала и конца каждой линии"
        )
        refs_layout.addWidget(self._btn_auto_refs)

        root.addWidget(self._refs_group)
        self._refs_group.setEnabled(False)

        # --- Detected PDF lines section ---
        det_group = QGroupBox("Линии PDF (detected)")
        det_layout = QVBoxLayout(det_group)
        det_layout.setContentsMargins(4, 4, 4, 4)
        det_layout.setSpacing(2)

        _DET_COLS = 5
        self._det_table = QTableWidget(0, _DET_COLS)
        self._det_table.setHorizontalHeaderLabels(
            ["ID", "Orient.", "Pos pts", "Span", "Cells"]
        )
        self._det_table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self._det_table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self._det_table.verticalHeader().setVisible(False)
        self._det_table.setAlternatingRowColors(True)
        self._det_table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        dh = self._det_table.horizontalHeader()
        dh.setStretchLastSection(True)
        dh.setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        det_layout.addWidget(self._det_table)

        self._lbl_det_count = QLabel("PDF линий: 0")
        det_layout.addWidget(self._lbl_det_count)
        root.addWidget(det_group)

        # --- signals ---
        self._table.itemSelectionChanged.connect(self._on_selection_changed)
        self._table.cellChanged.connect(self._on_cell_changed)
        self._btn_add_h.clicked.connect(lambda: self.interactive_add_line_requested.emit("h"))
        self._btn_add_v.clicked.connect(lambda: self.interactive_add_line_requested.emit("v"))
        self._btn_delete.clicked.connect(self._delete_selected)
        self._combo_start_line.currentIndexChanged.connect(self._on_start_line_combo_changed)
        self._combo_end_line.currentIndexChanged.connect(self._on_end_line_combo_changed)
        self._btn_pick_start.clicked.connect(lambda: self.pick_line_requested.emit("start"))
        self._btn_pick_end.clicked.connect(lambda: self.pick_line_requested.emit("end"))
        self._det_table.itemSelectionChanged.connect(self._on_det_selection_changed)

    # --- public API ---

    def line_intersection_snap_enabled(self) -> bool:
        """If True, apply snap_endpoints_to_refs after line geometry edits (default UI: off)."""
        return not self._chk_disable_auto_ref_snap.isChecked()

    def apply_external_visibility_checks(
        self,
        *,
        stamp_fields_visible: bool,
        template_grid_visible: bool,
        pdf_grid_visible: bool,
    ) -> None:
        """Синхронизация чекбоксов из панели «Границы» без сигналов."""
        for chk in (
            self._chk_hide_fields,
            self._chk_show_template_grid,
            self._chk_show_detected_grid,
        ):
            chk.blockSignals(True)
        try:
            self._chk_hide_fields.setChecked(not stamp_fields_visible)
            self._chk_show_template_grid.setChecked(template_grid_visible)
            self._chk_show_detected_grid.setChecked(pdf_grid_visible)
        finally:
            for chk in (
                self._chk_hide_fields,
                self._chk_show_template_grid,
                self._chk_show_detected_grid,
            ):
                chk.blockSignals(False)

    def set_detected_lines(self, det_lines) -> None:
        """Populate the detected PDF lines table from a list of DetectedGridLineInfo."""
        self._det_table.setRowCount(0)
        if not det_lines:
            self._lbl_det_count.setText("PDF линий: 0")
            return
        self._det_table.setRowCount(len(det_lines))
        for r, dgl in enumerate(det_lines):
            self._det_table.setItem(r, 0, QTableWidgetItem(dgl.id))
            self._det_table.setItem(r, 1, QTableWidgetItem(dgl.orientation.upper()))
            self._det_table.setItem(r, 2, _NumSortItem(f"{dgl.pos_pts:.1f}"))
            span = dgl.span_hi_pts - dgl.span_lo_pts
            self._det_table.setItem(r, 3, _NumSortItem(f"{span:.0f}"))
            self._det_table.setItem(r, 4, QTableWidgetItem(str(dgl.cell_count)))
        self._lbl_det_count.setText(f"PDF линий: {len(det_lines)}")

    def _on_det_selection_changed(self) -> None:
        rows = self._det_table.selectionModel().selectedRows()
        if not rows:
            return
        r = rows[0].row()
        det_id_item = self._det_table.item(r, 0)
        if det_id_item:
            self.line_selected.emit(det_id_item.text())

    def set_lines(self, lines: Sequence[TemplateGridLine]) -> None:
        prev_sel_id = self.selected_line_id()
        self._lines = list(lines)
        self._refresh_table()
        if prev_sel_id:
            self.select_line_by_id(prev_sel_id)
        self._refresh_refs_panel()

    def get_lines(self) -> list[TemplateGridLine]:
        return list(self._lines)

    def selected_line_id(self) -> str | None:
        rows = self._table.selectionModel().selectedRows()
        if not rows:
            return None
        r = rows[0].row()
        if 0 <= r < len(self._lines):
            return self._lines[r].id
        return None

    def selected_line_index(self) -> int:
        rows = self._table.selectionModel().selectedRows()
        if not rows:
            return -1
        return rows[0].row()

    def select_line_by_id(self, line_id: str) -> None:
        """Select a row in the table matching the given line ID."""
        for i, gl in enumerate(self._lines):
            if gl.id == line_id:
                self._table.selectRow(i)
                self._table.scrollToItem(
                    self._table.item(i, 0),
                    QAbstractItemView.ScrollHint.EnsureVisible,
                )
                return

    def set_picked_line(self, which: str, line_id: str) -> None:
        """Called by main_window after user picks a line on canvas.
        which = "start" or "end".
        """
        combo = self._combo_start_line if which == "start" else self._combo_end_line
        for i in range(combo.count()):
            if combo.itemData(i) == line_id:
                combo.setCurrentIndex(i)
                return

    # --- internal ---

    def _refresh_table(self) -> None:
        self._block_signals = True
        self._table.setRowCount(0)
        self._table.setRowCount(len(self._lines))
        for i, gl in enumerate(self._lines):
            idx_item = _NumSortItem(str(i + 1))
            idx_item.setFlags(idx_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
            self._table.setItem(i, _COL_IDX, idx_item)

            self._table.setItem(i, _COL_ID, QTableWidgetItem(gl.id))
            orient_item = QTableWidgetItem(gl.orientation)
            orient_item.setFlags(orient_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
            self._table.setItem(i, _COL_ORIENT, orient_item)

            self._table.setItem(i, _COL_POS, _NumSortItem(f"{gl.pos_mm:.2f}"))
            self._table.setItem(i, _COL_START, _NumSortItem(f"{gl.start_mm:.2f}"))
            self._table.setItem(i, _COL_END, _NumSortItem(f"{gl.end_mm:.2f}"))

            bnd_item = QTableWidgetItem(gl.boundary or "")
            self._table.setItem(i, _COL_BOUNDARY, bnd_item)

            if gl.boundary:
                for c in range(_NUM_COLS):
                    it = self._table.item(i, c)
                    if it:
                        f = it.font()
                        f.setBold(True)
                        it.setFont(f)

        self._block_signals = False
        self._lbl_count.setText(f"Линий: {len(self._lines)}")

    def _on_selection_changed(self) -> None:
        sel = self._table.selectionModel().selectedRows()
        self._btn_delete.setEnabled(bool(sel))
        self._refresh_refs_panel()
        lid = self.selected_line_id()
        if lid is not None:
            self.line_selected.emit(lid)

    def _refresh_refs_panel(self) -> None:
        """Update the intersection refs section for the currently selected line."""
        idx = self.selected_line_index()
        if idx < 0 or idx >= len(self._lines):
            self._refs_group.setEnabled(False)
            self._refs_lbl.setText("(выберите линию)")
            return

        self._refs_group.setEnabled(True)
        gl = self._lines[idx]
        self._refs_lbl.setText(f"Линия: {gl.id} ({gl.orientation})")

        self._block_signals = True
        perp_orient = "v" if gl.orientation == "h" else "h"
        perp_lines = [("(нет)", "")] + [
            (f"{p.id}  ({p.pos_mm:.1f}mm)", p.id)
            for p in self._lines if p.orientation == perp_orient
        ]

        for combo in (self._combo_start_line, self._combo_end_line):
            combo.clear()
            for label, data in perp_lines:
                combo.addItem(label, data)

        for combo, ref_id in [
            (self._combo_start_line, gl.start_line_id),
            (self._combo_end_line, gl.end_line_id),
        ]:
            selected_idx = 0
            if ref_id:
                for i in range(combo.count()):
                    if combo.itemData(i) == ref_id:
                        selected_idx = i
                        break
            combo.setCurrentIndex(selected_idx)

        self._block_signals = False

    def _on_start_line_combo_changed(self, _idx: int) -> None:
        if self._block_signals:
            return
        self._apply_ref_combo("start_line_id", self._combo_start_line)

    def _on_end_line_combo_changed(self, _idx: int) -> None:
        if self._block_signals:
            return
        self._apply_ref_combo("end_line_id", self._combo_end_line)

    def _apply_ref_combo(self, attr: str, combo: QComboBox) -> None:
        idx = self.selected_line_index()
        if idx < 0 or idx >= len(self._lines):
            return
        val = combo.currentData() or None
        self._lines[idx] = replace(self._lines[idx], **{attr: val})
        self.lines_changed.emit()

    def _on_cell_changed(self, row: int, col: int) -> None:
        if self._block_signals:
            return
        if row < 0 or row >= len(self._lines):
            return
        gl = self._lines[row]
        try:
            if col == _COL_ID:
                new_id = self._table.item(row, col).text().strip()
                if new_id and new_id != gl.id:
                    self._lines[row] = replace(gl, id=new_id)
            elif col == _COL_POS:
                self._lines[row] = replace(gl, pos_mm=round(float(self._table.item(row, col).text()), 3))
            elif col == _COL_START:
                self._lines[row] = replace(gl, start_mm=round(float(self._table.item(row, col).text()), 3))
            elif col == _COL_END:
                self._lines[row] = replace(gl, end_mm=round(float(self._table.item(row, col).text()), 3))
            elif col == _COL_BOUNDARY:
                raw = self._table.item(row, col).text().strip().lower()
                bnd = raw if raw in ("top", "bottom", "left", "right") else None
                self._lines[row] = replace(gl, boundary=bnd)
            else:
                return
        except (ValueError, TypeError):
            self._refresh_table()
            return
        self.lines_changed.emit()

    def _next_new_line_id(self, orientation: str) -> str:
        used_ids = {gl.id for gl in self._lines}
        prefix = "h" if orientation == "h" else "v"
        n = 1
        while f"{prefix}_new_{n}" in used_ids:
            n += 1
        return f"{prefix}_new_{n}"

    def append_line_mm(
        self,
        orientation: str,
        pos_mm: float,
        start_mm: float,
        end_mm: float,
    ) -> None:
        """Append a line with known template mm (used after two-click placement on canvas)."""
        from pdf_parsing_v2_engine.models import TemplateGridLine

        new_id = self._next_new_line_id(orientation)
        lo, hi = (start_mm, end_mm) if start_mm <= end_mm else (end_mm, start_mm)
        self._lines.append(
            TemplateGridLine(
                id=new_id,
                orientation=orientation,
                pos_mm=float(pos_mm),
                start_mm=float(lo),
                end_mm=float(hi),
            ),
        )
        self._refresh_table()
        self._table.selectRow(len(self._lines) - 1)
        self.lines_changed.emit()

    def delete_selected_line(self) -> None:
        """Remove the selected row (same as the Удалить button)."""
        self._delete_selected()

    def _delete_selected(self) -> None:
        rows = self._table.selectionModel().selectedRows()
        if not rows:
            return
        r = rows[0].row()
        if 0 <= r < len(self._lines):
            del self._lines[r]
            self._refresh_table()
            self.lines_changed.emit()


# ---------------------------------------------------------------------------
# FieldBindingsPanel
# ---------------------------------------------------------------------------
class FieldBindingsPanel(QWidget):
    """Per-field binding editor: 4 combo-boxes mapping to grid line IDs."""

    bindings_changed = Signal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self._field: FieldDef | None = None
        self._lines: list[TemplateGridLine] = []
        self._block_signals = False
        self._build_ui()

    def _build_ui(self) -> None:
        root = QVBoxLayout(self)
        root.setContentsMargins(2, 2, 2, 2)
        root.setSpacing(4)

        lbl = QLabel("Привязки поля")
        lbl.setStyleSheet("font-weight: bold;")
        root.addWidget(lbl)

        self._lbl_outside_hint = QLabel("")
        self._lbl_outside_hint.setWordWrap(True)
        self._lbl_outside_hint.setStyleSheet("color: #666; font-size: 11px;")
        self._lbl_outside_hint.hide()
        root.addWidget(self._lbl_outside_hint)

        self._lbl_field = QLabel("(нет выделенного поля)")
        root.addWidget(self._lbl_field)

        form = QFormLayout()
        form.setContentsMargins(0, 0, 0, 0)
        self._combo_top = QComboBox()
        self._combo_bottom = QComboBox()
        self._combo_left = QComboBox()
        self._combo_right = QComboBox()
        form.addRow("Top (H):", self._combo_top)
        form.addRow("Bottom (H):", self._combo_bottom)
        form.addRow("Left (V):", self._combo_left)
        form.addRow("Right (V):", self._combo_right)
        root.addLayout(form)

        self._btn_autobind_one = QPushButton("Привязать это поле")
        self._btn_autobind_one.setToolTip("Автоматически найти ближайшие линии для текущего поля")
        self._btn_autobind_one.setEnabled(False)
        root.addWidget(self._btn_autobind_one)

        self._status_label = QLabel("")
        root.addWidget(self._status_label)
        root.addStretch()

        for combo in (self._combo_top, self._combo_bottom, self._combo_left, self._combo_right):
            combo.currentIndexChanged.connect(self._on_combo_changed)
            combo.setEnabled(False)

        self._btn_autobind_one.setToolTip(
            self._btn_autobind_one.toolTip() + " Не применяется к полям вне штампа.",
        )

    # --- public API ---

    def set_lines(self, lines: Sequence[TemplateGridLine]) -> None:
        self._lines = list(lines)
        self._repopulate_combos()
        self._sync_combos_to_field()

    def set_field(self, fd: FieldDef | None) -> None:
        self._field = fd
        if fd is None:
            self._lbl_outside_hint.hide()
            self._lbl_field.setText("(нет выделенного поля)")
            for c in (self._combo_top, self._combo_bottom, self._combo_left, self._combo_right):
                c.setEnabled(False)
            self._btn_autobind_one.setEnabled(False)
            self._status_label.setText("")
        elif getattr(fd, "outside_stamp", False):
            self._lbl_outside_hint.setText(
                "Вне штампа: вертикальная привязка автоматическая (origin: верхний/нижний колонтитул). "
                "Горизонталь: через bbox_mm и stretch left/right. Линии каркаса не участвуют.",
            )
            self._lbl_outside_hint.show()
            self._lbl_field.setText(f"Поле: {fd.id}")
            for c in (self._combo_top, self._combo_bottom, self._combo_left, self._combo_right):
                c.setEnabled(False)
            self._btn_autobind_one.setEnabled(False)
            self._status_label.setText("")
            self._block_signals = True
            for combo in (self._combo_top, self._combo_bottom, self._combo_left, self._combo_right):
                combo.setCurrentIndex(0)
            self._block_signals = False
        else:
            self._lbl_outside_hint.hide()
            self._lbl_field.setText(f"Поле: {fd.id}")
            for c in (self._combo_top, self._combo_bottom, self._combo_left, self._combo_right):
                c.setEnabled(True)
            self._btn_autobind_one.setEnabled(True)
            self._sync_combos_to_field()

    def get_bindings(self) -> tuple[str | None, str | None, str | None, str | None]:
        """Return (bound_top, bound_bottom, bound_left, bound_right)."""
        def _val(combo: QComboBox) -> str | None:
            t = combo.currentData()
            return t if t else None
        return _val(self._combo_top), _val(self._combo_bottom), _val(self._combo_left), _val(self._combo_right)

    # --- internal ---

    def _repopulate_combos(self) -> None:
        self._block_signals = True
        h_ids = [("(нет)", "")] + [
            (f"{gl.id}  ({gl.pos_mm:.1f}mm)", gl.id)
            for gl in self._lines if gl.orientation == "h"
        ]
        v_ids = [("(нет)", "")] + [
            (f"{gl.id}  ({gl.pos_mm:.1f}mm)", gl.id)
            for gl in self._lines if gl.orientation == "v"
        ]
        for combo, items in [
            (self._combo_top, h_ids),
            (self._combo_bottom, h_ids),
            (self._combo_left, v_ids),
            (self._combo_right, v_ids),
        ]:
            combo.clear()
            for label, data in items:
                combo.addItem(label, data)
        self._block_signals = False

    def _sync_combos_to_field(self) -> None:
        if self._field is None:
            return
        self._block_signals = True
        for combo, bound_id in [
            (self._combo_top, self._field.bound_top),
            (self._combo_bottom, self._field.bound_bottom),
            (self._combo_left, self._field.bound_left),
            (self._combo_right, self._field.bound_right),
        ]:
            idx = 0
            if bound_id:
                for i in range(combo.count()):
                    if combo.itemData(i) == bound_id:
                        idx = i
                        break
            combo.setCurrentIndex(idx)
        self._block_signals = False
        self._update_status()

    def _on_combo_changed(self, _idx: int) -> None:
        if self._block_signals:
            return
        self._update_status()
        self.bindings_changed.emit()

    def _update_status(self) -> None:
        if self._field is None:
            self._status_label.setText("")
            return
        bt, bb, bl, br = self.get_bindings()
        count = sum(1 for b in (bt, bb, bl, br) if b)
        if count == 4:
            self._status_label.setText("✓ Все 4 привязки назначены")
            self._status_label.setStyleSheet("color: green;")
        elif count > 0:
            self._status_label.setText(f"⚠ {count}/4 привязок назначено")
            self._status_label.setStyleSheet("color: #b8860b;")
        else:
            self._status_label.setText("✗ Нет привязок")
            self._status_label.setStyleSheet("color: red;")
