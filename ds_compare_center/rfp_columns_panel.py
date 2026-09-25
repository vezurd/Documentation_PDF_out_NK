"""WYSIWYG Step4 Excel column templates for the DS/RFP control center."""

from __future__ import annotations

from typing import Any, Callable

from PySide6.QtCore import QMimeData, QPoint, Qt, Signal
from PySide6.QtGui import QDrag, QMouseEvent
from PySide6.QtWidgets import (
    QCheckBox,
    QComboBox,
    QFrame,
    QHBoxLayout,
    QInputDialog,
    QLabel,
    QMessageBox,
    QPushButton,
    QScrollArea,
    QSizePolicy,
    QSpinBox,
    QVBoxLayout,
    QWidget,
)

from RFQ.tags_rfp_compare.step4.step4_6_save_match_result_to_excel import (
    HEADER_FILL_COLORS,
)
from RFQ.tags_rfp_compare.step4.step4_excel_columns import (
    DEFAULT_TEMPLATE_ID,
    MAX_COLUMN_WIDTH,
    MIN_COLUMN_WIDTH,
    builtin_column_settings,
    create_template,
    delete_template,
    load_step4_excel_column_state,
    normalize_template_columns,
    rename_template,
    set_active_template,
    update_template_columns,
)

_MIME_COL_INDEX = "application/x-step4-col-index"
_DEFAULT_HINT = (
    "Шаблон по умолчанию только для просмотра и отката. "
    "Создайте свой, чтобы менять столбцы."
)
_USER_HINT = (
    "Перетащите карточку, чтобы менять порядок. Тяните правый край для ширины. "
    "Двойной щелчок по заголовку — подпись. «Сохранить» записывает шаблон."
)
_GROUP_COLORS = (
    "#2E7D32",
    "#1565C0",
    "#6A1B9A",
    "#E65100",
    "#00838F",
    "#C62828",
    "#4527A0",
    "#AD1457",
)
_CARD_HEADER_H = 54
_EDGE_PX = 4


def _card_pixel_width(excel_width: int) -> int:
    return max(52, int(excel_width) * 8)


def _group_bar_color(group_id: str) -> str:
    if not group_id:
        return "#888888"
    acc = 0
    for ch in group_id:
        acc = (acc * 33 + ord(ch)) & 0xFFFFFFFF
    return _GROUP_COLORS[acc % len(_GROUP_COLORS)]


class _ColumnCard(QFrame):
    """One vertical column card (header colour matches the Excel section)."""

    select_requested = Signal(int, object)
    reorder_requested = Signal(int, int)
    width_changed = Signal(int, int)
    output_changed = Signal(int, bool)
    header_edit_requested = Signal(int)

    def __init__(
        self,
        index: int,
        setting: dict[str, Any],
        *,
        readonly: bool,
        parent: QWidget | None = None,
    ) -> None:
        super().__init__(parent)
        self._index = index
        self._readonly = readonly
        self._selected = False
        self._resize_active = False
        self._drag_start: QPoint | None = None
        self._press_on_edge = False
        self.setFrameShape(QFrame.Shape.Box)
        self.setLineWidth(1)
        self.setAcceptDrops(not readonly)
        self.setMouseTracking(True)
        self._build(setting)
        self._apply_width(int(setting.get("width") or MIN_COLUMN_WIDTH))
        self._refresh_chrome()

    def column_index(self) -> int:
        return self._index

    def set_selected(self, selected: bool) -> None:
        self._selected = selected
        self._refresh_chrome()

    def _build(self, setting: dict[str, Any]) -> None:
        layout = QVBoxLayout(self)
        layout.setContentsMargins(4, 4, 4, 4)
        layout.setSpacing(3)

        self._group_bar = QLabel(self)
        self._group_bar.setFixedHeight(10)
        self._group_bar.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(self._group_bar)

        fill = HEADER_FILL_COLORS.get(str(setting.get("section") or "rfp"), "#dbbcdb")
        self._header = QLabel(str(setting.get("header_label") or ""), self)
        self._header.setWordWrap(True)
        self._header.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self._header.setFixedHeight(_CARD_HEADER_H)
        self._header.setStyleSheet(
            f"background:{fill}; color:#000000; font-weight:bold; padding:2px;"
        )
        layout.addWidget(self._header)

        self._off_sheet = QLabel("не на листе", self)
        self._off_sheet.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self._off_sheet.setStyleSheet("color:#666666; font-size:10px;")
        layout.addWidget(self._off_sheet)

        width_row = QHBoxLayout()
        width_row.setContentsMargins(0, 0, 0, 0)
        width_row.addWidget(QLabel("Шир.", self))
        self._width_spin = QSpinBox(self)
        self._width_spin.setRange(MIN_COLUMN_WIDTH, MAX_COLUMN_WIDTH)
        self._width_spin.setValue(int(setting.get("width") or MIN_COLUMN_WIDTH))
        self._width_spin.setEnabled(not self._readonly)
        self._width_spin.valueChanged.connect(self._on_spin_width)
        width_row.addWidget(self._width_spin)
        layout.addLayout(width_row)

        self._on_sheet = QCheckBox("На лист", self)
        self._on_sheet.setChecked(bool(setting.get("output")))
        self._on_sheet.setEnabled(not self._readonly)
        self._on_sheet.toggled.connect(self._on_output_toggled)
        layout.addWidget(self._on_sheet)

        self._code = QLabel(str(setting.get("col_name") or ""), self)
        self._code.setStyleSheet("color:#888888; font-size:10px;")
        self._code.setWordWrap(True)
        self._code.setTextInteractionFlags(Qt.TextInteractionFlag.NoTextInteraction)
        layout.addWidget(self._code)
        layout.addStretch(1)

        self._apply_output_look(bool(setting.get("output")))
        self._apply_group_bar(setting)

    def _apply_group_bar(self, setting: dict[str, Any]) -> None:
        gid = str(setting.get("group_id") or "").strip()
        if not gid:
            self._group_bar.setText("")
            self._group_bar.setStyleSheet("background:transparent;")
            return
        color = _group_bar_color(gid)
        collapsed = bool(setting.get("group_collapsed"))
        self._group_bar.setText("скрыта" if collapsed else "")
        if collapsed:
            self._group_bar.setStyleSheet(
                f"color:#ffffff; font-size:8px; border:1px dashed {color}; background:{color};"
            )
        else:
            self._group_bar.setStyleSheet(
                f"color:#ffffff; font-size:8px; background:{color};"
            )

    def _apply_output_look(self, on_sheet: bool) -> None:
        self._off_sheet.setVisible(not on_sheet)
        if on_sheet:
            self.setStyleSheet("")
        else:
            self.setStyleSheet("QFrame { background:#f4f4f4; }")

    def _apply_width(self, excel_width: int) -> None:
        px = _card_pixel_width(excel_width)
        self.setFixedWidth(px)

    def _refresh_chrome(self) -> None:
        if self._selected:
            self.setLineWidth(2)
            extra = "QFrame { border: 2px solid #1565C0; }"
            current = self.styleSheet() or ""
            if extra not in current:
                self.setStyleSheet(current + extra)
        else:
            self.setLineWidth(1)
            current = self.styleSheet() or ""
            self.setStyleSheet(current.replace("QFrame { border: 2px solid #1565C0; }", ""))

    def set_header_fill(self, hex_color: str) -> None:
        self._header.setStyleSheet(
            f"background:{hex_color}; color:#000000; font-weight:bold; padding:2px;"
        )

    def mouseDoubleClickEvent(self, event: QMouseEvent) -> None:
        if self._readonly:
            return
        if self._header.geometry().contains(event.position().toPoint()):
            self.header_edit_requested.emit(self._index)
            event.accept()
            return
        super().mouseDoubleClickEvent(event)

    def mousePressEvent(self, event: QMouseEvent) -> None:
        if event.button() != Qt.MouseButton.LeftButton:
            return super().mousePressEvent(event)
        pos = event.position().toPoint()
        self._press_on_edge = (not self._readonly) and pos.x() >= self.width() - _EDGE_PX
        self._resize_active = self._press_on_edge
        self._drag_start = pos
        self.select_requested.emit(self._index, event.modifiers())
        event.accept()

    def mouseMoveEvent(self, event: QMouseEvent) -> None:
        pos = event.position().toPoint()
        if self._readonly:
            return
        if self._resize_active:
            excel_w = max(
                MIN_COLUMN_WIDTH,
                min(MAX_COLUMN_WIDTH, max(52, pos.x()) // 8),
            )
            if self._width_spin.value() != excel_w:
                self._width_spin.blockSignals(True)
                self._width_spin.setValue(excel_w)
                self._width_spin.blockSignals(False)
                self._apply_width(excel_w)
                self.width_changed.emit(self._index, excel_w)
            event.accept()
            return
        if (
            self._drag_start is not None
            and not self._press_on_edge
            and (event.buttons() & Qt.MouseButton.LeftButton)
            and (pos - self._drag_start).manhattanLength() >= 8
        ):
            self._start_drag()
            event.accept()
            return
        if pos.x() >= self.width() - _EDGE_PX:
            self.setCursor(Qt.CursorShape.SizeHorCursor)
        else:
            self.setCursor(Qt.CursorShape.OpenHandCursor)
        super().mouseMoveEvent(event)

    def mouseReleaseEvent(self, event: QMouseEvent) -> None:
        self._resize_active = False
        self._drag_start = None
        self._press_on_edge = False
        super().mouseReleaseEvent(event)

    def _start_drag(self) -> None:
        drag = QDrag(self)
        mime = QMimeData()
        mime.setData(_MIME_COL_INDEX, str(self._index).encode("ascii"))
        drag.setMimeData(mime)
        drag.exec(Qt.DropAction.MoveAction)

    def dragEnterEvent(self, event) -> None:  # noqa: ANN001
        if self._readonly or not event.mimeData().hasFormat(_MIME_COL_INDEX):
            event.ignore()
            return
        event.acceptProposedAction()

    def dropEvent(self, event) -> None:  # noqa: ANN001
        if self._readonly or not event.mimeData().hasFormat(_MIME_COL_INDEX):
            event.ignore()
            return
        raw = bytes(event.mimeData().data(_MIME_COL_INDEX)).decode("ascii")
        try:
            src = int(raw)
        except ValueError:
            event.ignore()
            return
        dest = self._index
        if event.position().x() > self.width() / 2:
            dest += 1
        self.reorder_requested.emit(src, dest)
        event.acceptProposedAction()

    def _on_spin_width(self, value: int) -> None:
        self._apply_width(value)
        self.width_changed.emit(self._index, value)

    def _on_output_toggled(self, checked: bool) -> None:
        self._apply_output_look(checked)
        self.output_changed.emit(self._index, checked)


class _StripHost(QWidget):
    """Horizontal host that accepts drops past the last card."""

    reorder_requested = Signal(int, int)

    def __init__(
        self,
        readonly_fn: Callable[[], bool],
        count_fn: Callable[[], int],
        parent: QWidget | None = None,
    ) -> None:
        super().__init__(parent)
        self._readonly_fn = readonly_fn
        self._count_fn = count_fn
        self.setAcceptDrops(True)

    def dragEnterEvent(self, event) -> None:  # noqa: ANN001
        if self._readonly_fn() or not event.mimeData().hasFormat(_MIME_COL_INDEX):
            event.ignore()
            return
        event.acceptProposedAction()

    def dropEvent(self, event) -> None:  # noqa: ANN001
        if self._readonly_fn() or not event.mimeData().hasFormat(_MIME_COL_INDEX):
            event.ignore()
            return
        raw = bytes(event.mimeData().data(_MIME_COL_INDEX)).decode("ascii")
        try:
            src = int(raw)
        except ValueError:
            event.ignore()
            return
        self.reorder_requested.emit(src, self._count_fn())
        event.acceptProposedAction()


class RfpColumnsPanel(QWidget):
    """Named Step4 Excel column templates as a horizontal strip of cards."""

    def __init__(
        self,
        parent: QWidget | None = None,
        *,
        on_saved: Callable[[], None] | None = None,
    ) -> None:
        super().__init__(parent)
        self._on_saved = on_saved
        self._columns: list[dict[str, Any]] = []
        self._cards: list[_ColumnCard] = []
        self._selected: set[int] = set()
        self._anchor = 0
        self._updating_combo = False
        self._section_by_name = self._load_sections()

        self._combo = QComboBox(self)
        self._hint = QLabel(_DEFAULT_HINT, self)
        self._hint.setWordWrap(True)

        self._btn_create = QPushButton("Создать из этого", self)
        self._btn_save = QPushButton("Сохранить", self)
        self._btn_rename = QPushButton("Переименовать", self)
        self._btn_delete = QPushButton("Удалить", self)
        self._btn_group = QPushButton("Сгруппировать", self)
        self._btn_ungroup = QPushButton("Разгруппировать", self)
        self._chk_hide = QCheckBox("Скрыть группу", self)

        self._btn_create.clicked.connect(self._create_from_current)
        self._btn_save.clicked.connect(self._save_active)
        self._btn_rename.clicked.connect(self._rename_active)
        self._btn_delete.clicked.connect(self._delete_active)
        self._btn_group.clicked.connect(self._group_selected)
        self._btn_ungroup.clicked.connect(self._ungroup_selected)
        self._chk_hide.toggled.connect(self._on_hide_toggled)
        self._combo.currentIndexChanged.connect(self._on_combo_changed)

        top = QHBoxLayout()
        top.addWidget(QLabel("Шаблон:", self))
        top.addWidget(self._combo, stretch=1)
        top.addWidget(self._btn_create)
        top.addWidget(self._btn_save)
        top.addWidget(self._btn_rename)
        top.addWidget(self._btn_delete)

        group_row = QHBoxLayout()
        group_row.addWidget(self._btn_group)
        group_row.addWidget(self._btn_ungroup)
        group_row.addWidget(self._chk_hide)
        group_row.addStretch(1)

        self._scroll = QScrollArea(self)
        self._scroll.setWidgetResizable(False)
        self._scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        self._scroll.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self._scroll.setMinimumHeight(220)
        self._scroll.setSizePolicy(
            QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred
        )

        self._strip = _StripHost(self._is_readonly, lambda: len(self._columns), self)
        self._strip.reorder_requested.connect(self._reorder)
        self._strip_layout = QHBoxLayout(self._strip)
        self._strip_layout.setContentsMargins(6, 8, 6, 8)
        self._strip_layout.setSpacing(4)
        self._strip_layout.addStretch(1)
        self._scroll.setWidget(self._strip)

        layout = QVBoxLayout(self)
        layout.addLayout(top)
        layout.addWidget(self._hint)
        layout.addLayout(group_row)
        layout.addWidget(self._scroll, stretch=1)

        self.reload_from_disk()

    @staticmethod
    def _load_sections() -> dict[str, str]:
        from RFQ.tags_rfp_compare.step4.step4_6_save_match_result_to_excel import (
            OUTPUT_COLUMNS_CONFIG,
        )

        return {defn.col_name: defn.section for defn in OUTPUT_COLUMNS_CONFIG}

    def _is_readonly(self) -> bool:
        return self._current_template_id() == DEFAULT_TEMPLATE_ID

    def _current_template_id(self) -> str:
        data = self._combo.currentData()
        if data is None:
            return DEFAULT_TEMPLATE_ID
        return str(data)

    def reload_from_disk(self) -> None:
        """Reload combo + strip from config without writing column edits."""
        state = load_step4_excel_column_state()
        active_id = str(state.get("active_id") or DEFAULT_TEMPLATE_ID)
        templates = list(state.get("templates") or [])
        known = {DEFAULT_TEMPLATE_ID} | {str(t.get("id")) for t in templates}
        if active_id not in known:
            active_id = DEFAULT_TEMPLATE_ID
        self._updating_combo = True
        self._combo.blockSignals(True)
        self._combo.clear()
        self._combo.addItem("По умолчанию", DEFAULT_TEMPLATE_ID)
        for tmpl in templates:
            self._combo.addItem(str(tmpl.get("name") or tmpl.get("id")), tmpl.get("id"))
        idx = self._combo.findData(active_id)
        self._combo.setCurrentIndex(max(0, idx))
        self._combo.blockSignals(False)
        self._updating_combo = False
        self._load_columns_for_id(active_id, state)
        self._apply_readonly()

    def _load_columns_for_id(
        self,
        template_id: str,
        state: dict[str, Any] | None = None,
    ) -> None:
        if state is None:
            state = load_step4_excel_column_state()
        if template_id == DEFAULT_TEMPLATE_ID:
            columns = builtin_column_settings()
        else:
            tmpl = next(
                (
                    item
                    for item in state.get("templates") or []
                    if item.get("id") == template_id
                ),
                None,
            )
            columns = normalize_template_columns(
                None if tmpl is None else tmpl.get("columns")
            )
        for item in columns:
            item["section"] = self._section_by_name.get(item["col_name"], "rfp")
        self._columns = columns
        self._selected.clear()
        self._rebuild_strip()

    def _apply_readonly(self) -> None:
        readonly = self._is_readonly()
        self._btn_save.setEnabled(not readonly)
        self._btn_rename.setEnabled(not readonly)
        self._btn_delete.setEnabled(not readonly)
        self._btn_group.setEnabled(not readonly)
        self._btn_ungroup.setEnabled(not readonly)
        self._chk_hide.setEnabled(not readonly)
        self._hint.setText(_DEFAULT_HINT if readonly else _USER_HINT)

    def _rebuild_strip(self) -> None:
        while self._strip_layout.count() > 1:
            item = self._strip_layout.takeAt(0)
            widget = item.widget()
            if widget is not None:
                widget.setParent(None)
                widget.deleteLater()
        self._cards = []
        readonly = self._is_readonly()
        for idx, setting in enumerate(self._columns):
            card = _ColumnCard(idx, setting, readonly=readonly, parent=self._strip)
            fill = HEADER_FILL_COLORS.get(setting.get("section", "rfp"), "#dbbcdb")
            card.set_header_fill(fill)
            card.set_selected(idx in self._selected)
            card.select_requested.connect(self._on_card_select)
            card.reorder_requested.connect(self._reorder)
            card.width_changed.connect(self._on_width_changed)
            card.output_changed.connect(self._on_output_changed)
            card.header_edit_requested.connect(self._on_header_edit)
            self._strip_layout.insertWidget(idx, card)
            self._cards.append(card)
        self._strip.adjustSize()
        total_w = self._strip_layout.sizeHint().width()
        total_h = max(200, self._strip_layout.sizeHint().height())
        self._strip.resize(max(total_w, 80), total_h)

    def _on_combo_changed(self, _index: int) -> None:
        if self._updating_combo:
            return
        tid = self._current_template_id()
        set_active_template(tid)
        self._load_columns_for_id(tid)
        self._apply_readonly()

    def _on_card_select(self, index: int, modifiers: Qt.KeyboardModifiers) -> None:
        if modifiers & Qt.KeyboardModifier.ControlModifier:
            if index in self._selected:
                self._selected.discard(index)
            else:
                self._selected.add(index)
            self._anchor = index
        elif modifiers & Qt.KeyboardModifier.ShiftModifier:
            lo, hi = sorted((self._anchor, index))
            self._selected = set(range(lo, hi + 1))
        else:
            self._selected = {index}
            self._anchor = index
        for card in self._cards:
            card.set_selected(card.column_index() in self._selected)
        self._sync_hide_checkbox()

    def _sync_hide_checkbox(self) -> None:
        if not self._selected:
            return
        flags = [
            bool(self._columns[i].get("group_collapsed"))
            for i in sorted(self._selected)
            if self._columns[i].get("group_id")
        ]
        if not flags:
            return
        self._chk_hide.blockSignals(True)
        self._chk_hide.setChecked(any(flags))
        self._chk_hide.blockSignals(False)

    def _on_width_changed(self, index: int, width: int) -> None:
        if 0 <= index < len(self._columns):
            self._columns[index]["width"] = int(width)
        self._strip.adjustSize()
        self._strip.resize(
            max(self._strip_layout.sizeHint().width(), 80),
            max(200, self._strip_layout.sizeHint().height()),
        )

    def _on_output_changed(self, index: int, output: bool) -> None:
        if 0 <= index < len(self._columns):
            self._columns[index]["output"] = bool(output)

    def _on_header_edit(self, index: int) -> None:
        if self._is_readonly() or not (0 <= index < len(self._columns)):
            return
        current = str(self._columns[index].get("header_label") or "")
        text, ok = QInputDialog.getText(
            self,
            "Подпись столбца",
            "Название в Excel:",
            text=current,
        )
        if not ok:
            return
        label = text.strip()
        if not label:
            QMessageBox.warning(self, "Подпись столбца", "Пустая подпись не допускается.")
            return
        self._columns[index]["header_label"] = label
        self._rebuild_strip()

    def _reorder(self, src: int, dest: int) -> None:
        if self._is_readonly():
            return
        n = len(self._columns)
        if src < 0 or src >= n:
            return
        dest = max(0, min(n, dest))
        if dest == src or dest == src + 1:
            return
        item = self._columns.pop(src)
        if dest > src:
            dest -= 1
        self._columns.insert(dest, item)
        self._selected = {dest}
        self._anchor = dest
        self._rebuild_strip()

    def _selected_sorted(self) -> list[int]:
        return sorted(i for i in self._selected if 0 <= i < len(self._columns))

    def _fresh_group_id(self) -> str:
        existing = {str(item.get("group_id") or "") for item in self._columns}
        n = 1
        while f"g{n}" in existing:
            n += 1
        return f"g{n}"

    def _group_selected(self) -> None:
        if self._is_readonly():
            return
        idxs = self._selected_sorted()
        if len(idxs) < 2:
            QMessageBox.information(
                self,
                "Группа столбцов",
                "Выберите несколько соседних столбцов (Ctrl или Shift).",
            )
            return
        if idxs[-1] - idxs[0] + 1 != len(idxs):
            QMessageBox.warning(
                self, "Группа столбцов", "Группа только из соседних столбцов."
            )
            return
        gid = self._fresh_group_id()
        collapsed = self._chk_hide.isChecked()
        for i in idxs:
            self._columns[i]["group_id"] = gid
            self._columns[i]["group_collapsed"] = collapsed
        self._rebuild_strip()

    def _ungroup_selected(self) -> None:
        if self._is_readonly():
            return
        ids = {
            str(self._columns[i].get("group_id") or "")
            for i in self._selected_sorted()
        }
        ids.discard("")
        if not ids:
            for i in self._selected_sorted():
                self._columns[i]["group_id"] = ""
                self._columns[i]["group_collapsed"] = False
            self._rebuild_strip()
            return
        for item in self._columns:
            if str(item.get("group_id") or "") in ids:
                item["group_id"] = ""
                item["group_collapsed"] = False
        self._rebuild_strip()

    def _on_hide_toggled(self, checked: bool) -> None:
        if self._is_readonly():
            return
        ids = {
            str(self._columns[i].get("group_id") or "")
            for i in self._selected_sorted()
        }
        ids.discard("")
        if not ids:
            return
        for item in self._columns:
            if str(item.get("group_id") or "") in ids:
                item["group_collapsed"] = bool(checked)
        self._rebuild_strip()

    def _columns_payload(self) -> list[dict[str, Any]]:
        payload = []
        for item in self._columns:
            payload.append(
                {
                    "col_name": item["col_name"],
                    "output": bool(item.get("output")),
                    "width": int(item.get("width") or MIN_COLUMN_WIDTH),
                    "header_label": str(item.get("header_label") or ""),
                    "group_id": str(item.get("group_id") or ""),
                    "group_collapsed": bool(item.get("group_collapsed")),
                }
            )
        return payload

    def _select_template_id(self, template_id: str) -> None:
        self._updating_combo = True
        idx = self._combo.findData(template_id)
        if idx >= 0:
            self._combo.setCurrentIndex(idx)
        self._updating_combo = False

    def _create_from_current(self) -> None:
        name, ok = QInputDialog.getText(self, "Новый шаблон", "Имя шаблона:")
        if not ok:
            return
        label = name.strip()
        if not label:
            QMessageBox.warning(self, "Новый шаблон", "Имя не должно быть пустым.")
            return
        tid = create_template(label, self._columns_payload())
        self.reload_from_disk()
        self._select_template_id(tid)
        self._load_columns_for_id(tid)
        self._apply_readonly()
        if self._on_saved is not None:
            self._on_saved()

    def _save_active(self) -> None:
        tid = self._current_template_id()
        if tid == DEFAULT_TEMPLATE_ID:
            return
        try:
            update_template_columns(tid, self._columns_payload())
        except ValueError as exc:
            QMessageBox.warning(self, "Сохранить шаблон", str(exc))
            return
        if self._on_saved is not None:
            self._on_saved()

    def _rename_active(self) -> None:
        tid = self._current_template_id()
        if tid == DEFAULT_TEMPLATE_ID:
            return
        current = self._combo.currentText()
        name, ok = QInputDialog.getText(
            self, "Переименовать", "Имя шаблона:", text=current
        )
        if not ok:
            return
        label = name.strip()
        if not label:
            QMessageBox.warning(self, "Переименовать", "Имя не должно быть пустым.")
            return
        try:
            rename_template(tid, label)
        except ValueError as exc:
            QMessageBox.warning(self, "Переименовать", str(exc))
            return
        self.reload_from_disk()
        self._select_template_id(tid)
        self._load_columns_for_id(tid)
        if self._on_saved is not None:
            self._on_saved()

    def _delete_active(self) -> None:
        tid = self._current_template_id()
        if tid == DEFAULT_TEMPLATE_ID:
            return
        answer = QMessageBox.question(
            self,
            "Удалить шаблон",
            "Удалить текущий шаблон столбцов? Будет выбран шаблон по умолчанию.",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        if answer != QMessageBox.StandardButton.Yes:
            return
        try:
            delete_template(tid)
        except ValueError as exc:
            QMessageBox.warning(self, "Удалить шаблон", str(exc))
            return
        set_active_template(DEFAULT_TEMPLATE_ID)
        self.reload_from_disk()
        if self._on_saved is not None:
            self._on_saved()
