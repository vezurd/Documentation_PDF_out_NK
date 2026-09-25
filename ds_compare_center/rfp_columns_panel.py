"""WYSIWYG Step4 Excel column templates: sheet row, unused row, groups."""

from __future__ import annotations

import zipfile
from typing import Any, Callable

from PySide6.QtCore import QEvent, QMimeData, QPoint, Qt, QTimer, Signal
from PySide6.QtGui import QDrag, QMouseEvent
from PySide6.QtWidgets import (
    QCheckBox,
    QApplication,
    QComboBox,
    QDialog,
    QDialogButtonBox,
    QFileDialog,
    QFrame,
    QGridLayout,
    QHBoxLayout,
    QInputDialog,
    QLabel,
    QLineEdit,
    QLayout,
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
    format_column_share,
    load_step4_excel_column_state,
    normalize_template_columns,
    rename_template,
    set_active_template,
    update_template_columns,
    column_xlsx_shares,
)

_MIME = "application/x-step4-col-move"
_DEFAULT_HINT = (
    "Шаблон по умолчанию только для просмотра и отката. "
    "Создайте свой, чтобы менять столбцы."
)
_USER_HINT = (
    "Верхний ряд пишется на лист. Нижний — столбцы не на листе. "
    "Перетащите карточку между рядами. «+» разворачивает скрытую группу, "
    "рядом подписан столбец, который останется. "
    "У края группы: курсор ближе к ней — столбец войдёт в группу, дальше — встанет снаружи."
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


def _center(unit: dict[str, Any]) -> float:
    return (float(unit["left"]) + float(unit["right"])) / 2.0


def plan_sheet_insertion(
    units: list[dict[str, Any]],
    cursor_x: int,
) -> tuple[int, str]:
    """Pick a landing spot on the sheet row.

    Each unit is one visible column or one collapsed group block:

    - ``left`` / ``right``: pixel edges in the row
    - ``start``: visible index of the first column in the unit
    - ``end``: visible index after the last column
    - ``group_id``: ``""`` when the unit is not in a group

    The cursor joins the nearer unit's group. On the outer side of a row
    (no neighbour) a cursor past the unit edge stays outside the group.

    Args:
        units: Visual units left to right, indices already excluding the drag source.
        cursor_x: Horizontal cursor position in the same coordinates as ``left``.

    Returns:
        Visible insert index and the ``group_id`` to assign (``""`` = outside).
    """
    if not units:
        return 0, ""

    def dist_to(unit: dict[str, Any]) -> int:
        if cursor_x < int(unit["left"]):
            return int(unit["left"]) - cursor_x
        if cursor_x > int(unit["right"]):
            return cursor_x - int(unit["right"])
        return 0

    nearest_i = min(range(len(units)), key=lambda i: (dist_to(units[i]), i))
    unit = units[nearest_i]
    center = _center(unit)
    prev_unit = units[nearest_i - 1] if nearest_i else None
    next_unit = units[nearest_i + 1] if nearest_i + 1 < len(units) else None

    if cursor_x < int(unit["left"]):
        return int(unit["start"]), ""
    if cursor_x > int(unit["right"]):
        return int(unit["end"]), ""

    if cursor_x <= center and prev_unit is not None:
        prev_center = _center(prev_unit)
        if abs(cursor_x - prev_center) < abs(cursor_x - center):
            return int(unit["start"]), str(prev_unit.get("group_id") or "")
        return int(unit["start"]), str(unit.get("group_id") or "")
    if cursor_x > center and next_unit is not None:
        next_center = _center(next_unit)
        if abs(cursor_x - next_center) < abs(cursor_x - center):
            return int(next_unit["start"]), str(next_unit.get("group_id") or "")
        return int(unit["end"]), str(unit.get("group_id") or "")

    insert_at = int(unit["start"]) if cursor_x <= center else int(unit["end"])
    return insert_at, str(unit.get("group_id") or "")


def insertion_line_x(units: list[dict[str, Any]], insert_at: int) -> int:
    """X of the vertical drop line for ``insert_at``."""
    if not units:
        return 8
    for unit in units:
        if int(unit["start"]) == insert_at:
            return int(unit["left"]) - 2
    return int(units[-1]["right"]) + 2


def heal_column_groups(columns: list[dict[str, Any]]) -> None:
    """Keep a group id only on a contiguous run of on-sheet columns.

    A run of one column is cleared. The same id after a gap becomes a new id.
    ``group_collapsed`` is copied onto every member of the run.

    Args:
        columns: Template rows in sheet order, then unused. Mutated in place.
    """
    seen: dict[str, str] = {}
    seq = 1
    index = 0
    while index < len(columns):
        gid = str(columns[index].get("group_id") or "").strip()
        on_sheet = bool(columns[index].get("output"))
        if not gid or not on_sheet:
            columns[index]["group_id"] = ""
            columns[index]["group_collapsed"] = False
            index += 1
            continue
        end = index + 1
        while (
            end < len(columns)
            and bool(columns[end].get("output"))
            and str(columns[end].get("group_id") or "").strip() == gid
        ):
            end += 1
        if end - index < 2:
            columns[index]["group_id"] = ""
            columns[index]["group_collapsed"] = False
            index = end
            continue
        if gid in seen:
            while f"g{seq}" in seen.values() or any(
                str(item.get("group_id") or "") == f"g{seq}" for item in columns
            ):
                seq += 1
            gid = f"g{seq}"
            seq += 1
        seen[str(columns[index].get("group_id") or "")] = gid
        collapsed = any(bool(columns[pos].get("group_collapsed")) for pos in range(index, end))
        for pos in range(index, end):
            columns[pos]["group_id"] = gid
            columns[pos]["group_collapsed"] = collapsed
        index = end


class _ColumnCard(QFrame):
    """One vertical column card. Header colour matches the Excel section."""

    select_requested = Signal(int, object)
    width_changed = Signal(int, int)
    output_changed = Signal(int, bool)
    header_edit_requested = Signal(int)
    properties_requested = Signal(int)
    drag_started = Signal(str)

    def __init__(
        self,
        index: int,
        setting: dict[str, Any],
        *,
        readonly: bool,
        drop_handler: Callable[[str, QPoint, str, str], None] | None = None,
        parent: QWidget | None = None,
    ) -> None:
        super().__init__(parent)
        self._index = index
        self._readonly = readonly
        self._selected = False
        self._resize_active = False
        self._drag_start: QPoint | None = None
        self._press_on_edge = False
        self._drop_handler = drop_handler
        self._col_name = str(setting.get("col_name") or "")
        self.setObjectName("column-card")
        self.setFrameShape(QFrame.Shape.Box)
        self.setLineWidth(1)
        self.setAcceptDrops(drop_handler is not None and not readonly)
        self.setMouseTracking(True)
        self._build(setting)
        self._apply_width(int(setting.get("width") or MIN_COLUMN_WIDTH))
        self.setMinimumHeight(max(150, self.sizeHint().height()))
        self._refresh_chrome()

    def column_index(self) -> int:
        return self._index

    def col_name(self) -> str:
        return self._col_name

    def set_selected(self, selected: bool) -> None:
        self._selected = selected
        self._refresh_chrome()

    def set_stays(self, stays: bool) -> None:
        self._stays.setVisible(stays)

    def set_share(self, text: str) -> None:
        self._share.setText(text)
        self._share.setVisible(bool(text))

    def _build(self, setting: dict[str, Any]) -> None:
        layout = QVBoxLayout(self)
        layout.setContentsMargins(4, 4, 4, 4)
        layout.setSpacing(3)

        self._stays = QLabel("останется", self)
        self._stays.setObjectName("stays-badge")
        self._stays.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self._stays.setStyleSheet(
            "background:#1B5E20; color:#ffffff; font-size:10px; font-weight:bold; padding:1px;"
        )
        self._stays.setVisible(False)
        layout.addWidget(self._stays)

        fill = HEADER_FILL_COLORS.get(str(setting.get("section") or "rfp"), "#dbbcdb")
        self._header = QLabel(str(setting.get("header_label") or ""), self)
        self._header.setWordWrap(True)
        self._header.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self._header.setFixedHeight(_CARD_HEADER_H)
        self._header.setStyleSheet(
            f"background:{fill}; color:#000000; font-weight:bold; padding:2px;"
        )
        layout.addWidget(self._header)

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

        self._code = QLabel(self._col_name, self)
        self._code.setStyleSheet("color:#888888; font-size:10px;")
        self._code.setWordWrap(True)
        layout.addWidget(self._code)
        self._share = QLabel("", self)
        self._share.setObjectName("share-label")
        self._share.setWordWrap(True)
        self._share.setStyleSheet("color:#333333; font-size:10px;")
        self._share.setVisible(False)
        layout.addWidget(self._share)
        self._comments_off = QLabel("комм. выкл.", self)
        self._comments_off.setObjectName("comments-off")
        self._comments_off.setStyleSheet("color:#B71C1C; font-size:10px; font-weight:bold;")
        self._comments_off.setVisible(not bool(setting.get("write_comments", True)))
        layout.addWidget(self._comments_off)
        layout.addStretch(1)
        self.installEventFilter(self)
        for child in self.findChildren(QWidget):
            child.installEventFilter(self)

    def _apply_width(self, excel_width: int) -> None:
        self.setFixedWidth(_card_pixel_width(excel_width))

    def _refresh_chrome(self) -> None:
        border = "border: 2px solid #1565C0;" if self._selected else "border: 1px solid #bbbbbb;"
        self.setStyleSheet(f"QFrame#column-card {{ {border} background:#ffffff; }}")

    def mouseDoubleClickEvent(self, event: QMouseEvent) -> None:
        self.properties_requested.emit(self._index)
        event.accept()

    def eventFilter(self, watched, event) -> bool:  # noqa: ANN001
        if event.type() == QEvent.Type.MouseButtonDblClick:
            self.properties_requested.emit(self._index)
            return True
        return super().eventFilter(watched, event)

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
            excel_w = max(MIN_COLUMN_WIDTH, min(MAX_COLUMN_WIDTH, max(52, pos.x()) // 8))
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
        self.setCursor(
            Qt.CursorShape.SizeHorCursor
            if pos.x() >= self.width() - _EDGE_PX
            else Qt.CursorShape.OpenHandCursor
        )
        super().mouseMoveEvent(event)

    def mouseReleaseEvent(self, event: QMouseEvent) -> None:
        self._resize_active = False
        self._drag_start = None
        self._press_on_edge = False
        super().mouseReleaseEvent(event)

    def _start_drag(self) -> None:
        drag = QDrag(self)
        mime = QMimeData()
        mime.setData(_MIME, f"col:{self._index}".encode("ascii"))
        drag.setMimeData(mime)
        self.drag_started.emit(f"col:{self._index}")
        drag.exec(Qt.DropAction.MoveAction)

    def _forward(self, event, kind: str) -> None:
        if self._drop_handler is None or not event.mimeData().hasFormat(_MIME):
            event.ignore()
            return
        top_left = self.mapTo(self._drop_root(), QPoint(0, 0))
        pos = top_left + event.position().toPoint()
        token = bytes(event.mimeData().data(_MIME)).decode("utf-8")
        self._drop_handler(kind, pos, "sheet" if self._on_sheet.isChecked() else "unused", token)
        event.acceptProposedAction()

    def _drop_root(self) -> QWidget:
        widget: QWidget = self
        while widget.parentWidget() is not None and widget.objectName() not in {
            "sheet-row",
            "unused-row",
        }:
            parent = widget.parentWidget()
            if parent is None:
                break
            widget = parent
        return widget

    def dragEnterEvent(self, event) -> None:  # noqa: ANN001
        self._forward(event, "enter")

    def dragMoveEvent(self, event) -> None:  # noqa: ANN001
        self._forward(event, "move")

    def dropEvent(self, event) -> None:  # noqa: ANN001
        self._forward(event, "drop")

    def _on_spin_width(self, value: int) -> None:
        self._apply_width(value)
        self.width_changed.emit(self._index, value)

    def _on_output_toggled(self, checked: bool) -> None:
        self.output_changed.emit(self._index, checked)


class _GroupBox(QFrame):
    """Outline around one contiguous group, with +/− and an optional compact title."""

    toggle_requested = Signal(str)

    def __init__(
        self,
        group_id: str,
        *,
        collapsed: bool,
        titles: list[str],
        readonly: bool,
        drop_handler: Callable[[str, QPoint, str, str], None] | None = None,
        parent: QWidget | None = None,
    ) -> None:
        super().__init__(parent)
        self.group_id = group_id
        self.collapsed = collapsed
        self._drop_handler = drop_handler
        self.setAcceptDrops(drop_handler is not None and not readonly)
        self.setObjectName("group-box")
        color = _group_bar_color(group_id)
        self.setStyleSheet(
            f"QFrame#group-box {{ border: 2px solid {color}; background:#fafafa; }}"
        )
        outer = QVBoxLayout(self)
        outer.setContentsMargins(8, 8, 8, 8)
        outer.setSpacing(6)
        outer.setSizeConstraint(QLayout.SizeConstraint.SetMinimumSize)
        head = QHBoxLayout()
        head.setContentsMargins(0, 0, 0, 0)
        head.setSpacing(8)
        head.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignTop)
        self._btn = QPushButton("+" if collapsed else "−", self)
        self._btn.setObjectName("group-toggle")
        self._btn.setFixedSize(34, 34)
        self._btn.setEnabled(not readonly)
        self._btn.setToolTip(
            "Показать столбцы группы" if collapsed else "Скрыть столбцы группы в Excel"
        )
        if readonly:
            self._btn.setToolTip("Создайте свой шаблон, чтобы менять группы.")
        self._btn.clicked.connect(lambda: self.toggle_requested.emit(group_id))
        head.addWidget(self._btn, 0, Qt.AlignmentFlag.AlignTop)
        self._caption = QLabel("группа", self)
        self._caption.setObjectName("group-caption")
        self._caption.setStyleSheet(f"color:{color}; font-weight:bold;")
        self._caption.setMinimumHeight(34)
        self._caption.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
        head.addWidget(self._caption, 0, Qt.AlignmentFlag.AlignVCenter)
        head.addStretch(1)
        outer.addLayout(head)
        if collapsed and titles:
            hidden = "скроется: " + " · ".join(title for title in titles if title)
            detail = QLabel(hidden, self)
            detail.setObjectName("group-hidden")
            detail.setWordWrap(True)
            detail.setStyleSheet(f"color:{color};")
            outer.addWidget(detail)
        self._cards = QHBoxLayout()
        self._cards.setSpacing(4)
        outer.addLayout(self._cards)
        self.setMinimumWidth(168)

    def add_card(self, card: QWidget) -> None:
        self._cards.addWidget(card)

    def mousePressEvent(self, event: QMouseEvent) -> None:
        if event.button() == Qt.MouseButton.LeftButton and self.childAt(event.position().toPoint()) is None:
            self._drag_origin = event.position().toPoint()
        else:
            self._drag_origin = event.position().toPoint()
        super().mousePressEvent(event)

    def mouseMoveEvent(self, event: QMouseEvent) -> None:
        origin = getattr(self, "_drag_origin", None)
        if (
            origin is not None
            and not self._btn.geometry().contains(origin)
            and (event.buttons() & Qt.MouseButton.LeftButton)
            and (event.position().toPoint() - origin).manhattanLength() >= 8
        ):
            drag = QDrag(self)
            mime = QMimeData()
            mime.setData(_MIME, f"group:{self.group_id}".encode("utf-8"))
            drag.setMimeData(mime)
            self._drag_origin = None
            drag.exec(Qt.DropAction.MoveAction)
            event.accept()
            return
        super().mouseMoveEvent(event)

    def _forward_drop(self, event, kind: str) -> None:  # noqa: ANN001
        if self._drop_handler is None or not event.mimeData().hasFormat(_MIME):
            event.ignore()
            return
        root: QWidget = self
        while root.parentWidget() is not None and root.objectName() != "sheet-row":
            parent = root.parentWidget()
            if parent is None:
                break
            root = parent
        pos = self.mapTo(root, QPoint(0, 0)) + event.position().toPoint()
        token = bytes(event.mimeData().data(_MIME)).decode("utf-8")
        self._drop_handler(kind, pos, "sheet", token)
        event.acceptProposedAction()

    def dragEnterEvent(self, event) -> None:  # noqa: ANN001
        self._forward_drop(event, "enter")

    def dragMoveEvent(self, event) -> None:  # noqa: ANN001
        self._forward_drop(event, "move")

    def dropEvent(self, event) -> None:  # noqa: ANN001
        self._forward_drop(event, "drop")


class _RowFrame(QFrame):
    """Horizontal drop row. Forwards empty-area drags to the panel."""

    def __init__(
        self,
        name: str,
        on_drag: Callable[[str, QPoint, str, str], None],
        parent: QWidget | None = None,
    ) -> None:
        super().__init__(parent)
        self.setObjectName(name)
        self._row_kind = "sheet" if name == "sheet-row" else "unused"
        self._on_drag = on_drag
        self.setAcceptDrops(True)
        self.setSizePolicy(QSizePolicy.Policy.Minimum, QSizePolicy.Policy.Fixed)

    def _emit(self, event, kind: str) -> None:
        if not event.mimeData().hasFormat(_MIME):
            event.ignore()
            return
        token = bytes(event.mimeData().data(_MIME)).decode("utf-8")
        self._on_drag(kind, event.position().toPoint(), self._row_kind, token)
        event.acceptProposedAction()

    def dragEnterEvent(self, event) -> None:  # noqa: ANN001
        self._emit(event, "enter")

    def dragMoveEvent(self, event) -> None:  # noqa: ANN001
        self._emit(event, "move")

    def dropEvent(self, event) -> None:  # noqa: ANN001
        self._emit(event, "drop")


class _UnusedFlow(_RowFrame):
    """Unused columns wrap inside the viewport width. No horizontal overflow."""

    def __init__(
        self,
        name: str,
        on_drag: Callable[[str, QPoint, str, str], None],
        parent: QWidget | None = None,
    ) -> None:
        super().__init__(name, on_drag, parent)
        self._order: list[_ColumnCard] = []
        self._reflowing = False
        self._grid = QGridLayout(self)
        self._grid.setContentsMargins(6, 6, 6, 6)
        self._grid.setSpacing(4)
        self._grid.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignTop)
        self.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Minimum)

    def set_order(self, cards: list[_ColumnCard]) -> None:
        self._order = list(cards)
        self.reflow()

    def resizeEvent(self, event) -> None:  # noqa: ANN001
        super().resizeEvent(event)
        self.reflow()

    def _viewport_width(self) -> int:
        parent = self.parentWidget()
        while parent is not None and not isinstance(parent, QScrollArea):
            parent = parent.parentWidget()
        if isinstance(parent, QScrollArea):
            width = parent.viewport().width()
            if width > 40:
                return width
        return max(self.width(), 80)

    def reflow(self) -> None:
        if self._reflowing:
            return
        self._reflowing = True
        try:
            cards = [card for card in self._order if card.parent() in {self, None}]
            while self._grid.count():
                self._grid.takeAt(0)
            for col in range(self._grid.columnCount()):
                self._grid.setColumnStretch(col, 0)
                self._grid.setColumnMinimumWidth(col, 0)
            width = self._viewport_width()
            self.setMinimumWidth(0)
            self.setMaximumWidth(width)
            margin = 6
            spacing = 4
            limit = width - margin
            x = margin
            col = 0
            row = 0
            for card in cards:
                card.setParent(self)
                card.show()
                card_w = max(card.width(), 52)
                if col > 0 and x + card_w > limit:
                    x = margin
                    row += 1
                    col = 0
                layout = card.layout()
                if layout is not None and layout.hasHeightForWidth():
                    card_h = layout.heightForWidth(card_w)
                else:
                    card_h = card.sizeHint().height()
                card_h = max(card_h, card.minimumSizeHint().height(), 160)
                card.setMinimumHeight(card_h)
                card.setMaximumHeight(card_h)
                self._grid.addWidget(
                    card,
                    row,
                    col,
                    Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignTop,
                )
                self._grid.setColumnMinimumWidth(col, 0)
                x += card_w + spacing
                col += 1
            self._grid.activate()
            content_h = max(self._grid.totalMinimumSize().height(), 80)
            if self.minimumHeight() != content_h:
                self.setMinimumHeight(content_h)
        finally:
            self._reflowing = False


class _ColumnEditor(QDialog):
    """Larger copy of a column card: header, width, sheet flag, comments, share."""

    def __init__(
        self,
        setting: dict[str, Any],
        share_text: str,
        *,
        readonly: bool,
        parent: QWidget | None = None,
    ) -> None:
        super().__init__(parent)
        self.setWindowTitle("Столбец")
        self.setMinimumWidth(340)
        layout = QVBoxLayout(self)
        layout.setSpacing(8)
        fill = HEADER_FILL_COLORS.get(str(setting.get("section") or "rfp"), "#dbbcdb")
        self._header = QLineEdit(str(setting.get("header_label") or ""), self)
        self._header.setMinimumHeight(72)
        self._header.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self._header.setStyleSheet(
            f"background:{fill}; color:#000000; font-weight:bold; font-size:16px; padding:8px;"
        )
        layout.addWidget(self._header)

        width_row = QHBoxLayout()
        width_row.addWidget(QLabel("Ширина в Excel", self))
        self._width = QSpinBox(self)
        self._width.setRange(MIN_COLUMN_WIDTH, MAX_COLUMN_WIDTH)
        self._width.setValue(int(setting.get("width") or MIN_COLUMN_WIDTH))
        width_row.addWidget(self._width)
        layout.addLayout(width_row)

        self._on_sheet = QCheckBox("На лист", self)
        self._on_sheet.setChecked(bool(setting.get("output")))
        layout.addWidget(self._on_sheet)

        self._comments = QCheckBox("Писать комментарии", self)
        self._comments.setChecked(bool(setting.get("write_comments", True)))
        layout.addWidget(self._comments)

        share = QLabel(share_text or "Доля в файле появится после кнопки «Доля в xlsx».", self)
        share.setWordWrap(True)
        share.setStyleSheet("font-size:14px;")
        layout.addWidget(share)
        code = QLabel(str(setting.get("col_name") or ""), self)
        code.setStyleSheet("color:#888888;")
        layout.addWidget(code)

        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
            if not readonly
            else QDialogButtonBox.StandardButton.Close,
            self,
        )
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)
        if readonly:
            self._header.setReadOnly(True)
            self._width.setEnabled(False)
            self._on_sheet.setEnabled(False)
            self._comments.setEnabled(False)

    def values(self) -> tuple[str, int, bool, bool]:
        return (
            self._header.text().strip(),
            int(self._width.value()),
            self._on_sheet.isChecked(),
            self._comments.isChecked(),
        )


class RfpColumnsPanel(QWidget):
    """Named Step4 templates: on-sheet row, unused row, group +/−."""

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
        self._units: list[dict[str, Any]] = []
        self._shares: dict[str, dict[str, float]] = {}

        self._combo = QComboBox(self)
        self._hint = QLabel(_DEFAULT_HINT, self)
        self._hint.setWordWrap(True)

        self._btn_create = QPushButton("Создать из этого", self)
        self._btn_save = QPushButton("Сохранить", self)
        self._btn_rename = QPushButton("Переименовать", self)
        self._btn_delete = QPushButton("Удалить", self)
        self._btn_group = QPushButton("Сгруппировать", self)
        self._btn_ungroup = QPushButton("Разгруппировать", self)
        self._chk_hide = QCheckBox("Скрыть в Excel", self)
        self._btn_share = QPushButton("Доля в xlsx", self)
        self._btn_share.setToolTip(
            "По финальному файлу: доля столбца в архиве, отдельно комментарии."
        )

        self._btn_create.clicked.connect(self._create_from_current)
        self._btn_save.clicked.connect(self._save_active)
        self._btn_rename.clicked.connect(self._rename_active)
        self._btn_delete.clicked.connect(self._delete_active)
        self._btn_group.clicked.connect(self._group_selected)
        self._btn_ungroup.clicked.connect(self._ungroup_selected)
        self._chk_hide.toggled.connect(self._on_hide_toggled)
        self._btn_share.clicked.connect(self._load_xlsx_shares)
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
        group_row.addWidget(self._btn_share)
        group_row.addStretch(1)

        sheet_title = QLabel("На листе — эти столбцы пишутся в файл", self)
        sheet_title.setStyleSheet("font-weight:bold;")
        unused_title = QLabel("Не используются — на лист не попадают", self)
        unused_title.setStyleSheet("font-weight:bold;")

        self._sheet_scroll = QScrollArea(self)
        self._sheet_scroll.setObjectName("sheet-scroll")
        self._sheet_scroll.setWidgetResizable(False)
        self._sheet_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        self._sheet_scroll.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self._sheet_row = _RowFrame("sheet-row", self._on_row_drag, self._sheet_scroll)
        self._sheet_layout = QHBoxLayout(self._sheet_row)
        self._sheet_layout.setContentsMargins(6, 6, 28, 6)
        self._sheet_layout.setSpacing(4)
        self._sheet_layout.setSizeConstraint(QLayout.SizeConstraint.SetMinimumSize)
        self._sheet_layout.addStretch(1)
        self._sheet_row.setStyleSheet("QFrame#sheet-row { background:#ffffff; border: 1px solid #cccccc; }")
        self._sheet_scroll.setWidget(self._sheet_row)

        self._unused_scroll = QScrollArea(self)
        self._unused_scroll.setObjectName("unused-scroll")
        self._unused_scroll.setWidgetResizable(True)
        self._unused_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self._unused_scroll.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        self._unused_scroll.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        self._unused_row = _UnusedFlow("unused-row", self._on_row_drag, self._unused_scroll)
        self._unused_row.setStyleSheet("QFrame#unused-row { background:#f3f3f3; border: 1px solid #cccccc; }")
        self._unused_scroll.setWidget(self._unused_row)

        self._line = QFrame(self._sheet_row)
        self._line.setObjectName("drop-line")
        self._line.setFixedWidth(3)
        self._line.setStyleSheet("background:#1565C0;")
        self._line.hide()
        self._line_label = QLabel(self._sheet_row)
        self._line_label.setObjectName("drop-label")
        self._line_label.hide()

        layout = QVBoxLayout(self)
        layout.addLayout(top)
        layout.addWidget(self._hint)
        layout.addLayout(group_row)
        layout.addWidget(sheet_title)
        layout.addWidget(self._sheet_scroll)
        layout.addWidget(unused_title)
        layout.addWidget(self._unused_scroll, stretch=1)
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
        """Reload combo and both rows from config without writing column edits."""
        state = load_step4_excel_column_state()
        active_id = str(state.get("active_id") or DEFAULT_TEMPLATE_ID)
        templates = list(state.get("templates") or [])
        known = {DEFAULT_TEMPLATE_ID} | {str(item.get("id")) for item in templates}
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
        visible = [item for item in columns if item.get("output")]
        unused = [item for item in columns if not item.get("output")]
        self._columns = visible + unused
        self._selected.clear()
        self._rebuild_rows()

    def _apply_readonly(self) -> None:
        readonly = self._is_readonly()
        self._btn_save.setEnabled(not readonly)
        self._btn_rename.setEnabled(not readonly)
        self._btn_delete.setEnabled(not readonly)
        self._btn_group.setEnabled(not readonly)
        self._btn_ungroup.setEnabled(not readonly)
        self._chk_hide.setEnabled(not readonly)
        self._hint.setText(_DEFAULT_HINT if readonly else _USER_HINT)

    def _clear_layout(self, layout: QHBoxLayout) -> None:
        while layout.count() > 1:
            item = layout.takeAt(0)
            widget = item.widget()
            if widget is not None:
                widget.setParent(None)
                widget.deleteLater()

    def _make_card(self, index: int, setting: dict[str, Any], row: QWidget) -> _ColumnCard:
        card = _ColumnCard(
            index,
            setting,
            readonly=self._is_readonly(),
            drop_handler=self._on_row_drag,
            parent=row,
        )
        fill = HEADER_FILL_COLORS.get(setting.get("section", "rfp"), "#dbbcdb")
        card._header.setStyleSheet(
            f"background:{fill}; color:#000000; font-weight:bold; padding:2px;"
        )
        card.set_selected(index in self._selected)
        card.select_requested.connect(self._on_card_select)
        card.width_changed.connect(self._on_width_changed)
        card.output_changed.connect(self._on_output_changed)
        card.header_edit_requested.connect(self._on_header_edit)
        card.properties_requested.connect(self._on_properties)
        return card

    def _rebuild_rows(self) -> None:
        bar = self._sheet_scroll.horizontalScrollBar()
        saved_scroll = bar.value()
        self._hide_drop_line()
        self._clear_layout(self._sheet_layout)
        self._clear_unused_cards()
        self._cards = []
        visible_idx = [i for i, item in enumerate(self._columns) if item.get("output")]
        pos = 0
        while pos < len(visible_idx):
            master = visible_idx[pos]
            gid = str(self._columns[master].get("group_id") or "").strip()
            end = pos + 1
            if gid:
                while end < len(visible_idx) and str(
                    self._columns[visible_idx[end]].get("group_id") or ""
                ).strip() == gid:
                    end += 1
            if gid and end - pos >= 2:
                titles = [
                    str(self._columns[visible_idx[k]].get("header_label") or "")
                    for k in range(pos, end)
                ]
                collapsed = any(
                    bool(self._columns[visible_idx[k]].get("group_collapsed"))
                    for k in range(pos, end)
                )
                box = _GroupBox(
                    gid,
                    collapsed=collapsed,
                    titles=titles,
                    readonly=self._is_readonly(),
                    drop_handler=self._on_row_drag,
                    parent=self._sheet_row,
                )
                box.toggle_requested.connect(self._toggle_group)
                if not collapsed:
                    for k in range(pos, end):
                        card = self._make_card(visible_idx[k], self._columns[visible_idx[k]], box)
                        box.add_card(card)
                        self._cards.append(card)
                self._sheet_layout.insertWidget(self._sheet_layout.count() - 1, box)
                if collapsed and end < len(visible_idx):
                    pass
                pos = end
                continue
            card = self._make_card(master, self._columns[master], self._sheet_row)
            prev_collapsed = False
            if pos > 0:
                prev = self._columns[visible_idx[pos - 1]]
                prev_gid = str(prev.get("group_id") or "").strip()
                prev_collapsed = bool(prev_gid) and bool(prev.get("group_collapsed"))
            card.set_stays(prev_collapsed)
            self._sheet_layout.insertWidget(self._sheet_layout.count() - 1, card)
            self._cards.append(card)
            pos += 1

        for master, setting in enumerate(self._columns):
            if setting.get("output"):
                continue
            card = self._make_card(master, setting, self._unused_row)
            self._cards.append(card)
        self._apply_shares()
        self._fit_host()
        self._restore_sheet_scroll(saved_scroll)

    def _restore_sheet_scroll(self, value: int) -> None:
        """Keep the sheet viewport where it was after a row rebuild."""
        bar = self._sheet_scroll.horizontalScrollBar()

        def apply() -> None:
            bar.setValue(min(value, bar.maximum()))

        apply()
        QTimer.singleShot(0, apply)

    def _clear_unused_cards(self) -> None:
        for child in list(self._unused_row.children()):
            if isinstance(child, _ColumnCard):
                child.setParent(None)
                child.deleteLater()

    def _fit_host(self) -> None:
        for widget in self._sheet_row.findChildren(QWidget):
            widget.ensurePolished()
        self._sheet_layout.invalidate()
        self._sheet_layout.activate()
        hint = self._sheet_layout.totalMinimumSize()
        child_h = 0
        for index in range(self._sheet_layout.count()):
            widget = self._sheet_layout.itemAt(index).widget()
            if widget is None:
                continue
            child_h = max(child_h, widget.sizeHint().height(), widget.minimumSizeHint().height())
        row_h = max(200, child_h, hint.height(), self._sheet_layout.sizeHint().height()) + 36
        row_w = max(hint.width(), self._sheet_layout.sizeHint().width(), 200)
        self._sheet_row.setMinimumSize(row_w, row_h)
        self._sheet_row.resize(row_w, row_h)
        bar_h = self._sheet_scroll.horizontalScrollBar().sizeHint().height()
        self._sheet_scroll.setFixedHeight(row_h + bar_h + 4)
        for card in self._cards:
            if card.parent() is self._unused_row:
                card.ensurePolished()
                card.adjustSize()
        unused = [card for card in self._cards if card.parent() is self._unused_row]
        self._unused_row.set_order(unused)

    def _apply_shares(self) -> None:
        for card in self._cards:
            index = card.column_index()
            if not 0 <= index < len(self._columns):
                card.set_share("")
                continue
            label = str(self._columns[index].get("header_label") or "")
            share = self._shares.get(label)
            card.set_share("" if share is None else format_column_share(share))

    def _load_xlsx_shares(self) -> None:
        path, _selected = QFileDialog.getOpenFileName(
            self,
            "Финальный xlsx",
            "",
            "Excel (*.xlsx)",
        )
        if not path:
            return
        QApplication.setOverrideCursor(Qt.CursorShape.WaitCursor)
        try:
            shares = column_xlsx_shares(path)
        except (OSError, zipfile.BadZipFile) as exc:
            QMessageBox.warning(self, "Доля в xlsx", str(exc))
            return
        finally:
            QApplication.restoreOverrideCursor()
        if not shares:
            QMessageBox.information(self, "Доля в xlsx", "На первом листе нет заголовков.")
            return
        self._shares = shares
        self._apply_shares()
        self._fit_host()

    def _on_combo_changed(self, _index: int) -> None:
        if self._updating_combo:
            return
        template_id = self._current_template_id()
        set_active_template(template_id)
        self._load_columns_for_id(template_id)
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
        flags = [
            bool(self._columns[i].get("group_collapsed"))
            for i in sorted(self._selected)
            if 0 <= i < len(self._columns) and self._columns[i].get("group_id")
        ]
        if not flags:
            return
        self._chk_hide.blockSignals(True)
        self._chk_hide.setChecked(any(flags))
        self._chk_hide.blockSignals(False)

    def _on_width_changed(self, index: int, width: int) -> None:
        if 0 <= index < len(self._columns):
            self._columns[index]["width"] = int(width)
        self._fit_host()

    def _on_output_changed(self, index: int, output: bool) -> None:
        if self._is_readonly() or not (0 <= index < len(self._columns)):
            return
        name = str(self._columns[index].get("col_name") or "")
        self._columns[index]["output"] = bool(output)
        if not output:
            self._columns[index]["group_id"] = ""
            self._columns[index]["group_collapsed"] = False
        self._partition_and_heal()
        self._selected = {self._index_of_name(name)}
        self._anchor = next(iter(self._selected))
        self._rebuild_rows()

    def _index_of_name(self, col_name: str) -> int:
        for index, item in enumerate(self._columns):
            if item.get("col_name") == col_name:
                return index
        return 0

    def _on_header_edit(self, index: int) -> None:
        if self._is_readonly() or not (0 <= index < len(self._columns)):
            return
        current = str(self._columns[index].get("header_label") or "")
        text, ok = QInputDialog.getText(self, "Подпись столбца", "Название в Excel:", text=current)
        if not ok:
            return
        label = text.strip()
        if not label:
            QMessageBox.warning(self, "Подпись столбца", "Пустая подпись не допускается.")
            return
        self._columns[index]["header_label"] = label
        self._rebuild_rows()

    def _on_properties(self, index: int) -> None:
        if not 0 <= index < len(self._columns):
            return
        item = self._columns[index]
        label = str(item.get("header_label") or "")
        share = self._shares.get(label)
        dialog = _ColumnEditor(
            item,
            "" if share is None else format_column_share(share),
            readonly=self._is_readonly(),
            parent=self,
        )
        if dialog.exec() != QDialog.DialogCode.Accepted or self._is_readonly():
            return
        header, width, output, comments = dialog.values()
        if not header:
            QMessageBox.warning(self, "Столбец", "Пустая подпись не допускается.")
            return
        name = str(item.get("col_name") or "")
        item["header_label"] = header
        item["width"] = width
        item["write_comments"] = comments
        if bool(output) != bool(item.get("output")):
            self._on_output_changed(index, output)
            return
        self._selected = {self._index_of_name(name)}
        self._anchor = next(iter(self._selected))
        self._rebuild_rows()

    def _toggle_group(self, group_id: str) -> None:
        if self._is_readonly():
            return
        members = [
            item for item in self._columns if str(item.get("group_id") or "") == group_id
        ]
        if not members:
            return
        collapsed = not any(bool(item.get("group_collapsed")) for item in members)
        for item in members:
            item["group_collapsed"] = collapsed
        self._rebuild_rows()

    def _source_indices(self, token: str) -> list[int]:
        kind, _, payload = token.partition(":")
        if kind == "group":
            return [
                i
                for i, item in enumerate(self._columns)
                if str(item.get("group_id") or "") == payload and item.get("output")
            ]
        try:
            index = int(payload)
        except ValueError:
            return []
        if 0 <= index < len(self._columns):
            return [index]
        return []

    def _visible_units(self, skip: set[int]) -> list[dict[str, Any]]:
        """Geometry of on-sheet units, skipping dragged master indices."""
        self._sheet_row.adjustSize()
        units: list[dict[str, Any]] = []
        visible_pos = 0
        for widget_index in range(self._sheet_layout.count() - 1):
            widget = self._sheet_layout.itemAt(widget_index).widget()
            if widget is None:
                continue
            if isinstance(widget, _GroupBox):
                members = [
                    i
                    for i, item in enumerate(self._columns)
                    if str(item.get("group_id") or "") == widget.group_id and item.get("output")
                ]
                kept = [i for i in members if i not in skip]
                if not kept:
                    continue
                geo = widget.geometry()
                units.append(
                    {
                        "left": geo.left(),
                        "right": geo.right(),
                        "start": visible_pos,
                        "end": visible_pos + len(kept),
                        "group_id": widget.group_id,
                    }
                )
                visible_pos += len(kept)
                continue
            if isinstance(widget, _ColumnCard):
                if widget.column_index() in skip:
                    continue
                geo = widget.geometry()
                gid = str(self._columns[widget.column_index()].get("group_id") or "")
                units.append(
                    {
                        "left": geo.left(),
                        "right": geo.right(),
                        "start": visible_pos,
                        "end": visible_pos + 1,
                        "group_id": gid,
                    }
                )
                visible_pos += 1
        return units

    def _hide_drop_line(self) -> None:
        self._line.hide()
        self._line_label.hide()

    def _show_drop_line(self, x: int, group_id: str, near_group: bool) -> None:
        if group_id:
            text = "в группу"
            color = _group_bar_color(group_id)
        elif near_group:
            text = "вне группы"
            color = "#666666"
        else:
            text = ""
            color = "#1565C0"
        self._line.setStyleSheet(f"background:{color};")
        self._line.setGeometry(max(0, x), 4, 3, max(40, self._sheet_row.height() - 8))
        self._line.show()
        self._line.raise_()
        if text:
            self._line_label.setText(text)
            self._line_label.setStyleSheet(
                f"background:{color}; color:#ffffff; font-size:10px; padding:1px 3px;"
            )
            self._line_label.adjustSize()
            self._line_label.move(max(0, x + 6), 6)
            self._line_label.show()
            self._line_label.raise_()
        else:
            self._line_label.hide()

    def _on_row_drag(self, kind: str, pos: QPoint, row_kind: str, token: str) -> None:
        if self._is_readonly():
            self._hide_drop_line()
            return
        if kind == "drop":
            self._hide_drop_line()
            self._apply_drop(token, pos, row_kind)
            return
        if row_kind != "sheet":
            self._hide_drop_line()
            return
        skip = set(self._source_indices(token))
        units = self._visible_units(skip)
        insert_at, gid = plan_sheet_insertion(units, pos.x())
        near = any(str(unit.get("group_id") or "") for unit in units)
        self._show_drop_line(
            insertion_line_x(units, insert_at),
            gid,
            near_group=bool(gid) or near,
        )

    def _apply_drop(self, token: str, pos: QPoint, row_kind: str) -> None:
        indices = self._source_indices(token)
        if not indices:
            return
        names = [str(self._columns[i].get("col_name") or "") for i in indices]
        moving_group = token.startswith("group:")
        own_gid = token.split(":", 1)[1] if moving_group else ""
        if row_kind == "unused":
            self._move_names(
                names,
                visible_index=None,
                unused_at=self._unused_insert_at(pos, indices),
                group_id="",
            )
            return
        skip = set(indices)
        units = self._visible_units(skip)
        insert_at, gid = plan_sheet_insertion(units, pos.x())
        if moving_group and not gid:
            gid = own_gid
        self._move_names(names, visible_index=insert_at, unused_at=None, group_id=gid)

    def _unused_insert_at(self, pos: QPoint, skip: list[int]) -> int:
        cards = [
            card
            for card in self._unused_row._order
            if card.column_index() not in skip and card.parent() is self._unused_row
        ]
        x = pos.x()
        y = pos.y()
        at = 0
        for card in cards:
            geo = card.geometry()
            cy = (geo.top() + geo.bottom()) / 2
            cx = (geo.left() + geo.right()) / 2
            if cy < y - geo.height() / 2:
                at += 1
                continue
            if abs(cy - y) <= geo.height() / 2 + 6 and cx < x:
                at += 1
                continue
            break
        return at

    def _move_names(
        self,
        names: list[str],
        *,
        visible_index: int | None,
        unused_at: int | None,
        group_id: str,
    ) -> None:
        picked = [item for item in self._columns if item.get("col_name") in names]
        if not picked:
            return
        remain = [item for item in self._columns if item.get("col_name") not in names]
        visible = [item for item in remain if item.get("output")]
        unused = [item for item in remain if not item.get("output")]
        if visible_index is None:
            for item in picked:
                item["output"] = False
                item["group_id"] = ""
                item["group_collapsed"] = False
            at = 0 if unused_at is None else max(0, min(len(unused), unused_at))
            unused[at:at] = picked
        else:
            collapsed = False
            if group_id:
                collapsed = any(
                    str(item.get("group_id") or "") == group_id and bool(item.get("group_collapsed"))
                    for item in visible
                )
            for item in picked:
                item["output"] = True
                item["group_id"] = group_id
                item["group_collapsed"] = collapsed if group_id else False
            at = max(0, min(len(visible), visible_index))
            visible[at:at] = picked
        self._columns = visible + unused
        heal_column_groups(self._columns)
        focus = names[0]
        self._selected = {self._index_of_name(focus)}
        self._anchor = next(iter(self._selected))
        self._rebuild_rows()

    def _partition_and_heal(self) -> None:
        visible = [item for item in self._columns if item.get("output")]
        unused = [item for item in self._columns if not item.get("output")]
        self._columns = visible + unused
        heal_column_groups(self._columns)

    def _selected_sorted(self) -> list[int]:
        return sorted(i for i in self._selected if 0 <= i < len(self._columns))

    def _fresh_group_id(self) -> str:
        existing = {str(item.get("group_id") or "") for item in self._columns}
        number = 1
        while f"g{number}" in existing:
            number += 1
        return f"g{number}"

    def _group_selected(self) -> None:
        if self._is_readonly():
            return
        idxs = [i for i in self._selected_sorted() if self._columns[i].get("output")]
        if len(idxs) < 2:
            QMessageBox.information(
                self,
                "Группа столбцов",
                "Выберите несколько соседних столбцов верхнего ряда (Ctrl или Shift).",
            )
            return
        if idxs[-1] - idxs[0] + 1 != len(idxs):
            QMessageBox.warning(self, "Группа столбцов", "Группа только из соседних столбцов.")
            return
        gid = self._fresh_group_id()
        collapsed = self._chk_hide.isChecked()
        for index in idxs:
            self._columns[index]["group_id"] = gid
            self._columns[index]["group_collapsed"] = collapsed
        self._rebuild_rows()

    def _ungroup_selected(self) -> None:
        if self._is_readonly():
            return
        ids = {
            str(self._columns[i].get("group_id") or "")
            for i in self._selected_sorted()
        }
        ids.discard("")
        if not ids:
            self._rebuild_rows()
            return
        for item in self._columns:
            if str(item.get("group_id") or "") in ids:
                item["group_id"] = ""
                item["group_collapsed"] = False
        self._rebuild_rows()

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
        self._rebuild_rows()

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
                    "write_comments": bool(item.get("write_comments", True)),
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
        template_id = create_template(label, self._columns_payload())
        self.reload_from_disk()
        self._select_template_id(template_id)
        self._load_columns_for_id(template_id)
        self._apply_readonly()
        if self._on_saved is not None:
            self._on_saved()

    def _save_active(self) -> None:
        template_id = self._current_template_id()
        if template_id == DEFAULT_TEMPLATE_ID:
            return
        try:
            update_template_columns(template_id, self._columns_payload())
        except ValueError as exc:
            QMessageBox.warning(self, "Сохранить шаблон", str(exc))
            return
        if self._on_saved is not None:
            self._on_saved()

    def _rename_active(self) -> None:
        template_id = self._current_template_id()
        if template_id == DEFAULT_TEMPLATE_ID:
            return
        current = self._combo.currentText()
        name, ok = QInputDialog.getText(self, "Переименовать", "Имя шаблона:", text=current)
        if not ok:
            return
        label = name.strip()
        if not label:
            QMessageBox.warning(self, "Переименовать", "Имя не должно быть пустым.")
            return
        try:
            rename_template(template_id, label)
        except ValueError as exc:
            QMessageBox.warning(self, "Переименовать", str(exc))
            return
        self.reload_from_disk()
        self._select_template_id(template_id)
        self._load_columns_for_id(template_id)
        if self._on_saved is not None:
            self._on_saved()

    def _delete_active(self) -> None:
        template_id = self._current_template_id()
        if template_id == DEFAULT_TEMPLATE_ID:
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
            delete_template(template_id)
        except ValueError as exc:
            QMessageBox.warning(self, "Удалить шаблон", str(exc))
            return
        set_active_template(DEFAULT_TEMPLATE_ID)
        self.reload_from_disk()
        if self._on_saved is not None:
            self._on_saved()
