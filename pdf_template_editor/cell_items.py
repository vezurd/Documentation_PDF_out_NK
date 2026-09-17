"""FieldRectItem — перетаскиваемый прямоугольник поля на Canvas с resize handles."""

from __future__ import annotations

import copy

from PySide6.QtCore import Qt, QRectF, QPointF, QTimer
from PySide6.QtGui import QPen, QBrush, QColor, QPainter, QFont
from PySide6.QtWidgets import (
    QGraphicsRectItem,
    QGraphicsItem,
    QGraphicsSimpleTextItem,
    QGraphicsSceneHoverEvent,
    QGraphicsSceneMouseEvent,
    QStyleOptionGraphicsItem,
    QWidget,
)

from pdf_parsing_v2_engine.field_id_display import short_label_from_field_id
from pdf_parsing_v2_engine.models import FieldDef

_HANDLE_SIZE = 6.0
_MIN_ITEM_SIZE = _HANDLE_SIZE * 3

_COLOR_UNASSIGNED = QColor(160, 160, 160)
_COLOR_ASSIGNED = QColor(60, 120, 220)
_COLOR_LABEL = QColor(132, 72, 193)
_COLOR_DUP_ID = QColor(220, 40, 40)
_COLOR_OUTSIDE_STAMP = QColor(199, 93, 0)
_COLOR_ANCHOR = QColor(215, 45, 190)
_COLOR_OK = QColor(50, 180, 80)
_COLOR_WARNING = QColor(220, 180, 40)
_COLOR_ERROR = QColor(220, 50, 50)
_COLOR_HIGHLIGHT = QColor(255, 140, 0)
_COLOR_FILL_UNASSIGNED = QColor(160, 160, 160, 30)
_COLOR_FILL_ASSIGNED = QColor(60, 120, 220, 30)
_COLOR_FILL_LABEL = QColor(132, 72, 193, 28)
_COLOR_FILL_OUTSIDE_STAMP = QColor(199, 93, 0, 28)
_COLOR_FILL_BOUND_FULL    = QColor(0, 180, 0, 55)    # wireframe: все 4 стороны привязаны
_COLOR_FILL_BOUND_PARTIAL = QColor(220, 180, 0, 55)  # wireframe: 1–3 стороны привязаны


class _Handle(QGraphicsRectItem):
    """Small resize handle on corners / midpoints of a FieldRectItem.

    Handle indices (positions on parent rect):
        0=TL  1=TC  2=TR
        7=LC        3=RC
        6=BL  5=BC  4=BR
    """

    # Which edges each handle moves:
    _MOVES_LEFT   = {0, 6, 7}
    _MOVES_RIGHT  = {2, 3, 4}
    _MOVES_TOP    = {0, 1, 2}
    _MOVES_BOTTOM = {4, 5, 6}

    def __init__(self, index: int, cursor_shape: Qt.CursorShape, parent: "FieldRectItem"):
        super().__init__(-_HANDLE_SIZE / 2, -_HANDLE_SIZE / 2, _HANDLE_SIZE, _HANDLE_SIZE, parent)
        self._index = index
        self._dragging = False
        self._drag_start = QPointF()
        self._orig_scene_rect = QRectF()

        self.setBrush(QBrush(QColor(255, 255, 255)))
        self.setPen(QPen(QColor(0, 0, 0), 1))
        self.setFlag(QGraphicsItem.GraphicsItemFlag.ItemIsMovable, False)
        self.setFlag(QGraphicsItem.GraphicsItemFlag.ItemSendsGeometryChanges, False)
        self.setCursor(cursor_shape)
        self.setAcceptHoverEvents(True)
        self.setZValue(10)

    def mousePressEvent(self, event: QGraphicsSceneMouseEvent) -> None:
        if event.button() == Qt.MouseButton.LeftButton:
            self._dragging = True
            self._drag_start = event.scenePos()
            parent: FieldRectItem = self.parentItem()
            pos = parent.scenePos()
            r = parent.rect()
            self._orig_scene_rect = QRectF(pos.x() + r.x(), pos.y() + r.y(), r.width(), r.height())
            event.accept()
        else:
            super().mousePressEvent(event)

    def mouseMoveEvent(self, event: QGraphicsSceneMouseEvent) -> None:
        if not self._dragging:
            super().mouseMoveEvent(event)
            return

        dx = event.scenePos().x() - self._drag_start.x()
        dy = event.scenePos().y() - self._drag_start.y()

        r = QRectF(self._orig_scene_rect)

        if self._index in self._MOVES_LEFT:
            r.setLeft(r.left() + dx)
        if self._index in self._MOVES_RIGHT:
            r.setRight(r.right() + dx)
        if self._index in self._MOVES_TOP:
            r.setTop(r.top() + dy)
        if self._index in self._MOVES_BOTTOM:
            r.setBottom(r.bottom() + dy)

        # normalize handles flip-drag; enforce minimum size
        r = r.normalized()
        if r.width() < _MIN_ITEM_SIZE:
            r.setWidth(_MIN_ITEM_SIZE)
        if r.height() < _MIN_ITEM_SIZE:
            r.setHeight(_MIN_ITEM_SIZE)

        parent: FieldRectItem = self.parentItem()
        parent.setPos(r.topLeft())
        parent.setRect(QRectF(0, 0, r.width(), r.height()))
        event.accept()

    def mouseReleaseEvent(self, event: QGraphicsSceneMouseEvent) -> None:
        if event.button() == Qt.MouseButton.LeftButton and self._dragging:
            self._dragging = False
            event.accept()
        else:
            super().mouseReleaseEvent(event)


class FieldRectItem(QGraphicsRectItem):
    """Draggable, resizable rectangle representing one stamp field on the canvas."""

    _id_dup_resolver = None  # type: ignore[var-annotated]  # Callable[[str], str] | None

    @classmethod
    def set_id_dup_resolver(cls, fn) -> None:
        cls._id_dup_resolver = fn

    def __init__(self, x: float, y: float, w: float, h: float, cell_index: int = 0, parent=None):
        super().__init__(0, 0, w, h, parent)
        self.setPos(x, y)
        self.cell_index: int = cell_index
        self.field_def: FieldDef | None = None
        self.is_assigned: bool = False
        self._test_status: str | None = None  # "ok" | "warning" | "error" | None
        self._highlight_timer: QTimer | None = None
        # Optional callback(item) called on position/size change — used for live
        # bbox display in PropertiesPanel and modified-state tracking in main_window.
        self._geo_cb: object = None
        # Optional callback(new_item, source_item) called when Ctrl+drag duplicates item.
        self._clone_cb: object = None
        self._id_duplicate_highlight: bool = False
        self._wireframe_mode: bool = False

        self.setFlag(QGraphicsItem.GraphicsItemFlag.ItemIsMovable, True)
        self.setFlag(QGraphicsItem.GraphicsItemFlag.ItemIsSelectable, True)
        self.setFlag(QGraphicsItem.GraphicsItemFlag.ItemSendsGeometryChanges, True)
        self.setAcceptHoverEvents(True)

        self._label = QGraphicsSimpleTextItem(str(cell_index), self)
        self._label.setFont(QFont("Arial", 8))
        self._label.setBrush(QBrush(QColor(0, 0, 0)))
        self._update_label_pos()

        self._handles: list[_Handle] = []
        self._create_handles()
        self._update_appearance()

    def assign_field(self, fd: FieldDef) -> None:
        self.field_def = fd
        self.is_assigned = True
        self._test_status = None
        self._id_duplicate_highlight = False
        self._update_appearance()
        self._update_label_text()

    def set_id_duplicate_highlight(self, on: bool) -> None:
        self._id_duplicate_highlight = bool(on)
        self._update_appearance()

    def unassign(self) -> None:
        self.field_def = None
        self.is_assigned = False
        self._test_status = None
        self._id_duplicate_highlight = False
        self._update_appearance()
        self._update_label_text()

    def set_test_status(self, status: str | None) -> None:
        self._test_status = status
        self._update_appearance()

    def set_wireframe_mode(self, enabled: bool) -> None:
        """Switch wireframe color mode (green/yellow by binding completeness)."""
        self._wireframe_mode = enabled
        self._update_appearance()

    def highlight(self, on: bool) -> None:
        if on:
            self.setPen(QPen(_COLOR_HIGHLIGHT, 3, Qt.PenStyle.DashLine))
            if self._highlight_timer is None:
                self._highlight_timer = QTimer()
                self._highlight_timer.setSingleShot(True)
                self._highlight_timer.timeout.connect(lambda: self.highlight(False))
            self._highlight_timer.start(1500)
        else:
            self._update_appearance()

    def _border_color(self) -> QColor:
        if self._id_duplicate_highlight:
            return _COLOR_DUP_ID
        if self._test_status == "ok":
            return _COLOR_OK
        if self._test_status == "warning":
            return _COLOR_WARNING
        if self._test_status == "error":
            return _COLOR_ERROR
        if self.is_assigned:
            if self.field_def and self.field_def.is_anchor:
                return _COLOR_ANCHOR
            if self.field_def and self.field_def.outside_stamp:
                return _COLOR_OUTSIDE_STAMP
            if self.field_def and self.field_def.field_type == "label":
                return _COLOR_LABEL
            return _COLOR_ASSIGNED
        return _COLOR_UNASSIGNED

    def _fill_color(self) -> QColor:
        if not self.is_assigned:
            return _COLOR_FILL_UNASSIGNED
        if self.field_def and self.field_def.outside_stamp:
            return _COLOR_FILL_OUTSIDE_STAMP
        if self._wireframe_mode and self.field_def:
            fd = self.field_def
            bound_count = sum(
                1 for b in (fd.bound_top, fd.bound_bottom, fd.bound_left, fd.bound_right) if b
            )
            return _COLOR_FILL_BOUND_FULL if bound_count == 4 else _COLOR_FILL_BOUND_PARTIAL
        if self.field_def and self.field_def.field_type == "label":
            return _COLOR_FILL_LABEL
        return _COLOR_FILL_ASSIGNED

    def _update_appearance(self) -> None:
        pen = QPen(self._border_color(), 2)
        if self.field_def and (
            self.field_def.field_type == "empty" or self.field_def.outside_stamp
        ) and self._test_status is None:
            pen.setStyle(Qt.PenStyle.DashLine)
        if self.isSelected():
            pen.setWidth(3)
        self.setPen(pen)
        self.setBrush(QBrush(self._fill_color()))

    def _update_label_text(self) -> None:
        if self.field_def:
            self._label.setText(short_label_from_field_id(self.field_def.id))
        else:
            self._label.setText(str(self.cell_index))

    def _update_label_pos(self) -> None:
        r = self.rect()
        self._label.setPos(r.x() + 2, r.y() + 1)

    # ---- handles ----

    def _create_handles(self) -> None:
        # order: TL, TC, TR, RC, BR, BC, BL, LC  (indices 0–7)
        cursors = [
            Qt.CursorShape.SizeFDiagCursor,   # 0 TL
            Qt.CursorShape.SizeVerCursor,      # 1 TC
            Qt.CursorShape.SizeBDiagCursor,    # 2 TR
            Qt.CursorShape.SizeHorCursor,      # 3 RC
            Qt.CursorShape.SizeFDiagCursor,    # 4 BR
            Qt.CursorShape.SizeVerCursor,      # 5 BC
            Qt.CursorShape.SizeBDiagCursor,    # 6 BL
            Qt.CursorShape.SizeHorCursor,      # 7 LC
        ]
        for i, c in enumerate(cursors):
            h = _Handle(i, c, self)
            h.setVisible(False)
            self._handles.append(h)
        self._position_handles()

    def _position_handles(self) -> None:
        r = self.rect()
        positions = [
            QPointF(r.left(), r.top()),
            QPointF(r.center().x(), r.top()),
            QPointF(r.right(), r.top()),
            QPointF(r.right(), r.center().y()),
            QPointF(r.right(), r.bottom()),
            QPointF(r.center().x(), r.bottom()),
            QPointF(r.left(), r.bottom()),
            QPointF(r.left(), r.center().y()),
        ]
        for h, p in zip(self._handles, positions):
            h.setPos(p)

    def _show_handles(self, show: bool) -> None:
        for h in self._handles:
            h.setVisible(show)

    # ---- hover ----

    def hoverEnterEvent(self, event: QGraphicsSceneHoverEvent) -> None:
        self._show_handles(True)
        super().hoverEnterEvent(event)

    def hoverLeaveEvent(self, event: QGraphicsSceneHoverEvent) -> None:
        self._show_handles(False)
        super().hoverLeaveEvent(event)

    def mousePressEvent(self, event: QGraphicsSceneMouseEvent) -> None:
        # Ctrl + drag duplicates the field: original stays, dragged one moves.
        if (
            event.button() == Qt.MouseButton.LeftButton
            and bool(event.modifiers() & Qt.KeyboardModifier.ControlModifier)
        ):
            self._duplicate_for_ctrl_drag()
        super().mousePressEvent(event)

    def _duplicate_for_ctrl_drag(self) -> None:
        scene = self.scene()
        if scene is None:
            return
        pos = self.scenePos()
        r = self.rect()
        next_idx = self._next_cell_index(scene)
        clone = FieldRectItem(pos.x() + r.x(), pos.y() + r.y(), r.width(), r.height(), cell_index=next_idx)
        if self.field_def is not None:
            new_fd = copy.deepcopy(self.field_def)
            fn = FieldRectItem._id_dup_resolver
            if fn is not None:
                new_fd.id = fn(new_fd.id)
            else:
                new_fd.id = _fallback_unique_field_id(new_fd.id, scene)
            clone.assign_field(new_fd)
        scene.addItem(clone)
        clone.setSelected(False)
        # Preserve existing geometry callback integration.
        clone._geo_cb = self._geo_cb
        if self._clone_cb is not None:
            self._clone_cb(clone, self)

    @staticmethod
    def _next_cell_index(scene) -> int:
        max_idx = 0
        for it in scene.items():
            if isinstance(it, FieldRectItem):
                if it.cell_index > max_idx:
                    max_idx = it.cell_index
        return max_idx + 1

    # ---- geometry changes ----

    def itemChange(self, change, value):
        if change == QGraphicsItem.GraphicsItemChange.ItemPositionHasChanged:
            self._position_handles()
            if self._geo_cb is not None:
                self._geo_cb(self)
        if change == QGraphicsItem.GraphicsItemChange.ItemSelectedHasChanged:
            self._update_appearance()
        return super().itemChange(change, value)

    def setRect(self, rect: QRectF) -> None:
        super().setRect(rect)
        self._position_handles()
        self._update_label_pos()
        if self._geo_cb is not None:
            self._geo_cb(self)

    def paint(self, painter: QPainter, option: QStyleOptionGraphicsItem, widget: QWidget | None = None) -> None:
        super().paint(painter, option, widget)
        fd = self.field_def
        if fd and fd.stretch_to_page:
            self._paint_stretch_edges(painter, fd.stretch_to_page)

    def _paint_stretch_edges(self, painter: QPainter, stretch: tuple[str, ...]) -> None:
        """Draw stretch_to_page edges with an orange dashed line overlay."""
        r = self.rect()
        pen = QPen(QColor(255, 120, 0, 180), 2.5, Qt.PenStyle.DotLine)
        painter.setPen(pen)
        for edge in stretch:
            if edge == "bottom":
                painter.drawLine(r.bottomLeft(), r.bottomRight())
            elif edge == "top":
                painter.drawLine(r.topLeft(), r.topRight())
            elif edge == "left":
                painter.drawLine(r.topLeft(), r.bottomLeft())
            elif edge == "right":
                painter.drawLine(r.topRight(), r.bottomRight())


def _fallback_unique_field_id(base_id: str, scene) -> str:
    used = {
        it.field_def.id
        for it in scene.items()
        if isinstance(it, FieldRectItem) and it.field_def
    }
    n = 2
    while True:
        cand = f"{base_id}_{n}"
        if cand not in used:
            return cand
        n += 1
