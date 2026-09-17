"""Interactive grid-line items on the stamp editor canvas.

GridLineGraphicsItem: a QGraphicsLineItem with two endpoint handles.
Handles are constrained to one axis (H-line endpoints move along X,
V-line endpoints move along Y). The line itself can be dragged along its
perpendicular axis (H-line: up/down, V-line: left/right).

pick_line_mode helpers: dim fields, highlight lines on hover.
"""

from __future__ import annotations

from typing import Callable

from PySide6.QtCore import Qt, QPointF, QLineF, Signal, QObject
from PySide6.QtGui import (
    QPen,
    QColor,
    QBrush,
    QFont,
    QPainter,
)
from PySide6.QtWidgets import (
    QGraphicsItem,
    QGraphicsLineItem,
    QGraphicsRectItem,
    QGraphicsSimpleTextItem,
    QGraphicsSceneMouseEvent,
    QGraphicsSceneHoverEvent,
    QStyleOptionGraphicsItem,
    QWidget,
)

_HANDLE_SIZE = 7.0
_Z_LINE = 500
_Z_HANDLE = 501
_Z_LABEL = 502

_COLOR_BOUNDARY = {
    "top": QColor(0, 120, 255),       # bright blue
    "bottom": QColor(16, 185, 80),     # vivid green
    "left": QColor(245, 130, 0),       # bright orange
    "right": QColor(220, 30, 60),      # crimson
}
_COLOR_INTERNAL = QColor(160, 50, 220)          # vivid purple — high contrast on B&W
_COLOR_SELECTED = QColor(255, 50, 180)          # hot pink
_COLOR_FIELD_BOUND = QColor(255, 180, 0)        # warm amber
_COLOR_HOVER = QColor(0, 210, 230)              # bright cyan


class _EndpointHandle(QGraphicsRectItem):
    """Small draggable square at one endpoint of a grid line.

    index 0 = start endpoint, 1 = end endpoint.
    For H-lines, handles drag along X only.
    For V-lines, handles drag along Y only.
    """

    def __init__(self, index: int, parent: "GridLineGraphicsItem"):
        hs = _HANDLE_SIZE
        super().__init__(-hs / 2, -hs / 2, hs, hs, parent)
        self._index = index
        self._dragging = False
        self._drag_start_scene = QPointF()
        self._orig_pos = QPointF()

        self.setBrush(QBrush(QColor(255, 255, 255)))
        self.setPen(QPen(QColor(0, 0, 0), 1))
        self.setFlag(QGraphicsItem.GraphicsItemFlag.ItemIsMovable, False)
        self.setFlag(QGraphicsItem.GraphicsItemFlag.ItemSendsGeometryChanges, False)
        self.setAcceptHoverEvents(True)
        self.setZValue(_Z_HANDLE)
        self.setVisible(False)

        orient = parent.orientation
        if orient == "h":
            self.setCursor(Qt.CursorShape.SizeHorCursor)
        else:
            self.setCursor(Qt.CursorShape.SizeVerCursor)

    def hoverEnterEvent(self, event: QGraphicsSceneHoverEvent) -> None:
        self.setBrush(QBrush(QColor(120, 200, 255)))
        super().hoverEnterEvent(event)

    def hoverLeaveEvent(self, event: QGraphicsSceneHoverEvent) -> None:
        self.setBrush(QBrush(QColor(255, 255, 255)))
        super().hoverLeaveEvent(event)

    def mousePressEvent(self, event: QGraphicsSceneMouseEvent) -> None:
        if event.button() == Qt.MouseButton.LeftButton:
            self._dragging = True
            self._drag_start_scene = event.scenePos()
            self._orig_pos = self.scenePos()
            event.accept()
        else:
            super().mousePressEvent(event)

    def mouseMoveEvent(self, event: QGraphicsSceneMouseEvent) -> None:
        if not self._dragging:
            super().mouseMoveEvent(event)
            return
        parent: GridLineGraphicsItem = self.parentItem()
        delta = event.scenePos() - self._drag_start_scene
        if parent.orientation == "h":
            new_scene_x = self._orig_pos.x() + delta.x()
            parent._on_handle_dragged(self._index, new_scene_x)
        else:
            new_scene_y = self._orig_pos.y() + delta.y()
            parent._on_handle_dragged(self._index, new_scene_y)
        event.accept()

    def mouseReleaseEvent(self, event: QGraphicsSceneMouseEvent) -> None:
        if event.button() == Qt.MouseButton.LeftButton and self._dragging:
            self._dragging = False
            parent: GridLineGraphicsItem = self.parentItem()
            parent._on_handle_released()
            event.accept()
        else:
            super().mouseReleaseEvent(event)


class GridLineGraphicsItem(QGraphicsLineItem):
    """Interactive grid line on the canvas with two endpoint handles.

    Attributes:
        line_id: TemplateGridLine.id
        orientation: "h" or "v"
        boundary: "top"/"bottom"/"left"/"right" or None
        geo_changed_cb: ``callable(item, commit, *, reassign_refs=False)`` —
            *commit* False: line-body drag in progress (light update).
            *commit* True + *reassign_refs* True: endpoint handle released — rerun
            auto intersection refs then snap spans.
            *commit* True + *reassign_refs* False: line body released / deselect.
    """

    def __init__(
        self,
        line_id: str,
        orientation: str,
        boundary: str | None,
        x1: float, y1: float, x2: float, y2: float,
        parent=None,
    ):
        super().__init__(x1, y1, x2, y2, parent)
        self.line_id = line_id
        self.orientation = orientation
        self.boundary = boundary
        self.geo_changed_cb: Callable | None = None
        self._is_highlighted = False
        self._body_drag_active = False
        self._body_drag_moved = False

        self.setFlag(QGraphicsItem.GraphicsItemFlag.ItemIsSelectable, True)
        self.setFlag(QGraphicsItem.GraphicsItemFlag.ItemIsMovable, True)
        self.setFlag(QGraphicsItem.GraphicsItemFlag.ItemSendsGeometryChanges, True)
        self.setAcceptHoverEvents(True)
        self.setZValue(_Z_LINE)

        self._handles = [
            _EndpointHandle(0, self),
            _EndpointHandle(1, self),
        ]
        self._label = QGraphicsSimpleTextItem(line_id, self)
        self._label.setFont(QFont("Consolas", 7))
        self._label.setZValue(_Z_LABEL)

        self._update_appearance()
        self._position_handles()
        self._position_label()

    # ---- appearance ----

    def _update_appearance(self) -> None:
        if self.isSelected():
            color = _COLOR_SELECTED
            width = 2.5
            style = Qt.PenStyle.SolidLine
        elif self._is_highlighted:
            color = _COLOR_FIELD_BOUND
            width = 2.0
            style = Qt.PenStyle.SolidLine
        elif self.boundary:
            color = _COLOR_BOUNDARY.get(self.boundary, _COLOR_INTERNAL)
            width = 2.0
            style = Qt.PenStyle.SolidLine
        else:
            color = _COLOR_INTERNAL
            width = 1.5
            style = Qt.PenStyle.DashDotLine

        pen = QPen(color, width)
        pen.setStyle(style)
        self.setPen(pen)
        self._label.setBrush(QBrush(color))
        show_handles = self.isSelected()
        for h in self._handles:
            h.setVisible(show_handles)

    def set_highlighted(self, on: bool) -> None:
        self._is_highlighted = on
        self._update_appearance()

    # ---- handle positioning ----

    def _position_handles(self) -> None:
        line = self.line()
        self._handles[0].setPos(line.p1())
        self._handles[1].setPos(line.p2())

    def _position_label(self) -> None:
        line = self.line()
        if self.orientation == "h":
            self._label.setPos(line.p1().x() + 2, line.p1().y() - 12)
        else:
            self._label.setPos(line.p1().x() + 2, line.p1().y() + 2)

    # ---- handle drag callback ----

    def _on_handle_dragged(self, index: int, new_coord: float) -> None:
        """Called by handle during drag. Constrains to one axis."""
        line = self.line()
        sp = self.scenePos()
        if self.orientation == "h":
            local_x = new_coord - sp.x()
            if index == 0:
                self.setLine(QLineF(local_x, line.y1(), line.x2(), line.y2()))
            else:
                self.setLine(QLineF(line.x1(), line.y1(), local_x, line.y2()))
        else:
            local_y = new_coord - sp.y()
            if index == 0:
                self.setLine(QLineF(line.x1(), local_y, line.x2(), line.y2()))
            else:
                self.setLine(QLineF(line.x1(), line.y1(), line.x2(), local_y))
        self._position_handles()
        self._position_label()

    def _on_handle_released(self) -> None:
        """Called when handle drag finishes."""
        if self.geo_changed_cb:
            self.geo_changed_cb(self, True, reassign_refs=True)

    # ---- line body drag (perpendicular axis only) ----

    def mousePressEvent(self, event: QGraphicsSceneMouseEvent) -> None:
        if event.button() == Qt.MouseButton.LeftButton:
            self._body_drag_active = True
            self._body_drag_moved = False
        super().mousePressEvent(event)

    def mouseReleaseEvent(self, event: QGraphicsSceneMouseEvent) -> None:
        if event.button() == Qt.MouseButton.LeftButton and self._body_drag_active:
            self._body_drag_active = False
            if self._body_drag_moved and self.geo_changed_cb:
                self.geo_changed_cb(self, True)
        super().mouseReleaseEvent(event)

    def itemChange(self, change, value):
        if change == QGraphicsItem.GraphicsItemChange.ItemPositionChange:
            old_pos = self.pos()
            new_pos = value
            if self.orientation == "h":
                return QPointF(old_pos.x(), new_pos.y())
            else:
                return QPointF(new_pos.x(), old_pos.y())
        if change == QGraphicsItem.GraphicsItemChange.ItemPositionHasChanged:
            self._position_label()
            if self.geo_changed_cb and self._body_drag_active:
                self._body_drag_moved = True
                self.geo_changed_cb(self, False)
        if change == QGraphicsItem.GraphicsItemChange.ItemSelectedHasChanged:
            if not value and self._body_drag_moved and self.geo_changed_cb:
                self.geo_changed_cb(self, True)
            if not value:
                self._body_drag_active = False
                self._body_drag_moved = False
            self._update_appearance()
        return super().itemChange(change, value)

    # ---- hover ----

    def hoverEnterEvent(self, event: QGraphicsSceneHoverEvent) -> None:
        if not self.isSelected():
            pen = self.pen()
            pen.setColor(_COLOR_HOVER)
            pen.setWidth(max(pen.width(), 2))
            self.setPen(pen)
        super().hoverEnterEvent(event)

    def hoverLeaveEvent(self, event: QGraphicsSceneHoverEvent) -> None:
        self._update_appearance()
        super().hoverLeaveEvent(event)

    # ---- scene coords for mm conversion ----

    def scene_endpoints(self) -> tuple[float, float, float, float]:
        """Return (x1, y1, x2, y2) in scene coordinates."""
        line = self.line()
        sp = self.scenePos()
        return (
            sp.x() + line.x1(), sp.y() + line.y1(),
            sp.x() + line.x2(), sp.y() + line.y2(),
        )


# ---------------------------------------------------------------------------
#  DetectedGridLineItem — read-only magenta item for PDF-detected grid lines
# ---------------------------------------------------------------------------

_COLOR_DETECTED = QColor(255, 0, 255, 200)
_COLOR_DETECTED_SEL = QColor(255, 80, 0, 255)
_Z_DETECTED = 903


class DetectedGridLineItem(QGraphicsLineItem):
    """Read-only, selectable item representing a detected PDF grid line."""

    def __init__(
        self,
        line_id: str,
        orientation: str,
        pos_pts: float,
        span_lo_pts: float,
        span_hi_pts: float,
        cell_count: int,
        dpi_scale: float,
    ) -> None:
        self.line_id = line_id
        self.orientation = orientation
        self.pos_pts = pos_pts
        self.span_lo_pts = span_lo_pts
        self.span_hi_pts = span_hi_pts
        self.cell_count = cell_count
        self.matched_tpl_id: str = ""

        if orientation == "h":
            super().__init__(
                span_lo_pts * dpi_scale, pos_pts * dpi_scale,
                span_hi_pts * dpi_scale, pos_pts * dpi_scale,
            )
        else:
            super().__init__(
                pos_pts * dpi_scale, span_lo_pts * dpi_scale,
                pos_pts * dpi_scale, span_hi_pts * dpi_scale,
            )

        self.setFlags(
            QGraphicsItem.GraphicsItemFlag.ItemIsSelectable
        )
        self.setAcceptHoverEvents(True)
        self.setZValue(_Z_DETECTED)
        self.setToolTip(
            f"{line_id} | {orientation} pos={pos_pts:.1f}pts "
            f"span=[{span_lo_pts:.0f}..{span_hi_pts:.0f}] cells={cell_count}"
        )
        self._update_pen()

    def _update_pen(self) -> None:
        if self.isSelected():
            self.setPen(QPen(_COLOR_DETECTED_SEL, 2.5, Qt.PenStyle.SolidLine))
        else:
            self.setPen(QPen(_COLOR_DETECTED, 1.5, Qt.PenStyle.SolidLine))

    def itemChange(self, change, value):
        if change == QGraphicsItem.GraphicsItemChange.ItemSelectedHasChanged:
            self._update_pen()
        return super().itemChange(change, value)

    def hoverEnterEvent(self, event: QGraphicsSceneHoverEvent) -> None:
        self.setPen(QPen(_COLOR_HOVER, 2.5, Qt.PenStyle.SolidLine))
        super().hoverEnterEvent(event)

    def hoverLeaveEvent(self, event: QGraphicsSceneHoverEvent) -> None:
        self._update_pen()
        super().hoverLeaveEvent(event)
