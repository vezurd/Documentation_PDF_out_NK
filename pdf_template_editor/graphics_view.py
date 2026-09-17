"""QGraphicsView + QGraphicsScene: PDF-подложка, zoom-to-cursor, pan."""

from __future__ import annotations

from typing import TYPE_CHECKING

import fitz
from PySide6.QtCore import Qt, Signal, QPoint, QPointF, QRectF, QTimer
from PySide6.QtGui import QImage, QPixmap, QWheelEvent, QMouseEvent, QKeyEvent, QPen, QColor
from PySide6.QtWidgets import (
    QGraphicsView,
    QGraphicsScene,
    QGraphicsPixmapItem,
    QGraphicsRectItem,
    QGraphicsLineItem,
)

if TYPE_CHECKING:
    from pdf_parsing_v2_engine.models import FrameInfo

_DPI = 150
_ZOOM_FACTOR = 1.15
_MIN_ZOOM = 0.1
_MAX_ZOOM = 20.0


class StampGraphicsView(QGraphicsView):
    """PDF page viewer with zoom-to-cursor and pan."""

    mouse_scene_pos_changed = Signal(QPointF)
    zoom_changed = Signal(float)
    rect_drawn = Signal(QRectF)
    # Emitted in place_mode on left-click; carries the QRectF of the preview
    # rectangle (scene coords) so the caller knows where to place the field.
    place_confirmed = Signal(QRectF)
    # Two-click placement for template grid lines (+H / +V): orientation "h"|"v",
    # then scene endpoints (same as draw_mode rect corners).
    wire_line_place_finished = Signal(str, QPointF, QPointF)

    def __init__(self, parent=None):
        super().__init__(parent)
        self._scene = QGraphicsScene(self)
        self.setScene(self._scene)

        self._pixmap_item: QGraphicsPixmapItem | None = None
        self._dpi: int = _DPI
        self._page_width_pts: float = 0.0
        self._page_height_pts: float = 0.0
        self._current_zoom: float = 1.0
        self._frame_info: FrameInfo | None = None
        self._space_held: bool = False
        self._fitz_page: fitz.Page | None = None
        # Панорамирование средней кнопкой: вручную (не ScrollHandDrag), иначе поля перехватывают «левый» клик
        self._mid_pan_active: bool = False
        self._mid_pan_last: QPoint | None = None
        self._draw_mode: bool = False
        self._draw_start_scene: QPointF | None = None
        self._draw_preview_item: QGraphicsRectItem | None = None
        # place_mode: single-click to place a field with a preset size
        self._place_mode: bool = False
        self._place_w_px: float = 80.0
        self._place_h_px: float = 24.0
        self._place_preview_item: QGraphicsRectItem | None = None
        # Two-click line placement for wireframe (+H / +V)
        self._wire_line_orient: str | None = None
        self._wire_line_first_scene: QPointF | None = None
        self._wire_line_preview: QGraphicsLineItem | None = None
        self._wire_line_min_scene_dist: float = 3.0

        self.setTransformationAnchor(QGraphicsView.ViewportAnchor.AnchorUnderMouse)
        self.setResizeAnchor(QGraphicsView.ViewportAnchor.AnchorUnderMouse)
        from PySide6.QtGui import QPainter
        self.setRenderHint(QPainter.RenderHint.SmoothPixmapTransform, True)
        self.setDragMode(QGraphicsView.DragMode.RubberBandDrag)
        self.setViewportUpdateMode(QGraphicsView.ViewportUpdateMode.SmartViewportUpdate)

    @property
    def current_fitz_page(self) -> fitz.Page | None:
        return self._fitz_page

    @property
    def frame_info(self) -> FrameInfo | None:
        return self._frame_info

    def load_page(self, fitz_page: fitz.Page, frame_info: FrameInfo | None = None) -> None:
        self._clear_draw_preview()
        self.set_wire_line_place_mode(None)
        self._fitz_page = fitz_page
        self._frame_info = frame_info
        self._page_width_pts = fitz_page.rect.width
        self._page_height_pts = fitz_page.rect.height

        pix = fitz_page.get_pixmap(dpi=self._dpi)
        fmt = QImage.Format.Format_RGB888 if pix.n == 3 else QImage.Format.Format_RGBA8888
        qimg = QImage(pix.samples, pix.width, pix.height, pix.stride, fmt)
        qpix = QPixmap.fromImage(qimg)

        if self._pixmap_item is not None:
            self._scene.removeItem(self._pixmap_item)
        self._pixmap_item = QGraphicsPixmapItem(qpix)
        self._pixmap_item.setZValue(-1000)
        self._scene.addItem(self._pixmap_item)
        self._scene.setSceneRect(QRectF(qpix.rect()))
        self.resetTransform()
        self._current_zoom = 1.0
        self.zoom_changed.emit(self._current_zoom)
        # После layout: основная работа у штампа (правый нижний угол рамки / страницы)
        QTimer.singleShot(0, self._scroll_to_stamp_area)

    def _scroll_to_stamp_area(self) -> None:
        """Прокрутить вид к правому нижнему углу рамки (штамп) или страницы."""
        fi = self._frame_info
        if fi is not None:
            # Нижний правый угол рамки в fitz displayed (y вниз): x=x1, y=page_height−y0
            pt = self.pts_to_scene(QPointF(fi.x1, fi.page_height - fi.y0))
        else:
            r = self._scene.sceneRect()
            pt = QPointF(r.right(), r.bottom())
        self.centerOn(pt)

    def clear_page(self) -> None:
        self._clear_draw_preview()
        self._clear_place_preview()
        self.set_wire_line_place_mode(None)
        if self._pixmap_item is not None:
            self._scene.removeItem(self._pixmap_item)
            self._pixmap_item = None
        self._fitz_page = None
        self._frame_info = None

    def set_draw_mode(self, enabled: bool) -> None:
        self._draw_mode = bool(enabled)
        self._clear_draw_preview()
        if self._draw_mode:
            self.setDragMode(QGraphicsView.DragMode.NoDrag)
            self.viewport().setCursor(Qt.CursorShape.CrossCursor)
        else:
            if self._space_held:
                self.setDragMode(QGraphicsView.DragMode.ScrollHandDrag)
                self.viewport().setCursor(Qt.CursorShape.OpenHandCursor)
            else:
                self.setDragMode(QGraphicsView.DragMode.RubberBandDrag)
                self.viewport().unsetCursor()

    def set_place_mode(
        self,
        enabled: bool,
        preview_w_px: float = 80.0,
        preview_h_px: float = 24.0,
    ) -> None:
        """Enable/disable place-mode: click once to place a field with preset size.

        While active the cursor is a crosshair and a dashed preview rectangle
        follows the mouse. On left-click `place_confirmed(QRectF)` is emitted
        with scene coords of the preview, then place-mode exits automatically.
        Escape cancels without placing.
        """
        self._place_mode = bool(enabled)
        self._place_w_px = float(preview_w_px)
        self._place_h_px = float(preview_h_px)
        self._clear_place_preview()
        if self._place_mode:
            # place_mode is mutually exclusive with draw_mode
            self._draw_mode = False
            self._clear_draw_preview()
            self.setDragMode(QGraphicsView.DragMode.NoDrag)
            self.viewport().setCursor(Qt.CursorShape.CrossCursor)
        else:
            if self._space_held:
                self.setDragMode(QGraphicsView.DragMode.ScrollHandDrag)
                self.viewport().setCursor(Qt.CursorShape.OpenHandCursor)
            else:
                self.setDragMode(QGraphicsView.DragMode.RubberBandDrag)
                self.viewport().unsetCursor()

    def set_wire_line_place_mode(self, orientation: str | None) -> None:
        """Two-click mode for +H/+V grid lines: first click start, second end.

        Pass ``None`` to cancel. Mutually exclusive with draw_mode and place_mode.
        """
        self._wire_line_orient = orientation if orientation in ("h", "v") else None
        self._wire_line_first_scene = None
        self._clear_wire_line_preview()
        if self._wire_line_orient:
            self._draw_mode = False
            self._clear_draw_preview()
            self._place_mode = False
            self._clear_place_preview()
            self.setDragMode(QGraphicsView.DragMode.NoDrag)
            self.viewport().setCursor(Qt.CursorShape.CrossCursor)
        else:
            if self._space_held:
                self.setDragMode(QGraphicsView.DragMode.ScrollHandDrag)
                self.viewport().setCursor(Qt.CursorShape.OpenHandCursor)
            else:
                self.setDragMode(QGraphicsView.DragMode.RubberBandDrag)
                self.viewport().unsetCursor()

    def _clear_wire_line_preview(self) -> None:
        if self._wire_line_preview is not None:
            if self._wire_line_preview.scene() is self._scene:
                self._scene.removeItem(self._wire_line_preview)
            self._wire_line_preview = None

    def _clear_place_preview(self) -> None:
        if self._place_preview_item is not None:
            if self._place_preview_item.scene() is self._scene:
                self._scene.removeItem(self._place_preview_item)
            self._place_preview_item = None

    def _update_place_preview(self, scene_pos: QPointF) -> None:
        """Move or create the place-mode ghost rectangle around scene_pos."""
        x = scene_pos.x() - self._place_w_px / 2
        y = scene_pos.y() - self._place_h_px / 2
        rect = QRectF(x, y, self._place_w_px, self._place_h_px)
        if self._place_preview_item is None:
            self._place_preview_item = QGraphicsRectItem(rect)
            self._place_preview_item.setPen(
                QPen(QColor(50, 140, 230, 210), 1.5, Qt.PenStyle.DashLine)
            )
            self._place_preview_item.setBrush(
                QColor(50, 140, 230, 30)
            )
            self._place_preview_item.setZValue(985)
            self._scene.addItem(self._place_preview_item)
        else:
            self._place_preview_item.setRect(rect)

    def _clear_draw_preview(self) -> None:
        self._draw_start_scene = None
        if self._draw_preview_item is not None:
            if self._draw_preview_item.scene() is self._scene:
                self._scene.removeItem(self._draw_preview_item)
            self._draw_preview_item = None

    # ---- coordinate helpers ----

    def scene_to_pts(self, scene_pos: QPointF) -> QPointF:
        """Convert scene (pixel) position to PDF points."""
        if self._page_width_pts == 0:
            return QPointF(0, 0)
        sx = self._page_width_pts / (self._dpi / 72.0 * self._page_width_pts / self._page_width_pts)
        scale = 72.0 / self._dpi
        return QPointF(scene_pos.x() * scale, scene_pos.y() * scale)

    def pts_to_scene(self, pts: QPointF) -> QPointF:
        scale = self._dpi / 72.0
        return QPointF(pts.x() * scale, pts.y() * scale)

    def scene_to_mm_from_frame(self, scene_pos: QPointF) -> tuple[float, float] | None:
        """Convert scene position to mm from frame bottom-right (template coords)."""
        if self._frame_info is None:
            return None
        from pdf_parsing_v2_engine.frame_detector import SCALE
        pts = self.scene_to_pts(scene_pos)
        fi = self._frame_info
        fitz_x, fitz_y = pts.x(), pts.y()
        pm_x = fitz_x
        pm_y = fi.page_height - fitz_y
        mm_from_right = (fi.x1 - pm_x) / SCALE
        mm_from_bottom = (pm_y - fi.y0) / SCALE
        return (mm_from_right, mm_from_bottom)

    # ---- zoom ----

    def wheelEvent(self, event: QWheelEvent) -> None:
        if event.modifiers() & Qt.KeyboardModifier.ControlModifier:
            super().wheelEvent(event)
            return
        angle = event.angleDelta().y()
        if angle == 0:
            return
        factor = _ZOOM_FACTOR if angle > 0 else 1.0 / _ZOOM_FACTOR
        new_zoom = self._current_zoom * factor
        if new_zoom < _MIN_ZOOM or new_zoom > _MAX_ZOOM:
            return
        self.scale(factor, factor)
        self._current_zoom = new_zoom
        self.zoom_changed.emit(self._current_zoom)

    # ---- pan with space or middle mouse ----

    def keyPressEvent(self, event: QKeyEvent) -> None:
        if event.key() == Qt.Key.Key_Escape and self._wire_line_orient:
            self.set_wire_line_place_mode(None)
            event.accept()
            return
        if event.key() == Qt.Key.Key_Escape and self._place_mode:
            self.set_place_mode(False)
            event.accept()
            return
        if event.key() == Qt.Key.Key_Space and not event.isAutoRepeat():
            self._space_held = True
            self.setDragMode(QGraphicsView.DragMode.ScrollHandDrag)
            self.viewport().setCursor(Qt.CursorShape.OpenHandCursor)
            return
        super().keyPressEvent(event)

    def keyReleaseEvent(self, event: QKeyEvent) -> None:
        if event.key() == Qt.Key.Key_Space and not event.isAutoRepeat():
            self._space_held = False
            self.setDragMode(QGraphicsView.DragMode.RubberBandDrag)
            self.viewport().unsetCursor()
            return
        super().keyReleaseEvent(event)

    def mousePressEvent(self, event: QMouseEvent) -> None:
        if event.button() == Qt.MouseButton.MiddleButton:
            event.accept()
            self._mid_pan_active = True
            self._mid_pan_last = event.position().toPoint()
            self.viewport().grabMouse()
            self.viewport().setCursor(Qt.CursorShape.ClosedHandCursor)
            return
        if self._wire_line_orient and event.button() == Qt.MouseButton.LeftButton:
            event.accept()
            scene_pos = self.mapToScene(event.position().toPoint())
            if self._wire_line_first_scene is None:
                self._wire_line_first_scene = scene_pos
                self._wire_line_preview = QGraphicsLineItem(
                    scene_pos.x(), scene_pos.y(), scene_pos.x(), scene_pos.y(),
                )
                self._wire_line_preview.setPen(
                    QPen(QColor(80, 200, 120, 220), 1.5, Qt.PenStyle.DashLine),
                )
                self._wire_line_preview.setZValue(986)
                self._scene.addItem(self._wire_line_preview)
                return
            first = self._wire_line_first_scene
            dist = ((scene_pos.x() - first.x()) ** 2 + (scene_pos.y() - first.y()) ** 2) ** 0.5
            orient = self._wire_line_orient
            self.set_wire_line_place_mode(None)
            if dist >= self._wire_line_min_scene_dist:
                self.wire_line_place_finished.emit(orient, first, scene_pos)
            return
        if self._place_mode and event.button() == Qt.MouseButton.LeftButton:
            event.accept()
            scene_pos = self.mapToScene(event.position().toPoint())
            x = scene_pos.x() - self._place_w_px / 2
            y = scene_pos.y() - self._place_h_px / 2
            placed_rect = QRectF(x, y, self._place_w_px, self._place_h_px)
            self.set_place_mode(False)
            self.place_confirmed.emit(placed_rect)
            return
        if self._draw_mode and event.button() == Qt.MouseButton.LeftButton:
            event.accept()
            self._draw_start_scene = self.mapToScene(event.position().toPoint())
            self._clear_draw_preview()
            self._draw_start_scene = self.mapToScene(event.position().toPoint())
            self._draw_preview_item = QGraphicsRectItem(QRectF(self._draw_start_scene, self._draw_start_scene))
            self._draw_preview_item.setPen(QPen(QColor(70, 130, 220, 200), 1.5, Qt.PenStyle.DashLine))
            self._draw_preview_item.setBrush(Qt.BrushStyle.NoBrush)
            self._draw_preview_item.setZValue(980)
            self._scene.addItem(self._draw_preview_item)
            return
        super().mousePressEvent(event)

    def mouseReleaseEvent(self, event: QMouseEvent) -> None:
        if event.button() == Qt.MouseButton.MiddleButton and self._mid_pan_active:
            event.accept()
            self._mid_pan_active = False
            self._mid_pan_last = None
            self.viewport().releaseMouse()
            if not self._space_held:
                self.viewport().unsetCursor()
                self.setDragMode(QGraphicsView.DragMode.RubberBandDrag)
            else:
                self.viewport().setCursor(Qt.CursorShape.OpenHandCursor)
            return
        if self._draw_mode and event.button() == Qt.MouseButton.LeftButton:
            event.accept()
            start = self._draw_start_scene
            end = self.mapToScene(event.position().toPoint())
            self._clear_draw_preview()
            if start is None:
                return
            rect = QRectF(start, end).normalized()
            if rect.width() >= 4.0 and rect.height() >= 4.0:
                self.rect_drawn.emit(rect)
            return
        super().mouseReleaseEvent(event)

    def mouseMoveEvent(self, event: QMouseEvent) -> None:
        scene_pos = self.mapToScene(event.position().toPoint())
        self.mouse_scene_pos_changed.emit(scene_pos)
        if self._place_mode:
            self._update_place_preview(scene_pos)
            event.accept()
            return
        if (
            self._wire_line_orient
            and self._wire_line_first_scene is not None
            and self._wire_line_preview is not None
        ):
            self._wire_line_preview.setLine(
                self._wire_line_first_scene.x(),
                self._wire_line_first_scene.y(),
                scene_pos.x(),
                scene_pos.y(),
            )
            event.accept()
            return
        if self._draw_mode and self._draw_start_scene is not None and self._draw_preview_item is not None:
            rect = QRectF(self._draw_start_scene, scene_pos).normalized()
            self._draw_preview_item.setRect(rect)
            event.accept()
            return
        if self._mid_pan_active and self._mid_pan_last is not None:
            cur = event.position().toPoint()
            delta = cur - self._mid_pan_last
            self._mid_pan_last = cur
            h = self.horizontalScrollBar()
            v = self.verticalScrollBar()
            h.setValue(h.value() - delta.x())
            v.setValue(v.value() - delta.y())
            event.accept()
            return
        super().mouseMoveEvent(event)
