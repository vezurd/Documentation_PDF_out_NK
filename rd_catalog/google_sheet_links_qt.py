"""Qt click/cursor bind for Google Sheets cell hrefs."""

from __future__ import annotations

from collections.abc import Callable

from PySide6.QtCore import QEvent, QObject, Qt, QTimer
from PySide6.QtGui import QCursor, QMouseEvent
from PySide6.QtWidgets import QApplication, QTableWidget

from rd_catalog.google_sheet_links import open_google_sheet_url
from rd_catalog.monitor_qt import ROLE_HREF


class _GoogleHrefClickFilter(QObject):
    """Single-click opens Sheets; double-click stays the copy editor."""

    def __init__(
        self,
        table: QTableWidget,
        href_at: Callable[[int, int], str],
        *,
        on_ctrl_click: Callable[[int, int], None] | None = None,
        on_href: Callable[[str], bool] | None = None,
    ) -> None:
        super().__init__(table)
        self._table = table
        self._href_at = href_at
        self._on_ctrl_click = on_ctrl_click
        self._on_href = on_href or open_google_sheet_url
        self._pending_href = ""
        self._open_timer = QTimer(self)
        self._open_timer.setSingleShot(True)
        self._open_timer.timeout.connect(self._flush_pending_href)
        viewport = table.viewport()
        viewport.setMouseTracking(True)
        viewport.installEventFilter(self)

    def eventFilter(self, watched: QObject, event: QEvent) -> bool:
        try:
            table = self._table
            viewport = table.viewport()
        except (RuntimeError, AttributeError):
            return False
        if watched is not viewport:
            return False
        if event.type() == QEvent.Type.MouseMove:
            self._update_cursor(event)
            return False
        if event.type() == QEvent.Type.MouseButtonDblClick:
            self._cancel_pending()
            return False
        if event.type() != QEvent.Type.MouseButtonPress:
            return False
        if not isinstance(event, QMouseEvent):
            return False
        if event.button() != Qt.MouseButton.LeftButton:
            return False
        index = self._table.indexAt(event.position().toPoint())
        if not index.isValid():
            self._cancel_pending()
            return False
        if (
            self._on_ctrl_click is not None
            and event.modifiers() & Qt.KeyboardModifier.ControlModifier
        ):
            self._cancel_pending()
            self._on_ctrl_click(index.row(), index.column())
            return True
        if event.modifiers() & Qt.KeyboardModifier.ControlModifier:
            self._cancel_pending()
            return False
        href = self._href_at(index.row(), index.column())
        if not href:
            self._cancel_pending()
            return False
        self._pending_href = href
        self._open_timer.start(int(QApplication.doubleClickInterval()))
        return False

    def _cancel_pending(self) -> None:
        self._open_timer.stop()
        self._pending_href = ""

    def _flush_pending_href(self) -> None:
        href = self._pending_href
        self._pending_href = ""
        if href:
            self._on_href(href)

    def _update_cursor(self, event: QEvent) -> None:
        if not isinstance(event, QMouseEvent):
            return
        try:
            index = self._table.indexAt(event.position().toPoint())
            href = ""
            if index.isValid():
                href = self._href_at(index.row(), index.column())
            viewport = self._table.viewport()
        except (RuntimeError, AttributeError):
            return
        if href:
            viewport.setCursor(QCursor(Qt.CursorShape.PointingHandCursor))
        else:
            viewport.unsetCursor()


def href_from_table_item(table: QTableWidget, row: int, column: int) -> str:
    """Return ``ROLE_HREF`` stored on a painted table item."""

    if row < 0 or column < 0:
        return ""
    item = table.item(row, column)
    if item is None:
        return ""
    value = item.data(ROLE_HREF)
    return str(value or "")


def attach_google_href_clicks(
    table: QTableWidget,
    href_at: Callable[[int, int], str] | None = None,
    *,
    on_ctrl_click: Callable[[int, int], None] | None = None,
    on_href: Callable[[str], bool] | None = None,
) -> QObject:
    """Install click-to-open-Google on ``table``.

    Args:
        table: Catalog table whose cells may carry a Sheets URL.
        href_at: ``(row, column) → url``. Defaults to ``ROLE_HREF``.
        on_ctrl_click: Optional Ctrl+left handler (Комплекты Подсказки).
        on_href: Override for tests; default opens the system browser.

    Returns:
        Event filter; keep a Python reference for the widget lifetime.

    Single-click opens Sheets after the platform double-click interval.
    The press is not consumed, so double-click still starts the read-only
    copy editor. Ctrl+click still runs ``on_ctrl_click`` when set.
    """

    resolver = href_at or (
        lambda row, column: href_from_table_item(table, row, column)
    )
    return _GoogleHrefClickFilter(
        table,
        resolver,
        on_ctrl_click=on_ctrl_click,
        on_href=on_href,
    )
