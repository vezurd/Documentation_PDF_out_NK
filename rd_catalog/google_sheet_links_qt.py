"""Qt click/cursor bind for Google Sheets cell hrefs."""

from __future__ import annotations

from collections.abc import Callable

from PySide6.QtCore import QEvent, QObject, Qt
from PySide6.QtGui import QCursor, QMouseEvent
from PySide6.QtWidgets import QTableWidget

from rd_catalog.google_sheet_links import open_google_sheet_url
from rd_catalog.monitor_qt import ROLE_HREF


class _GoogleHrefClickFilter(QObject):
    """Shift+click opens Sheets; other clicks stay copy/select."""

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
        if event.type() != QEvent.Type.MouseButtonPress:
            return False
        if not isinstance(event, QMouseEvent):
            return False
        if event.button() != Qt.MouseButton.LeftButton:
            return False
        index = self._table.indexAt(event.position().toPoint())
        if not index.isValid():
            return False
        modifiers = event.modifiers()
        if (
            self._on_ctrl_click is not None
            and modifiers & Qt.KeyboardModifier.ControlModifier
        ):
            self._on_ctrl_click(index.row(), index.column())
            return True
        if modifiers & Qt.KeyboardModifier.ControlModifier:
            return False
        if not (modifiers & Qt.KeyboardModifier.ShiftModifier):
            return False
        href = self._href_at(index.row(), index.column())
        if not href:
            return False
        self._table.selectRow(index.row())
        self._table.setCurrentCell(index.row(), index.column())
        self._on_href(href)
        return True

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
    """Install Shift+click-to-open-Google on ``table``.

    Args:
        table: Catalog table whose cells may carry a Sheets URL.
        href_at: ``(row, column) → url``. Defaults to ``ROLE_HREF``.
        on_ctrl_click: Optional Ctrl+left handler (Комплекты Подсказки).
        on_href: Override for tests; default opens the system browser.

    Returns:
        Event filter; keep a Python reference for the widget lifetime.

    Plain click and double-click are not consumed (copy editor / select).
    Shift+left-click opens Sheets. Ctrl+click still runs ``on_ctrl_click``.
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
