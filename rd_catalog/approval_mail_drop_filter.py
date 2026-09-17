"""Application-wide drag filter so Outlook/``.msg`` drops hit the mail tab.

Qt only. Does not use Outlook ``Selection``. Installed on ``QApplication``
so drops on Комплекты / heatmap still switch to «Письма о согласовании».
"""

from __future__ import annotations

from collections.abc import Callable

from PySide6.QtCore import QEvent, QObject
from PySide6.QtGui import QCursor
from PySide6.QtWidgets import QApplication, QWidget

from rd_catalog.outlook_drop_qt import mime_has_approval_mail


class ApprovalMailDropFilter(QObject):
    """Accept Outlook/Explorer mail drops over one catalog window."""

    def __init__(
        self,
        window: QWidget,
        on_drop: Callable[[object], None],
    ) -> None:
        """Bind the filter to ``window``.

        Args:
            window: Catalog window that owns the mail tab.
            on_drop: Called with ``QMimeData`` after a successful drop.
        """

        super().__init__(window)
        self._window = window
        self._on_drop = on_drop

    def eventFilter(self, watched: QObject, event: QEvent) -> bool:  # noqa: N802
        """Intercept drag/drop of approval letters over ``window``.

        Args:
            watched: Object that received the event.
            event: Qt event.

        Returns:
            True when the event was a mail drop on this window and is consumed.
        """

        del watched
        event_type = event.type()
        if event_type not in (
            QEvent.Type.DragEnter,
            QEvent.Type.DragMove,
            QEvent.Type.Drop,
        ):
            return False
        if not self._is_over_window():
            return False
        mime = getattr(event, "mimeData", lambda: None)()
        if not mime_has_approval_mail(mime):
            return False
        event.acceptProposedAction()
        if event_type == QEvent.Type.Drop:
            self._on_drop(mime)
        return True

    def _is_over_window(self) -> bool:
        widget = QApplication.widgetAt(QCursor.pos())
        while widget is not None:
            if widget is self._window:
                return True
            widget = widget.parentWidget()
        return False
