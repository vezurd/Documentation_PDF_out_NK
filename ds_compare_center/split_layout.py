"""Shared left/right (controls + Job monitor) layout for control-center tabs."""

from __future__ import annotations

from collections.abc import Callable
from math import ceil

from PySide6.QtCore import QSize, Qt, QTimer, QUrl, Signal
from PySide6.QtGui import QResizeEvent, QShowEvent, QTextDocument, QTextOption, QWheelEvent
from PySide6.QtWidgets import (
    QFrame,
    QLabel,
    QScrollArea,
    QSizePolicy,
    QSplitter,
    QTextBrowser,
    QVBoxLayout,
    QWidget,
)

from main_v2.job_monitor import JobMonitorPanel

_LEFT_WIDTH = 520
_CONSOLE_WIDTH = 770
_CONSOLE_MIN_WIDTH = 420
_LEFT_MIN_WIDTH = 240


def shrink_h(widget: QWidget) -> QWidget:
    """Allow widget to follow a narrow splitter pane (ignore content-based width)."""
    widget.setMinimumWidth(0)
    policy = widget.sizePolicy()
    policy.setHorizontalPolicy(QSizePolicy.Policy.Expanding)
    widget.setSizePolicy(policy)
    return widget


class WrappingLabel(QLabel):
    """QLabel whose height follows wrapped text when the pane width changes."""

    def __init__(
        self,
        text: str = "",
        parent: QWidget | None = None,
        *,
        selectable: bool = True,
    ) -> None:
        super().__init__(text, parent)
        self.setWordWrap(True)
        self.setMinimumWidth(0)
        if selectable:
            self.setTextInteractionFlags(
                Qt.TextInteractionFlag.TextSelectableByMouse
                | Qt.TextInteractionFlag.TextSelectableByKeyboard
            )
        policy = QSizePolicy(
            QSizePolicy.Policy.Expanding,
            QSizePolicy.Policy.Preferred,
        )
        policy.setHeightForWidth(True)
        self.setSizePolicy(policy)

    def hasHeightForWidth(self) -> bool:
        return True

    def heightForWidth(self, width: int) -> int:
        if width <= 0:
            return super().sizeHint().height()
        return super().heightForWidth(width)

    def sizeHint(self) -> QSize:
        width = self.width()
        if width <= 0:
            width = super().sizeHint().width()
        return QSize(width, self.heightForWidth(width))

    def minimumSizeHint(self) -> QSize:
        return QSize(0, self.heightForWidth(self.width() or 1))

    def setText(self, text: str) -> None:
        super().setText(text)
        self.updateGeometry()

    def resizeEvent(self, event: QResizeEvent) -> None:
        super().resizeEvent(event)
        # Width changed → wrap height may change; notify parent layouts.
        self.updateGeometry()


class WrappingPlainText(QTextBrowser):
    """Read-only text that wraps long paths, stays selectable, optional file links.

    Unlike ``QLabel``, can break inside ``C:\\Users\\...`` tokens.
    Height is recomputed after wrap/resize so all lines stay visible (no inner scroll).
    Clicking an ``<a href="file:...">`` emits ``path_link_activated`` with a local path.
    """

    path_link_activated = Signal(str)

    def __init__(
        self,
        text: str = "",
        parent: QWidget | None = None,
    ) -> None:
        super().__init__(parent)
        self._fit_scheduled = False
        self.setReadOnly(True)
        self.setAcceptRichText(True)
        self.setUndoRedoEnabled(False)
        self.setOpenLinks(False)
        self.setOpenExternalLinks(False)
        self.setFrameShape(QFrame.Shape.NoFrame)
        self.setFrameShadow(QFrame.Shadow.Plain)
        self.setLineWidth(0)
        self.setMidLineWidth(0)
        self.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.setLineWrapMode(QTextBrowser.LineWrapMode.WidgetWidth)
        self.setWordWrapMode(QTextOption.WrapMode.WrapAnywhere)
        self.setMinimumWidth(0)
        self.setViewportMargins(0, 0, 0, 0)
        self.document().setDocumentMargin(0)
        self.setStyleSheet(
            "QTextBrowser { background: transparent; border: none; outline: none; padding: 0; }"
            "QTextBrowser a { color: #0645ad; text-decoration: underline; }"
        )
        policy = QSizePolicy(
            QSizePolicy.Policy.Expanding,
            QSizePolicy.Policy.Minimum,
        )
        policy.setHeightForWidth(True)
        self.setSizePolicy(policy)
        self.document().contentsChanged.connect(self._schedule_fit)
        self.anchorClicked.connect(self._on_anchor_clicked)
        if text:
            self.setPlainText(text)

    def hasHeightForWidth(self) -> bool:
        return True

    def heightForWidth(self, width: int) -> int:
        return self._document_height_for_width(width)

    def sizeHint(self) -> QSize:
        width = self.viewport().width() or self.width() or 200
        return QSize(width, self._document_height_for_width(width))

    def minimumSizeHint(self) -> QSize:
        width = max(self.viewport().width() or self.width(), 1)
        return QSize(0, self._document_height_for_width(width))

    def setText(self, text: str) -> None:
        """Match ``QLabel.setText`` API used by panels (plain text)."""
        self.setPlainText(text)
        self._schedule_fit()

    def setHtmlText(self, html: str) -> None:
        """Set HTML (e.g. with ``file:`` links) and refit height."""
        self.setHtml(html)
        self._schedule_fit()

    def showEvent(self, event: QShowEvent) -> None:
        super().showEvent(event)
        self._schedule_fit()

    def resizeEvent(self, event: QResizeEvent) -> None:
        super().resizeEvent(event)
        self._schedule_fit()

    def wheelEvent(self, event: QWheelEvent) -> None:
        # Do not scroll inside the field; let the parent QScrollArea handle it.
        event.ignore()

    def _on_anchor_clicked(self, url: QUrl) -> None:
        local = url.toLocalFile()
        if not local:
            local = url.toString(QUrl.UrlFormattingOption.PreferLocalFile)
        if local:
            self.path_link_activated.emit(local)

    def _schedule_fit(self) -> None:
        """Defer fit until after layout/style polish (width and font are final)."""
        if self._fit_scheduled:
            return
        self._fit_scheduled = True
        QTimer.singleShot(0, self._fit_to_document)

    def _document_height_for_width(self, width: int) -> int:
        if width <= 0:
            return max(int(self.fontMetrics().height()), 1)
        self.ensurePolished()
        doc = QTextDocument()
        # Match the font actually used by this widget (incl. stylesheet).
        doc.setDefaultFont(self.document().defaultFont())
        doc.setDocumentMargin(0)
        doc.setPlainText(self.toPlainText())
        option = doc.defaultTextOption()
        option.setWrapMode(QTextOption.WrapMode.WrapAnywhere)
        doc.setDefaultTextOption(option)
        doc.setTextWidth(float(width))
        margins = self.contentsMargins()
        # +2: rounding / last-line descent; keeps last glyphs from being clipped.
        return max(
            ceil(doc.size().height())
            + margins.top()
            + margins.bottom()
            + 2,
            1,
        )

    def _fit_to_document(self) -> None:
        """Recompute fixed height from current wrap width; call after layout."""
        self._fit_scheduled = False
        self.ensurePolished()
        width = self.viewport().width()
        if width <= 0:
            width = self.width()
        if width <= 0:
            return

        self.document().setTextWidth(float(width))
        height = self._document_height_for_width(width)
        # Prefer live document layout if it reports a larger need (font/style).
        layout = self.document().documentLayout()
        if layout is not None:
            live = ceil(layout.documentSize().height()) + 2
            margins = self.contentsMargins()
            live += margins.top() + margins.bottom()
            height = max(height, live)

        if self.minimumHeight() != height or self.maximumHeight() != height:
            self.setFixedHeight(height)
            self.updateGeometry()
            # Parent may reflow width after our height change — refit once more.
            self._schedule_fit()
        bar = self.verticalScrollBar()
        if bar is not None:
            bar.setValue(0)


def build_side_by_side(
    parent: QWidget,
    *,
    build_left: Callable[[QWidget, QVBoxLayout], None],
    show_stop: bool = False,
) -> tuple[QSplitter, JobMonitorPanel]:
    """Horizontal splitter: scrollable left column + full-height Job monitor.

    Args:
        parent: Owner widget.
        build_left: ``(left_root, left_layout) -> None`` that fills the left column.
        show_stop: Whether Job monitor shows Stop button.

    Returns:
        ``(splitter, monitor)`` — add ``splitter`` to the parent layout.
    """
    splitter = QSplitter(Qt.Orientation.Horizontal, parent)

    scroll = QScrollArea(splitter)
    scroll.setWidgetResizable(True)
    scroll.setFrameShape(QScrollArea.Shape.NoFrame)
    scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
    scroll.setMinimumWidth(_LEFT_MIN_WIDTH)
    scroll.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)

    left = QWidget(scroll)
    left.setSizePolicy(QSizePolicy.Policy.Ignored, QSizePolicy.Policy.Expanding)
    left_layout = QVBoxLayout(left)
    left_layout.setContentsMargins(0, 0, 4, 0)
    build_left(left, left_layout)
    scroll.setWidget(left)

    monitor = JobMonitorPanel(splitter, show_stop=show_stop)
    monitor.setMinimumWidth(_CONSOLE_MIN_WIDTH)
    monitor.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)

    splitter.addWidget(scroll)
    splitter.addWidget(monitor)
    splitter.setStretchFactor(0, 1)
    splitter.setStretchFactor(1, 1)
    splitter.setSizes([_LEFT_WIDTH, _CONSOLE_WIDTH])
    splitter.setChildrenCollapsible(False)
    return splitter, monitor
