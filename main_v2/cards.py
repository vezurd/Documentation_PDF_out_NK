"""Reusable Qt group widgets for main_v2 shell UI."""

from __future__ import annotations

from collections.abc import Callable

from PySide6.QtWidgets import (
    QGridLayout,
    QGroupBox,
    QHBoxLayout,
    QPushButton,
    QWidget,
)


class WorkflowToolCard(QGroupBox):
    """A titled card with Run / Settings / Open last result actions.

    Wires three callbacks supplied by the host window; implementations may be
    placeholders until a runner API is connected.
    """

    def __init__(
        self,
        title: str,
        *,
        on_run: Callable[[], None],
        on_settings: Callable[[], None],
        on_open_last: Callable[[], None],
        parent: QWidget | None = None,
    ) -> None:
        """Builds the card layout and connects buttons.

        Args:
            title: Group box title (e.g. "RFP / RFQ").
            on_run: Invoked when the Run button is clicked.
            on_settings: Invoked when the Settings button is clicked.
            on_open_last: Invoked when Open last result is clicked.
            parent: Optional parent widget.
        """
        super().__init__(title, parent)
        grid = QGridLayout(self)
        row = QHBoxLayout()
        btn_run = QPushButton("Запуск")
        btn_settings = QPushButton("Настройки")
        btn_open = QPushButton("Открыть последний результат")
        btn_run.clicked.connect(on_run)
        btn_settings.clicked.connect(on_settings)
        btn_open.clicked.connect(on_open_last)
        row.addWidget(btn_run)
        row.addWidget(btn_settings)
        row.addWidget(btn_open)
        row.addStretch(1)
        grid.addLayout(row, 0, 0)
