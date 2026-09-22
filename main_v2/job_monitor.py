"""Reusable job status panel: log tail, elapsed time, indeterminate progress."""

from __future__ import annotations

import time
import webbrowser
from pathlib import Path

from PySide6.QtCore import Qt, QTimer, Signal, Slot
from PySide6.QtGui import QGuiApplication, QTextCursor
from PySide6.QtWidgets import (
    QGridLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QPlainTextEdit,
    QProgressBar,
    QPushButton,
    QVBoxLayout,
    QWidget,
)


class JobMonitorPanel(QWidget):
    """Status header, indeterminate progress, and live log for long-running jobs.

    Attributes:
        stop_requested: Emitted when the user clicks ``Остановить``.

    Typical integration: connect :attr:`stop_requested` to
    :meth:`ProcessJobRunner.request_stop`, and wire
    :meth:`append_log` / :meth:`finish_job` to runner signals.
    """

    stop_requested = Signal()

    def __init__(self, parent: QWidget | None = None, *, show_stop: bool = True) -> None:
        super().__init__(parent)
        self._result_path: str | None = None
        self._start_monotonic: float | None = None

        self._status_label = QLabel(self)
        self._status_label.setWordWrap(True)

        self._elapsed_label = QLabel("Elapsed: 0:00", self)

        self._progress = QProgressBar(self)
        self._progress.setRange(0, 100)
        self._progress.setValue(0)

        self._log = QPlainTextEdit(self)
        self._log.setReadOnly(True)
        self._log.setLineWrapMode(QPlainTextEdit.LineWrapMode.WidgetWidth)

        self._btn_copy = QPushButton("Копировать лог", self)
        self._btn_clear = QPushButton("Очистить", self)
        self._btn_open = QPushButton("Открыть результат", self)
        self._btn_open.setEnabled(False)
        self._btn_stop = QPushButton("Остановить", self)
        self._btn_stop.setEnabled(False)
        self._btn_stop.setVisible(show_stop)

        self._btn_copy.clicked.connect(self._copy_log)
        self._btn_clear.clicked.connect(self.clear_log)
        self._btn_open.clicked.connect(self._open_result)
        self._btn_stop.clicked.connect(self.stop_requested.emit)

        self._elapsed_timer = QTimer(self)
        self._elapsed_timer.setInterval(500)
        self._elapsed_timer.timeout.connect(self._refresh_elapsed)

        row_btns = QHBoxLayout()
        row_btns.addWidget(self._btn_copy)
        row_btns.addWidget(self._btn_clear)
        row_btns.addWidget(self._btn_open)
        row_btns.addStretch(1)
        row_btns.addWidget(self._btn_stop)

        grid = QGridLayout()
        grid.addWidget(QLabel("Status:", self), 0, 0, alignment=Qt.AlignmentFlag.AlignTop)
        grid.addWidget(self._status_label, 0, 1)
        grid.addWidget(QLabel("Time:", self), 1, 0)
        grid.addWidget(self._elapsed_label, 1, 1)
        grid.addWidget(QLabel("Progress:", self), 2, 0)
        grid.addWidget(self._progress, 2, 1)

        outer = QVBoxLayout(self)
        box = QGroupBox("Job monitor", self)
        bl = QVBoxLayout(box)
        bl.addLayout(grid)
        bl.addWidget(self._log, stretch=1)
        bl.addLayout(row_btns)
        outer.addWidget(box)

    def start_job(self, title: str) -> None:
        """Set status text, reset elapsed timer, and start indeterminate progress.

        Args:
            title: Short description shown in the status line.
        """
        self._status_label.setText(title)
        self._start_monotonic = time.monotonic()
        self._elapsed_label.setText("Elapsed: 0:00")
        self._elapsed_timer.start()
        self.set_running(True)
        self._result_path = None
        self._btn_open.setEnabled(False)

    def append_log(self, text: str) -> None:
        """Append a chunk of process output to the log view.

        Args:
            text: Raw text (may be partial lines).
        """
        self._log.moveCursor(QTextCursor.MoveOperation.End)
        self._log.insertPlainText(text)
        self._log.moveCursor(QTextCursor.MoveOperation.End)

    def clear_log(self) -> None:
        """Clear the log text."""
        self._log.clear()

    def finish_job(
        self,
        success: bool,
        message: str,
        result_path: str | None = None,
    ) -> None:
        """Stop elapsed updates, show outcome, and fix progress bar range.

        Args:
            success: When ``True``, progress value is set to 100; otherwise 0.
            message: Final status line text.
            result_path: Optional path for ``Открыть результат``.
        """
        self._elapsed_timer.stop()
        self._refresh_elapsed()
        self._status_label.setText(message)
        self._result_path = result_path
        self.set_running(False)
        self._progress.setValue(100 if success else 0)
        self._btn_open.setEnabled(result_path is not None)

    def set_running(self, running: bool) -> None:
        """Toggle indeterminate progress and stop button availability.

        Args:
            running: When ``True``, progress bar is indeterminate and stop is
                enabled (if visible).
        """
        if running:
            self._progress.setRange(0, 0)
            self._btn_stop.setEnabled(self._btn_stop.isVisible())
        else:
            self._progress.setRange(0, 100)
            self._btn_stop.setEnabled(False)

    def set_progress_fraction(self, current: int, total: int) -> None:
        """Show determinate progress when ``total > 0``."""
        if total <= 0:
            return
        self._progress.setRange(0, total)
        self._progress.setValue(min(max(current, 0), total))

    @Slot()
    def _copy_log(self) -> None:
        clip = QGuiApplication.clipboard()
        if clip is not None:
            clip.setText(self._log.toPlainText())

    @Slot()
    def _open_result(self) -> None:
        """Open ``result_path``: file in the default app, directory in Explorer."""
        if not self._result_path:
            return
        path = Path(self._result_path)
        try:
            is_file = path.is_file()
        except OSError:
            is_file = False
        if is_file:
            webbrowser.open(path.as_uri())
            return
        # Directory, missing path, or UNC flake — open via Explorer helper.
        from utils.path import open_dir

        open_dir(str(path))

    @Slot()
    def _refresh_elapsed(self) -> None:
        if self._start_monotonic is None:
            self._elapsed_label.setText("Elapsed: 0:00")
            return
        sec = int(time.monotonic() - self._start_monotonic)
        mm, ss = divmod(sec, 60)
        hh, mm = divmod(mm, 60)
        if hh:
            self._elapsed_label.setText(f"Elapsed: {hh:d}:{mm:02d}:{ss:02d}")
        else:
            self._elapsed_label.setText(f"Elapsed: {mm:d}:{ss:02d}")
