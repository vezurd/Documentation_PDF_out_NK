"""Preview and progress dialogs for bulk MTO export copy.

The confirmation dialog is a batch preview (not a reuse of
:class:`rd_catalog.robot_mto_dialog.RobotMtoSyncDialog`, which is one kit
plus a live xlsx compare). Progress/cancel mirrors that pair's worker
lifecycle: a ``QThread``, ``finished``, and a close-guard while running.
"""

from __future__ import annotations

from collections.abc import Sequence

from PySide6.QtCore import Slot
from PySide6.QtGui import QCloseEvent
from PySide6.QtWidgets import (
    QDialog,
    QDialogButtonBox,
    QLabel,
    QMessageBox,
    QPlainTextEdit,
    QProgressBar,
    QPushButton,
    QVBoxLayout,
    QWidget,
)

from rd_catalog.mto_export import ExportTarget
from rd_catalog.mto_export_copy import (
    ExportCopyReport,
    ExportPreviewGroup,
    ExportSelection,
    format_export_copy_report,
    format_preview_body,
    group_export_preview,
    preview_copy_count,
)
from rd_catalog.mto_export_thread import MtoExportCopyThread


class MtoExportPreviewDialog(QDialog):
    """Show grouped add/replace/skip rows before one batch confirmation."""

    def __init__(
        self,
        selections: Sequence[ExportSelection],
        target: ExportTarget,
        parent: QWidget | None = None,
    ) -> None:
        """Build the combined preview.

        Args:
            selections: Currently visible heatmap rows.
            target: Destination folder.
            parent: Optional Qt parent.
        """

        super().__init__(parent)
        self._selections = tuple(selections)
        self._target = target
        self.groups: tuple[ExportPreviewGroup, ...] = group_export_preview(
            self._selections
        )
        copy_n = preview_copy_count(self._selections)
        self.setWindowTitle("Копирование MTO в папку робота")
        self.setMinimumSize(720, 520)
        self.resize(840, 640)
        layout = QVBoxLayout(self)
        hint = QLabel(
            "Будут скопированы только строки «добавится» и «заменится». "
            "«данные совпадают» и «нет файла» показаны, но не копируются. "
            "Старый файл в папке назначения уходит в "
            "«_old_<имя>_ГГГГ.ММ.ДД_ЧЧ.ММ». РД и SQ не изменяются.",
            self,
        )
        hint.setWordWrap(True)
        layout.addWidget(hint)
        details = QPlainTextEdit(self)
        details.setReadOnly(True)
        details.setPlainText(format_preview_body(self.groups, target=target))
        layout.addWidget(details, 1)
        buttons = QDialogButtonBox(self)
        confirm = (
            f"Подтвердить копирование ({copy_n})"
            if copy_n
            else "Закрыть"
        )
        self._confirm_button = buttons.addButton(
            confirm, QDialogButtonBox.ButtonRole.AcceptRole
        )
        if copy_n:
            buttons.addButton("Отмена", QDialogButtonBox.ButtonRole.RejectRole)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)


class MtoExportProgressDialog(QDialog):
    """Progress, cancel, and the final report for a running copy thread."""

    def __init__(
        self,
        thread: MtoExportCopyThread,
        parent: QWidget | None = None,
    ) -> None:
        """Attach to an already constructed (not necessarily started) thread.

        Args:
            thread: Batch copy worker.
            parent: Optional Qt parent.
        """

        super().__init__(parent)
        self._thread = thread
        self.report: ExportCopyReport | None = None
        self.setWindowTitle("Копирование MTO")
        self.setMinimumSize(560, 320)
        self.setModal(True)
        layout = QVBoxLayout(self)
        self._status = QLabel("Копирование файлов MTO…", self)
        self._status.setWordWrap(True)
        layout.addWidget(self._status)
        self._progress = QProgressBar(self)
        total = len(thread.plan.items) or 1
        self._progress.setRange(0, total)
        self._progress.setValue(0)
        layout.addWidget(self._progress)
        self._log = QPlainTextEdit(self)
        self._log.setReadOnly(True)
        layout.addWidget(self._log, 1)
        self._cancel_button = QPushButton("Отмена", self)
        self._cancel_button.clicked.connect(self._on_cancel_clicked)
        layout.addWidget(self._cancel_button)
        thread.progress.connect(self._on_progress)
        thread.log.connect(self._append_log)
        thread.error.connect(self._append_log)
        thread.finished.connect(self._on_finished)

    def start_copy(self) -> None:
        """Start the worker if it is not already running."""

        if not self._thread.isRunning():
            self._thread.start()

    @Slot(int, int)
    def _on_progress(self, completed: int, total: int) -> None:
        if total > 0:
            self._progress.setRange(0, total)
        self._progress.setValue(completed)
        self._status.setText(f"Копирование MTO: {completed} из {total}")

    @Slot(str)
    def _append_log(self, message: str) -> None:
        self._log.appendPlainText(message)

    @Slot()
    def _on_cancel_clicked(self) -> None:
        self._cancel_button.setEnabled(False)
        self._thread.request_cancel()
        self._status.setText("Отмена копирования…")

    @Slot()
    def _on_finished(self) -> None:
        self.report = self._thread.report
        self._cancel_button.setEnabled(False)
        if self._thread.failure or self.report is None:
            self._status.setText("Ошибка копирования MTO")
            if self._thread.failure:
                self._log.appendPlainText(self._thread.failure)
        elif self.report is not None:
            self._log.setPlainText(format_export_copy_report(self.report))
            self._status.setText("Копирование MTO завершено")
            total = len(self._thread.plan.items) or 1
            self._progress.setValue(total)
        self.accept()

    def closeEvent(self, event: QCloseEvent) -> None:
        """Avoid destroying a running copy worker."""

        if self._thread.isRunning():
            self._thread.request_cancel()
            if not self._thread.wait(4000):
                QMessageBox.information(
                    self,
                    "Копирование MTO",
                    "Копирование ещё идёт. Повторите закрытие через несколько секунд.",
                )
                event.ignore()
                return
        super().closeEvent(event)
