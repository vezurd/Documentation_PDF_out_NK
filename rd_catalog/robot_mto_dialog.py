"""Confirmation dialog for copying an RD/SQ MTO into the robot folder."""

from __future__ import annotations

from datetime import datetime
from pathlib import Path

from PySide6.QtCore import Slot
from PySide6.QtGui import QCloseEvent
from PySide6.QtWidgets import (
    QDialog,
    QDialogButtonBox,
    QHBoxLayout,
    QLabel,
    QMessageBox,
    QPlainTextEdit,
    QPushButton,
    QVBoxLayout,
    QWidget,
)

from rd_catalog.path_actions import open_path
from rd_catalog.robot_mto_sync import RobotMtoSyncPlan, plan_summary_text
from rd_catalog.robot_mto_thread import RobotMtoCompareThread


class RobotMtoSyncDialog(QDialog):
    """Show source and robot file details before copying."""

    def __init__(
        self,
        plan: RobotMtoSyncPlan,
        parent: QWidget | None = None,
        *,
        now: datetime | None = None,
        runtime_dir: str | Path | None = None,
    ) -> None:
        """Build the confirmation dialog.

        Args:
            plan: Copy/replace plan for one title+mark kit.
            parent: Optional Qt parent.
            now: Timestamp used to preview the archive folder name.
            runtime_dir: Local catalog runtime folder for the compare workbook.
        """

        super().__init__(parent)
        self._plan = plan
        self._runtime_dir = Path(runtime_dir) if runtime_dir else None
        self._compare_thread: RobotMtoCompareThread | None = None
        self._report_path: str | None = None
        self.setWindowTitle(
            "Заменить MTO у робота" if plan.is_replace else "Добавить MTO у робота"
        )
        self.setMinimumSize(720, 560)
        self.resize(860, 680)
        self._build(now)

    def _build(self, now: datetime | None) -> None:
        layout = QVBoxLayout(self)
        hint = QLabel(
            "Проверьте источник и файл робота. РД и SQ не изменяются. "
            "Старый файл робота будет перемещён в папку "
            "«_old_<имя>_ГГГГ.ММ.ДД_ЧЧ.ММ» в том же каталоге. "
            "«Сравнить» читает оба xlsx заново (не кэш скана).",
            self,
        )
        hint.setWordWrap(True)
        layout.addWidget(hint)

        details = QPlainTextEdit(self)
        details.setReadOnly(True)
        details.setPlainText(plan_summary_text(self._plan, now=now))
        layout.addWidget(details, 1)

        compare_label = QLabel("Сверка MTO", self)
        layout.addWidget(compare_label)
        self._compare_view = QPlainTextEdit(self)
        self._compare_view.setReadOnly(True)
        self._compare_view.setPlaceholderText(
            "Нажмите «Сравнить», чтобы прочитать оба xlsx заново "
            "и показать отличия источника от файла робота."
        )
        layout.addWidget(self._compare_view, 1)

        compare_row = QHBoxLayout()
        self._compare_button = QPushButton("Сравнить", self)
        self._compare_button.clicked.connect(self._start_compare)
        compare_row.addWidget(self._compare_button)
        self._open_report_button = QPushButton("Открыть файл сравнения", self)
        self._open_report_button.setEnabled(False)
        self._open_report_button.clicked.connect(self._open_report)
        compare_row.addWidget(self._open_report_button)
        compare_row.addStretch(1)
        layout.addLayout(compare_row)

        buttons = QDialogButtonBox(self)
        confirm_label = (
            "Подтвердить замену" if self._plan.is_replace else "Подтвердить добавление"
        )
        buttons.addButton(confirm_label, QDialogButtonBox.ButtonRole.AcceptRole)
        buttons.addButton("Отмена", QDialogButtonBox.ButtonRole.RejectRole)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def _compare_output_dir(self) -> Path | None:
        if self._runtime_dir is None:
            return None
        return self._runtime_dir / "mto_preview"

    def _start_compare(self) -> None:
        if self._compare_thread is not None:
            return
        output_dir = self._compare_output_dir()
        if output_dir is None:
            QMessageBox.warning(
                self,
                "Сверка MTO",
                "Не задана папка runtime для файла сравнения.",
            )
            return
        thread = RobotMtoCompareThread(self._plan, output_dir, self)
        thread.error.connect(self._on_compare_error)
        thread.finished.connect(self._on_compare_finished)
        self._compare_thread = thread
        self._compare_button.setEnabled(False)
        self._open_report_button.setEnabled(False)
        self._compare_view.setPlainText("Читаем xlsx источника и робота…")
        thread.start()

    @Slot(str)
    def _on_compare_error(self, message: str) -> None:
        self._compare_view.setPlainText(message)

    @Slot()
    def _on_compare_finished(self) -> None:
        thread = self._compare_thread
        self._compare_thread = None
        self._compare_button.setEnabled(True)
        if thread is None:
            return
        thread.deleteLater()
        if thread.failure or thread.result is None:
            if not self._compare_view.toPlainText().strip():
                self._compare_view.setPlainText(
                    thread.failure or "Не удалось выполнить сверку."
                )
            return
        preview = thread.result
        self._compare_view.setPlainText(preview.summary)
        self._report_path = preview.report_path
        self._open_report_button.setEnabled(bool(preview.report_path))

    def _open_report(self) -> None:
        if not self._report_path:
            return
        ok, message = open_path(self._report_path)
        if not ok:
            QMessageBox.warning(self, "Файл сравнения", message)

    def closeEvent(self, event: QCloseEvent) -> None:
        """Avoid destroying a running compare worker."""

        if self._compare_thread is not None:
            if not self._compare_thread.wait(4000):
                QMessageBox.information(
                    self,
                    "Сверка MTO",
                    "Сверка ещё читает файлы. Повторите закрытие через несколько секунд.",
                )
                event.ignore()
                return
        super().closeEvent(event)
