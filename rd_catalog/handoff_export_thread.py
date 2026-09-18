"""Background Qt thread for official-kit dump copy (MTO + BBB)."""

from __future__ import annotations

import threading
import traceback
from datetime import datetime

from PySide6.QtCore import QThread, Signal

from rd_catalog.handoff_export import (
    HandoffCopyPlan,
    HandoffCopyReport,
    execute_handoff_copy,
)


class HandoffExportThread(QThread):
    """Copy planned dump files on a worker thread; never under RD/SQ roots."""

    progress = Signal(int, int)
    log = Signal(str)
    error = Signal(str)

    def __init__(
        self,
        plan: HandoffCopyPlan,
        parent=None,
        *,
        now: datetime | None = None,
    ) -> None:
        """Store an immutable dump copy request.

        Args:
            plan: Validated copy plan and preview rows.
            parent: Optional Qt parent.
            now: Clock for the перечень filename.
        """

        super().__init__(parent)
        self._plan = plan
        self.plan = plan
        self._now = now
        self._cancel = threading.Event()
        self.report: HandoffCopyReport | None = None
        self.failure: str | None = None

    def request_cancel(self) -> None:
        """Request cooperative cancellation between copy items."""

        self._cancel.set()
        self.log.emit("Запрошена отмена выгрузки комплектов…")

    def run(self) -> None:
        """Copy on a worker thread; never on the GUI thread."""

        self.log.emit(
            "Выгрузка комплектов: "
            f"{len(self._plan.items)} файл(ов) → {self._plan.dest_root}"
        )
        try:
            self.report = execute_handoff_copy(
                self._plan,
                now=self._now,
                cancel=self._cancel.is_set,
                progress=self._emit_progress,
            )
        except Exception as exc:
            self.failure = f"{type(exc).__name__}: {exc}"
            self.error.emit(f"{self.failure}\n{traceback.format_exc()}")
            return
        report = self.report
        if report.cancelled:
            self.log.emit("Выгрузка комплектов прервана.")
            return
        failed = len(report.failed)
        self.log.emit(
            "Выгрузка комплектов: "
            f"скопировано {report.copied}, "
            f"пропущено {report.skipped}, "
            f"ошибок {failed}"
            + (f", перечень {report.list_path}" if report.list_path else "")
        )

    def _emit_progress(self, completed: int, total: int) -> None:
        self.progress.emit(completed, total)
