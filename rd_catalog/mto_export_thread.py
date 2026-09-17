"""Background Qt thread for bulk MTO export copy.

Mirrors :class:`rd_catalog.robot_mto_thread.RobotMtoSyncThread`: a ``QThread``
that calls the existing copy executor. Progress and cooperative cancel are
added because a batch can be hundreds of kits; the single-kit thread has
neither. Archiving and the write guard stay inside
:func:`rd_catalog.mto_export_copy.execute_export_copy` →
:func:`rd_catalog.robot_mto_sync.execute_robot_mto_sync`.
"""

from __future__ import annotations

import threading
import traceback
from datetime import datetime

from PySide6.QtCore import QThread, Signal

from rd_catalog.mto_export import ExportPlan
from rd_catalog.mto_export_copy import ExportCopyReport, execute_export_copy


class MtoExportCopyThread(QThread):
    """Copy planned MTO files on a worker thread; writes only under the target."""

    progress = Signal(int, int)
    log = Signal(str)
    error = Signal(str)

    def __init__(
        self,
        plan: ExportPlan,
        parent=None,
        *,
        now: datetime | None = None,
    ) -> None:
        """Store an immutable batch copy request.

        Args:
            plan: Validated copy/replace plan.
            parent: Optional Qt parent.
            now: Timestamp for archive folder names.
        """

        super().__init__(parent)
        self._plan = plan
        self.plan = plan
        self._now = now
        self._cancel = threading.Event()
        self.report: ExportCopyReport | None = None
        self.failure: str | None = None

    def request_cancel(self) -> None:
        """Request cooperative cancellation between copy items."""

        self._cancel.set()
        self.log.emit("Запрошена отмена копирования MTO…")

    def run(self) -> None:
        """Copy on a worker thread; never on the GUI thread."""

        self.log.emit(
            f"Копирование MTO: {len(self._plan.items)} файл(ов) → {self._plan.target.name}"
        )
        try:
            self.report = execute_export_copy(
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
            self.log.emit("Копирование MTO прервано.")
        self.log.emit(
            "Копирование MTO: "
            f"скопировано {len(report.copied)}, "
            f"архив {len(report.archived)}, "
            f"пропущено {len(report.skipped)}, "
            f"ошибок {len(report.failed)}"
        )

    def _emit_progress(self, completed: int, total: int) -> None:
        self.progress.emit(completed, total)
