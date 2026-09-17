"""Background Qt threads for robot MTO copy and live comparison."""

from __future__ import annotations

import traceback
from datetime import datetime
from pathlib import Path

from PySide6.QtCore import QThread, Signal

from rd_catalog.robot_mto_sync import (
    RobotMtoComparePreview,
    RobotMtoSyncError,
    RobotMtoSyncPlan,
    RobotMtoSyncResult,
    execute_robot_mto_sync,
    preview_robot_mto_compare,
)


class RobotMtoSyncThread(QThread):
    """Archive the previous robot MTO and copy the planned source."""

    log = Signal(str)
    error = Signal(str)

    def __init__(
        self,
        plan: RobotMtoSyncPlan,
        parent=None,
        *,
        now: datetime | None = None,
    ) -> None:
        """Store an immutable copy request.

        Args:
            plan: Validated copy/replace plan.
            parent: Optional Qt parent.
            now: Timestamp for archive folder names.
        """

        super().__init__(parent)
        self._plan = plan
        self.plan = plan
        self._now = now
        self.result: RobotMtoSyncResult | None = None
        self.failure: str | None = None

    def run(self) -> None:
        """Copy on a worker thread; writes only under ``robot_root``."""

        source = self._plan.source
        self.log.emit(
            f"Копирование MTO {source.source_label} → робот: {source.name}"
        )
        try:
            self.result = execute_robot_mto_sync(self._plan, now=self._now)
        except RobotMtoSyncError as exc:
            self.failure = str(exc)
            self.error.emit(self.failure)
            return
        except Exception as exc:
            self.failure = f"{type(exc).__name__}: {exc}"
            self.error.emit(f"{self.failure}\n{traceback.format_exc()}")
            return
        result = self.result
        for item in result.archived:
            self.log.emit(f"Старый файл робота отложен: {item.archived_path}")
        action = "добавлен" if result.added else "заменён"
        self.log.emit(f"MTO робота {action}: {result.destination_path}")


class RobotMtoCompareThread(QThread):
    """Read both MTO xlsx files and build a local comparison preview."""

    error = Signal(str)

    def __init__(
        self,
        plan: RobotMtoSyncPlan,
        output_dir: str | Path,
        parent=None,
    ) -> None:
        """Store an immutable live-compare request.

        Args:
            plan: Copy/replace plan.
            output_dir: Local runtime folder for the preview workbook.
            parent: Optional Qt parent.
        """

        super().__init__(parent)
        self._plan = plan
        self._output_dir = Path(output_dir)
        self.result: RobotMtoComparePreview | None = None
        self.failure: str | None = None

    def run(self) -> None:
        """Compare on a worker thread; never write to RD/SQ/robot roots."""

        try:
            self.result = preview_robot_mto_compare(
                self._plan, self._output_dir
            )
        except Exception as exc:
            self.failure = f"{type(exc).__name__}: {exc}"
            self.error.emit(f"{self.failure}\n{traceback.format_exc()}")
