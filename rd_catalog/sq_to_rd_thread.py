"""Background Qt thread for moving an SQ kit folder into RD."""

from __future__ import annotations

import traceback

from PySide6.QtCore import QThread, Signal

from rd_catalog.sq_to_rd import (
    SqToRdError,
    SqToRdPlan,
    SqToRdResult,
    execute_sq_to_rd_transfer,
)


class SqToRdThread(QThread):
    """Move the planned SQ folder on a worker thread."""

    log = Signal(str)
    error = Signal(str)

    def __init__(self, plan: SqToRdPlan, parent=None) -> None:
        """Store an immutable move request.

        Args:
            plan: Validated SQ→RD move plan.
            parent: Optional Qt parent.
        """

        super().__init__(parent)
        self._plan = plan
        self.plan = plan
        self.result: SqToRdResult | None = None
        self.failure: str | None = None

    def run(self) -> None:
        """Move SQ children into the new RD transfer folder."""

        plan = self._plan
        self.log.emit(
            f"Перенос SQ → РД: {plan.title}-{plan.mark} · {plan.transfer_name}"
        )
        try:
            self.result = execute_sq_to_rd_transfer(plan)
        except SqToRdError as exc:
            self.failure = str(exc)
            self.error.emit(self.failure)
            return
        except Exception as exc:
            self.failure = f"{type(exc).__name__}: {exc}"
            self.error.emit(f"{self.failure}\n{traceback.format_exc()}")
            return
        result = self.result
        self.log.emit(f"Создана передача: {result.destination_folder}")
        for name in result.moved_names:
            self.log.emit(f"  перемещено: {name}")
        if result.source_removed:
            self.log.emit(f"Пустая папка SQ удалена: {result.source_folder}")
        else:
            self.log.emit(f"Папка SQ сохранена (не пуста): {result.source_folder}")
