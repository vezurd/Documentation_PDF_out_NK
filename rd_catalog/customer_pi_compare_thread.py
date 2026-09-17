"""Background comparison of customer AutoMTO files with one RD MTO."""

from __future__ import annotations

import traceback
from collections.abc import Sequence
from pathlib import Path

from PySide6.QtCore import QThread, Signal

from rd_catalog.customer_pi_auto_mto import (
    AutoMtoCompareResult,
    AutoMtoFile,
    compare_auto_mto_to_rd,
)


class AutoMtoCompareThread(QThread):
    """Read RD and AutoMTO workbooks without blocking the Qt GUI."""

    result_ready = Signal(object)
    error = Signal(str)

    def __init__(
        self,
        files: Sequence[AutoMtoFile],
        rd_path: str | Path,
        *,
        parent=None,
    ) -> None:
        super().__init__(parent)
        self._files = tuple(files)
        self._rd_path = str(rd_path)

    def run(self) -> None:
        """Execute exact single/composite matching on the worker thread."""

        try:
            result: AutoMtoCompareResult = compare_auto_mto_to_rd(
                self._files,
                self._rd_path,
            )
        except Exception as exc:  # pragma: no cover - defensive Qt boundary
            self.error.emit(f"{type(exc).__name__}: {exc}")
            traceback.print_exc()
            return
        self.result_ready.emit(result)
