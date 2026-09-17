"""Background MTO content compare for disputed transfer packages."""

from __future__ import annotations

import traceback
from dataclasses import dataclass

from PySide6.QtCore import QThread, Signal

from rd_catalog.customer_pi_auto_mto import MtoPairCompareResult
from rd_catalog.kits import TransferReviewMtoPair
from rd_catalog.transfer_review_compare import (
    cache_entry_key,
    compare_transfer_mto_pair,
)


@dataclass(frozen=True, slots=True)
class TransferReviewCompareOutcome:
    """One disputed-package MTO pair after content compare."""

    pair: TransferReviewMtoPair
    cache_key: str
    result: MtoPairCompareResult


class TransferReviewCompareThread(QThread):
    """Read disputed RD MTO workbooks without blocking the Qt GUI."""

    result_ready = Signal(object)
    error = Signal(str)

    def __init__(
        self,
        jobs: tuple[TransferReviewMtoPair, ...],
        parent=None,
    ) -> None:
        super().__init__(parent)
        self._jobs = jobs

    def run(self) -> None:
        """Compare every queued pair on the worker thread."""

        try:
            for pair in self._jobs:
                if self.isInterruptionRequested():
                    return
                result = compare_transfer_mto_pair(
                    pair.left_path, pair.right_path
                )
                self.result_ready.emit(
                    TransferReviewCompareOutcome(
                        pair=pair,
                        cache_key=cache_entry_key(
                            pair.left_path,
                            pair.left_mtime_ns,
                            pair.right_path,
                            pair.right_mtime_ns,
                        ),
                        result=result,
                    )
                )
        except Exception as exc:  # pragma: no cover - defensive Qt boundary
            self.error.emit(f"{type(exc).__name__}: {exc}")
            traceback.print_exc()
