"""Background pairwise AN content compare against AutoMTO and RD MTO."""

from __future__ import annotations

import traceback
from dataclasses import dataclass

from PySide6.QtCore import QThread, Signal

from rd_catalog.an_compare import compare_an_file
from rd_catalog.customer_pi_auto_mto import MtoPairCompareResult


@dataclass(frozen=True, slots=True)
class AnContentCompareOutcome:
    """One AN file compared to the kit AutoMTO and/or RD MTO workbooks."""

    an_path: str
    vs_auto: MtoPairCompareResult | None
    vs_rd: MtoPairCompareResult | None


class AnContentCompareThread(QThread):
    """Read AN / PI / RD workbooks without blocking the Qt GUI."""

    result_ready = Signal(object)
    error = Signal(str)

    def __init__(
        self,
        an_path: str,
        *,
        auto_path: str = "",
        rd_path: str = "",
        parent=None,
    ) -> None:
        super().__init__(parent)
        self._an_path = str(an_path)
        self._auto_path = str(auto_path or "")
        self._rd_path = str(rd_path or "")

    def run(self) -> None:
        """Execute pairwise matching on the worker thread."""

        try:
            vs_auto, vs_rd = compare_an_file(
                self._an_path,
                auto_path=self._auto_path,
                rd_path=self._rd_path,
            )
        except Exception as exc:  # pragma: no cover - defensive Qt boundary
            self.error.emit(f"{type(exc).__name__}: {exc}")
            traceback.print_exc()
            return
        self.result_ready.emit(
            AnContentCompareOutcome(
                an_path=self._an_path,
                vs_auto=vs_auto,
                vs_rd=vs_rd,
            )
        )
