"""Background Qt thread for Google kits CSV export."""

from __future__ import annotations

import traceback

from PySide6.QtCore import QThread, Signal

from rd_catalog.config import CatalogConfig
from rd_catalog.google_kits import GoogleKitsLoadResult, fetch_google_kits


class KitLoadThread(QThread):
    """Fetch the Google kits sheet away from the GUI thread."""

    log = Signal(str)
    error = Signal(str)

    def __init__(self, config: CatalogConfig, parent=None) -> None:
        """Store an immutable fetch request.

        Args:
            config: Resolved catalog configuration.
            parent: Optional Qt parent.
        """

        super().__init__(parent)
        self._config = config
        self.result: GoogleKitsLoadResult | None = None
        self.failure: str | None = None

    def run(self) -> None:
        """Download/parse kits; never write to UNC or the Google sheet."""

        try:
            self.log.emit("Загрузка комплектов из Google-таблицы…")
            self.result = fetch_google_kits(self._config)
            if self.result.error:
                self.failure = self.result.error
                self.error.emit(self.result.error)
                return
            if self.result.warning:
                self.log.emit(self.result.warning)
            self.log.emit(
                f"Google: {self.result.stats.kept} комплектов "
                f"(пропущено {self.result.stats.skipped}); "
                f"Выдача РД ПД: {self.result.issuance_stats.kept} "
                f"(пропущено {self.result.issuance_stats.skipped}); "
                f"источник {self.result.source}"
            )
        except Exception as exc:
            self.failure = f"{type(exc).__name__}: {exc}"
            self.error.emit(f"{self.failure}\n{traceback.format_exc()}")
