"""Background thread that writes confirmed F/D/E patches to KSB ИД."""

from __future__ import annotations

import traceback
from collections.abc import Sequence
from pathlib import Path

from PySide6.QtCore import QThread, Signal

from rd_catalog.config import CatalogConfig
from rd_catalog.google_f_write import (
    GoogleSheetsRestClient,
    GoogleWriteError,
    GoogleWriteResult,
    JournalWriteJob,
    SheetsClient,
    execute_journal_writes,
    service_account_path,
)
from rd_catalog.google_kits import GoogleKitsLoadResult, fetch_google_kits


class GoogleFWriteThread(QThread):
    """Write journal jobs away from the GUI thread; no UNC writes."""

    log = Signal(str)
    error = Signal(str)

    def __init__(
        self,
        config: CatalogConfig,
        jobs: Sequence[JournalWriteJob],
        parent=None,
        *,
        client: SheetsClient | None = None,
        credentials_path: str | Path | None = None,
    ) -> None:
        """Store an immutable write batch.

        Args:
            config: Catalog config (spreadsheet id / first-tab pin).
            jobs: Confirmed writes in preview order.
            parent: Optional Qt parent.
            client: Injected Sheets client (tests); live uses the REST client.
            credentials_path: Optional service-account JSON.
        """

        super().__init__(parent)
        self._config = config
        self._jobs = tuple(jobs)
        self._client = client
        self._credentials_path = credentials_path
        self.results: tuple[GoogleWriteResult, ...] = ()
        self.kits_result: GoogleKitsLoadResult | None = None
        self.failure: str | None = None

    def run(self) -> None:
        """Authorize, write each job, emit per-kit log lines."""

        try:
            client = self._client
            if client is None:
                path = self._credentials_path or service_account_path()
                self.log.emit(f"Авторизация сервисного аккаунта: {path}")
                client = GoogleSheetsRestClient(path)
            self.log.emit(f"Запись в КСБ ИД: {len(self._jobs)} записей…")
            results = execute_journal_writes(client, self._config, self._jobs)
            self.results = results
            errors = [item for item in results if item.error]
            for item in results:
                if item.error:
                    self.log.emit(
                        f"{item.title}-{item.mark}: ошибка — {item.error}"
                    )
                else:
                    de = "D/E да" if item.update_de else "только F"
                    self.log.emit(
                        f"{item.title}-{item.mark}: строка {item.row_index} "
                        f"({item.sheet_title}), {de}"
                    )
            if errors and len(errors) == len(results):
                self.failure = errors[0].error
                self.error.emit(self.failure)
                return
            self.log.emit("Обновление комплектов Google после записи…")
            self.kits_result = fetch_google_kits(
                self._config, include_issuance=False
            )
            if self.kits_result.error:
                self.log.emit(
                    f"Таблица записана, но выгрузка комплектов не удалась: "
                    f"{self.kits_result.error}"
                )
            elif self.kits_result.warning:
                self.log.emit(self.kits_result.warning)
        except GoogleWriteError as exc:
            self.failure = str(exc)
            self.error.emit(self.failure)
        except Exception as exc:
            self.failure = f"{type(exc).__name__}: {exc}"
            self.error.emit(f"{self.failure}\n{traceback.format_exc()}")
