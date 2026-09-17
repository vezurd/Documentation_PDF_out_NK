"""Background Qt thread for customer PI xlsb reload and АвтоМто catalog build."""

from __future__ import annotations

import traceback
from pathlib import Path

from PySide6.QtCore import QThread, Signal

from rd_catalog.customer_pi import (
    CustomerPiStore,
    build_from_xlsb,
    default_pickle_path,
    load_customer_pi,
)
from rd_catalog.customer_pi_auto_mto import (
    STATUS_WRITING,
    STATUS_WRITTEN,
    AutoMtoRebuildResult,
    default_auto_mto_dir,
    format_spec_progress_summary,
    list_spec_progress_rows,
    overlay_written_specs,
    rebuild_auto_mto_catalog,
)

JOB_RELOAD_XLSB = "reload_xlsb"
JOB_BUILD_CATALOG = "build_catalog"
JOB_BOTH = "both"


class CustomerPiThread(QThread):
    """Reload the PI pickle and/or rewrite АвтоМто xlsx away from the GUI.

    Never writes the source ``.xlsb``. Pickle and catalog stay under
    ``rd_catalog/База заказчика``.
    """

    log = Signal(str)
    error = Signal(str)
    progress = Signal(int, int, str)
    specs_ready = Signal(object)
    spec_status = Signal(str, str)

    def __init__(
        self,
        *,
        job: str,
        xlsb_path: str | Path | None = None,
        pickle_path: str | Path | None = None,
        dest_dir: str | Path | None = None,
        parent=None,
    ) -> None:
        """Store an immutable job.

        Args:
            job: ``reload_xlsb``, ``build_catalog``, or ``both``.
            xlsb_path: Customer EDMS ``.xlsb`` (required for reload jobs).
            pickle_path: Destination pickle; default packaged dump.
            dest_dir: АвтоМто folder; default packaged catalog.
            parent: Optional Qt parent.
        """

        super().__init__(parent)
        self.job = job
        self._job = job
        self._xlsb_path = Path(xlsb_path) if xlsb_path else None
        self._pickle_path = Path(pickle_path) if pickle_path else default_pickle_path()
        self._dest_dir = Path(dest_dir) if dest_dir else default_auto_mto_dir()
        self.store: CustomerPiStore | None = None
        self.rebuild: AutoMtoRebuildResult | None = None
        self.failure: str | None = None
        self._last_progress_spec = ""

    def run(self) -> None:
        """Parse xlsb and/or write MTO workbooks; cooperative cancel on catalog."""

        try:
            if self._job not in {JOB_RELOAD_XLSB, JOB_BUILD_CATALOG, JOB_BOTH}:
                raise ValueError(f"unknown customer PI job: {self._job!r}")
            if self._job in {JOB_RELOAD_XLSB, JOB_BOTH}:
                if self._xlsb_path is None:
                    raise ValueError("не задан путь к xlsb")
                if self._xlsb_path.suffix.casefold() == ".xlsb" and self._xlsb_path == self._pickle_path:
                    raise ValueError("pickle path must not overwrite the xlsb")
                self.log.emit(f"Чтение xlsb: {self._xlsb_path}")
                self.store = build_from_xlsb(
                    self._xlsb_path,
                    pickle_path=self._pickle_path,
                )
                self.log.emit(
                    f"Pickle обновлён: {len(self.store)} строк BCC → {self._pickle_path}"
                )
            else:
                self.log.emit(f"Чтение pickle: {self._pickle_path}")
                self.store = load_customer_pi(self._pickle_path)
                self.log.emit(f"Pickle: {len(self.store)} строк BCC")
            assert self.store is not None
            idle_dest = (
                self._dest_dir if self._job == JOB_RELOAD_XLSB else None
            )
            spec_rows = list_spec_progress_rows(self.store, dest=idle_dest)
            self.specs_ready.emit(spec_rows)
            if self._job in {JOB_BUILD_CATALOG, JOB_BOTH}:
                self.log.emit(f"Сборка каталога АвтоМТО → {self._dest_dir}")
                self._last_progress_spec = ""
                self.rebuild = rebuild_auto_mto_catalog(
                    self.store,
                    self._dest_dir,
                    progress=self._on_progress,
                    is_cancelled=self.isInterruptionRequested,
                )
                if self._last_progress_spec:
                    self.spec_status.emit(self._last_progress_spec, STATUS_WRITTEN)
                final_rows = overlay_written_specs(spec_rows, self.rebuild.written)
                self.specs_ready.emit(final_rows)
                self.log.emit(
                    f"АвтоМТО: записано {len(self.rebuild.written)}, "
                    f"пропущено {len(self.rebuild.skipped)}, "
                    f"удалено {len(self.rebuild.removed)}"
                )
                self.log.emit(format_spec_progress_summary(final_rows))
                for note in self.rebuild.skipped:
                    self.log.emit(f"Пропуск: {note.splitlines()[0]}")
            else:
                self.log.emit(format_spec_progress_summary(spec_rows))
        except Exception as exc:
            self.failure = f"{type(exc).__name__}: {exc}"
            self.error.emit(f"{self.failure}\n{traceback.format_exc()}")

    def _on_progress(self, index: int, total: int, spec: str) -> None:
        previous = self._last_progress_spec
        if previous and previous != spec:
            self.spec_status.emit(previous, STATUS_WRITTEN)
        self.spec_status.emit(spec, STATUS_WRITING)
        self._last_progress_spec = spec
        self.progress.emit(index, total, spec)
        if index == 1 or index == total or index % 25 == 0:
            self.log.emit(f"АвтоМТО {index}/{total}: {spec}")
