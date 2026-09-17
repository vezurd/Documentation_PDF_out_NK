"""QThread wrapper that runs ``run_v2_pipeline`` off the GUI thread."""

from __future__ import annotations

import traceback

from PySide6.QtCore import QThread, Signal

from pdf_v2_monitor.monitor_callback import MonitorCallback


class PipelineThread(QThread):
    """Runs the full v2 pipeline in a background thread.

    Signals:
        sig_phase: emitted when a pipeline phase starts (``str`` phase name).
        sig_finished: emitted when pipeline is complete (``str`` result_dir, ``float`` elapsed).
        sig_error: emitted on uncaught pipeline exception (``str`` message).
    """

    sig_error = Signal(str)

    def __init__(
        self,
        pdf_path: str,
        cfg: dict,
        project: str | None,
        callback: MonitorCallback,
        parent=None,
    ) -> None:
        super().__init__(parent)
        self._pdf_path = pdf_path
        self._cfg = dict(cfg)
        self._project = project
        self._callback = callback

    def run(self) -> None:
        try:
            from pdf_parsing_v2.v2_pipeline import run_v2_pipeline

            run_v2_pipeline(
                self._pdf_path,
                self._cfg,
                project=self._project,
                extraction_callback=self._callback,
            )
        except Exception as e:
            self.sig_error.emit(f"{type(e).__name__}: {e}\n{traceback.format_exc()}")
