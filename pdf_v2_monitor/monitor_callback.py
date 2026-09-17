"""Qt bridge for pipeline callbacks emitted from the main process."""

from __future__ import annotations

from PySide6.QtCore import QObject, Signal


class MonitorCallback(QObject):
    """Implements ``ParallelCallback`` (and ``ExtractionCallback``) protocol.

    Emits Qt signals that the ``MonitorWindow`` connects to in the GUI thread.
    """

    sig_file_start = Signal(str, int, int)
    sig_file_done = Signal(str, int, int, float, object)
    sig_file_error = Signal(str, int, str)
    sig_tag_file_start = Signal(str, int, int)
    sig_tag_file_done = Signal(str, int, float)
    sig_tag_file_error = Signal(str, int, str)
    sig_batch_progress = Signal(int, int)
    sig_phase_changed = Signal(str)
    sig_stage_start = Signal(str, str, object, object)
    sig_stage_done = Signal(str, str, float, object, object)
    sig_stage_error = Signal(str, str, str, object, object)
    sig_pipeline_finished = Signal(str, float, object)

    def on_file_start(self, file_path: str, index: int, total: int) -> None:
        self.sig_file_start.emit(file_path, index, total)

    def on_file_done(
        self,
        file_path: str,
        index: int,
        n_pages: int,
        elapsed_sec: float,
        file_detail: dict | None = None,
    ) -> None:
        self.sig_file_done.emit(
            file_path,
            index,
            n_pages,
            elapsed_sec,
            file_detail if file_detail is not None else {},
        )

    def on_file_error(
        self, file_path: str, index: int, error: Exception
    ) -> None:
        self.sig_file_error.emit(file_path, index, str(error))

    def on_tag_file_start(self, file_path: str, index: int, total: int) -> None:
        self.sig_tag_file_start.emit(file_path, index, total)

    def on_tag_file_done(
        self, file_path: str, index: int, elapsed_sec: float
    ) -> None:
        self.sig_tag_file_done.emit(file_path, index, elapsed_sec)

    def on_tag_file_error(
        self, file_path: str, index: int, error: Exception
    ) -> None:
        self.sig_tag_file_error.emit(file_path, index, str(error))

    def on_batch_progress(self, completed: int, total: int) -> None:
        self.sig_batch_progress.emit(completed, total)

    def on_phase_changed(self, phase: str) -> None:
        self.sig_phase_changed.emit(phase)

    def on_stage_start(
        self,
        category: str,
        stage: str,
        file_name: str | None = None,
        detail: dict | None = None,
    ) -> None:
        self.sig_stage_start.emit(category, stage, file_name, detail or {})

    def on_stage_done(
        self,
        category: str,
        stage: str,
        elapsed_sec: float,
        file_name: str | None = None,
        detail: dict | None = None,
    ) -> None:
        self.sig_stage_done.emit(category, stage, elapsed_sec, file_name, detail or {})

    def on_stage_error(
        self,
        category: str,
        stage: str,
        error: Exception,
        file_name: str | None = None,
        detail: dict | None = None,
    ) -> None:
        self.sig_stage_error.emit(category, stage, str(error), file_name, detail or {})

    def on_pipeline_finished(
        self,
        result_dir: str,
        total_elapsed: float,
        summary: dict,
    ) -> None:
        self.sig_pipeline_finished.emit(result_dir, total_elapsed, summary)
