"""Run child Python jobs via ``QProcess`` without blocking the GUI thread."""

from __future__ import annotations

import sys
from pathlib import Path
from PySide6.QtCore import QObject, QProcess, QProcessEnvironment, Signal


def build_aggregate_tags_argv(
    project_root: str | Path,
    *,
    asbuild_config_path: str | Path | None = None,
) -> list[str]:
    """Build the command line for ``RFQ/tags_rfp_compare/agregate_tags.py``.

    Mirrors ``main.py`` helpers ``_agregate_tags_subprocess_work`` and
    ``_agregate_tags_asbuild_subprocess_work``: ``sys.executable -u`` on the
    script, optional trailing path to the as-build JSON config.

    Args:
        project_root: Repository root (directory containing ``RFQ/``).
        asbuild_config_path: When set, passed as a single extra CLI argument
            (as-build mode). When ``None``, standard RFQ/MTO/RFP run.

    Returns:
        Argument list suitable for :meth:`ProcessJobRunner.start`.
    """
    root = Path(project_root).resolve()
    script = root / "RFQ" / "tags_rfp_compare" / "agregate_tags.py"
    argv: list[str] = [sys.executable, "-u", str(script)]
    if asbuild_config_path is not None:
        argv.append(str(Path(asbuild_config_path).resolve()))
    return argv


def aggregate_tags_env_overlay(project_root: str | Path) -> dict[str, str]:
    """Return environment entries required for the aggregate_tags child process.

    Sets ``PYTHONPATH`` to the project root, matching ``main.py`` subprocess
    wiring.

    Args:
        project_root: Repository root directory.

    Returns:
        Mapping merged into the child environment by :class:`ProcessJobRunner`.
    """
    root = str(Path(project_root).resolve())
    return {"PYTHONPATH": root}


class ProcessJobRunner(QObject):
    """Launches ``QProcess`` with streamed stdout/stderr on the GUI thread.

    All ``QProcess`` slots run on the Qt event loop; output is read
    incrementally via ``readyReadStandardOutput`` / ``readyReadStandardError``.

    Signals:
        started: Emitted when the process starts successfully.
        log_received: Raw text chunk from stdout or stderr (stderr prefixed).
        finished: ``(exit_code, exit_status)`` where ``exit_status`` is
            :attr:`QProcess.ExitStatus.NormalExit` or ``CrashExit``.
        error: Human-readable message (start failure, already running, etc.).
    """

    started = Signal()
    log_received = Signal(str)
    finished = Signal(int, int)
    error = Signal(str)

    def __init__(self, parent: QObject | None = None) -> None:
        super().__init__(parent)
        self._process = QProcess(self)
        self._process.setProcessChannelMode(QProcess.ProcessChannelMode.SeparateChannels)
        self._process.started.connect(self._on_started)
        self._process.readyReadStandardOutput.connect(self._on_stdout)
        self._process.readyReadStandardError.connect(self._on_stderr)
        self._process.finished.connect(self._on_finished)
        self._process.errorOccurred.connect(self._on_error_occurred)

    def process(self) -> QProcess:
        """Return the owned ``QProcess`` (for advanced wiring)."""
        return self._process

    def start(
        self,
        argv: list[str],
        *,
        cwd: str | Path | None = None,
        env: dict[str, str] | None = None,
    ) -> bool:
        """Start a child process.

        Args:
            argv: Full argument vector; ``argv[0]`` is the program to execute.
            cwd: Working directory; defaults to ``None`` (Qt/runtime default).
            env: If set, merged onto ``os.environ.copy()`` for the child.

        Returns:
            ``True`` if ``QProcess.start`` was invoked, ``False`` if a process
            was already running or ``argv`` was invalid.
        """
        if self._process.state() != QProcess.ProcessState.NotRunning:
            self.error.emit("Process already running")
            return False
        if len(argv) < 1:
            self.error.emit("Empty argv")
            return False
        program = argv[0]
        arguments = [str(a) for a in argv[1:]]
        if env:
            qenv = QProcessEnvironment.systemEnvironment()
            for k, v in env.items():
                qenv.insert(k, str(v))
            self._process.setProcessEnvironment(qenv)
        else:
            self._process.setProcessEnvironment(QProcessEnvironment.systemEnvironment())
        if cwd is not None:
            self._process.setWorkingDirectory(str(Path(cwd).resolve()))
        self._process.start(program, arguments)
        if self._process.state() == QProcess.ProcessState.NotRunning:
            # Start failed synchronously; errorOccurred may also fire.
            return False
        return True

    def request_stop(self) -> None:
        """Ask the child to stop (terminate, then kill if still running)."""
        if self._process.state() == QProcess.ProcessState.NotRunning:
            return
        self._process.terminate()
        if not self._process.waitForFinished(3000):
            self._process.kill()

    def _on_started(self) -> None:
        self.started.emit()

    def _on_stdout(self) -> None:
        data = bytes(self._process.readAllStandardOutput()).decode(
            "utf-8",
            errors="replace",
        )
        if data:
            self.log_received.emit(data)

    def _on_stderr(self) -> None:
        data = bytes(self._process.readAllStandardError()).decode(
            "utf-8",
            errors="replace",
        )
        if data:
            self.log_received.emit(f"[stderr] {data}")

    def _on_finished(self, exit_code: int, status: QProcess.ExitStatus) -> None:
        self._drain_streams()
        # PySide6 ExitStatus is an Enum; int(status) raises TypeError on some builds.
        status_int = int(getattr(status, "value", status))
        self.finished.emit(exit_code, status_int)

    def _drain_streams(self) -> None:
        self._on_stdout()
        self._on_stderr()

    def _on_error_occurred(self, proc_error: QProcess.ProcessError) -> None:
        msg = self._process.errorString() or str(proc_error)
        self.error.emit(msg)
