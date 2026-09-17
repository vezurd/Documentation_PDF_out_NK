"""Qt controller that runs ``python -m rd_catalog_web`` as a child process."""

from __future__ import annotations

import os
import queue
import subprocess
import sys
import threading
from pathlib import Path

from PySide6.QtCore import QObject, QTimer, Signal

from rd_catalog_web.urls import (
    DEFAULT_HOST,
    DEFAULT_PORT,
    lan_share_url,
    listen_urls,
    local_open_url,
)

_PROJECT_ROOT = Path(__file__).resolve().parents[1]
_CREATE_NO_WINDOW = 0x08000000
# Printed by uvicorn after the socket is bound. Do not match listen_urls()
# lines that appear *before* uvicorn.run().
_READY_MARKERS = ("uvicorn running", "application startup complete")


class WebServerProcess(QObject):
    """Long-running uvicorn child. Not a catalog worker (scan stays free)."""

    ready = Signal()
    failed = Signal(str)
    stopped = Signal()
    log_line = Signal(str)

    def __init__(self, parent: QObject | None = None) -> None:
        """Create a stopped controller.

        Args:
            parent: Optional Qt parent (the catalog window).
        """

        super().__init__(parent)
        self.host = DEFAULT_HOST
        self.port = DEFAULT_PORT
        self._process: subprocess.Popen[str] | None = None
        self._ready = False
        self._out_queue: queue.SimpleQueue[tuple[str, str]] = queue.SimpleQueue()
        self._watch = QTimer(self)
        self._watch.setInterval(400)
        self._watch.timeout.connect(self._check_exit)
        self._drain = QTimer(self)
        self._drain.setInterval(80)
        self._drain.timeout.connect(self._drain_output)

    @property
    def running(self) -> bool:
        """True while the child process is alive."""

        process = self._process
        return process is not None and process.poll() is None

    @property
    def is_ready(self) -> bool:
        """True after uvicorn accepted connections."""

        return self._ready and self.running

    def local_url(self) -> str:
        """Return the loopback URL for this PC's browser."""

        return local_open_url(self.host, self.port)

    def share_url(self) -> str:
        """Return the LAN URL for colleagues (not ``0.0.0.0``)."""

        return lan_share_url(self.host, self.port)

    def display_urls(self) -> list[str]:
        """Return loopback plus LAN URLs (never ``0.0.0.0``)."""

        return listen_urls(self.host, self.port)

    def start(self) -> None:
        """Spawn the monitor if it is not already running."""

        if self.running:
            if self._ready:
                self.ready.emit()
            return
        self._ready = False
        env = os.environ.copy()
        root = str(_PROJECT_ROOT)
        pythonpath = env.get("PYTHONPATH", "")
        env["PYTHONPATH"] = root if not pythonpath else root + os.pathsep + pythonpath
        env["PYTHONUTF8"] = "1"
        env["PYTHONIOENCODING"] = "utf-8"
        kwargs: dict[str, object] = {
            "stdout": subprocess.PIPE,
            "stderr": subprocess.STDOUT,
            "text": True,
            "encoding": "utf-8",
            "errors": "replace",
            "cwd": root,
            "env": env,
        }
        if sys.platform == "win32":
            kwargs["creationflags"] = _CREATE_NO_WINDOW
        try:
            process = subprocess.Popen(
                [
                    sys.executable,
                    "-m",
                    "rd_catalog_web",
                    "--host",
                    self.host,
                    "--port",
                    str(self.port),
                ],
                **kwargs,
            )
        except OSError as exc:
            self.failed.emit(str(exc))
            return
        self._process = process
        thread = threading.Thread(target=self._pump_stdout, daemon=True)
        thread.start()
        self._drain.start()
        self._watch.start()

    def stop(self, *, wait_ms: int = 2000) -> None:
        """Terminate the child process if it is running.

        Args:
            wait_ms: Milliseconds to wait after terminate before kill.
        """

        self._watch.stop()
        self._drain.stop()
        process = self._process
        self._process = None
        self._ready = False
        if process is None:
            return
        if process.poll() is None:
            process.terminate()
            try:
                process.wait(timeout=max(wait_ms, 1) / 1000)
            except subprocess.TimeoutExpired:
                process.kill()
                try:
                    process.wait(timeout=1)
                except subprocess.TimeoutExpired:
                    pass
        self.stopped.emit()

    def _pump_stdout(self) -> None:
        process = self._process
        if process is None or process.stdout is None:
            return
        for raw in process.stdout:
            line = raw.rstrip()
            if line:
                self._out_queue.put(("log", line))
            lowered = line.casefold()
            if any(marker in lowered for marker in _READY_MARKERS):
                self._out_queue.put(("ready", ""))

    def _drain_output(self) -> None:
        while True:
            try:
                kind, payload = self._out_queue.get_nowait()
            except queue.Empty:
                break
            if kind == "log":
                self.log_line.emit(payload)
            elif kind == "ready" and not self._ready and self.running:
                self._ready = True
                self.ready.emit()

    def _check_exit(self) -> None:
        process = self._process
        if process is None:
            self._watch.stop()
            self._drain.stop()
            return
        code = process.poll()
        if code is None:
            return
        self._watch.stop()
        self._drain.stop()
        self._drain_output()
        self._process = None
        was_ready = self._ready
        self._ready = False
        if code != 0 and not was_ready:
            hint = (
                "pip install fastapi uvicorn"
                if code == 2
                else f"процесс завершился с кодом {code}"
            )
            self.failed.emit(hint)
            return
        self.stopped.emit()
