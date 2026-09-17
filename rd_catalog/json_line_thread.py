"""Qt thread that pumps a child-process JSON-line job without holding the GIL.

Shared by RD↔robot MTO compare and export pair compare. Subclasses supply the
``python -m`` module name, the job JSON payload, and how to read a ``done``
line. Cooperative cancel is a flag file checked by the child between units of
work.
"""

from __future__ import annotations

import json
import os
import subprocess
import sys
import threading
from pathlib import Path
from typing import Any

from PySide6.QtCore import QThread, Signal

from rd_catalog.config import CatalogConfig

_PROJECT_ROOT = Path(__file__).resolve().parents[1]
_BELOW_NORMAL_PRIORITY_CLASS = 0x00004000


class JsonLineWorkerThread(QThread):
    """Start ``python -m <cli_module> --job <file>`` and forward JSON lines."""

    progress = Signal(int, int)
    log = Signal(str)
    error = Signal(str)

    _cli_module: str = ""
    _job_prefix: str = "job"
    _cancel_prefix: str = "cancel"
    _cancel_log: str = "Запрошена отмена…"
    _process_failure: str = "Процесс завершился"

    def __init__(
        self,
        config: CatalogConfig,
        *,
        batch_size: int = 1,
        parent=None,
    ) -> None:
        """Store configuration shared by every JSON-line child job.

        Args:
            config: Resolved catalog configuration.
            batch_size: Worker batch size written into the job file.
            parent: Optional Qt parent.
        """

        super().__init__(parent)
        self._config = config
        self._batch_size = int(batch_size)
        self._cancel_event = threading.Event()
        self._cancel_path: Path | None = None
        self._process: subprocess.Popen[str] | None = None
        self.comparison_count = 0
        self.planned = 0
        self.skipped_cache_hits = 0
        self.cancelled = False
        self.failure: str | None = None

    def request_cancel(self) -> None:
        """Request cooperative cancellation at the next child checkpoint."""

        self._cancel_event.set()
        if self._cancel_path is not None:
            try:
                self._cancel_path.write_text("1", encoding="utf-8")
            except OSError:
                pass
        self.log.emit(self._cancel_log)

    def _job_payload(self, cancel_path: str) -> dict[str, Any]:
        """Return the JSON object written next to the cancel flag.

        Args:
            cancel_path: Absolute path of the cooperative-cancel flag file.

        Returns:
            Job mapping understood by the child ``--job`` entry point.
        """

        raise NotImplementedError

    def _handle_done(self, payload: dict[str, Any]) -> None:
        """Apply subclass-specific fields from a ``t=done`` line.

        Args:
            payload: Decoded JSON object from the child.
        """

        return None

    def run(self) -> None:
        """Launch the worker process and pump stdout until exit."""

        runtime_dir = Path(self._config.runtime_dir)
        runtime_dir.mkdir(parents=True, exist_ok=True)
        stamp = f"{os.getpid()}_{threading.get_ident()}"
        job_path = runtime_dir / f"{self._job_prefix}_{stamp}.json"
        cancel_path = runtime_dir / f"{self._cancel_prefix}_{stamp}.flag"
        self._cancel_path = cancel_path
        job_path.write_text(
            json.dumps(
                self._job_payload(str(cancel_path)),
                ensure_ascii=False,
                indent=2,
            ),
            encoding="utf-8",
        )
        env = os.environ.copy()
        pythonpath = env.get("PYTHONPATH", "")
        root = str(_PROJECT_ROOT)
        env["PYTHONPATH"] = (
            root if not pythonpath else root + os.pathsep + pythonpath
        )
        env["PYTHONUTF8"] = "1"
        env["PYTHONIOENCODING"] = "utf-8"
        popen_kwargs: dict[str, object] = {
            "stdout": subprocess.PIPE,
            "stderr": subprocess.PIPE,
            "text": True,
            "encoding": "utf-8",
            "errors": "replace",
            "cwd": root,
            "env": env,
        }
        if sys.platform == "win32":
            popen_kwargs["creationflags"] = _BELOW_NORMAL_PRIORITY_CLASS
        try:
            process = subprocess.Popen(
                [
                    sys.executable,
                    "-m",
                    self._cli_module,
                    "--job",
                    str(job_path),
                ],
                **popen_kwargs,  # type: ignore[arg-type]
            )
            self._process = process
            assert process.stdout is not None
            stderr_chunks: list[str] = []

            def drain_stderr() -> None:
                if process.stderr is not None:
                    stderr_chunks.append(process.stderr.read())

            stderr_thread = threading.Thread(target=drain_stderr, daemon=True)
            stderr_thread.start()
            for raw_line in process.stdout:
                if self._cancel_event.is_set() and not cancel_path.exists():
                    try:
                        cancel_path.write_text("1", encoding="utf-8")
                    except OSError:
                        pass
                line = raw_line.strip()
                if not line:
                    continue
                try:
                    payload = json.loads(line)
                except json.JSONDecodeError:
                    self.log.emit(line)
                    continue
                self._handle_payload(payload)
            stderr_thread.join(timeout=5)
            stderr_text = "".join(stderr_chunks).strip()
            exit_code = process.wait()
            if self.failure is None and exit_code not in (0, 1):
                detail = stderr_text or f"код {exit_code}"
                self.failure = f"{self._process_failure}: {detail}"
                self.error.emit(self.failure)
            elif stderr_text and self.failure is None:
                self.log.emit(stderr_text)
        except Exception as exc:
            self.failure = f"{type(exc).__name__}: {exc}"
            self.error.emit(self.failure)
        finally:
            self._process = None
            for path in (job_path, cancel_path):
                try:
                    path.unlink(missing_ok=True)
                except OSError:
                    pass

    def _handle_payload(self, payload: object) -> None:
        if not isinstance(payload, dict):
            return
        kind = payload.get("t")
        if kind == "log":
            self.log.emit(str(payload.get("m") or ""))
            return
        if kind == "progress":
            self.progress.emit(
                int(payload.get("completed") or 0),
                int(payload.get("total") or 0),
            )
            return
        if kind == "error":
            message = str(payload.get("m") or "")
            self.failure = message.split("\n", 1)[0] or message
            self.error.emit(message)
            return
        if kind == "done":
            self.comparison_count = int(payload.get("comparison_count") or 0)
            self.planned = int(payload.get("planned") or 0)
            self.skipped_cache_hits = int(payload.get("skipped_cache_hits") or 0)
            self.cancelled = bool(payload.get("cancelled"))
            failure = payload.get("failure")
            if failure:
                self.failure = str(failure)
            self._handle_done(payload)
