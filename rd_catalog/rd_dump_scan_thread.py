"""Qt thread that pumps a child-process RD dump scan without holding the GIL."""

from __future__ import annotations

import json
import os
import subprocess
import sys
import threading
from pathlib import Path

from PySide6.QtCore import QThread, Signal

from rd_catalog.config import CatalogConfig
from rd_catalog.rd_dump_scan import RdDumpScanProgress, rd_dump_job_from_config

_PROJECT_ROOT = Path(__file__).resolve().parents[1]


class RdDumpScanThread(QThread):
    """Start ``python -m rd_catalog.rd_dump_scan_cli`` and forward JSON lines."""

    progress = Signal(object)
    log = Signal(str)
    error = Signal(str)

    def __init__(self, config: CatalogConfig, parent=None) -> None:
        """Store an immutable RD dump scan request.

        Args:
            config: Resolved catalog configuration (``rd_root``, ``db_path``,
                ``runtime_dir``).
            parent: Optional Qt parent.
        """

        super().__init__(parent)
        self._config = config
        self._cancel_event = threading.Event()
        self._cancel_path: Path | None = None
        self._process: subprocess.Popen[str] | None = None
        self.failure: str | None = None
        self.accepted = 0
        self.files_seen = 0
        self.scanned_at = ""
        self.cancelled = False

    def request_cancel(self) -> None:
        """Request cooperative cancellation at the next walker checkpoint."""

        self._cancel_event.set()
        if self._cancel_path is not None:
            try:
                self._cancel_path.write_text("1", encoding="utf-8")
            except OSError:
                pass
        self.log.emit("Запрошена отмена скана РД (xlsx). Завершаем текущую операцию…")

    def run(self) -> None:
        """Launch the RD dump scan worker process and pump stdout until exit."""

        runtime_dir = Path(self._config.runtime_dir)
        runtime_dir.mkdir(parents=True, exist_ok=True)
        stamp = f"{os.getpid()}_{threading.get_ident()}"
        job_path = runtime_dir / f"rd_dump_scan_job_{stamp}.json"
        cancel_path = runtime_dir / f"rd_dump_scan_cancel_{stamp}.flag"
        self._cancel_path = cancel_path
        job_path.write_text(
            json.dumps(
                rd_dump_job_from_config(self._config, cancel_path=str(cancel_path)),
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
        try:
            process = subprocess.Popen(
                [
                    sys.executable,
                    "-m",
                    "rd_catalog.rd_dump_scan_cli",
                    "--job",
                    str(job_path),
                ],
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
                encoding="utf-8",
                errors="replace",
                cwd=root,
                env=env,
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
                self.failure = f"Процесс скана РД (xlsx) завершился: {detail}"
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
                RdDumpScanProgress(
                    path=str(payload.get("path") or ""),
                    files_seen=int(payload.get("files_seen") or 0),
                    message=str(payload.get("message") or ""),
                )
            )
            return
        if kind == "error":
            message = str(payload.get("m") or "")
            self.failure = message.split("\n", 1)[0] or message
            self.error.emit(message)
            return
        if kind == "done":
            self.accepted = int(payload.get("accepted") or 0)
            self.files_seen = int(payload.get("files_seen") or 0)
            self.scanned_at = str(payload.get("scanned_at") or "")
            self.cancelled = bool(payload.get("cancelled"))
            failure = payload.get("failure")
            if failure:
                self.failure = str(failure)
