"""Qt thread that pumps the RD file-stamp census without holding the GIL."""

from __future__ import annotations

import json
import os
import subprocess
import sys
import threading
from pathlib import Path

from PySide6.QtCore import QThread, Signal

from rd_catalog.config import CatalogConfig
from rd_catalog.rd_freshness import RdFileStamp

_PROJECT_ROOT = Path(__file__).resolve().parents[1]


class RdFreshnessThread(QThread):
    """Start ``python -m rd_catalog.rd_freshness_cli`` and read its JSON lines."""

    log = Signal(str)
    error = Signal(str)

    def __init__(
        self,
        config: CatalogConfig,
        baseline: tuple[RdFileStamp, ...],
        parent=None,
    ) -> None:
        """Store the census request.

        Args:
            config: Resolved catalog configuration.
            baseline: Present RD stamps already in the catalog.
            parent: Optional Qt parent.
        """

        super().__init__(parent)
        self._config = config
        self._baseline = baseline
        self._cancel_event = threading.Event()
        self._cancel_path: Path | None = None
        self._process: subprocess.Popen[str] | None = None
        self.failure: str | None = None
        self.seen = 0
        self.folders: tuple[str, ...] = ()
        self.cancelled = False

    def request_cancel(self) -> None:
        """Ask the child walk to stop at the next directory."""

        self._cancel_event.set()
        if self._cancel_path is not None:
            try:
                self._cancel_path.write_text("1", encoding="utf-8")
            except OSError:
                pass

    def run(self) -> None:
        """Launch the census process and pump stdout until it exits."""

        runtime_dir = Path(self._config.runtime_dir)
        runtime_dir.mkdir(parents=True, exist_ok=True)
        stamp = f"{os.getpid()}_{threading.get_ident()}"
        job_path = runtime_dir / f"rd_freshness_job_{stamp}.json"
        cancel_path = runtime_dir / f"rd_freshness_cancel_{stamp}.flag"
        self._cancel_path = cancel_path
        job_path.write_text(
            json.dumps(
                {
                    "rd_root": str(self._config.rd_root),
                    "sq_root": str(self._config.sq_root or ""),
                    "skip_dirs": list(self._config.skip_dirs),
                    "cancel_path": str(cancel_path),
                    "baseline": [
                        {
                            "path": item.path,
                            "path_key": item.path_key,
                            "size": item.size,
                            "mtime_ns": item.mtime_ns,
                        }
                        for item in self._baseline
                    ],
                },
                ensure_ascii=False,
            ),
            encoding="utf-8",
        )
        env = os.environ.copy()
        pythonpath = env.get("PYTHONPATH", "")
        root = str(_PROJECT_ROOT)
        env["PYTHONPATH"] = root if not pythonpath else root + os.pathsep + pythonpath
        env["PYTHONUTF8"] = "1"
        env["PYTHONIOENCODING"] = "utf-8"
        creationflags = 0
        if sys.platform == "win32":
            creationflags = subprocess.BELOW_NORMAL_PRIORITY_CLASS
        try:
            process = subprocess.Popen(
                [sys.executable, "-m", "rd_catalog.rd_freshness_cli", "--job", str(job_path)],
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
                encoding="utf-8",
                errors="replace",
                cwd=root,
                env=env,
                creationflags=creationflags,
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
            exit_code = process.wait()
            stderr_text = "".join(stderr_chunks).strip()
            if self.failure is None and exit_code not in (0, 1):
                detail = stderr_text or f"код {exit_code}"
                self.failure = f"Список РД завершился: {detail}"
                self.error.emit(self.failure)
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
        if kind == "error":
            message = str(payload.get("m") or "")
            self.failure = message.split("\n", 1)[0] or message
            self.error.emit(message)
            return
        if kind == "done":
            self.seen = int(payload.get("seen") or 0)
            self.cancelled = bool(payload.get("cancelled"))
            raw_folders = payload.get("folders") or []
            self.folders = tuple(
                str(item) for item in raw_folders if str(item or "").strip()
            )
