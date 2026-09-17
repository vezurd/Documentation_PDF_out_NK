"""Qt thread that pumps a child-process catalog scan without holding the GIL."""

from __future__ import annotations

import json
import os
import subprocess
import sys
import threading
from collections.abc import Iterable
from pathlib import Path

from PySide6.QtCore import QThread, Signal

from rd_catalog.config import CatalogConfig
from rd_catalog.models import ScanProgress, SourceKind
from rd_catalog.scan import ScanSubtree, normalize_scan_subtrees
from rd_catalog.scan_job import config_to_job_dict

_PROJECT_ROOT = Path(__file__).resolve().parents[1]


class ScanThread(QThread):
    """Start ``python -m rd_catalog.scan_cli`` and forward JSON status lines."""

    progress = Signal(object)
    log = Signal(str)
    error = Signal(str)

    def __init__(
        self,
        config: CatalogConfig,
        sources: Iterable[SourceKind],
        parent=None,
        *,
        robot_subtree: str | Path | None = None,
        rd_subtree: ScanSubtree | None = None,
        sq_subtree: str | Path | None = None,
    ) -> None:
        """Store an immutable scan request.

        Args:
            config: Resolved catalog configuration.
            sources: Source subset to scan.
            parent: Optional Qt parent.
            robot_subtree: Optional folder under ``robot_root`` for a scoped
                robot rescan after a kit MTO copy.
            rd_subtree: Optional folder or folders under ``rd_root`` after an
                SQ→RD move (the kit gate, e.g. «Для передачи», not the new
                ``NN_рев`` package) or a kit point rescan.
            sq_subtree: Optional folder under ``sq_root`` after an SQ→RD move.
        """

        super().__init__(parent)
        self._config = config
        self._sources = tuple(sources)
        self._robot_subtree = (
            str(robot_subtree) if robot_subtree is not None else None
        )
        self._rd_subtrees = normalize_scan_subtrees(rd_subtree)
        self._rd_subtree = (
            self._rd_subtrees[0] if self._rd_subtrees else None
        )
        self._sq_subtree = str(sq_subtree) if sq_subtree is not None else None
        self._cancel_event = threading.Event()
        self._cancel_path: Path | None = None
        self._process: subprocess.Popen[str] | None = None
        self.run_id: int | None = None
        self.comparison_count = 0
        self.failure: str | None = None
        self.touched_mto_keys: list[tuple[str, str]] = []
        self.compare_enqueue_scope: str = "none"
        self.mto_compare_skipped: bool = True

    def request_cancel(self) -> None:
        """Request cooperative cancellation at the next scanner checkpoint."""

        self._cancel_event.set()
        if self._cancel_path is not None:
            try:
                self._cancel_path.write_text("1", encoding="utf-8")
            except OSError:
                pass
        self.log.emit("Запрошена отмена. Завершаем текущую операцию…")

    def run(self) -> None:
        """Launch the scan worker process and pump its stdout until exit."""

        runtime_dir = Path(self._config.runtime_dir)
        runtime_dir.mkdir(parents=True, exist_ok=True)
        stamp = f"{os.getpid()}_{threading.get_ident()}"
        job_path = runtime_dir / f"scan_job_{stamp}.json"
        cancel_path = runtime_dir / f"scan_cancel_{stamp}.flag"
        self._cancel_path = cancel_path
        job_path.write_text(
            json.dumps(
                config_to_job_dict(
                    self._config,
                    self._sources,
                    robot_subtree=self._robot_subtree,
                    rd_subtree=self._rd_subtrees or None,
                    sq_subtree=self._sq_subtree,
                    cancel_path=str(cancel_path),
                ),
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
                [sys.executable, "-m", "rd_catalog.scan_cli", "--job", str(job_path)],
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
                self.failure = f"Процесс скана завершился: {detail}"
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
            try:
                source = SourceKind(str(payload.get("source")))
            except ValueError:
                return
            self.progress.emit(
                ScanProgress(
                    source=source,
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
            run_id = payload.get("run_id")
            self.run_id = int(run_id) if run_id is not None else None
            self.comparison_count = int(payload.get("comparison_count") or 0)
            raw_keys = payload.get("touched_mto_keys") or []
            self.touched_mto_keys = [
                (str(key[0]), str(key[1]))
                for key in raw_keys
                if isinstance(key, (list, tuple)) and len(key) == 2
            ]
            self.mto_compare_skipped = bool(
                payload.get("mto_compare_skipped", True)
            )
            self.compare_enqueue_scope = str(
                payload.get("compare_enqueue_scope") or "none"
            )
            failure = payload.get("failure")
            if failure:
                self.failure = str(failure)
