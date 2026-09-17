"""Run Python callables off the GUI thread with streamed stdout/stderr text."""

from __future__ import annotations

import sys
import threading
import traceback
from collections.abc import Callable, Mapping
from dataclasses import dataclass
from io import TextIOBase
from typing import Any

from PySide6.QtCore import QObject, Signal


@dataclass(frozen=True)
class ActionResult:
    """Structured return value from a job callable.

    Attributes:
        success: Whether the action completed successfully.
        message: Short human-readable summary for the UI.
        result_path: Optional filesystem path (e.g. for ``JobMonitorPanel``).
    """

    success: bool
    message: str
    result_path: str | None = None


class _MirrorTextStream(TextIOBase):
    """Forward writes to a callback and optionally to an original stream."""

    def __init__(
        self,
        base: Any,
        on_write: Callable[[str], None],
    ) -> None:
        super().__init__()
        self._base = base
        self._on_write = on_write

    def writable(self) -> bool:
        return True

    def write(self, s: str) -> int:
        if not isinstance(s, str):
            s = str(s)
        self._on_write(s)
        if self._base is not None:
            try:
                return self._base.write(s)
            except Exception:
                return len(s)
        return len(s)

    def flush(self) -> None:
        if self._base is not None:
            try:
                self._base.flush()
            except Exception:
                pass

    @property
    def encoding(self) -> str:
        enc = getattr(self._base, "encoding", None) if self._base is not None else None
        return enc if isinstance(enc, str) else "utf-8"


class FunctionJobRunner(QObject):
    """Runs a callable in a worker thread without blocking the Qt event loop.

    Stdout and stderr are temporarily mirrored: output is forwarded to the
    previous process streams (if any) and emitted through :attr:`log_received`
    in chunks as :meth:`write` is called. Restoration happens in a ``finally``
    block so ``sys.stdout`` / ``sys.stderr`` are not left redirected.

    Only one job per runner instance is allowed; a second :meth:`start` emits
    :attr:`error` and returns ``False``.

    Signals:
        started: Emitted when the worker thread begins executing the callable.
        log_received: Text chunk (stdout/stderr interleaved as produced).
        finished: ``(success, message, result_path)`` when the callable returns
            or raises. ``result_path`` is ``None`` when not applicable.
        error: For non-startable jobs (already running) or unexpected return
            types. On uncaught exceptions, traceback text is emitted here and
            :attr:`finished` is emitted with ``success=False``.
    """

    started = Signal()
    log_received = Signal(str)
    finished = Signal(bool, str, object)
    error = Signal(str)

    def __init__(self, parent: QObject | None = None) -> None:
        super().__init__(parent)
        self._lock = threading.Lock()
        self._running = False

    def is_running(self) -> bool:
        """Return whether a worker thread is currently running a job."""
        with self._lock:
            return self._running

    def start(
        self,
        title: str,
        fn: Callable[..., ActionResult | str | None],
        *args: Any,
        **kwargs: Any,
    ) -> bool:
        """Start ``fn(*args, **kwargs)`` in a background thread.

        Args:
            title: Short label for logs (parent may also pass it to
                :meth:`JobMonitorPanel.start_job`).
            fn: Callable returning :class:`ActionResult`, a ``str`` (treated as
                ``result_path`` on success), or ``None`` (success, no path).
            *args: Positional arguments for ``fn``.
            **kwargs: Keyword arguments for ``fn``.

        Returns:
            ``True`` if a thread was started, ``False`` if a job was already
            running (also emits :attr:`error`).
        """
        with self._lock:
            if self._running:
                self.error.emit("Job already running")
                return False
            self._running = True

        thread = threading.Thread(
            target=self._thread_main,
            name="FunctionJobRunner",
            args=(title, fn, args, kwargs),
            daemon=False,
        )
        thread.start()
        return True

    def _thread_main(
        self,
        title: str,
        fn: Callable[..., Any],
        args: tuple[Any, ...],
        kwargs: Mapping[str, Any],
    ) -> None:
        try:
            self.started.emit()
            self.log_received.emit(f"--- start: {title} ---\n")
            old_out = sys.stdout
            old_err = sys.stderr
            try:
                sys.stdout = _MirrorTextStream(
                    old_out if self._is_text_stream(old_out) else None,
                    self.log_received.emit,
                )
                sys.stderr = _MirrorTextStream(
                    old_err if self._is_text_stream(old_err) else None,
                    self._emit_stderr_chunk,
                )
                raw = fn(*args, **kwargs)
            finally:
                try:
                    sys.stdout.flush()
                    sys.stderr.flush()
                except Exception:
                    pass
                sys.stdout = old_out
                sys.stderr = old_err

            invalid_msg, success, message, result_path = self._normalize_result(raw)
            if invalid_msg is not None:
                self.error.emit(invalid_msg)
                self.finished.emit(False, invalid_msg, None)
                return

            self.log_received.emit(f"--- end: {title} ---\n")
            self.finished.emit(success, message, result_path)
        except Exception:
            tb = traceback.format_exc()
            self.error.emit(tb)
            exc = sys.exc_info()[1]
            short = (
                f"{type(exc).__name__}: {exc}"
                if exc is not None
                else "Unhandled exception"
            )
            self.finished.emit(False, short, None)
        finally:
            with self._lock:
                self._running = False

    @staticmethod
    def _is_text_stream(stream: Any) -> bool:
        return hasattr(stream, "write") and hasattr(stream, "flush")

    def _emit_stderr_chunk(self, chunk: str) -> None:
        self.log_received.emit(f"[stderr] {chunk}")

    @staticmethod
    def _normalize_result(
        raw: Any,
    ) -> tuple[str | None, bool, str, str | None]:
        """Returns ``(invalid_message, success, message, result_path)``.

        When ``invalid_message`` is not ``None``, the return type was invalid.
        """
        if isinstance(raw, ActionResult) or all(
            hasattr(raw, name) for name in ("success", "message", "result_path")
        ):
            return (
                None,
                bool(raw.success),
                str(raw.message),
                str(raw.result_path) if raw.result_path is not None else None,
            )
        if isinstance(raw, str):
            return None, True, "Finished.", raw
        if raw is None:
            return None, True, "Finished.", None
        msg = (
            f"Callable must return ActionResult, str, or None; "
            f"got {type(raw).__name__}."
        )
        return msg, False, msg, None
