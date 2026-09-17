"""Nested operation timing for RD catalog hang profiling.

Wrap coarse GUI and domain operations only — not per-cell paint, not
every helper. Disabled until :func:`configure_perf_log` (or a sink is
set). ``RD_CATALOG_PERF=0`` / ``false`` / ``off`` turns everything off.

The journal is ``runtime_dir/perf.log`` (rotated at 8 MiB) and an optional
sink (the Qt «Журнал» tab). This module is Qt-free and must not import
``window`` / ``pipeline``.
"""

from __future__ import annotations

import os
import re
import threading
import time
from collections.abc import Callable, Iterator, Mapping
from contextlib import contextmanager
from datetime import datetime
from pathlib import Path
from typing import Any

PERF_LOG_NAME = "perf.log"
PERF_ENV_VAR = "RD_CATALOG_PERF"
_MAX_BYTES = 8 * 1024 * 1024
_TS_PREFIX = re.compile(r"^\d{4}-\d{2}-\d{2} ")
_FIELD_MAX = 120

_lock = threading.Lock()
_local = threading.local()
_path: Path | None = None
_sink: Callable[[str], None] | None = None


def _env_disabled() -> bool:
    raw = os.environ.get(PERF_ENV_VAR, "").strip().casefold()
    return raw in {"0", "false", "off", "no"}


def _stack() -> list[str]:
    stack = getattr(_local, "stack", None)
    if stack is None:
        stack = []
        _local.stack = stack
    return stack


def _timestamp() -> str:
    now = datetime.now()
    return now.strftime("%Y-%m-%d %H:%M:%S.") + f"{now.microsecond // 1000:03d}"


def _thread_label() -> str:
    name = threading.current_thread().name
    if name == "MainThread":
        return "gui"
    return name.replace(" ", "_")[:24]


def _format_fields(fields: Mapping[str, Any]) -> str:
    parts: list[str] = []
    for key, value in fields.items():
        if value is None or value == "":
            continue
        if isinstance(value, bool):
            text = "1" if value else "0"
        else:
            text = str(value).replace("\n", " ").replace("\r", " ")
        if len(text) > _FIELD_MAX:
            text = text[: _FIELD_MAX - 3] + "..."
        parts.append(f"{key}={text}")
    return (" " + " ".join(parts)) if parts else ""


def _write_file(line: str) -> None:
    path = _path
    if path is None:
        return
    try:
        path.parent.mkdir(parents=True, exist_ok=True)
        if path.is_file() and path.stat().st_size >= _MAX_BYTES:
            backup = path.with_name(path.name + ".1")
            try:
                backup.unlink(missing_ok=True)
            except OSError:
                pass
            try:
                path.replace(backup)
            except OSError:
                pass
        with path.open("a", encoding="utf-8") as handle:
            handle.write(line + "\n")
    except OSError:
        return


def _emit(line: str) -> None:
    if _env_disabled():
        return
    sink = _sink
    with _lock:
        _write_file(line)
    if sink is not None:
        try:
            sink(line)
        except Exception:
            return


def is_perf_enabled() -> bool:
    """Return whether spans currently write to a file and/or sink."""

    if _env_disabled():
        return False
    return _path is not None or _sink is not None


def configure_perf_log(runtime_dir: str | Path | None) -> Path | None:
    """Set the journal file to ``runtime_dir/perf.log``.

    Does not create directories until the first line is written. ``None``
    disables the file sink (an optional callback may still receive lines).

    Args:
        runtime_dir: Catalog runtime directory, or ``None`` to disable file
            output.

    Returns:
        The journal path when a directory was given, otherwise ``None``.
    """

    global _path
    if runtime_dir is None:
        _path = None
        return None
    _path = Path(runtime_dir) / PERF_LOG_NAME
    return _path


def set_perf_sink(callback: Callable[[str], None] | None) -> None:
    """Install or clear an extra line sink (typically the Qt journal tab)."""

    global _sink
    _sink = callback


def reset_perf_log_for_tests() -> None:
    """Clear file path, sink, and the calling thread's nest stack."""

    global _path, _sink
    _path = None
    _sink = None
    _local.stack = []


def stamp_log_line(message: str) -> str:
    """Prefix ``message`` with a local timestamp unless it already has one.

    Args:
        message: Raw journal text, possibly already stamped by ``perf_span``.

    Returns:
        A single line without a trailing newline, or ``""`` when empty.
    """

    text = message.rstrip()
    if not text:
        return ""
    if _TS_PREFIX.match(text):
        return text
    return f"{_timestamp()} {text}"


def perf_note(name: str, **fields: Any) -> None:
    """Write one unpaired NOTE line. Never raises."""

    if not is_perf_enabled():
        return
    try:
        depth = len(_stack())
        indent = "  " * depth
        line = (
            f"{_timestamp()} [{_thread_label()}] {indent}NOTE  {name}"
            f"{_format_fields(fields)}"
        )
        _emit(line)
    except Exception:
        return


@contextmanager
def perf_span(name: str, **fields: Any) -> Iterator[None]:
    """Log BEGIN on enter and END with ``elapsed_ms`` on exit.

    Nested spans indent. Exceptions still emit END (plus ``exc=``) and
    are re-raised. Never raises its own errors.

    Args:
        name: Dotted label, e.g. ``gui.refresh`` or ``pipeline.rebuild``.
        **fields: Optional scalar context (paths, counts, flags).
    """

    if not is_perf_enabled():
        yield
        return
    stack = _stack()
    extra = _format_fields(fields)
    try:
        depth = len(stack)
        stack.append(name)
        _emit(
            f"{_timestamp()} [{_thread_label()}] {'  ' * depth}BEGIN {name}{extra}"
        )
    except Exception:
        pass
    started = time.perf_counter()
    error = ""
    try:
        yield
    except BaseException as exc:
        error = type(exc).__name__
        raise
    finally:
        elapsed_ms = int((time.perf_counter() - started) * 1000)
        try:
            if stack and stack[-1] == name:
                stack.pop()
            elif stack:
                stack.pop()
            depth = len(stack)
            suffix = extra
            if error:
                suffix += f" exc={error}"
            _emit(
                f"{_timestamp()} [{_thread_label()}] {'  ' * depth}END   {name}"
                f" elapsed_ms={elapsed_ms}{suffix}"
            )
        except Exception:
            return
