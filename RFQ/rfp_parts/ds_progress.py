"""Live stdout progress for DS cockpit jobs (RFP · Сбор частей tab).

Lines use prefix ``[ds progress]`` so the Job monitor can update the bar.
"""

from __future__ import annotations

import time
from collections.abc import Callable
from typing import Protocol

DS_PROGRESS_PREFIX = "[ds progress]"
DS_PROGRESS_SLOW_THRESHOLD_SEC = 30.0
DS_PROGRESS_HEARTBEAT_SEC = 2.0


def ds_progress_format_start(label: str) -> str:
    """Format a phase START line."""
    return f"{DS_PROGRESS_PREFIX} START: {label}"


def ds_progress_format_done(label: str, elapsed_sec: float, total_sec: float) -> str:
    """Format a phase DONE line."""
    return (
        f"{DS_PROGRESS_PREFIX} DONE: {label} — "
        f"{elapsed_sec:.2f} с; всего {total_sec:.2f} с"
    )


def ds_progress_format_message(text: str) -> str:
    """Format an informational line (no timing)."""
    return f"{DS_PROGRESS_PREFIX} {text}"


def ds_progress_format_file_start(
    label: str,
    index: int,
    total: int,
    relpath: str,
    total_sec: float,
) -> str:
    """Format start of parsing one workbook."""
    name = relpath.replace("\\", "/").split("/")[-1] if relpath else "—"
    return (
        f"{DS_PROGRESS_PREFIX} PROGRESS: {label} {index}/{total} "
        f"{name}; всего {total_sec:.1f} с"
    )


def ds_progress_format_file_done(
    label: str,
    index: int,
    total: int,
    relpath: str,
    file_sec: float,
    total_sec: float,
) -> str:
    """Format completion of one workbook."""
    name = relpath.replace("\\", "/").split("/")[-1] if relpath else "—"
    return (
        f"{DS_PROGRESS_PREFIX} PROGRESS: {label} {index}/{total} "
        f"{name} — {file_sec:.1f} с; всего {total_sec:.1f} с"
    )


def ds_progress_format_slow_file(
    label: str,
    index: int,
    total: int,
    relpath: str,
    file_sec: float,
) -> str:
    """Format a slow workbook line."""
    name = relpath.replace("\\", "/").split("/")[-1] if relpath else "—"
    return (
        f"{DS_PROGRESS_PREFIX} SLOW: {label} {index}/{total} "
        f"{name} — {file_sec:.1f} с"
    )


def ds_progress_format_fraction(current: int, total: int) -> str:
    """Machine-readable fraction for the GUI progress bar."""
    return f"{DS_PROGRESS_PREFIX} FRACTION: {current}/{total}"


class EmitFn(Protocol):
    def __call__(self, message: str) -> None: ...


class DsProgressSession:
    """Phase timing with START/DONE lines on stdout."""

    def __init__(self, emit: EmitFn) -> None:
        self._emit = emit
        self._baseline = time.perf_counter()

    def elapsed_total(self) -> float:
        return time.perf_counter() - self._baseline

    def phase_start(self, label: str) -> None:
        self._emit(ds_progress_format_start(label))

    def phase_done(self, label: str, phase_sec: float) -> None:
        self._emit(
            ds_progress_format_done(label, phase_sec, self.elapsed_total())
        )

    def message(self, text: str) -> None:
        self._emit(ds_progress_format_message(text))

    def fraction(self, current: int, total: int) -> None:
        if total <= 0:
            return
        self._emit(ds_progress_format_fraction(current, total))


class DsHeartbeat:
    """Emit a line at least every ``interval_sec`` while a long write runs.

    ``emit`` is the same callback jobs already wrap with ``[ds progress]``.
    The first automatic tick waits for ``interval_sec`` so a fast phase stays
    quiet; use ``force=True`` before a blocking save.
    """

    def __init__(
        self,
        emit: Callable[[str], None] | None,
        label: str,
        *,
        interval_sec: float = DS_PROGRESS_HEARTBEAT_SEC,
    ) -> None:
        self._emit = emit
        self._label = label
        self._interval = max(0.5, interval_sec)
        self._t0 = time.perf_counter()
        self._last = self._t0

    def elapsed(self) -> float:
        return time.perf_counter() - self._t0

    def tick(self, detail: str, *, force: bool = False) -> None:
        if self._emit is None:
            return
        now = time.perf_counter()
        if not force and now - self._last < self._interval:
            return
        self._last = now
        self._emit(f"{self._label}: {detail}; {now - self._t0:.1f} с")

    def finish(self, detail: str = "готово") -> None:
        if self._emit is None:
            return
        elapsed = time.perf_counter() - self._t0
        self._emit(f"{self._label}: {detail} — {elapsed:.1f} с")


class DsFileProgressTracker:
    """Callback for ``build_ds_baseline`` / hybrid RFP loops."""

    def __init__(
        self,
        emit: EmitFn,
        *,
        label: str,
        slow_threshold_sec: float = DS_PROGRESS_SLOW_THRESHOLD_SEC,
        emit_every: int = 1,
    ) -> None:
        self._emit = emit
        self._label = label
        self._slow = slow_threshold_sec
        self._every = max(1, emit_every)
        self._session = DsProgressSession(emit)
        self._total_files = 0

    def bind_total(self, total: int) -> None:
        self._total_files = total
        self._session.message(f"{self._label}: найдено файлов {total}")

    def baseline_callback(
        self,
    ) -> Callable[[int, int, str, float | None], None]:
        """Return ``(index, total, relpath, file_elapsed)`` hook."""

        def hook(
            index: int,
            total: int,
            relpath: str,
            file_elapsed: float | None,
        ) -> None:
            if file_elapsed is None:
                if index == 1 or index % self._every == 0 or index == total:
                    self._session.fraction(index - 1, total)
                    self._emit(
                        ds_progress_format_file_start(
                            self._label,
                            index,
                            total,
                            relpath,
                            self._session.elapsed_total(),
                        )
                    )
                return
            if (
                index % self._every == 0
                or index == total
                or file_elapsed >= self._slow
            ):
                self._session.fraction(index, total)
                self._emit(
                    ds_progress_format_file_done(
                        self._label,
                        index,
                        total,
                        relpath,
                        file_elapsed,
                        self._session.elapsed_total(),
                    )
                )
            if file_elapsed >= self._slow:
                self._emit(
                    ds_progress_format_slow_file(
                        self._label,
                        index,
                        total,
                        relpath,
                        file_elapsed,
                    )
                )

        return hook

    def rfp_callback(
        self,
    ) -> Callable[[int, int, str, float | None], None]:
        return self.baseline_callback()
