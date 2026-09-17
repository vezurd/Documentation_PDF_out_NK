"""Debounced persistence of Qt layout snapshots into ``pdf_v2_config.json``."""

from __future__ import annotations

import json
from typing import Any, Callable

from PySide6.QtCore import QObject, QTimer


class DebouncedLayoutSaver(QObject):
    """Compare JSON snapshot every ``interval_ms``; write only when changed."""

    def __init__(
        self,
        parent: QObject | None,
        *,
        interval_ms: int = 1000,
        snapshot: Callable[[], dict[str, Any]],
        on_flush: Callable[[dict[str, Any]], None],
    ) -> None:
        super().__init__(parent)
        self._interval_ms = max(100, int(interval_ms))
        self._snapshot = snapshot
        self._on_flush = on_flush
        self._last_written_json = ""
        self._timer = QTimer(self)
        self._timer.setSingleShot(True)
        self._timer.timeout.connect(self._flush_if_changed)

    def schedule(self) -> None:
        self._timer.stop()
        self._timer.start(self._interval_ms)

    def flush_now(self) -> None:
        self._timer.stop()
        self._flush_if_changed()

    def sync_last_written_from_snapshot(self, snap: dict[str, Any]) -> None:
        self._last_written_json = json.dumps(snap, sort_keys=True, ensure_ascii=False)

    def _flush_if_changed(self) -> None:
        snap = self._snapshot()
        blob = json.dumps(snap, sort_keys=True, ensure_ascii=False)
        if blob == self._last_written_json:
            return
        self._on_flush(snap)
        self._last_written_json = blob
