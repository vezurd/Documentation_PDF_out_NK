"""In-memory context-menu usage counters with delayed JSON persist.

Clicks never touch the disk. A background timer (and window close) writes
``runtime_dir/context_menu_usage.json``, re-reads it, and refreshes the
cached bold set. The Qt helper applies that set on the next menu popup.

Never UNC. Qt-free: do not import ``window`` / widgets.
"""

from __future__ import annotations

import json
import math
import threading
from collections.abc import Mapping
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

CONTEXT_MENU_USAGE_FILENAME = "context_menu_usage.json"
CONTEXT_MENU_USAGE_VERSION = 1
FLUSH_INTERVAL_SEC = 120
HOT_MIN_COUNT = 3
HOT_SHARE_OF_MAX = 0.35
HOT_MAX_PER_MENU = 8

MENU_KITS = "kits"
MENU_KIT_PACKAGE = "kit_package"
MENU_HEATMAP = "heatmap"
MENU_MTO_WORKLIST = "mto_worklist"
MENU_AN = "an"
MENU_RD_DUMP = "rd_dump"
MENU_HANDOFF_EXPORT = "handoff_export"
MENU_MTO_READINESS = "mto_readiness"
MENU_ROBOT_MTO_ACCEPT = "robot_mto_accept"
MENU_ISSUANCE_JOURNAL = "issuance_journal"
MENU_APPROVAL_MAIL = "approval_mail"
MENU_DOC_TREE = "doc_tree"
MENU_HISTORY = "history"
MENU_COLLISION = "collision"

_KEY_SEP = "::"


@dataclass(frozen=True, slots=True)
class ActionStat:
    """One tracked context-menu command."""

    menu_id: str
    label: str
    count: int
    last_used: str


def normalize_action_label(label: str) -> str:
    """Strip Qt mnemonics and collapse whitespace."""

    text = str(label or "").replace("&", "").strip()
    return " ".join(text.split())


def action_key(menu_id: str, label: str) -> str:
    """Return the stable ``menu::label`` identity."""

    return f"{str(menu_id).strip()}{_KEY_SEP}{normalize_action_label(label)}"


def compute_hot_keys(
    stats: Mapping[str, ActionStat],
    *,
    min_count: int = HOT_MIN_COUNT,
    share_of_max: float = HOT_SHARE_OF_MAX,
    max_per_menu: int = HOT_MAX_PER_MENU,
) -> frozenset[str]:
    """Return keys that should render bold in their own menu.

    Args:
        stats: ``action_key`` → counts.
        min_count: Ignore one-off clicks.
        share_of_max: Keep items at least this fraction of the menu's max.
        max_per_menu: Cap so a flat-high menu does not bold everything.

    Returns:
        Frozen set of ``menu::label`` keys.
    """

    by_menu: dict[str, list[tuple[str, int]]] = {}
    for key, stat in stats.items():
        if stat.count <= 0:
            continue
        by_menu.setdefault(stat.menu_id, []).append((key, stat.count))
    hot: set[str] = set()
    for items in by_menu.values():
        items.sort(key=lambda item: (-item[1], item[0]))
        max_count = items[0][1]
        if max_count < min_count:
            continue
        threshold = max(min_count, math.ceil(max_count * share_of_max))
        taken = 0
        for key, count in items:
            if count < threshold or taken >= max_per_menu:
                break
            hot.add(key)
            taken += 1
    return frozenset(hot)


def usage_path(runtime_dir: str | Path) -> Path:
    """Return the JSON path under *runtime_dir*."""

    return Path(runtime_dir) / CONTEXT_MENU_USAGE_FILENAME


def _utc_now() -> str:
    return datetime.now(timezone.utc).isoformat()


def _stat_from_json(menu_id: str, label: str, item: Any) -> ActionStat | None:
    if isinstance(item, int):
        count = item
        last_used = ""
    elif isinstance(item, dict):
        try:
            count = int(item.get("count") or 0)
        except (TypeError, ValueError):
            return None
        last_used = str(item.get("last_used") or "").strip()
        stored_label = normalize_action_label(str(item.get("label") or label))
        if stored_label:
            label = stored_label
    else:
        return None
    label = normalize_action_label(label)
    menu_id = str(menu_id or "").strip()
    if not menu_id or not label or count <= 0:
        return None
    return ActionStat(
        menu_id=menu_id,
        label=label,
        count=count,
        last_used=last_used,
    )


def load_usage_stats(path: Path) -> dict[str, ActionStat]:
    """Read counters from *path*. Corrupt or missing files yield empty."""

    try:
        raw = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError):
        return {}
    if not isinstance(raw, dict):
        return {}
    menus = raw.get("menus")
    if not isinstance(menus, dict):
        return {}
    stats: dict[str, ActionStat] = {}
    for menu_id, entries in menus.items():
        if not isinstance(entries, dict):
            continue
        for label, item in entries.items():
            stat = _stat_from_json(str(menu_id), str(label), item)
            if stat is None:
                continue
            stats[action_key(stat.menu_id, stat.label)] = stat
    return stats


def dump_usage_payload(stats: Mapping[str, ActionStat]) -> dict[str, Any]:
    """Build the JSON object written to disk (human-readable ranks)."""

    menus: dict[str, dict[str, Any]] = {}
    for stat in stats.values():
        bucket = menus.setdefault(stat.menu_id, {})
        bucket[stat.label] = {
            "count": stat.count,
            "last_used": stat.last_used,
        }
    ordered_menus: dict[str, dict[str, Any]] = {}
    for menu_id in sorted(menus):
        entries = menus[menu_id]
        ranked = sorted(
            entries.items(),
            key=lambda item: (-int(item[1]["count"]), item[0]),
        )
        ordered_menus[menu_id] = dict(ranked)
    hot_keys = compute_hot_keys(stats)
    hot_by_menu: dict[str, list[str]] = {}
    for key in sorted(hot_keys):
        menu_id, sep, label = key.partition(_KEY_SEP)
        if not sep:
            continue
        hot_by_menu.setdefault(menu_id, []).append(label)
    return {
        "version": CONTEXT_MENU_USAGE_VERSION,
        "updated": _utc_now(),
        "hot": {menu_id: hot_by_menu[menu_id] for menu_id in sorted(hot_by_menu)},
        "menus": ordered_menus,
    }


class ContextMenuUsageStore:
    """Accumulate clicks and flush them off the GUI click path."""

    def __init__(
        self,
        path: str | Path,
        *,
        flush_interval_sec: float = FLUSH_INTERVAL_SEC,
    ) -> None:
        """Load existing counters; do not create the file until a flush.

        Args:
            path: JSON file under the catalog runtime directory.
            flush_interval_sec: Delay after the first dirty click.
        """

        self.path = Path(path)
        self._flush_interval_sec = max(float(flush_interval_sec), 0.05)
        self._lock = threading.Lock()
        self._flush_lock = threading.Lock()
        self._stats: dict[str, ActionStat] = load_usage_stats(self.path)
        self._hot: frozenset[str] = compute_hot_keys(self._stats)
        self._dirty = False
        self._timer: threading.Timer | None = None

    @classmethod
    def from_runtime_dir(
        cls,
        runtime_dir: str | Path,
        *,
        flush_interval_sec: float = FLUSH_INTERVAL_SEC,
    ) -> ContextMenuUsageStore:
        """Build a store for ``runtime_dir/context_menu_usage.json``."""

        return cls(
            usage_path(runtime_dir),
            flush_interval_sec=flush_interval_sec,
        )

    def hot_keys(self) -> frozenset[str]:
        """Return the last flushed (or loaded) bold set. No I/O."""

        with self._lock:
            return self._hot

    def stats_snapshot(self) -> dict[str, ActionStat]:
        """Copy current in-memory counters (including unflushed clicks)."""

        with self._lock:
            return dict(self._stats)

    def note(self, menu_id: str, label: str) -> str:
        """Increment one command. Does not read or write the file.

        Args:
            menu_id: Menu identity (``MENU_KITS``, …).
            label: Visible action text.

        Returns:
            The action key, or ``""`` when the label is empty.
        """

        key = action_key(menu_id, label)
        if key.endswith(_KEY_SEP) or not str(menu_id).strip():
            return ""
        normalized = normalize_action_label(label)
        now = _utc_now()
        arm = False
        with self._lock:
            previous = self._stats.get(key)
            count = 1 if previous is None else previous.count + 1
            self._stats[key] = ActionStat(
                menu_id=str(menu_id).strip(),
                label=normalized,
                count=count,
                last_used=now,
            )
            self._dirty = True
            if self._timer is None:
                arm = True
        if arm:
            self._arm_timer()
        return key

    def flush(self) -> bool:
        """Write pending counts, re-read the file, refresh the bold set.

        Returns:
            True when a write was attempted.
        """

        with self._flush_lock:
            return self._flush_unlocked()

    def shutdown(self) -> None:
        """Cancel the timer and persist leftovers. Safe to call twice."""

        with self._lock:
            timer = self._timer
            self._timer = None
        if timer is not None:
            timer.cancel()
        self.flush()

    def _arm_timer(self) -> None:
        timer = threading.Timer(self._flush_interval_sec, self._on_flush_timer)
        timer.daemon = True
        with self._lock:
            if self._timer is not None:
                return
            self._timer = timer
        timer.start()

    def _on_flush_timer(self) -> None:
        with self._lock:
            self._timer = None
        self.flush()

    def _flush_unlocked(self) -> bool:
        with self._lock:
            if not self._dirty:
                return False
            snapshot = dict(self._stats)
            self._dirty = False
            timer = self._timer
            self._timer = None
        if timer is not None:
            timer.cancel()
        payload = dump_usage_payload(snapshot)
        try:
            self._write_payload(payload)
            loaded = load_usage_stats(self.path)
        except OSError:
            with self._lock:
                self._dirty = True
                need_arm = self._timer is None
            if need_arm:
                self._arm_timer()
            return False
        with self._lock:
            if loaded:
                for key, disk_stat in loaded.items():
                    memory = self._stats.get(key)
                    if memory is None or disk_stat.count > memory.count:
                        self._stats[key] = disk_stat
            self._hot = compute_hot_keys(self._stats)
            need_arm = self._dirty and self._timer is None
        if need_arm:
            self._arm_timer()
        return True

    def _write_payload(self, payload: Mapping[str, Any]) -> None:
        self.path.parent.mkdir(parents=True, exist_ok=True)
        text = json.dumps(payload, ensure_ascii=False, indent=2) + "\n"
        tmp_path = self.path.with_name(self.path.name + ".tmp")
        tmp_path.write_text(text, encoding="utf-8")
        tmp_path.replace(self.path)


_store: ContextMenuUsageStore | None = None


def configure_context_menu_usage(
    runtime_dir: str | Path | None,
    *,
    flush_interval_sec: float = FLUSH_INTERVAL_SEC,
) -> ContextMenuUsageStore | None:
    """Install the process-wide store used by Qt menus.

    Args:
        runtime_dir: Catalog runtime folder, or ``None`` to drop the store.
        flush_interval_sec: Delay after the first dirty click.

    Returns:
        The active store, or ``None`` when *runtime_dir* is omitted.
    """

    global _store
    previous = _store
    if previous is not None:
        previous.shutdown()
        _store = None
    if runtime_dir is None:
        return None
    _store = ContextMenuUsageStore.from_runtime_dir(
        runtime_dir,
        flush_interval_sec=flush_interval_sec,
    )
    return _store


def get_context_menu_usage() -> ContextMenuUsageStore | None:
    """Return the store from :func:`configure_context_menu_usage`."""

    return _store


def shutdown_context_menu_usage() -> None:
    """Flush and drop the process-wide store (window close / tests)."""

    global _store
    store = _store
    _store = None
    if store is not None:
        store.shutdown()


def reset_context_menu_usage_for_tests() -> None:
    """Drop the global store without writing (unit tests)."""

    global _store
    store = _store
    _store = None
    if store is None:
        return
    with store._lock:
        timer = store._timer
        store._timer = None
        store._dirty = False
    if timer is not None:
        timer.cancel()
