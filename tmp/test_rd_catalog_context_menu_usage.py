"""Local checks for delayed context-menu usage counters (no Qt, no UNC)."""

from __future__ import annotations

import json
import sys
import tempfile
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from rd_catalog.context_menu_usage import (
    CONTEXT_MENU_USAGE_FILENAME,
    CONTEXT_MENU_USAGE_VERSION,
    ContextMenuUsageStore,
    MENU_HEATMAP,
    MENU_KITS,
    action_key,
    compute_hot_keys,
    configure_context_menu_usage,
    get_context_menu_usage,
    load_usage_stats,
    normalize_action_label,
    reset_context_menu_usage_for_tests,
    shutdown_context_menu_usage,
)


def main() -> None:
    """Verify keys, hot ranking, delayed write, and reload."""

    reset_context_menu_usage_for_tests()
    assert normalize_action_label("  Открыть &файл · РД  ") == "Открыть файл · РД"
    assert action_key(MENU_KITS, "Открыть файл · РД") == (
        f"{MENU_KITS}::Открыть файл · РД"
    )

    from rd_catalog.context_menu_usage import ActionStat

    def _stat(menu: str, label: str, count: int) -> ActionStat:
        return ActionStat(menu, label, count, "")

    stats = {
        action_key(MENU_KITS, "folder"): _stat(MENU_KITS, "folder", 10),
        action_key(MENU_KITS, "file"): _stat(MENU_KITS, "file", 9),
        action_key(MENU_KITS, "copy"): _stat(MENU_KITS, "copy", 2),
        action_key(MENU_KITS, "ban"): _stat(MENU_KITS, "ban", 1),
        action_key(MENU_HEATMAP, "kits"): _stat(MENU_HEATMAP, "kits", 5),
        action_key(MENU_HEATMAP, "pin"): _stat(MENU_HEATMAP, "pin", 1),
    }
    hot = compute_hot_keys(stats)
    assert action_key(MENU_KITS, "folder") in hot
    assert action_key(MENU_KITS, "file") in hot
    assert action_key(MENU_KITS, "copy") not in hot
    assert action_key(MENU_KITS, "ban") not in hot
    assert action_key(MENU_HEATMAP, "kits") in hot
    assert action_key(MENU_HEATMAP, "pin") not in hot

    low = {
        action_key(MENU_KITS, "once"): _stat(MENU_KITS, "once", 2),
        action_key(MENU_KITS, "twice"): _stat(MENU_KITS, "twice", 1),
    }
    assert compute_hot_keys(low) == frozenset()

    with tempfile.TemporaryDirectory(prefix="rd_catalog_menu_") as temp:
        runtime = Path(temp) / "runtime"
        path = runtime / CONTEXT_MENU_USAGE_FILENAME
        store = ContextMenuUsageStore(path, flush_interval_sec=3600)
        assert not path.exists()
        store.note(MENU_KITS, "Открыть содержащую папку · РД")
        store.note(MENU_KITS, "Открыть содержащую папку · РД")
        store.note(MENU_KITS, "Открыть содержащую папку · РД")
        store.note(MENU_KITS, "Скрыть титул–марку (бан-фильтр)")
        assert not path.exists()
        assert store.hot_keys() == frozenset()
        assert store.flush()
        assert path.is_file()
        payload = json.loads(path.read_text(encoding="utf-8"))
        assert payload["version"] == CONTEXT_MENU_USAGE_VERSION
        folder_label = "Открыть содержащую папку · РД"
        assert payload["menus"][MENU_KITS][folder_label]["count"] == 3
        assert folder_label in payload["hot"][MENU_KITS]
        assert store.hot_keys() == frozenset(
            {action_key(MENU_KITS, folder_label)}
        )

        store.note(MENU_KITS, folder_label)
        assert (
            store.stats_snapshot()[action_key(MENU_KITS, folder_label)].count
            == 4
        )
        reloaded = load_usage_stats(path)
        assert reloaded[action_key(MENU_KITS, folder_label)].count == 3
        store.flush()
        reloaded = load_usage_stats(path)
        assert reloaded[action_key(MENU_KITS, folder_label)].count == 4

        broken = ContextMenuUsageStore(
            runtime / "missing.json", flush_interval_sec=3600
        )
        assert broken.hot_keys() == frozenset()
        bad = runtime / "bad.json"
        bad.write_text("{not json", encoding="utf-8")
        corrupt = ContextMenuUsageStore(bad, flush_interval_sec=3600)
        assert corrupt.stats_snapshot() == {}

        store.shutdown()
        broken.shutdown()
        corrupt.shutdown()

        configured = configure_context_menu_usage(runtime, flush_interval_sec=3600)
        assert configured is not None
        assert get_context_menu_usage() is configured
        configured.note(MENU_HEATMAP, "Показать в „Комплекты“")
        shutdown_context_menu_usage()
        assert get_context_menu_usage() is None
        disk = load_usage_stats(path)
        assert action_key(MENU_HEATMAP, "Показать в „Комплекты“") in disk

        before = disk[action_key(MENU_KITS, folder_label)].count
        fast = ContextMenuUsageStore(path, flush_interval_sec=0.05)
        fast.note(MENU_KITS, folder_label)
        deadline = time.time() + 2.0
        while time.time() < deadline:
            disk_after = load_usage_stats(path)
            if disk_after[action_key(MENU_KITS, folder_label)].count > before:
                break
            time.sleep(0.02)
        else:
            raise AssertionError("timer did not persist usage JSON")
        fast.shutdown()

    reset_context_menu_usage_for_tests()
    print("rd_catalog context_menu_usage ok")


if __name__ == "__main__":
    main()
