"""Read-only: what changes for the robot if export follows the approved contour.

Compares three answers per kit, without touching the network:
  now_robot  - the MTO file currently sitting in robot_root;
  today_pick - what plan_robot_mto_sync would copy today (overlay-current);
  contour    - what resolve_approved_contour says is the approved MTO.
"""

from __future__ import annotations

import json
import sys
from collections import Counter
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from rd_catalog.config import load_config  # noqa: E402
from rd_catalog.db import CatalogDatabase  # noqa: E402
from rd_catalog.pipeline import resolve_all_approved_contours  # noqa: E402
from rd_catalog.robot_mto_sync import (  # noqa: E402
    RobotMtoSyncError,
    collect_kit_mto_candidates,
    collect_robot_mto_files,
)
from rd_catalog.robot_mto_sync import _freshness_key  # noqa: E402

config = load_config()
database = CatalogDatabase(config.db_path)
records = database.list_files()
overlay_ids = {
    int(row["file_id"])
    for row in database.current_overlay()
    if row.get("detected_current")
}

banned: set[tuple[str, str]] = set()
banned_path = Path(config.runtime_dir) / "banned_title_marks.json"
if banned_path.exists():
    raw = json.loads(banned_path.read_text(encoding="utf-8"))
    items = raw.get("pairs", raw) if isinstance(raw, dict) else raw
    for item in items:
        if isinstance(item, dict):
            banned.add((str(item.get("title", "")).casefold(), str(item.get("mark", "")).casefold()))
        elif isinstance(item, (list, tuple)) and len(item) == 2:
            banned.add((str(item[0]).casefold(), str(item[1]).casefold()))

contours = {
    (c.title.casefold(), c.mark.casefold()): c
    for c in resolve_all_approved_contours(
        database, records=records, detected_current_ids=overlay_ids
    )
}

verdicts: Counter[str] = Counter()
changes: list[str] = []
downgrades: list[str] = []

for key, contour in sorted(contours.items()):
    if key in banned:
        continue
    title, mark = contour.title, contour.mark

    candidates = collect_kit_mto_candidates(
        title=title,
        mark=mark,
        records=records,
        detected_current_ids=overlay_ids,
        discover_siblings=False,
    )
    today = max(candidates, key=_freshness_key).path if candidates else ""
    robot = collect_robot_mto_files(title=title, mark=mark, records=records)
    now_robot = robot[0].path if robot else ""
    wanted = contour.mto_path

    if not wanted:
        verdicts["контур не даёт MTO — выгрузка должна отказать"] += 1
        continue
    if not today:
        verdicts["сегодня выгрузка невозможна, контур даёт файл"] += 1
        changes.append(f"{title}/{mark}: сегодня нечего копировать, контур -> {Path(wanted).name}")
        continue

    if Path(today).name.casefold() == Path(wanted).name.casefold() and today.casefold() == wanted.casefold():
        verdicts["совпадает — поведение не изменится"] += 1
    else:
        verdicts["ВЫБОР ИЗМЕНИТСЯ"] += 1
        line = (
            f"{title}/{mark}\n"
            f"    сегодня : {Path(today).parent.name}\\{Path(today).name}\n"
            f"    контур  : {Path(wanted).parent.name}\\{Path(wanted).name}\n"
            f"    у робота: {Path(now_robot).name if now_robot else '<нет>'}\n"
            f"    доверие : {contour.confidence} ({contour.match_reason})"
        )
        changes.append(line)
        if now_robot and Path(now_robot).name.casefold() == Path(wanted).name.casefold():
            downgrades.append(f"{title}/{mark}: у робота уже правильный {Path(now_robot).name}, "
                              f"сегодняшняя кнопка подменила бы его на {Path(today).name}")

print("=== влияние перехода на согласованный контур ===")
total = sum(verdicts.values())
for verdict, count in verdicts.most_common():
    print(f"  {count:4d}  {verdict}")
print(f"  ---- всего комплектов: {total}")

print()
print(f"=== комплекты, где сегодняшняя кнопка ИСПОРТИЛА БЫ правильный файл у робота: {len(downgrades)} ===")
for line in downgrades:
    print("  " + line)

print()
print("=== все изменения выбора ===")
for line in changes:
    print("  " + line)
