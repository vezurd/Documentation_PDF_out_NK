"""Read-only acceptance of the four export rules against the live database.

Nothing is written: no JSON store is saved, no file is copied, the database is
only read.
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
from rd_catalog.kits import kit_identity_key  # noqa: E402
from rd_catalog.mto_export import (  # noqa: E402
    EXPORT_RULES,
    default_robot_export_target,
    resolve_export_selections,
    target_files_from_records,
)
from rd_catalog.models import SourceKind  # noqa: E402

RULE_ORDER = ("approved", "latest_tdo", "latest_issued", "latest_no_as_build")
CHECK_KITS = (("1600", "SOT"), ("3240", "KSB1"), ("6400", "SOT"), ("6100", "SOS"))

config = load_config()
database = CatalogDatabase(config.db_path)
records = database.list_files()
overlay_ids = {
    int(row["file_id"]) for row in database.current_overlay() if row.get("detected_current")
}

banned: set[tuple[str, str]] = set()
banned_path = Path(config.runtime_dir) / "banned_title_marks.json"
if banned_path.exists():
    raw = json.loads(banned_path.read_text(encoding="utf-8"))
    items = raw.get("pairs", raw) if isinstance(raw, dict) else raw
    for item in items:
        if isinstance(item, dict):
            banned.add(kit_identity_key(str(item.get("title", "")), str(item.get("mark", ""))))
        elif isinstance(item, (list, tuple)) and len(item) == 2:
            banned.add(kit_identity_key(str(item[0]), str(item[1])))

target = default_robot_export_target(config)
robot_records = [r for r in records if r.source is SourceKind.ROBOT and r.present]
target_files = target_files_from_records(robot_records)
print(f"цель: «{target.name}»  правило по умолчанию: {target.rule}")
print(f"файлов у робота найдено: {len(target_files)}")
print(f"забаненных комплектов: {len(banned)}")
print()

assert set(RULE_ORDER) == set(EXPORT_RULES), sorted(EXPORT_RULES)

per_rule: dict[str, dict[tuple[str, str], object]] = {}

for rule in RULE_ORDER:
    selections = resolve_export_selections(
        database,
        records=records,
        detected_current_ids=overlay_ids,
        rule=rule,
        target=target,
        target_files=target_files,
        pins=(),
        rd_root=config.rd_root,
    )
    live = [s for s in selections if kit_identity_key(s.title, s.mark) not in banned]
    per_rule[rule] = {kit_identity_key(s.title, s.mark): s for s in live}

    states = Counter(s.state for s in live)
    conf = Counter(s.confidence for s in live if s.source_path)
    print(f"--- правило «{rule}»  комплектов: {len(live)}")
    for state, count in states.most_common():
        print(f"      {state:<12} {count}")
    if conf:
        print(f"      доверие: {dict(conf)}")
    print()

print("=== контрольные комплекты ===")
for title, mark in CHECK_KITS:
    key = kit_identity_key(title, mark)
    print(f"  {title}/{mark}")
    for rule in RULE_ORDER:
        sel = per_rule[rule].get(key)
        if sel is None:
            print(f"      {rule:<20} <нет строки>")
            continue
        name = Path(sel.source_path).name if sel.source_path else "—"
        pkg = Path(sel.package_path).name if sel.package_path else "—"
        print(f"      {rule:<20} {sel.state:<10} {name}")
        print(f"      {'':<20} из папки {pkg}")
    print()

print("=== расхождения «Согласованное» против «Последнее выданное» ===")
approved = per_rule["approved"]
latest = per_rule["latest_issued"]
differ = 0
for key, sel in sorted(approved.items()):
    other = latest.get(key)
    if other is None:
        continue
    a, b = sel.source_path, other.source_path
    if a and b and a.casefold() != b.casefold():
        differ += 1
print(f"  расходятся по файлу: {differ}")
