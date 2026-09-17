"""Rebuild the pipeline on a COPY of the live database, then measure the rules.

The user's working database is never opened for writing: everything happens in
``tmp/rd_catalog_copy.sqlite3``.
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
from rd_catalog.models import SourceKind  # noqa: E402
from rd_catalog.mto_export import (  # noqa: E402
    default_robot_export_target,
    resolve_export_selections,
    target_files_from_records,
)
from rd_catalog.pipeline import rebuild_pipeline  # noqa: E402

RULES = ("approved", "latest_tdo", "latest_issued", "latest_no_as_build")

config = load_config()
copy_path = ROOT / "tmp" / "rd_catalog_copy.sqlite3"
database = CatalogDatabase(copy_path)
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


def package_stats() -> tuple[int, int]:
    rows = database.list_kit_packages()
    junk = sum(1 for row in rows if "_рев." not in Path(row.package_path).name.casefold()
               and not row.is_grey)
    return len(rows), junk


before_total, before_junk = package_stats()
print(f"до пересборки : папок передач {before_total}, из них без «_рев.» в имени {before_junk}")

rebuild_pipeline(
    database,
    records=records,
    detected_current_ids=overlay_ids,
    rd_root=config.rd_root,
)

after_total, after_junk = package_stats()
print(f"после         : папок передач {after_total}, из них без «_рев.» в имени {after_junk}")
print()

target = default_robot_export_target(config)
target_files = target_files_from_records(
    [r for r in records if r.source is SourceKind.ROBOT and r.present]
)

for rule in RULES:
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
    states = Counter(s.state for s in live)
    print(f"--- «{rule}»  комплектов {len(live)}: " + ", ".join(
        f"{state}={count}" for state, count in states.most_common()
    ))

print()
for title, mark in (("1600", "SOT"), ("3140", "KSB1"), ("3240", "KSB1")):
    key = kit_identity_key(title, mark)
    print(f"  {title}/{mark}")
    for rule in RULES:
        sel = next(
            (
                s
                for s in resolve_export_selections(
                    database,
                    records=records,
                    detected_current_ids=overlay_ids,
                    rule=rule,
                    target=target,
                    target_files=target_files,
                    pins=(),
                    rd_root=config.rd_root,
                )
                if kit_identity_key(s.title, s.mark) == key
            ),
            None,
        )
        if sel is None:
            print(f"      {rule:<20} <нет строки>")
            continue
        print(f"      {rule:<20} {Path(sel.source_path).name or '—'}")
        print(f"      {'':<20} из {Path(sel.package_path).name or '—'}")
