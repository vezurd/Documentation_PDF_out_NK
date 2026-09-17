"""Read-only: what each heatmap filter would put into a robot folder.

For every predicate of ``iter_mto_files_for_cells`` this reports how many kits
get exactly one MTO file, how many get several (ambiguous), and how many get
nothing at all. The "nothing" bucket is what the new third column must colour
loudly so it cannot be missed.
"""

from __future__ import annotations

import json
import sys
from collections import defaultdict
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from rd_catalog.config import load_config  # noqa: E402
from rd_catalog.db import CatalogDatabase  # noqa: E402
from rd_catalog.kits import kit_identity_key  # noqa: E402
from rd_catalog.pipeline import iter_mto_files_for_cells  # noqa: E402

PREDICATES = ("code_a", "tdo_passed", "current", "current_ifc", "exclude_as_build")

config = load_config()
database = CatalogDatabase(config.db_path)
records = database.list_files()

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

# The universe the cockpit must show: every kit the catalog knows about.
universe: set[tuple[str, str]] = set()
for row in database.list_kit_pipelines():
    key = kit_identity_key(row.title, row.mark)
    if key not in banned:
        universe.add(key)

print(f"комплектов в конвейере (без забаненных): {len(universe)}")
print()

for predicate in PREDICATES:
    files = iter_mto_files_for_cells(
        database,
        records=records,
        predicate=predicate,
        rd_root=config.rd_root,
    )
    by_kit: dict[tuple[str, str], set[str]] = defaultdict(set)
    for record in files:
        title = str(record.data.get("title") or "").strip()
        mark = str(record.data.get("mark") or "").strip()
        key = kit_identity_key(title, mark)
        if key in banned:
            continue
        by_kit[key].add(Path(record.path).name)

    single = sum(1 for names in by_kit.values() if len(names) == 1)
    many = sum(1 for names in by_kit.values() if len(names) > 1)
    empty = len(universe - set(by_kit))

    print(f"--- предикат «{predicate}»")
    print(f"      ровно один файл  : {single}")
    print(f"      несколько файлов : {many}   <- нужен выбор")
    print(f"      НИ ОДНОГО файла  : {empty}   <- красить как пропуск")
    if many:
        shown = 0
        for key, names in sorted(by_kit.items()):
            if len(names) > 1 and shown < 5:
                print(f"        {key[0]}/{key[1]}: {', '.join(sorted(names))}")
                shown += 1
    print()
