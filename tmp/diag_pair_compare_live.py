"""End-to-end check of pair comparison on a small sample of real xlsx files.

Writes only to ``tmp/rd_catalog_copy.sqlite3``. Reads a handful of real
workbooks, so it is slow per pair; the sample size is deliberately small.
"""

from __future__ import annotations

import sys
import time
from collections import Counter
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from rd_catalog.config import load_config  # noqa: E402
from rd_catalog.db import CatalogDatabase  # noqa: E402
from rd_catalog.models import SourceKind  # noqa: E402
from rd_catalog.mto_export import (  # noqa: E402
    default_robot_export_target,
    resolve_export_selections,
    target_files_from_records,
)
from rd_catalog.mto_pair_compare import (  # noqa: E402
    build_pair_compare_plan,
    execute_pair_compare,
    pair_pool_status,
    pairs_from_export_selections,
)

SAMPLE = 6

config = load_config()
database = CatalogDatabase(ROOT / "tmp" / "rd_catalog_copy.sqlite3")
records = database.list_files()
overlay_ids = {
    int(row["file_id"]) for row in database.current_overlay() if row.get("detected_current")
}
target = default_robot_export_target(config)
target_files = target_files_from_records(
    [r for r in records if r.source is SourceKind.ROBOT and r.present]
)
selections = resolve_export_selections(
    database,
    records=records,
    detected_current_ids=overlay_ids,
    rule="latest_issued",
    target=target,
    target_files=target_files,
    pins=(),
    rd_root=config.rd_root,
)

states = Counter(s.state for s in selections)
print("состояния до сравнения:", dict(states.most_common()))

pairs = pairs_from_export_selections(selections)[:SAMPLE]
print(f"выборка пар: {len(pairs)}")

plan = build_pair_compare_plan(database, pairs=pairs, records=records)
print(f"к расчёту {len(plan.pending)}, из кэша {len(plan.cached)}")

started = time.monotonic()
outcome = execute_pair_compare(database, plan=plan, records=records)
elapsed = time.monotonic() - started
print(f"записано в базу: {outcome.written}, в сессии: {len(outcome.session_verdicts)}")
print(f"время: {elapsed:.1f} c на {len(plan.pending)} пар")
print()

session = {key: result.content_status for key, result in outcome.session_verdicts.items()}
status = pair_pool_status(
    database, pairs=pairs, records=records, session_verdicts=session
)
print("пул после расчёта:", status)
print("  очередь пуста:", status.is_drained)
print()

for pair in pairs:
    row = None
    for key, result in outcome.session_verdicts.items():
        del key
        row = row
    print(f"  {pair.title}/{pair.mark}")
    print(f"      {Path(pair.left_path).name}")
    print(f"      {Path(pair.right_path).name}")

print()
print("вердикты:")
verdicts = Counter(result.content_status for result in outcome.session_verdicts.values())
for name, count in verdicts.most_common():
    print(f"   {count:3d}  {name}")
for result in outcome.session_verdicts.values():
    if result.error:
        print("   ошибка:", result.error[:160])

# Second pass must be a pure cache hit for successful pairs.
plan2 = build_pair_compare_plan(database, pairs=pairs, records=records)
print()
print(f"повторный проход: к расчёту {len(plan2.pending)}, из кэша {len(plan2.cached)}")
