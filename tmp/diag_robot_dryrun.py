"""Dry run of «Обновить MTO у робота» — which file would be copied, no UNC, no writes.

``discover_siblings=False`` keeps the planner purely on DB records, so nothing
touches the network and nothing is written anywhere.
"""

from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from rd_catalog.config import load_config  # noqa: E402
from rd_catalog.db import CatalogDatabase  # noqa: E402
from rd_catalog.models import SourceKind  # noqa: E402
from rd_catalog.robot_mto_sync import (  # noqa: E402
    RobotMtoSyncError,
    collect_kit_mto_candidates,
    collect_robot_mto_files,
    plan_robot_mto_sync,
)

KITS = [("1600", "SOT"), ("6400", "SOT"), ("2225", "KSB")]

config = load_config()
database = CatalogDatabase(config.db_path)
records = database.list_files()
overlay_ids = {
    int(row["file_id"])
    for row in database.current_overlay()
    if row.get("detected_current")
}

print(f"records: {len(records)}  overlay-current ids: {len(overlay_ids)}")
print(f"robot_root: {config.robot_root}")
print()

for title, mark in KITS:
    print(f"===== {title}/{mark} =====")
    candidates = collect_kit_mto_candidates(
        title=title,
        mark=mark,
        records=records,
        detected_current_ids=overlay_ids,
        discover_siblings=False,
    )
    print(f"  candidates ({len(candidates)}):")
    for ref in candidates:
        print(f"    [{ref.source.value}] rev={ref.revision_text!r} {ref.path}")

    robot = collect_robot_mto_files(title=title, mark=mark, records=records)
    print(f"  robot now ({len(robot)}):")
    for ref in robot:
        print(f"    rev={ref.revision_text!r} {ref.path}")

    try:
        plan = plan_robot_mto_sync(
            title=title,
            mark=mark,
            records=records,
            detected_current_ids=overlay_ids,
            robot_root=config.robot_root,
            robot_flat_structure=getattr(config, "robot_flat_structure", False),
            discover_siblings=False,
        )
    except RobotMtoSyncError as exc:
        print(f"  PLAN: refused -> {exc}")
    except Exception as exc:  # noqa: BLE001 - diagnostic script
        print(f"  PLAN: error {type(exc).__name__}: {exc}")
    else:
        source = plan.source
        print(f"  PLAN source : rev={source.revision_text!r} {source.path}")
        print(f"  PLAN dest   : {plan.destination_path}")
        existing = getattr(plan, "existing_robot", None) or getattr(plan, "existing_path", None)
        print(f"  PLAN replaces: {existing}")
    print()

# Всё, что есть на диске по MTO для этих комплектов, независимо от overlay.
print("===== every present RD MTO (context) =====")
for title, mark in KITS:
    print(f"--- {title}/{mark} ---")
    for record in records:
        if not record.present or record.source is not SourceKind.RD:
            continue
        if str(record.data.get("file_kind") or "") != "mto_xlsx":
            continue
        if str(record.data.get("title") or "") != title:
            continue
        if str(record.data.get("mark") or "").casefold() != mark.casefold():
            continue
        flag = "CURRENT" if record.id in overlay_ids else "       "
        print(
            f"  {flag} rev={record.data.get('revision')!r}"
            f"-{record.data.get('appendix')!r}"
            f" NN={record.data.get('transfer_sequence')!r} {record.data.get('transfer_name')!r}"
        )
