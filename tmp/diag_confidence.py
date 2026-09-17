"""Read-only estimate: how many kits would need a human decision about «which MTO».

Proxy for the planned resolver confidence tiers:
  high   - approved cycle points at a real (non-grey) package that itself holds an
           MTO, and the kit has no overlay_mto conflicts;
  medium - the MTO has to be inherited from an earlier package, still no conflicts;
  low    - overlay_mto conflicts, or the cycle only reaches a grey issuance stub,
           or no MTO anywhere in the kit.
"""

from __future__ import annotations

import json
import os
import sqlite3
from pathlib import Path

DB = Path(os.environ["LOCALAPPDATA"]) / "Documentation_PDF_out_NK" / "rd_catalog" / "rd_catalog.sqlite"
OUT = Path(__file__).resolve().parent / "diag_confidence.txt"

conn = sqlite3.connect(f"file:{DB}?mode=ro", uri=True)
conn.row_factory = sqlite3.Row

banned: set[tuple[str, str]] = set()
banned_path = DB.parent / "banned_title_marks.json"
if banned_path.exists():
    raw = json.loads(banned_path.read_text(encoding="utf-8"))
    items = raw.get("pairs", raw) if isinstance(raw, dict) else raw
    for item in items:
        if isinstance(item, dict):
            banned.add((str(item.get("title", "")).casefold(), str(item.get("mark", "")).casefold()))
        elif isinstance(item, (list, tuple)) and len(item) == 2:
            banned.add((str(item[0]).casefold(), str(item[1]).casefold()))


def ok(title: str, mark: str) -> bool:
    return (str(title).casefold(), str(mark).casefold()) not in banned


packages_by_id = {row["id"]: row for row in conn.execute("SELECT * FROM kit_package")}
packages_by_kit: dict[tuple[str, str], list[sqlite3.Row]] = {}
for row in packages_by_id.values():
    if row["source"] != "rd" or row["is_grey"]:
        continue
    packages_by_kit.setdefault((row["title"].casefold(), row["mark"].casefold()), []).append(row)

cycles_by_kit: dict[tuple[str, str], list[sqlite3.Row]] = {}
for row in conn.execute("SELECT * FROM kit_cycle"):
    cycles_by_kit.setdefault((row["title"].casefold(), row["mark"].casefold()), []).append(row)

# Kits touched by MTO-overlay conflicts (path -> title/mark via file_entry).
mto_conflict_kits: set[tuple[str, str]] = set()
paths_to_kit: dict[str, tuple[str, str]] = {}
for row in conn.execute("SELECT path_key, title, mark FROM file_entry WHERE title IS NOT NULL"):
    paths_to_kit[row["path_key"]] = (str(row["title"]).casefold(), str(row["mark"] or "").casefold())
for row in conn.execute(
    "SELECT scope, kind, path_keys_json FROM current_collision "
    "WHERE kind IN ('transfer_order_conflict','transfer_mtime_conflict','dup_same_revision')"
):
    if row["scope"] != "overlay_mto":
        continue
    for path_key in json.loads(row["path_keys_json"] or "[]"):
        kit = paths_to_kit.get(path_key)
        if kit:
            mto_conflict_kits.add(kit)

pipes = [r for r in conn.execute("SELECT * FROM kit_pipeline") if ok(r["title"], r["mark"])]

tiers: dict[str, list[str]] = {"high": [], "medium": [], "low": []}
reasons: dict[str, int] = {}

for pipe in pipes:
    kit = (pipe["title"].casefold(), pipe["mark"].casefold())
    label = f"{pipe['title']}/{pipe['mark']}"
    cycles = cycles_by_kit.get(kit, [])
    code_cycles = [c for c in cycles if c["code_event_id"] is not None]
    chosen = (code_cycles or cycles)[-1] if (code_cycles or cycles) else None

    if kit in mto_conflict_kits:
        tiers["low"].append(label + " [конфликт передач по MTO]")
        reasons["конфликт передач по MTO"] = reasons.get("конфликт передач по MTO", 0) + 1
        continue
    if chosen is None:
        tiers["low"].append(label + " [нет цикла в журнале]")
        reasons["нет цикла в журнале"] = reasons.get("нет цикла в журнале", 0) + 1
        continue

    pkg = packages_by_id.get(chosen["package_id"]) if chosen["package_id"] else None
    if pkg is None or pkg["is_grey"] or pkg["source"] != "rd":
        kit_packages = packages_by_kit.get(kit, [])
        if not kit_packages:
            tiers["low"].append(label + " [журнал есть, папок РД нет]")
            reasons["журнал есть, папок РД нет"] = reasons.get("журнал есть, папок РД нет", 0) + 1
        else:
            tiers["low"].append(label + " [цикл не сматчен на папку]")
            reasons["цикл не сматчен на папку"] = reasons.get("цикл не сматчен на папку", 0) + 1
        continue

    if (pkg["mto_revision_text"] or "").strip():
        tiers["high"].append(label)
        continue

    sequence = pkg["sequence"] if pkg["sequence"] is not None else -1
    earlier = [
        p
        for p in packages_by_kit.get(kit, [])
        if (p["sequence"] if p["sequence"] is not None else -1) <= sequence
        and (p["mto_revision_text"] or "").strip()
    ]
    if earlier:
        tiers["medium"].append(label)
    else:
        tiers["low"].append(label + " [MTO нет ни в одной передаче до согласованной]")
        reasons["MTO нет ни в одной передаче"] = reasons.get("MTO нет ни в одной передаче", 0) + 1

total = len(pipes)
lines = [f"Комплектов (без забаненных): {total}", ""]
for tier, caption in (
    ("high", "high — робот решает сам, вопросов нет"),
    ("medium", "medium — робот решает сам (MTO унаследован из более ранней передачи)"),
    ("low", "low — нужен человек хотя бы один раз"),
):
    items = tiers[tier]
    lines.append(f"{caption}: {len(items)}  ({100 * len(items) / max(total, 1):.0f}%)")
lines.append("")
lines.append("Причины попадания в low:")
for reason, count in sorted(reasons.items(), key=lambda item: -item[1]):
    lines.append(f"  {count:4d}  {reason}")
lines.append("")
lines.append("Список low:")
lines.extend("  " + item for item in sorted(tiers["low"]))

OUT.write_text("\n".join(lines), encoding="utf-8")
print("\n".join(lines[:14]))
print(f"\n(полный список: {OUT})")
