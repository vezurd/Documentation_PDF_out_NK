"""Read-only estimate v2 of resolver confidence, using the rule we actually plan.

Difference from v1: the approved package is resolved with fallbacks (cycle -> folder
revision -> filename revision -> last package), and only conflicts *inside* the
approved contour (packages with sequence <= approved) count as ambiguity. Conflicts
introduced by later packages are exactly what the approved-contour rule resolves,
so they must not trigger a question.
"""

from __future__ import annotations

import json
import os
import sqlite3
from collections import Counter
from pathlib import Path

DB = Path(os.environ["LOCALAPPDATA"]) / "Documentation_PDF_out_NK" / "rd_catalog" / "rd_catalog.sqlite"
OUT = Path(__file__).resolve().parent / "diag_confidence2.txt"

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


def norm(text: str | None) -> str:
    return (text or "").strip().casefold()


def eq_rev(left: str | None, right: str | None) -> bool:
    a, b = norm(left), norm(right)
    if not a or not b:
        return False
    if a == b:
        return True
    try:
        ah, _, at = a.partition("-an")
        bh, _, bt = b.partition("-an")
        return (int(ah), int(at or 0)) == (int(bh), int(bt or 0))
    except ValueError:
        return False


packages_by_id = {row["id"]: row for row in conn.execute("SELECT * FROM kit_package")}
rd_packages: dict[tuple[str, str], list[sqlite3.Row]] = {}
for row in packages_by_id.values():
    if row["source"] != "rd" or row["is_grey"]:
        continue
    rd_packages.setdefault((norm(row["title"]), norm(row["mark"])), []).append(row)
for items in rd_packages.values():
    items.sort(key=lambda r: (r["sequence"] if r["sequence"] is not None else -1, r["id"]))

cycles: dict[tuple[str, str], list[sqlite3.Row]] = {}
for row in conn.execute("SELECT * FROM kit_cycle"):
    cycles.setdefault((norm(row["title"]), norm(row["mark"])), []).append(row)

# Folder revision of each transfer, from file_entry.
folder_rev: dict[tuple[str, str, str], str] = {}
for row in conn.execute(
    "SELECT title, mark, transfer_name, transfer_revision, transfer_appendix "
    "FROM file_entry WHERE source='rd' AND present=1 AND transfer_name IS NOT NULL"
):
    rev = (row["transfer_revision"] or "").strip()
    app = (row["transfer_appendix"] or "").strip()
    if not rev:
        continue
    folder_rev[(norm(row["title"]), norm(row["mark"]), norm(row["transfer_name"]))] = (
        f"{rev}-AN{app}" if app else rev
    )

# MTO-overlay conflicts, mapped to the packages they involve.
path_to_pkg: dict[str, tuple[tuple[str, str], int]] = {}
for row in conn.execute(
    "SELECT path_key, title, mark, transfer_sequence FROM file_entry "
    "WHERE source='rd' AND title IS NOT NULL"
):
    path_to_pkg[row["path_key"]] = (
        (norm(row["title"]), norm(row["mark"])),
        row["transfer_sequence"] if row["transfer_sequence"] is not None else -1,
    )

conflicts: dict[tuple[str, str], list[tuple[str, list[int]]]] = {}
for row in conn.execute(
    "SELECT scope, kind, path_keys_json FROM current_collision WHERE scope='overlay_mto'"
):
    kit: tuple[str, str] | None = None
    seqs: list[int] = []
    for path_key in json.loads(row["path_keys_json"] or "[]"):
        found = path_to_pkg.get(path_key)
        if not found:
            continue
        kit, seq = found
        seqs.append(seq)
    if kit and seqs:
        conflicts.setdefault(kit, []).append((row["kind"], seqs))

pipes = [r for r in conn.execute("SELECT * FROM kit_pipeline") if (norm(r["title"]), norm(r["mark"])) not in banned]

tiers = Counter()
reasons = Counter()
match_reasons = Counter()
low_list: list[str] = []
detail: list[str] = []

for pipe in pipes:
    kit = (norm(pipe["title"]), norm(pipe["mark"]))
    label = f"{pipe['title']}/{pipe['mark']}"
    approved_rev = pipe["code_revision_text"] or pipe["official_revision_text"] or ""
    kit_packages = rd_packages.get(kit, [])

    if not kit_packages:
        tiers["low"] += 1
        reasons["папок РД нет вообще"] += 1
        low_list.append(f"{label} [папок РД нет]")
        continue

    # --- resolve the approved package -------------------------------------
    approved = None
    how = ""
    kit_cycles = cycles.get(kit, [])
    code_cycles = [c for c in kit_cycles if c["code_event_id"] is not None]
    chosen_cycle = (code_cycles or kit_cycles)[-1] if (code_cycles or kit_cycles) else None
    if chosen_cycle is not None and chosen_cycle["package_id"]:
        candidate = packages_by_id.get(chosen_cycle["package_id"])
        if candidate is not None and candidate["source"] == "rd" and not candidate["is_grey"]:
            approved, how = candidate, f"цикл ({chosen_cycle['match_reason']})"
    if approved is None and approved_rev:
        by_folder = [
            p
            for p in kit_packages
            if eq_rev(folder_rev.get((kit[0], kit[1], norm(p["transfer_name"]))), approved_rev)
        ]
        if by_folder:
            approved, how = by_folder[-1], "имя папки рев.*"
    if approved is None and approved_rev:
        by_files = [p for p in kit_packages if eq_rev(p["revision_text"], approved_rev)]
        if by_files:
            approved, how = by_files[-1], "ревизия файлов"
    if approved is None:
        approved, how = kit_packages[-1], "последняя передача (запасной вариант)"
    match_reasons[how] += 1

    seq = approved["sequence"] if approved["sequence"] is not None else -1

    # --- find the MTO inside the approved contour -------------------------
    contour = [
        p
        for p in kit_packages
        if (p["sequence"] if p["sequence"] is not None else -1) <= seq
        and (p["mto_revision_text"] or "").strip()
    ]
    if not contour:
        tiers["low"] += 1
        reasons["MTO нет ни в одной передаче контура"] += 1
        low_list.append(f"{label} [нет MTO в контуре, согл.рев={approved_rev}]")
        continue

    # --- ambiguity strictly inside the contour ----------------------------
    inside = [
        (kind, seqs)
        for kind, seqs in conflicts.get(kit, [])
        if all(s <= seq for s in seqs) and len(set(seqs)) > 0
    ]
    blocking = [kind for kind, seqs in inside if kind in {"dup_same_revision", "transfer_order_conflict"}]

    if blocking:
        tiers["low"] += 1
        reasons["конфликт внутри согласованного контура"] += 1
        low_list.append(f"{label} [конфликт в контуре: {', '.join(sorted(set(blocking)))}]")
        continue

    if (approved["mto_revision_text"] or "").strip():
        tiers["high"] += 1
    else:
        tiers["medium"] += 1
        detail.append(f"{label}: MTO унаследован из NN{contour[-1]['sequence']} ({how})")

total = len(pipes)
lines = [
    f"Комплектов (без забаненных): {total}",
    "",
    f"high   — робот решает сам, вопросов нет                : {tiers['high']:3d}  ({100*tiers['high']/max(total,1):.0f}%)",
    f"medium — робот решает сам, MTO унаследован из ранней NN: {tiers['medium']:3d}  ({100*tiers['medium']/max(total,1):.0f}%)",
    f"low    — нужен человек хотя бы один раз                : {tiers['low']:3d}  ({100*tiers['low']/max(total,1):.0f}%)",
    "",
    "Как определился согласованный пакет:",
]
for reason, count in match_reasons.most_common():
    lines.append(f"  {count:4d}  {reason}")
lines.append("")
lines.append("Причины попадания в low:")
for reason, count in reasons.most_common():
    lines.append(f"  {count:4d}  {reason}")
lines.append("")
lines.append("Список low:")
lines.extend("  " + item for item in sorted(low_list))
lines.append("")
lines.append("medium (унаследованный MTO):")
lines.extend("  " + item for item in sorted(detail))

OUT.write_text("\n".join(lines), encoding="utf-8")
print("\n".join(lines[: 14 + len(match_reasons) + len(reasons)]))
print(f"\n(полный список: {OUT})")
