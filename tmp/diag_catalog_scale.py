"""Read-only: how widespread are the 1600/SOT-style inconsistencies across the catalog."""

from __future__ import annotations

import json
import os
import sqlite3
from pathlib import Path

DB = Path(os.environ["LOCALAPPDATA"]) / "Documentation_PDF_out_NK" / "rd_catalog" / "rd_catalog.sqlite"
OUT = Path(__file__).resolve().parent / "diag_catalog_scale.txt"

conn = sqlite3.connect(f"file:{DB}?mode=ro", uri=True)
conn.row_factory = sqlite3.Row
lines: list[str] = []


def emit(text: str = "") -> None:
    lines.append(text)


banned_path = DB.parent / "banned_title_marks.json"
banned: set[tuple[str, str]] = set()
if banned_path.exists():
    raw = json.loads(banned_path.read_text(encoding="utf-8"))
    items = raw.get("pairs", raw) if isinstance(raw, dict) else raw
    for item in items:
        if isinstance(item, dict):
            banned.add((str(item.get("title", "")).casefold(), str(item.get("mark", "")).casefold()))
        elif isinstance(item, (list, tuple)) and len(item) == 2:
            banned.add((str(item[0]).casefold(), str(item[1]).casefold()))
emit(f"banned pairs: {len(banned)}")


def allowed(title: str, mark: str) -> bool:
    return (str(title).casefold(), str(mark).casefold()) not in banned


# --- 1. Packages whose MTO revision lags the package revision -----------------
rows = conn.execute(
    "SELECT title, mark, sequence, transfer_name, revision_text, mto_revision_text "
    "FROM kit_package WHERE source='rd' AND is_grey=0"
).fetchall()
rows = [r for r in rows if allowed(r["title"], r["mark"])]
lag = [r for r in rows if r["mto_revision_text"] and r["revision_text"] and r["mto_revision_text"] != r["revision_text"]]
no_mto_pkg = [r for r in rows if not r["mto_revision_text"]]
emit()
emit("=== 1. kit_package: MTO revision vs package revision ===")
emit(f"  RD packages total: {len(rows)}")
emit(f"  MTO rev != package rev: {len(lag)}  ({100*len(lag)/max(len(rows),1):.1f}%)")
emit(f"  package without MTO at all: {len(no_mto_pkg)}")
emit(f"  distinct kits with a lagging MTO: {len({(r['title'], r['mark']) for r in lag})}")
emit("  examples:")
for r in lag[:15]:
    emit(f"    {r['title']}/{r['mark']} NN={r['sequence']} '{r['transfer_name']}' pkg={r['revision_text']} mto={r['mto_revision_text']}")

# --- 2. Cells with an approval letter but no MTO file -------------------------
cells = conn.execute("SELECT * FROM kit_revision_cell").fetchall()
cells = [c for c in cells if allowed(c["title"], c["mark"])]
emit()
emit("=== 2. kit_revision_cell ===")
emit(f"  cells total: {len(cells)}  kits: {len({(c['title'], c['mark']) for c in cells})}")
code_a_no_mto = [c for c in cells if c["pipeline_status"] == "code_a" and not c["has_mto"]]
emit(f"  status=code_a but has_mto=0: {len(code_a_no_mto)} cells / {len({(c['title'], c['mark']) for c in code_a_no_mto})} kits")
for c in code_a_no_mto[:15]:
    emit(f"    {c['title']}/{c['mark']} rev={c['revision_text']} is_current={c['is_current']} ifc={c['is_current_ifc']}")

# --- 3. Kits where the pipeline official revision has no MTO ------------------
pipes = conn.execute("SELECT * FROM kit_pipeline").fetchall()
pipes = [p for p in pipes if allowed(p["title"], p["mark"])]
cells_by_kit: dict[tuple[str, str], list[sqlite3.Row]] = {}
for c in cells:
    cells_by_kit.setdefault((c["title"].casefold(), c["mark"].casefold()), []).append(c)

official_no_mto = []
official_missing_cell = []
for p in pipes:
    official = (p["official_revision_text"] or "").strip()
    if not official:
        continue
    kit_cells = cells_by_kit.get((p["title"].casefold(), p["mark"].casefold()), [])
    match = [c for c in kit_cells if (c["revision_text"] or "").casefold() == official.casefold()]
    if not match:
        official_missing_cell.append(p)
    elif not any(c["has_mto"] for c in match):
        official_no_mto.append(p)

emit()
emit("=== 3. kit_pipeline: official revision without an MTO file ===")
emit(f"  kits with pipeline row: {len(pipes)}")
emit(f"  official rev has a cell but no MTO: {len(official_no_mto)}")
for p in official_no_mto[:20]:
    emit(f"    {p['title']}/{p['mark']} official={p['official_revision_text']} status={p['status']} code={p['code']} stale={p['code_stale']}")
emit(f"  official rev has no cell at all: {len(official_missing_cell)}")
for p in official_missing_cell[:10]:
    emit(f"    {p['title']}/{p['mark']} official={p['official_revision_text']} status={p['status']}")

# --- 4. Кits with agreed/code A but a transfer order conflict -----------------
conflicts = conn.execute(
    "SELECT title, mark, problem_kinds_json FROM kit_revision_cell"
).fetchall()
order_kits = {
    (c["title"], c["mark"])
    for c in conflicts
    if allowed(c["title"], c["mark"]) and "transfer_order_conflict" in (c["problem_kinds_json"] or "")
}
mtime_kits = {
    (c["title"], c["mark"])
    for c in conflicts
    if allowed(c["title"], c["mark"]) and "transfer_mtime_conflict" in (c["problem_kinds_json"] or "")
}
agreed_kits = {(p["title"], p["mark"]) for p in pipes if p["status"] == "agreed"}
emit()
emit("=== 4. transfer order / mtime conflicts ===")
emit(f"  kits with transfer_order_conflict: {len(order_kits)}")
emit(f"  kits with transfer_mtime_conflict: {len(mtime_kits)}")
emit(f"  kits 'Согласован' AND order conflict: {len(agreed_kits & order_kits)}")
for key in sorted(agreed_kits & order_kits)[:25]:
    emit(f"    {key[0]}/{key[1]}")

# --- 5. Latest package is NOT the highest revision ----------------------------
by_kit: dict[tuple[str, str], list[sqlite3.Row]] = {}
for r in rows:
    by_kit.setdefault((r["title"], r["mark"]), []).append(r)


def rank(text: str) -> tuple[int, int]:
    text = (text or "").strip()
    if not text:
        return (-2, 0)
    head, _, tail = text.partition("-AN")
    try:
        return (int(head), int(tail) if tail else 0)
    except ValueError:
        return (-1, 0)


inverted = []
for key, pkgs in by_kit.items():
    ordered = sorted(pkgs, key=lambda r: (r["sequence"] if r["sequence"] is not None else -1))
    if len(ordered) < 2:
        continue
    last = ordered[-1]
    best = max(ordered, key=lambda r: rank(r["revision_text"]))
    if rank(best["revision_text"]) > rank(last["revision_text"]):
        inverted.append((key, last, best))

emit()
emit("=== 5. kits where the LAST package (highest NN) is not the highest revision ===")
emit(f"  count: {len(inverted)} of {len(by_kit)} kits with RD packages")
for key, last, best in inverted[:30]:
    emit(
        f"    {key[0]}/{key[1]}: last NN={last['sequence']} '{last['transfer_name']}' rev={last['revision_text']}"
        f"  |  highest rev NN={best['sequence']} '{best['transfer_name']}' rev={best['revision_text']}"
    )

OUT.write_text("\n".join(lines), encoding="utf-8")
print(f"written: {OUT}")
