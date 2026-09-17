"""Read-only: are the «no MTO at the approved revision» rows real gaps or a keying artefact?

For every kit_revision_cell with an approval letter and has_mto=0, look at the RD
packages linked to that cell and check whether they physically contain an MTO xlsx
(under any filename revision).
"""

from __future__ import annotations

import json
import os
import sqlite3
from pathlib import Path

DB = Path(os.environ["LOCALAPPDATA"]) / "Documentation_PDF_out_NK" / "rd_catalog" / "rd_catalog.sqlite"
OUT = Path(__file__).resolve().parent / "diag_false_gaps.txt"

conn = sqlite3.connect(f"file:{DB}?mode=ro", uri=True)
conn.row_factory = sqlite3.Row
lines: list[str] = []

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

packages = {row["id"]: row for row in conn.execute("SELECT * FROM kit_package")}

cells = [
    row
    for row in conn.execute(
        "SELECT * FROM kit_revision_cell WHERE pipeline_status='code_a' AND has_mto=0"
    )
    if (row["title"].casefold(), row["mark"].casefold()) not in banned
]

artefact = 0
real_gap = 0
no_package = 0
examples_artefact: list[str] = []
examples_gap: list[str] = []

for cell in cells:
    ids = json.loads(cell["package_ids_json"] or "[]")
    rd = [packages[i] for i in ids if i in packages and packages[i]["source"] == "rd" and not packages[i]["is_grey"]]
    if not rd:
        no_package += 1
        continue
    with_mto = [p for p in rd if (p["mto_revision_text"] or "").strip()]
    label = (
        f"{cell['title']}/{cell['mark']} cell_rev={cell['revision_text']} -> "
        + ", ".join(
            f"NN{p['sequence']}('{p['transfer_name']}' pkg={p['revision_text']} mto={p['mto_revision_text'] or '-'})"
            for p in sorted(rd, key=lambda p: (p["sequence"] or -1))[:4]
        )
    )
    if with_mto:
        artefact += 1
        if len(examples_artefact) < 25:
            examples_artefact.append(label)
    else:
        real_gap += 1
        if len(examples_gap) < 20:
            examples_gap.append(label)

lines.append("=== cells: status=code_a AND has_mto=0 ===")
lines.append(f"  total: {len(cells)}")
lines.append(f"  package of that revision DOES contain an MTO (other filename rev) -> keying artefact: {artefact}")
lines.append(f"  package of that revision has NO MTO at all -> real gap: {real_gap}")
lines.append(f"  no RD package linked to the cell: {no_package}")
lines.append("")
lines.append("--- artefact examples (MTO exists in the package, other filename rev) ---")
lines.extend("  " + item for item in examples_artefact)
lines.append("")
lines.append("--- real-gap examples (no MTO in any package of that revision) ---")
lines.extend("  " + item for item in examples_gap)

OUT.write_text("\n".join(lines), encoding="utf-8")
print(f"written: {OUT}")
