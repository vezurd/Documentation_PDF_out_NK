"""Read-only dump of 1600/SOT MTO files, overlay flags and collisions."""

from __future__ import annotations

import os
import sqlite3
from pathlib import Path

TITLE = "1600"
MARK = "SOT"

DB = Path(os.environ["LOCALAPPDATA"]) / "Documentation_PDF_out_NK" / "rd_catalog" / "rd_catalog.sqlite"
OUT = Path(__file__).resolve().parent / "diag_1600_sot_files.txt"

lines: list[str] = []
conn = sqlite3.connect(f"file:{DB}?mode=ro", uri=True)
conn.row_factory = sqlite3.Row

lines.append("=== file_entry columns ===")
lines.append(", ".join(r[1] for r in conn.execute("PRAGMA table_info(file_entry)")))
lines.append("")
lines.append("=== overlay_state columns ===")
lines.append(", ".join(r[1] for r in conn.execute("PRAGMA table_info(overlay_state)")))
lines.append("")
lines.append("=== current_collision columns ===")
lines.append(", ".join(r[1] for r in conn.execute("PRAGMA table_info(current_collision)")))
lines.append("")

kind_col = "file_kind"

sql = """
SELECT f.*,
       (SELECT o.detected_current FROM overlay_state o WHERE o.file_id = f.id) AS cur,
       (SELECT o.document_key FROM overlay_state o WHERE o.file_id = f.id) AS dockey
FROM file_entry f
WHERE f.title=? COLLATE NOCASE AND f.mark=? COLLATE NOCASE
ORDER BY f.source, f.transfer_sequence, f.path
"""
cur = conn.execute(sql, (TITLE, MARK))
names = [d[0] for d in cur.description]
rows = cur.fetchall()

lines.append(f"=== all files ({len(rows)}) — MTO xlsx first ===")
show = [
    n
    for n in names
    if n
    in {
        "id",
        "source",
        "file_kind",
        "revision",
        "appendix",
        "present",
        "transfer_sequence",
        "transfer_name",
        "name",
        "cur",
    }
]


def is_mto(row: sqlite3.Row) -> bool:
    value = (row[kind_col] if kind_col else "") or ""
    return "mto" in str(value).lower() or ".xls" in row["path"].lower()


lines.append("--- MTO / xlsx ---")
for row in rows:
    if is_mto(row):
        lines.append("  " + " | ".join(f"{n}={row[n]!r}" for n in show))
lines.append("")
lines.append("--- overlay-current non-MTO (first 40) ---")
count = 0
for row in rows:
    if not is_mto(row) and row["cur"]:
        lines.append("  " + " | ".join(f"{n}={row[n]!r}" for n in show))
        count += 1
        if count >= 40:
            break
lines.append("")

lines.append("=== collisions mentioning 1600\\SOT ===")
cur = conn.execute("SELECT * FROM current_collision")
cnames = [d[0] for d in cur.description]
hits = 0
for row in cur.fetchall():
    blob = " ".join(str(row[n]) for n in cnames)
    if "1600" in blob and "SOT" in blob.upper():
        lines.append("  " + " | ".join(f"{n}={row[n]!r}" for n in cnames))
        hits += 1
lines.append(f"  -- {hits} collision row(s)")

OUT.write_text("\n".join(lines), encoding="utf-8")
print(f"written: {OUT}")
