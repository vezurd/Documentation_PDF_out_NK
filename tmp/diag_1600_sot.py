"""Read-only dump of the 1600/SOT kit state from the runtime rd_catalog SQLite.

Nothing is written to the database; the connection is opened in read-only URI mode.
Output goes to tmp/diag_1600_sot.txt (UTF-8) to avoid console codepage issues.
"""

from __future__ import annotations

import os
import sqlite3
from pathlib import Path

TITLE = "1600"
MARK = "SOT"

DB = Path(os.environ["LOCALAPPDATA"]) / "Documentation_PDF_out_NK" / "rd_catalog" / "rd_catalog.sqlite"
OUT = Path(__file__).resolve().parent / "diag_1600_sot.txt"

lines: list[str] = []


def emit(text: str = "") -> None:
    lines.append(text)


def dump(conn: sqlite3.Connection, caption: str, sql: str, params: tuple = ()) -> None:
    emit(f"=== {caption} ===")
    try:
        cur = conn.execute(sql, params)
    except sqlite3.Error as exc:
        emit(f"  SQL error: {exc}")
        emit()
        return
    cols = [d[0] for d in cur.description]
    rows = cur.fetchall()
    if not rows:
        emit("  (no rows)")
    for row in rows:
        emit("  " + " | ".join(f"{c}={row[c]!r}" for c in cols))
    emit(f"  -- {len(rows)} row(s)")
    emit()


conn = sqlite3.connect(f"file:{DB}?mode=ro", uri=True)
conn.row_factory = sqlite3.Row

emit(f"DB: {DB}")
emit(f"tables: {[r[0] for r in conn.execute('SELECT name FROM sqlite_master WHERE type=\"table\" ORDER BY name')]}")
emit()

dump(conn, "kit_pipeline", "SELECT * FROM kit_pipeline WHERE title=? AND mark=? COLLATE NOCASE", (TITLE, MARK))
dump(
    conn,
    "kit_package",
    "SELECT * FROM kit_package WHERE title=? AND mark=? COLLATE NOCASE ORDER BY sequence",
    (TITLE, MARK),
)
dump(
    conn,
    "kit_revision_cell",
    "SELECT * FROM kit_revision_cell WHERE title=? AND mark=? COLLATE NOCASE",
    (TITLE, MARK),
)
dump(
    conn,
    "google_kit",
    "SELECT * FROM google_kit WHERE title=? AND mark=? COLLATE NOCASE",
    (TITLE, MARK),
)
dump(
    conn,
    "google_event",
    """SELECT e.* FROM google_event e JOIN google_kit k ON e.kit_id = k.id
       WHERE k.title=? AND k.mark=? COLLATE NOCASE ORDER BY e.seq""",
    (TITLE, MARK),
)
dump(
    conn,
    "issuance_send",
    "SELECT * FROM issuance_send WHERE title=? AND mark=? COLLATE NOCASE",
    (TITLE, MARK),
)
dump(
    conn,
    "kit_cycle",
    "SELECT * FROM kit_cycle WHERE title=? AND mark=? COLLATE NOCASE",
    (TITLE, MARK),
)

# RD/SQ/robot files of the kit with overlay state.
dump(
    conn,
    "file_entry + overlay (RD/SQ/robot)",
    """SELECT f.id, f.source, f.kind, f.revision, f.appendix, f.present,
              f.transfer_sequence, f.transfer_name, f.transfer_is_as_build,
              f.mtime_ns, f.path,
              (SELECT o.detected_current FROM overlay_state o WHERE o.path_key = f.path_key) AS cur
       FROM file_entry f
       WHERE f.title_system=? AND f.mark=? COLLATE NOCASE
       ORDER BY f.source, f.transfer_sequence, f.path""",
    (TITLE, MARK),
)

OUT.write_text("\n".join(lines), encoding="utf-8")
print(f"written: {OUT} ({len(lines)} lines)")
