"""Read-only dump of 7120 SOT transfer-review facts from the local catalog DB."""

from __future__ import annotations

import json
import sqlite3
import sys
from datetime import datetime
from pathlib import Path

OUT = Path(__file__).with_name("diag_7120_sot_out.txt")

DB = Path.home() / "AppData/Local/Documentation_PDF_out_NK/rd_catalog/rd_catalog.sqlite"
GOOGLE = DB.parent / "google_kits_data.json"
ISSUANCE = DB.parent / "google_issuance_data.json"


def fmt_ns(ns: int | None) -> str:
    if not ns:
        return "—"
    return datetime.fromtimestamp(int(ns) / 1e9).strftime("%Y.%m.%d %H:%M:%S")


def emit(*parts: object) -> None:
    line = " ".join(str(p) for p in parts)
    sys.stdout.buffer.write((line + "\n").encode("utf-8", errors="backslashreplace"))


def main() -> None:
    buf: list[str] = []

    def print(*parts: object) -> None:  # noqa: A001
        line = " ".join(str(p) for p in parts)
        buf.append(line)
        emit(line)

    print("db", DB, "exists", DB.exists(), "size", DB.stat().st_size if DB.exists() else 0)
    con = sqlite3.connect(f"file:{DB.as_posix()}?mode=ro", uri=True)
    con.row_factory = sqlite3.Row
    cur = con.cursor()
    print("tables", [r[0] for r in cur.execute("SELECT name FROM sqlite_master WHERE type='table' ORDER BY 1")])
    print("\n=== kit_pipeline ===")
    for r in cur.execute(
        "SELECT * FROM kit_pipeline WHERE title='7120' AND lower(mark)='sot'"
    ):
        print(dict(r))
    print("\n=== kit_package ===")
    for r in cur.execute(
        """
        SELECT source, sequence, transfer_name, package_path, revision_text,
               mto_revision_text, is_current, is_as_build, is_grey,
               overlay_current_count, pdf_count, editable_count, max_mtime_ns
        FROM kit_package
        WHERE title='7120' AND lower(mark)='sot'
        ORDER BY source, sequence
        """
    ):
        d = dict(r)
        d["mtime"] = fmt_ns(d.get("max_mtime_ns"))
        print(d)
    print("\n=== file_entry columns ===")
    cols = [d[1] for d in cur.execute("PRAGMA table_info(file_entry)")]
    print(cols)
    print("\n=== present RD MTO/OD files ===")
    rows = list(
        cur.execute(
            """
            SELECT id, file_kind, name, present, parse_status, revision, appendix,
                   transfer_sequence, transfer_name, discipline_block, core_stem,
                   mtime_ns, extension, path
            FROM file_entry
            WHERE title='7120' AND lower(mark)='sot' AND source='rd' AND present=1
              AND (
                    lower(discipline_block) LIKE 'mto%'
                    OR lower(discipline_block) LIKE 'od%'
                    OR lower(file_kind) LIKE '%mto%'
                  )
            ORDER BY transfer_sequence, file_kind, name
            """
        )
    )
    print("count", len(rows))
    for r in rows:
        d = dict(r)
        d["mtime"] = fmt_ns(d.get("mtime_ns"))
        print(d)

    print("\n=== overlay_state / current_collision schema ===")
    print("overlay_state", [d[1] for d in cur.execute("PRAGMA table_info(overlay_state)")])
    print("current_collision", [d[1] for d in cur.execute("PRAGMA table_info(current_collision)")])
    print("\n=== overlay_state sample ===")
    for r in cur.execute("SELECT * FROM overlay_state LIMIT 3"):
        print(list(dict(r)))
        break
    print("\n=== collisions for 7120 SOT ===")
    try:
        for r in cur.execute(
            """
            SELECT * FROM current_collision
            WHERE document_key LIKE '%7120%' OR path_keys LIKE '%7120%'
               OR message LIKE '%7120%'
            """
        ):
            print({k: dict(r)[k] for k in dict(r)})
    except sqlite3.OperationalError as exc:
        print("collision query", exc)
        for r in cur.execute("SELECT * FROM current_collision LIMIT 1"):
            print("sample keys", list(dict(r)))

    print("\n=== google_kit / events ===")
    try:
        for r in cur.execute(
            "SELECT * FROM google_kit WHERE title='7120' AND lower(mark)='sot'"
        ):
            print("kit", dict(r))
        for r in cur.execute(
            """
            SELECT e.* FROM google_event e
            JOIN google_kit k ON k.id = e.kit_id
            WHERE k.title='7120' AND lower(k.mark)='sot'
            ORDER BY e.seq
            """
        ):
            print("event", dict(r))
    except sqlite3.OperationalError as exc:
        print("google query", exc)

    print("\n=== issuance_send ===")
    try:
        for r in cur.execute(
            """
            SELECT * FROM issuance_send
            WHERE title='7120' AND lower(mark)='sot'
            ORDER BY send_date, id
            """
        ):
            print(dict(r))
    except sqlite3.OperationalError as exc:
        print("issuance query", exc)

    print("\n=== kit_working / annulled ===")
    for table in ("kit_working_flag", "kit_annulled_flag"):
        try:
            for r in cur.execute(
                f"SELECT * FROM {table} WHERE title='7120' AND lower(mark)='sot'"
            ):
                print(table, dict(r))
        except sqlite3.OperationalError as exc:
            print(table, exc)

    print("\n=== overlay current ids hint ===")
    try:
        for r in cur.execute(
            """
            SELECT id, name, transfer_name, transfer_sequence, file_kind,
                   revision, appendix, detected_current, present
            FROM file_entry
            WHERE title='7120' AND lower(mark)='sot' AND source='rd' AND present=1
            """
        ):
            print(dict(r))
    except sqlite3.OperationalError:
        pass

    con.close()

    print("\n=== google json F cell ===")
    if GOOGLE.exists():
        data = json.loads(GOOGLE.read_text(encoding="utf-8"))
        rows = data.get("rows") or data
        for row in rows:
            if not isinstance(row, list) or len(row) < 6:
                continue
            if str(row[0]).strip() == "7120" and "SOT" in str(row[1]).upper().replace("КСБ", "KSB"):
                print(row[:6])

    print("\n=== issuance json ===")
    if ISSUANCE.exists():
        data = json.loads(ISSUANCE.read_text(encoding="utf-8"))
        rows = data.get("rows") or data
        for row in rows:
            if not isinstance(row, list) or len(row) < 7:
                continue
            mark = str(row[2] if len(row) > 2 else "").upper()
            title = str(row[4] if len(row) > 4 else "")
            if title.strip() == "7120" and "SOT" in mark:
                print(row[:7], "P=", row[15] if len(row) > 15 else "", "Q=", row[16] if len(row) > 16 else "")

    OUT.write_text("\n".join(buf) + "\n", encoding="utf-8")
    emit("wrote", OUT)


if __name__ == "__main__":
    main()
