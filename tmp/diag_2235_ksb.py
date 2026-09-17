"""Read-only dump of kit 2235-KSB from the runtime catalog SQLite."""

from __future__ import annotations

import json
import os
import sqlite3
import sys
from pathlib import Path

DB = (
    Path(os.environ["LOCALAPPDATA"])
    / "Documentation_PDF_out_NK"
    / "rd_catalog"
    / "rd_catalog.sqlite"
)


def _ascii(value: object) -> str:
    return ascii(value)


def main() -> None:
    out_path = Path(__file__).with_suffix(".txt")
    sys.stdout = out_path.open("w", encoding="utf-8")
    con = sqlite3.connect(f"file:{DB.as_posix()}?mode=ro", uri=True)
    con.row_factory = sqlite3.Row
    cur = con.cursor()

    print("=== kit_pipeline ===")
    for row in cur.execute(
        "SELECT * FROM kit_pipeline WHERE title = ? AND mark = ?",
        ("2235", "KSB"),
    ):
        print(dict(row))

    print("\n=== kit_package ===")
    pkg_cols = [item[1] for item in cur.execute("PRAGMA table_info(kit_package)")]
    print("cols", pkg_cols)
    for row in cur.execute(
        """
        SELECT *
        FROM kit_package
        WHERE title = ? AND mark = ?
        ORDER BY source, sequence
        """,
        ("2235", "KSB"),
    ):
        data = dict(row)
        print(
            {
                "id": data.get("id"),
                "source": data.get("source"),
                "seq": data.get("sequence"),
                "rev": data.get("revision_text"),
                "mto": data.get("mto_revision_text"),
                "overlay": data.get("overlay_current_count"),
                "grey": data.get("is_grey"),
                "current": data.get("is_current"),
                "as_build": data.get("is_as_build"),
                "name": data.get("transfer_name"),
                "path": (data.get("package_path") or "")[-140:],
            }
        )

    print("\n=== present RD files ===")
    for row in cur.execute(
        """
        SELECT id, present, file_kind, name, revision, appendix,
               transfer_sequence, transfer_name, transfer_is_as_build,
               parse_status, path
        FROM file_entry
        WHERE source = 'rd' AND title = ? AND mark = ? AND present = 1
        ORDER BY COALESCE(transfer_sequence, 0), file_kind, name
        """,
        ("2235", "KSB"),
    ):
        data = dict(row)
        print(
            f"{data['id']:>6} seq={data['transfer_sequence']} "
            f"asb={data['transfer_is_as_build']} {data['file_kind']:12} "
            f"rev={data['revision']}-{data['appendix']}  {data['transfer_name']}"
        )
        print(f"       {data['name']}")
        print(f"       ...{data['path'][-130:]}")

    print("\n=== overlay current RD ===")
    for row in cur.execute(
        """
        SELECT f.id, f.file_kind, f.name, f.revision, f.appendix,
               f.transfer_sequence, f.transfer_name, o.detected_current
        FROM overlay_state o
        JOIN file_entry f ON f.id = o.file_id
        WHERE f.source = 'rd' AND f.title = ? AND f.mark = ?
        ORDER BY f.file_kind, f.name
        """,
        ("2235", "KSB"),
    ):
        print(dict(row))

    print("\n=== google_kit / events ===")
    for row in cur.execute(
        "SELECT id, title, mark, sheet_revision_text, status_sheet FROM google_kit WHERE title = ? AND mark = ?",
        ("2235", "KSB"),
    ):
        print(dict(row))
        kit_id = dict(row).get("id")
        if kit_id is None:
            continue
        for event in cur.execute(
            "SELECT seq, event_date, stage, revision, raw FROM google_event WHERE kit_id = ? ORDER BY seq",
            (kit_id,),
        ):
            print("  ev", dict(event))

    print("\n=== issuance_send ===")
    send_cols = [item[1] for item in cur.execute("PRAGMA table_info(issuance_send)")]
    print("cols", send_cols)
    for row in cur.execute(
        "SELECT * FROM issuance_send WHERE title = ? AND mark = ? ORDER BY id",
        ("2235", "KSB"),
    ):
        data = dict(row)
        print(
            {
                k: data[k]
                for k in data
                if k
                in (
                    "id",
                    "revision_text",
                    "status",
                    "send_date",
                    "send_transmittal",
                    "note_raw",
                )
                or "rev" in k
                or "date" in k
            }
        )

    print("\n=== kit_working_flag ===")
    try:
        for row in cur.execute(
            "SELECT * FROM kit_working_flag WHERE title = ? AND mark = ?",
            ("2235", "KSB"),
        ):
            print(dict(row))
    except sqlite3.Error as exc:
        print(exc)

    print("\n=== kit_revision_cell (if any) ===")
    tables = [
        item[0]
        for item in cur.execute(
            "SELECT name FROM sqlite_master WHERE type='table' ORDER BY name"
        )
    ]
    print("tables", tables)
    if "kit_revision_cell" in tables:
        for row in cur.execute(
            """
            SELECT revision_text, pipeline_status, gap_kind, is_current
            FROM kit_revision_cell
            WHERE title = ? AND mark = ?
            ORDER BY revision_text
            """,
            ("2235", "KSB"),
        ):
            print(dict(row))


if __name__ == "__main__":
    main()
