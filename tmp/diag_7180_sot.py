"""Read-only dump of 7180 SOT approval vs F journal."""

from __future__ import annotations

import sqlite3
from pathlib import Path

OUT = Path(__file__).with_name("diag_7180_sot_out.txt")
DB = Path.home() / "AppData/Local/Documentation_PDF_out_NK/rd_catalog/rd_catalog.sqlite"


def main() -> None:
    lines: list[str] = []

    def emit(*parts: object) -> None:
        lines.append(" ".join(str(p) for p in parts))

    con = sqlite3.connect(f"file:{DB.as_posix()}?mode=ro", uri=True)
    con.row_factory = sqlite3.Row
    cur = con.cursor()

    emit("=== kit_pipeline ===")
    for r in cur.execute(
        "SELECT * FROM kit_pipeline WHERE title='7180' AND lower(mark)='sot'"
    ):
        emit(dict(r))

    emit("\n=== google_kit ===")
    kit = cur.execute(
        "SELECT * FROM google_kit WHERE title='7180' AND lower(mark)='sot'"
    ).fetchone()
    if kit is None:
        emit("missing")
        kit_id = None
    else:
        emit(dict(kit))
        kit_id = kit["id"]

    emit("\n=== google_event ===")
    if kit_id is not None:
        for r in cur.execute(
            "SELECT * FROM google_event WHERE kit_id=? ORDER BY seq",
            (kit_id,),
        ):
            emit(dict(r))

    emit("\n=== issuance_send ===")
    for r in cur.execute(
        """
        SELECT id, revision_text, status, send_date, incoming_control_date,
               send_transmittal, confirm_transmittal, note_raw, row_index
        FROM issuance_send
        WHERE title='7180' AND lower(mark)='sot'
        ORDER BY send_date_sortable, id
        """
    ):
        emit(dict(r))

    emit("\n=== issuance_review ===")
    try:
        for r in cur.execute(
            """
            SELECT * FROM issuance_review
            WHERE title='7180' AND lower(mark)='sot'
            """
        ):
            emit(dict(r))
    except sqlite3.OperationalError as exc:
        emit(exc)

    emit("\n=== kit_cycle ===")
    for r in cur.execute(
        """
        SELECT c.* FROM kit_cycle c
        WHERE c.title='7180' AND lower(c.mark)='sot'
        ORDER BY c.id
        """
    ):
        emit(dict(r))

    emit("\n=== kit_package rd ===")
    for r in cur.execute(
        """
        SELECT source, sequence, transfer_name, revision_text, mto_revision_text,
               is_current, is_grey
        FROM kit_package
        WHERE title='7180' AND lower(mark)='sot'
        ORDER BY source, sequence
        """
    ):
        emit(dict(r))

    emit("\n=== kit_revision_cell ===")
    for r in cur.execute(
        """
        SELECT revision_text, pipeline_status, letters, is_current
        FROM kit_revision_cell
        WHERE title='7180' AND lower(mark)='sot'
        ORDER BY revision_text
        """
    ):
        emit(dict(r))

    con.close()
    OUT.write_text("\n".join(lines) + "\n", encoding="utf-8")
    print(f"wrote {OUT}", flush=True)


if __name__ == "__main__":
    main()
