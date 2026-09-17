"""Compute actual «Проверить передачи» notes for 7120 SOT (read-only)."""

from __future__ import annotations

from datetime import datetime
from pathlib import Path

from rd_catalog.config import load_config
from rd_catalog.db import CatalogDatabase
from rd_catalog.issuance_review import latest_effective_issuance_kits
from rd_catalog.kits import (
    _agreed_cycle_package_notes,
    _transfer_review_notes,
    kit_identity_key,
)
from rd_catalog.models import SourceKind
from rd_catalog.pipeline import official_detected_current_ids
from rd_catalog.startup_hydrate import overlay_current_ids

OUT = Path(__file__).with_name("diag_7120_sot_notes.txt")
KEY = kit_identity_key("7120", "SOT")


def fmt_ns(ns: object) -> str:
    try:
        value = int(ns or 0)
    except (TypeError, ValueError):
        return "—"
    if value <= 0:
        return "—"
    return datetime.fromtimestamp(value / 1e9).strftime("%Y.%m.%d %H:%M:%S")


def main() -> None:
    lines: list[str] = []

    def emit(*parts: object) -> None:
        lines.append(" ".join(str(p) for p in parts))

    cfg = load_config()
    db = CatalogDatabase(cfg.db_path)
    records = db.list_files()
    overlay_rows = db.current_overlay()
    by_id = {record.id: record for record in records}
    detected = overlay_current_ids(overlay_rows, by_id, str(cfg.rd_root))
    pipelines = db.list_kit_pipelines()
    official = official_detected_current_ids(records, detected, pipelines)
    google_kits = {kit_identity_key(k.title, k.mark): k for k in db.list_google_kits()}
    issuance = {
        kit_identity_key(k.title, k.mark): k
        for k in latest_effective_issuance_kits(db)
    }
    google = google_kits.get(KEY)
    iss = issuance.get(KEY)

    emit("KEY", KEY)
    emit("google sheet", getattr(google, "sheet_revision_text", None))
    emit("issuance", getattr(iss, "revision_text", None), getattr(iss, "send_date", None))
    if google is not None:
        emit("last event", google.last_event)
        for event in google.events:
            emit(" F", event.date, event.stage, event.revision, event.appendix, event.raw)

    kit_records = [
        rec
        for rec in records
        if rec.source is SourceKind.RD
        and rec.present
        and str(rec.data.get("title") or "").strip() == "7120"
        and str(rec.data.get("mark") or "").strip().casefold() == "sot"
    ]
    emit("\n=== all present RD files ===")
    for rec in sorted(
        kit_records,
        key=lambda item: (
            int(item.data.get("transfer_sequence") or -1),
            str(item.data.get("file_kind") or ""),
            str(item.data.get("name") or ""),
        ),
    ):
        emit(
            f"id={rec.id} seq={rec.data.get('transfer_sequence')} "
            f"kind={rec.data.get('file_kind')} disc={rec.data.get('discipline_block')} "
            f"rev={rec.data.get('revision')}-{rec.data.get('appendix')} "
            f"mtime={fmt_ns(rec.data.get('mtime_ns'))} "
            f"disk={fmt_ns(rec.data.get('disk_mtime_ns'))} "
            f"override_stale={rec.data.get('mtime_override_stale')} "
            f"overlay={int(rec.id in detected)} official={int(rec.id in official)} "
            f"folder={rec.data.get('transfer_name')} name={rec.data.get('name')}"
        )

    notes, pairs = _transfer_review_notes(
        records,
        official,
        google_by_key=google_kits,
        issuance_by_key=issuance,
    )
    agreed, agreed_pairs = _agreed_cycle_package_notes(
        records,
        google_by_key=google_kits,
        issuance_by_key=issuance,
    )
    emit("\n=== inversion notes ===")
    for line in notes.get(KEY, ("<none>",)):
        emit(line)
    emit("pairs", pairs.get(KEY, ()))
    emit("\n=== agreed notes ===")
    for line in agreed.get(KEY, ("<none>",)):
        emit(line)
    emit("agreed pairs", agreed_pairs.get(KEY, ()))

    emit("\n=== collisions mentioning 7120 ===")
    import sqlite3, json as json_lib
    con = sqlite3.connect(f"file:{cfg.db_path.as_posix()}?mode=ro", uri=True)
    con.row_factory = sqlite3.Row
    for row in con.execute(
        """
        SELECT kind, message, document_key, path_keys_json
        FROM current_collision
        WHERE document_key LIKE '%7120%'
           OR path_keys_json LIKE '%7120%'
           OR message LIKE '%7120%'
        """
    ):
        emit(row["kind"], row["document_key"], row["message"], row["path_keys_json"][:400])
    emit("\n=== mtime overrides ===")
    for row in con.execute(
        "SELECT path_key, override_date, reason FROM file_mtime_override WHERE path_key LIKE '%7120%'"
    ):
        emit(dict(row))
    con.close()

    emit("\n=== overlay_state join for kit file ids ===")
    # already printed overlay flags above
    emit("overlay current count", sum(1 for rec in kit_records if rec.id in detected))
    emit("official current count", sum(1 for rec in kit_records if rec.id in official))
    emit(
        "kit file ids overlay",
        sorted(rec.id for rec in kit_records if rec.id in detected),
    )

    OUT.write_text("\n".join(lines) + "\n", encoding="utf-8")
    print(f"wrote {OUT}", flush=True)


if __name__ == "__main__":
    main()
