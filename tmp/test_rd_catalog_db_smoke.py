"""Local smoke checks for RD catalog SQLite APIs."""

from __future__ import annotations

import tempfile
import sys
import sqlite3
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from rd_catalog.an_index import AnMtoFile
from rd_catalog.db import SCHEMA_VERSION, CatalogDatabase, mto_file_stat_signature
from rd_catalog.issuance_review import send_identity_fingerprint
from rd_catalog.kits import GoogleKit, IssuanceKit, KitEvent, kit_identity_key
from rd_catalog.models import (
    CollisionKind,
    OverlayCollision,
    ReviewState,
    ScanRunStatus,
    ScanSummary,
    SourceError,
    SourceKind,
    SourceScanResult,
)
from rd_catalog.scan import scan_document_source
from rd_catalog.overlay import build_rd_overlays

_V3_TABLES = (
    "google_kit",
    "google_event",
    "issuance_send",
    "kit_package",
    "kit_cycle",
    "kit_pipeline",
    "kit_liquidity_review",
)
_V4_TABLES = _V3_TABLES + ("kit_revision_cell",)


def _table_names(path: Path) -> set[str]:
    connection = sqlite3.connect(path)
    try:
        rows = connection.execute(
            "SELECT name FROM sqlite_master WHERE type = 'table'"
        ).fetchall()
    finally:
        connection.close()
    return {str(row[0]) for row in rows}


def _sample_google_kit() -> GoogleKit:
    events = (
        KitEvent(
            raw="11.04.2025 отправлен на ТДО AGCC-BCC-TRM-000001",
            date="11.04.2025",
            stage="tdo_sent",
            stage_label="отпр. на ТДО",
            revision=None,
            appendix=None,
            transmittals=("AGCC-BCC-TRM-000001",),
            parsed=True,
        ),
        KitEvent(
            raw="20.05.2025 код А AGCC-BCC-TRM-000002",
            date="20.05.2025",
            stage="code_a",
            stage_label="код А",
            revision="01",
            appendix="02",
            transmittals=("AGCC-BCC-TRM-000002",),
            parsed=True,
        ),
    )
    return GoogleKit(
        title="9110",
        mark="KSB1",
        mark_raw="KSB1",
        title_system="9110-KSB1",
        sheet_revision="01",
        sheet_appendix="02",
        sheet_revision_text="01-AN02",
        status_sheet="Прошла ТДО",
        comment_raw="\n".join(event.raw for event in events),
        events=events,
        last_event=events[1],
        row_index=12,
    )


def _sample_issuance_send(*, row_index: int, transmittal: str, send_date: str) -> IssuanceKit:
    return IssuanceKit(
        title="9110",
        mark="KSB1",
        mark_raw="KSB1",
        title_system="9110-KSB1",
        revision="01",
        appendix="02",
        revision_text="01-AN02",
        status="На рассмотрении вх.контроля",
        send_date=send_date,
        send_date_sortable=".".join(reversed(send_date.split("."))),
        send_transmittal=transmittal,
        incoming_control_date="",
        incoming_control_date_sortable="",
        confirm_transmittal="",
        note_raw="",
        row_index=row_index,
    )


def main() -> None:
    """Verify reopen, schema, review history, and comparison cache."""

    assert SCHEMA_VERSION == 13
    with tempfile.TemporaryDirectory(prefix="rd_catalog_db_") as temp:
        legacy_path = Path(temp, "legacy_v1.sqlite")
        connection = sqlite3.connect(legacy_path)
        try:
            connection.execute(
                "CREATE TABLE schema_meta "
                "(key TEXT PRIMARY KEY, value TEXT NOT NULL)"
            )
            CatalogDatabase._migrate_0_to_1(connection)
            connection.execute("CREATE TABLE legacy_guard(value TEXT)")
            connection.execute("INSERT INTO legacy_guard VALUES ('preserved')")
            connection.commit()
        finally:
            connection.close()
        migrated = CatalogDatabase(legacy_path)
        migrated.initialize()
        assert migrated.schema_version() == SCHEMA_VERSION
        tables = _table_names(legacy_path)
        for table_name in _V4_TABLES:
            assert table_name in tables
        assert "kit_liquidity_review" in tables
        assert "kit_revision_cell" in tables
        assert "issuance_review" in tables
        assert "an_mto_file" in tables
        assert "rd_dump_mto_file" in tables
        assert "kit_working_flag" in tables
        assert "kit_annulled_flag" in tables
        assert "file_mtime_override" in tables
        connection = sqlite3.connect(legacy_path)
        try:
            assert connection.execute(
                "SELECT value FROM legacy_guard"
            ).fetchone() == ("preserved",)
        finally:
            connection.close()

        v3_path = Path(temp, "legacy_v3.sqlite")
        connection = sqlite3.connect(v3_path)
        try:
            connection.execute(
                "CREATE TABLE schema_meta "
                "(key TEXT PRIMARY KEY, value TEXT NOT NULL)"
            )
            CatalogDatabase._migrate_0_to_1(connection)
            CatalogDatabase._migrate_1_to_2(connection)
            CatalogDatabase._migrate_2_to_3(connection)
            connection.commit()
        finally:
            connection.close()
        v3_db = CatalogDatabase(v3_path)
        assert v3_db.schema_version() == 3
        for table_name in _V3_TABLES:
            assert table_name in _table_names(v3_path)
        assert "kit_revision_cell" not in _table_names(v3_path)
        review_id = v3_db.upsert_liquidity_review(
            "9110",
            "KSB1",
            "05_рев.0-AN02",
            "01-AN02",
            decision="confirmed_ok",
            sequence=5,
            comment="keep-across-v4",
            evidence_mtime_ns=1,
        )
        assert review_id > 0
        v3_db.initialize()
        assert v3_db.schema_version() == SCHEMA_VERSION
        v4_tables = _table_names(v3_path)
        assert "kit_revision_cell" in v4_tables
        assert "mto_pair_comparison" in v4_tables
        assert "issuance_review" in v4_tables
        assert "an_mto_file" in v4_tables
        assert "rd_dump_mto_file" in v4_tables
        assert "kit_working_flag" in v4_tables
        assert "kit_annulled_flag" in v4_tables
        assert "file_mtime_override" in v4_tables
        for table_name in _V4_TABLES:
            assert table_name in v4_tables
        reviews = v3_db.list_liquidity_reviews("9110", "KSB1")
        assert len(reviews) == 1
        assert reviews[0].comment == "keep-across-v4"
        connection = sqlite3.connect(v3_path)
        try:
            columns = {
                str(row[1])
                for row in connection.execute("PRAGMA table_info(kit_package)")
            }
        finally:
            connection.close()
        assert "is_as_build" in columns
        connection = sqlite3.connect(v3_path)
        try:
            pipeline_columns = {
                str(row[1])
                for row in connection.execute("PRAGMA table_info(kit_pipeline)")
            }
        finally:
            connection.close()
        for column_name in (
            "code_stale",
            "code_revision_text",
            "code_date",
            "review_as_build",
            "tdo_date",
            "working_as_build",
        ):
            assert column_name in pipeline_columns

        root = Path(temp, "rd")
        transfer = root / "2225" / "KSB" / "Для передачи" / "01_рев.01_2225-KSB"
        transfer.mkdir(parents=True)
        document = transfer / "AGCC.287-2225-KSB.MTO-0001_01_RU.xlsx"
        document.write_bytes(b"fixture")

        source_result = scan_document_source(
            root,
            SourceKind.RD,
            skip_dirs=("old", "tmp"),
        )
        summary = ScanSummary(sources={SourceKind.RD: source_result})
        summary.rd_pdf_overlay, summary.rd_mto_overlay = build_rd_overlays(
            summary.files
        )
        source_collision = OverlayCollision(
            CollisionKind.UNPARSED_FILE,
            "source collision",
            (source_result.files[0].path_key,),
            "mto:2225-ksb|mto-0001",
        )
        overlay_collision = OverlayCollision(
            CollisionKind.DUP_SAME_REVISION,
            "overlay collision",
            (source_result.files[0].path_key,),
            "mto:2225-ksb|mto-0001",
        )
        source_result.collisions.append(source_collision)
        summary.rd_mto_overlay.collisions.append(overlay_collision)

        db_path = Path(temp, "runtime", "rd_catalog.sqlite")
        database = CatalogDatabase(db_path)
        assert not db_path.exists()
        database.initialize()
        assert database.schema_version() == SCHEMA_VERSION
        assert "kit_liquidity_review" in _table_names(db_path)
        assert "kit_revision_cell" in _table_names(db_path)
        assert "mto_pair_comparison" in _table_names(db_path)
        assert "issuance_review" in _table_names(db_path)
        assert "an_mto_file" in _table_names(db_path)
        assert "rd_dump_mto_file" in _table_names(db_path)
        assert "kit_working_flag" in _table_names(db_path)
        assert "kit_annulled_flag" in _table_names(db_path)
        assert "file_mtime_override" in _table_names(db_path)
        database.store_scan(summary)
        collisions = database.list_current_collisions()
        assert {row["scope"] for row in collisions} == {"source", "overlay_mto"}
        assert all(row["paths"] == [str(document)] for row in collisions)

        failed_result = scan_document_source(
            root,
            SourceKind.RD,
            skip_dirs=("old", "tmp"),
        )
        failed_result.errors.append(SourceError(SourceKind.RD, str(root), "failed"))
        failed = ScanSummary(
            sources={SourceKind.RD: failed_result},
            status=ScanRunStatus.PARTIAL,
        )
        failed.rd_pdf_overlay, failed.rd_mto_overlay = build_rd_overlays(
            failed.files
        )
        database.store_scan(failed)
        assert database.list_current_collisions() == collisions

        cancelled_result = scan_document_source(
            root,
            SourceKind.RD,
            skip_dirs=("old", "tmp"),
        )
        cancelled_result.cancelled = True
        cancelled = ScanSummary(
            sources={SourceKind.RD: cancelled_result},
            status=ScanRunStatus.CANCELLED,
        )
        cancelled.rd_pdf_overlay, cancelled.rd_mto_overlay = build_rd_overlays(
            cancelled.files
        )
        database.store_scan(cancelled)
        assert database.list_current_collisions() == collisions

        replacement_result = scan_document_source(
            root,
            SourceKind.RD,
            skip_dirs=("old", "tmp"),
        )
        replacement = ScanSummary(sources={SourceKind.RD: replacement_result})
        replacement.rd_pdf_overlay, replacement.rd_mto_overlay = build_rd_overlays(
            replacement.files
        )
        database.store_scan(replacement)
        assert database.list_current_collisions() == []
        last_scan = database.last_scan_info()
        assert last_scan is not None
        assert last_scan["id"] == 4

        reopened = CatalogDatabase(db_path)
        assert reopened.schema_version() == SCHEMA_VERSION
        records = reopened.list_files()
        assert len(records) == 1
        file_id = records[0].id
        disk_mtime = int(records[0].data.get("mtime_ns") or 0)
        override = reopened.upsert_file_mtime_override(
            records[0].path_key, "24.03.2024", reason="code_b"
        )
        assert override.reason == "code_b"
        assert override.disk_mtime_ns == disk_mtime
        assert override.override_date == "24.03.2024"
        catalog = reopened.list_files()[0]
        assert catalog.data.get("mtime_override_applied") is True
        assert int(catalog.data.get("disk_mtime_ns") or 0) == disk_mtime
        assert int(catalog.data.get("mtime_ns") or 0) != disk_mtime
        assert reopened.upsert_file_mtime_override(
            records[0].path_key, "07.07.2023", reason="folder_mean"
        ).reason == "folder_mean"
        assert reopened.upsert_file_mtime_override(
            records[0].path_key, "08.07.2023", reason="manual"
        ).reason == "manual"
        override = reopened.upsert_file_mtime_override(
            records[0].path_key, "24.03.2024", reason="code_b"
        )
        assert override.reason == "code_b"
        signature = mto_file_stat_signature(catalog)
        assert signature["mtime_ns"] == disk_mtime
        reopened.store_scan(replacement)
        kept = reopened.list_files()[0]
        assert kept.data.get("mtime_override_applied") is True
        assert int(kept.data.get("disk_mtime_ns") or 0) == disk_mtime

        reopened.record_review(
            file_id,
            ReviewState.IGNORED,
            comment="test-only ignore",
        )
        history = reopened.review_history(file_id)
        assert history[-1]["action"] == ReviewState.IGNORED.value
        assert records[0].path_key in reopened.get_ignored_path_keys()

        comparison_id = reopened.record_mto_comparison(
            file_id,
            file_id,
            rd_fingerprint="rd",
            robot_fingerprint="robot",
            algorithm_version=1,
            status="content_equal",
            stats={"rows": 1},
            diff={},
        )
        cached = reopened.get_mto_comparison(file_id, file_id, 1)
        assert comparison_id > 0
        assert cached is not None
        assert cached["stats"] == {"rows": 1}

        review_id = reopened.upsert_liquidity_review(
            "9110",
            "KSB1",
            "05_рев.0-AN02_AGCC.287-9110-KSB1",
            "01-AN02",
            decision="confirmed_ok",
            sequence=5,
            comment="keep",
            evidence_mtime_ns=1_700_000_000_000_000_000,
        )
        assert review_id > 0
        working_flag_id = reopened.upsert_working_flag(
            "9110",
            "KSB1",
            "01-AN03",
            transfer_name="06_рев.01-AN03",
        )
        assert working_flag_id > 0
        sibling_flag_id = reopened.upsert_working_flag(
            "9110",
            "KSB1",
            "01-AN03",
            transfer_name="07_рев.01-AN03",
        )
        assert sibling_flag_id > 0
        assert sibling_flag_id != working_flag_id
        kit = _sample_google_kit()
        sends = (
            _sample_issuance_send(
                row_index=20,
                transmittal="AGCC-BCC-TRM-000001",
                send_date="01.09.2026",
            ),
            _sample_issuance_send(
                row_index=21,
                transmittal="AGCC-BCC-TRM-000002",
                send_date="15.09.2026",
            ),
        )
        issuance_review_id = reopened.upsert_issuance_review(
            "9110",
            "KSB1",
            "send",
            "issuance",
            send_identity_fingerprint(sends[1]),
            decision="annulled",
            comment="keep-issuance",
            revision_text=sends[1].revision_text,
            send_date=sends[1].send_date,
            send_date_sortable=sends[1].send_date_sortable,
            send_transmittal=sends[1].send_transmittal,
            match_state="matched",
        )
        assert issuance_review_id > 0
        reopened.replace_google_snapshot(
            (kit,),
            sends,
            loaded_at="2026-09-04T12:00:00+00:00",
            source="test",
            warning=None,
        )
        reviews = reopened.list_liquidity_reviews("9110", "KSB1")
        assert len(reviews) == 1
        assert reviews[0].decision == "confirmed_ok"
        assert reviews[0].transfer_name.startswith("05_рев")
        issuance_reviews = reopened.list_issuance_reviews("9110", "KSB1")
        assert len(issuance_reviews) == 1
        assert issuance_reviews[0].decision == "annulled"
        assert issuance_reviews[0].comment == "keep-issuance"
        stored_kits = reopened.list_google_kits()
        assert len(stored_kits) == 1
        assert stored_kits[0].title == "9110"
        assert stored_kits[0].mark == "KSB1"
        assert stored_kits[0].sheet_revision == "01"
        assert len(stored_kits[0].events) == 2
        events = reopened.list_google_events("9110", "KSB1")
        assert [event.stage for event in events] == ["tdo_sent", "code_a"]
        assert events[0].transmittals == ("AGCC-BCC-TRM-000001",)
        stored_sends = reopened.list_issuance_sends("9110", "KSB1")
        assert len(stored_sends) == 2
        assert {send.send_transmittal for send in stored_sends} == {
            "AGCC-BCC-TRM-000001",
            "AGCC-BCC-TRM-000002",
        }
        assert len(reopened.list_files()) == 1
        reopened.replace_kit_derived((), (), ())
        assert len(reopened.list_liquidity_reviews("9110", "KSB1")) == 1
        assert len(reopened.list_issuance_reviews("9110", "KSB1")) == 1
        flags = reopened.list_working_flags("9110", "KSB1")
        assert len(flags) == 2
        assert {flag.transfer_name for flag in flags} == {
            "06_рев.01-AN03",
            "07_рев.01-AN03",
        }
        assert {flag.sequence for flag in flags} == {6, 7}
        assert {flag.revision_text for flag in flags} == {"01-AN03"}
        reopened.replace_google_snapshot(
            (),
            (),
            loaded_at="2026-09-04T13:00:00+00:00",
            source="test-empty",
        )
        assert reopened.list_google_kits() == []
        assert reopened.list_issuance_sends() == []
        assert len(reopened.list_liquidity_reviews("9110", "KSB1")) == 1
        assert len(reopened.list_issuance_reviews("9110", "KSB1")) == 1
        assert len(reopened.list_working_flags("9110", "KSB1")) == 2
        annulled_id = reopened.upsert_annulled_flag(
            "9110",
            "KSB1",
            "01-AN03",
            transfer_name="06_рев.01-AN03",
        )
        assert annulled_id > 0
        leftover_working = reopened.list_working_flags("9110", "KSB1")
        assert {flag.transfer_name for flag in leftover_working} == {
            "07_рев.01-AN03",
        }
        annulled = reopened.list_annulled_flags("9110", "KSB1")
        assert len(annulled) == 1
        assert annulled[0].transfer_name == "06_рев.01-AN03"
        assert annulled[0].sequence == 6
        restored_working = reopened.upsert_working_flag(
            "9110",
            "KSB1",
            "01-AN03",
            transfer_name="06_рев.01-AN03",
        )
        assert restored_working > 0
        assert reopened.list_annulled_flags("9110", "KSB1") == []
        assert {flag.transfer_name for flag in reopened.list_working_flags("9110", "KSB1")} == {
            "06_рев.01-AN03",
            "07_рев.01-AN03",
        }
        reopened.upsert_annulled_flag(
            "9110",
            "KSB1",
            "01-AN03",
            transfer_name="08_рев.01-AN03",
        )
        reopened.replace_kit_derived((), (), ())
        reopened.replace_google_snapshot(
            (),
            (),
            loaded_at="2026-09-04T14:00:00+00:00",
            source="test-empty-annulled",
        )
        assert {
            flag.transfer_name for flag in reopened.list_annulled_flags("9110", "KSB1")
        } == {"08_рев.01-AN03"}
        assert {flag.transfer_name for flag in reopened.list_working_flags("9110", "KSB1")} == {
            "06_рев.01-AN03",
            "07_рев.01-AN03",
        }
        assert len(reopened.list_file_mtime_overrides()) == 1
        assert reopened.get_issuance_review(issuance_review_id) is not None
        assert len(reopened.list_files()) == 1
        still = reopened.list_files()[0]
        assert still.data.get("mtime_override_applied") is True
        connection = sqlite3.connect(db_path)
        try:
            connection.execute(
                "UPDATE file_entry SET mtime_ns = mtime_ns + 1000"
            )
            connection.commit()
        finally:
            connection.close()
        stale = reopened.list_files()[0]
        assert stale.data.get("mtime_override_stale") is True
        assert not stale.data.get("mtime_override_applied")
        assert reopened.delete_file_mtime_override(stale.path_key)
        cleared = reopened.list_files()[0]
        assert not cleared.data.get("mtime_override_stale")
        assert not cleared.data.get("mtime_override_applied")
        an_file = AnMtoFile(
            path=r"C:\an\AGCC.287-1600-POS.MTO-0001_02-AN01_RU.xlsx",
            path_key="c:\\an\\agcc.287-1600-pos.mto-0001_02-an01_ru.xlsx",
            title="1600",
            mark="POS",
            revision_text="02-AN01",
            core_stem="AGCC.287-1600-POS.MTO-0001",
            discipline_block="MTO-0001",
            name="AGCC.287-1600-POS.MTO-0001_02-AN01_RU.xlsx",
            parent_dir=r"C:\an",
            mtime_ns=0,
            size=0,
        )
        reopened.replace_an_snapshot(
            (an_file,),
            scanned_at="2026-09-13T00:00:00+00:00",
        )
        assert len(reopened.list_issuance_reviews("9110", "KSB1")) == 1
        assert reopened.get_issuance_review(issuance_review_id) is not None
        assert len(reopened.list_liquidity_reviews("9110", "KSB1")) == 1
        assert len(reopened.list_files()) == 1
        stored_an = reopened.list_an_files_by_kit()
        assert stored_an[kit_identity_key("1600", "POS")][0].path == an_file.path
        reopened.replace_rd_dump_snapshot(
            (an_file,),
            scanned_at="2026-09-13T00:00:00+00:00",
        )
        stored_dump = reopened.list_rd_dump_files_by_kit()
        assert stored_dump[kit_identity_key("1600", "POS")][0].path == an_file.path
        reopened.replace_google_snapshot(
            (),
            (),
            loaded_at="2026-09-17T00:00:00+00:00",
            source="test-keep-rd-dump",
        )
        reopened.replace_kit_derived((), (), ())
        assert (
            reopened.list_rd_dump_mto_files()[0].path == an_file.path
        )
        assert reopened.list_an_mto_files()[0].path == an_file.path
        assert "an_mto_file" in _table_names(db_path)
        assert "rd_dump_mto_file" in _table_names(db_path)
    test_purge_noncanonical_rd_keeps_canonical_missing()
    print("RD catalog DB smoke: OK")


def _insert_file_entry(
    connection: sqlite3.Connection,
    *,
    file_id: int,
    path: str,
    source: str,
    run_id: int,
    present: int,
) -> None:
    connection.execute(
        """
        INSERT INTO file_entry(
            id, path, path_key, source, file_kind, name, size, mtime_ns,
            first_seen_run_id, last_seen_run_id, present, review_state,
            parse_status, transfer_is_as_build
        ) VALUES (
            ?, ?, ?, ?, 'pdf', 'file.pdf', 1, 1, ?, ?, ?, 'pending', 'parsed', 0
        )
        """,
        (file_id, path, path.casefold(), source, run_id, run_id, present),
    )


def test_purge_noncanonical_rd_keeps_canonical_missing() -> None:
    """Drop leftover non-canonical RD; keep missing canonical RD and SQ."""

    with tempfile.TemporaryDirectory(prefix="rd_catalog_purge_") as temp:
        root = Path(temp, "rd")
        transfer = root / "2225" / "KSB" / "Для передачи" / "01_рев.01_2225-KSB"
        transfer.mkdir(parents=True)
        document = transfer / "AGCC.287-2225-KSB.OD-0001_01_RU.pdf"
        document.write_bytes(b"fixture")
        db_path = Path(temp, "runtime", "rd_catalog.sqlite")
        database = CatalogDatabase(db_path)
        database.initialize()
        source_result = scan_document_source(
            root, SourceKind.RD, skip_dirs=("old", "tmp")
        )
        summary = ScanSummary(sources={SourceKind.RD: source_result})
        summary.rd_pdf_overlay, summary.rd_mto_overlay = build_rd_overlays(
            summary.files
        )
        database.store_scan(summary)
        document.unlink()
        missing = scan_document_source(
            root, SourceKind.RD, skip_dirs=("old", "tmp")
        )
        missing_summary = ScanSummary(sources={SourceKind.RD: missing})
        missing_summary.rd_pdf_overlay, missing_summary.rd_mto_overlay = (
            build_rd_overlays(missing.files)
        )
        database.store_scan(missing_summary)
        canonical = next(record for record in database.list_files())
        assert not canonical.present
        assert canonical.source is SourceKind.RD

        probe = sqlite3.connect(db_path)
        try:
            run_id = int(
                probe.execute(
                    "SELECT id FROM scan_run ORDER BY id DESC LIMIT 1"
                ).fetchone()[0]
            )
        finally:
            probe.close()
        connection = sqlite3.connect(db_path)
        try:
            loose = str(root / "2225" / "KSB" / "loose.pdf")
            sq_path = str(Path(temp, "sq", "AGCC.287-2225-KSB.OD-0002_01_RU.pdf")
            )
            _insert_file_entry(
                connection,
                file_id=9001,
                path=loose,
                source="rd",
                run_id=run_id,
                present=0,
            )
            _insert_file_entry(
                connection,
                file_id=9002,
                path=str(root / "2225" / "KSB" / "still_present.pdf"),
                source="rd",
                run_id=run_id,
                present=1,
            )
            _insert_file_entry(
                connection,
                file_id=9003,
                path=sq_path,
                source="sq",
                run_id=run_id,
                present=1,
            )
            connection.execute(
                """
                INSERT INTO review_event(file_id, action, comment, created_at)
                VALUES (9001, 'detected_missing', NULL, 't')
                """
            )
            connection.execute(
                """
                INSERT INTO mto_comparison(
                    rd_file_id, robot_file_id, rd_fingerprint, robot_fingerprint,
                    algorithm_version, status, stats_json, diff_json, compared_at
                ) VALUES (9002, NULL, 'rd', NULL, 1, 'not_compared', '{}', '{}', 't')
                """
            )
            connection.commit()
        finally:
            connection.close()

        deleted = database.purge_noncanonical_rd_files(root)
        assert deleted == 2
        remaining = {(record.source, record.present, record.path) for record in database.list_files()}
        assert (SourceKind.RD, False, canonical.path) in remaining
        assert (SourceKind.SQ, True, sq_path) in remaining
        assert all("loose.pdf" not in path for _source, _present, path in remaining)
        assert all("still_present.pdf" not in path for _source, _present, path in remaining)
        leftover_conn = sqlite3.connect(db_path)
        try:
            assert leftover_conn.execute(
                "SELECT COUNT(*) FROM review_event WHERE file_id IN (9001, 9002)"
            ).fetchone()[0] == 0
            assert leftover_conn.execute(
                "SELECT COUNT(*) FROM mto_comparison WHERE rd_file_id IN (9001, 9002)"
            ).fetchone()[0] == 0
            assert leftover_conn.execute(
                "SELECT value FROM schema_meta WHERE key = 'noncanonical_rd_clean'"
            ).fetchone()[0] == "1"
        finally:
            leftover_conn.close()

        assert database.purge_noncanonical_rd_files(root, skip_if_clean=True) == 0
        again = str(root / "2225" / "KSB" / "again.pdf")
        leftover_conn = sqlite3.connect(db_path)
        try:
            _insert_file_entry(
                leftover_conn,
                file_id=9004,
                path=again,
                source="rd",
                run_id=run_id,
                present=1,
            )
            leftover_conn.commit()
        finally:
            leftover_conn.close()
        assert database.purge_noncanonical_rd_files(root, skip_if_clean=True) == 0
        assert any("again.pdf" in record.path for record in database.list_files())

        robot_summary = ScanSummary(
            sources={
                SourceKind.ROBOT: SourceScanResult(
                    source=SourceKind.ROBOT,
                    root=str(Path(temp, "robot")),
                )
            }
        )
        database.store_scan(robot_summary)
        assert database.purge_noncanonical_rd_files(root, skip_if_clean=True) == 0
        assert any("again.pdf" in record.path for record in database.list_files())

        again_scan = scan_document_source(
            root, SourceKind.RD, skip_dirs=("old", "tmp")
        )
        again_summary = ScanSummary(sources={SourceKind.RD: again_scan})
        again_summary.rd_pdf_overlay, again_summary.rd_mto_overlay = (
            build_rd_overlays(again_scan.files)
        )
        database.store_scan(again_summary)
        assert database.purge_noncanonical_rd_files(root, skip_if_clean=True) == 1
        assert all("again.pdf" not in record.path for record in database.list_files())

    test_store_scan_marks_void_folder_annulled()


def test_store_scan_marks_void_folder_annulled() -> None:
    """RD scan upserts kit_annulled_flag for an NN folder named Void."""

    with tempfile.TemporaryDirectory(prefix="rd_catalog_void_scan_") as raw:
        root = Path(raw, "rd")
        transfer = (
            root
            / "7560"
            / "SKUD"
            / "Для передачи"
            / "04_рев.0-AN02_AGCC.287-7560-SKUD_Void"
        )
        transfer.mkdir(parents=True)
        document = transfer / "AGCC.287-7560-SKUD.MTO-0001_0-AN02_RU.xlsx"
        document.write_bytes(b"mto")
        source_result = scan_document_source(root, SourceKind.RD, skip_dirs=())
        summary = ScanSummary(sources={SourceKind.RD: source_result})
        summary.rd_pdf_overlay, summary.rd_mto_overlay = build_rd_overlays(
            summary.files
        )
        database = CatalogDatabase(Path(raw, "rd_catalog.sqlite"))
        database.initialize()
        database.store_scan(summary)
        flags = database.list_annulled_flags("7560", "SKUD")
        assert len(flags) == 1
        assert flags[0].transfer_name.endswith("_Void")
        assert flags[0].revision_text == "0-AN02"


if __name__ == "__main__":
    main()
