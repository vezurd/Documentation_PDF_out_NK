"""Local checks for RD catalog issuance-review fingerprints and rematch."""

from __future__ import annotations

import sys
import tempfile
from datetime import datetime
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from rd_catalog.db import CatalogDatabase
from rd_catalog.issuance_review import (
    IssuanceJournalRow,
    add_manual_journal_row,
    apply_journal_decision,
    effective_issuance_sends,
    journal_row_matches_issuance,
    latest_effective_issuance_kits,
    list_issuance_journal,
    orphan_identity_fingerprint,
    pick_journal_row_for_issuance,
    send_evidence_fingerprint,
    send_identity_fingerprint,
    sheet_send_decision,
    validate_issuance_decision,
)
from rd_catalog.kits import (
    GoogleKit,
    IssuanceKit,
    KitEvent,
    format_event_date_sortable,
    format_revision,
)
from rd_catalog.models import FileRecord, ReviewState, SourceKind
from rd_catalog.pipeline import ingest_google_snapshot


_LOADED_AT = "2026-09-13T09:00:00+00:00"


def _mtime_ns(year: int, month: int, day: int) -> int:
    return int(datetime(year, month, day, 12, 0, 0).timestamp() * 1_000_000_000)


def _record(
    file_id: int,
    *,
    title: str,
    mark: str,
    revision: str,
    appendix: str | None,
    sequence: int,
    folder: str,
    mtime_ns: int,
    source: SourceKind = SourceKind.RD,
    file_kind: str = "pdf",
    present: bool = True,
    as_build: bool = False,
) -> FileRecord:
    rev_text = format_revision(revision, appendix)
    suffix = "xlsx" if file_kind == "mto_xlsx" else "pdf"
    discipline = "MTO-0001" if file_kind == "mto_xlsx" else "OD-0001"
    filename = f"AGCC.287-{title}-{mark}.{discipline}_{rev_text}_RU.{suffix}"
    path = (
        rf"\\bcc\eng\PrDoc\РД\{title}\06_{mark}\Для передачи\{folder}"
        rf"\PDF\{filename}"
    )
    return FileRecord(
        id=file_id,
        path=path,
        path_key=f"{source.value}/{title}-{mark}/{file_id}",
        source=source,
        present=present,
        review_state=ReviewState.ACKNOWLEDGED,
        first_seen_run_id=1,
        last_seen_run_id=1,
        data={
            "file_kind": file_kind,
            "name": filename,
            "parse_status": "parsed",
            "title": title,
            "mark": mark,
            "revision": revision,
            "appendix": appendix,
            "transfer_revision": revision,
            "transfer_appendix": appendix,
            "mtime_ns": mtime_ns,
            "transfer_is_as_build": 1 if as_build else 0,
            "transfer_sequence": sequence,
            "transfer_name": folder,
            "core_stem": f"agcc.287-{title}-{mark}.{file_id}",
            "discipline_block": discipline,
            "title_system": f"{title}-{mark}",
        },
    )


def _send(
    *,
    revision: str,
    appendix: str | None = None,
    send_date: str,
    transmittal: str = "",
    row_index: int = 1,
    note_raw: str = "",
    status: str = "Отправлен",
    incoming: str = "",
    confirm: str = "",
) -> IssuanceKit:
    return IssuanceKit(
        title="1600",
        mark="SOS",
        mark_raw="SOS",
        title_system="1600-SOS",
        revision=revision,
        appendix=appendix,
        revision_text=format_revision(revision, appendix),
        status=status,
        send_date=send_date,
        send_date_sortable=format_event_date_sortable(send_date),
        send_transmittal=transmittal,
        incoming_control_date=incoming,
        incoming_control_date_sortable=format_event_date_sortable(incoming),
        confirm_transmittal=confirm,
        note_raw=note_raw,
        row_index=row_index,
    )


def _open_db(temp: Path) -> CatalogDatabase:
    database = CatalogDatabase(temp / "rd_catalog.sqlite")
    database.initialize()
    return database


def _ingest(database: CatalogDatabase, sends: tuple[IssuanceKit, ...]) -> None:
    ingest_google_snapshot(
        database,
        (),
        sends,
        loaded_at=_LOADED_AT,
        source="test",
    )


def _annul_send(database: CatalogDatabase, send: IssuanceKit, comment: str) -> int:
    validate_issuance_decision(
        decision="annulled",
        revision_text=send.revision_text,
        comment=comment,
    )
    return database.upsert_issuance_review(
        send.title,
        send.mark,
        "send",
        "issuance",
        send_identity_fingerprint(send),
        decision="annulled",
        comment=comment,
        revision_text=send.revision_text,
        send_date=send.send_date,
        send_date_sortable=send.send_date_sortable,
        send_transmittal=send.send_transmittal,
        incoming_control_date=send.incoming_control_date,
        incoming_control_date_sortable=send.incoming_control_date_sortable,
        confirm_transmittal=send.confirm_transmittal,
        sheet_status=send.status,
        note=send.note_raw,
        evidence_fingerprint=send_evidence_fingerprint(send),
        match_state="matched",
    )


def test_annul_03_latest_is_02(temp: Path) -> None:
    database = _open_db(temp / "annul")
    send_02 = _send(
        revision="02",
        send_date="01.08.2026",
        transmittal="AGCC-BCC-TRM-000010",
        row_index=10,
    )
    send_03 = _send(
        revision="03",
        send_date="15.08.2026",
        transmittal="AGCC-BCC-TRM-000011",
        row_index=11,
    )
    _ingest(database, (send_02, send_03))
    _annul_send(database, send_03, "03 снята, действует 02")
    raw = database.list_issuance_sends("1600", "SOS")
    assert {item.revision_text for item in raw} == {"02", "03"}
    latest = latest_effective_issuance_kits(database, "1600", "SOS")
    assert len(latest) == 1
    assert latest[0].revision_text == "02"
    effective = effective_issuance_sends(database, "1600", "SOS")
    assert [item[1].revision_text for item in effective] == ["02"]


def test_rematch_after_note_and_row_index(temp: Path) -> None:
    database = _open_db(temp / "rematch_note")
    send_02 = _send(
        revision="02",
        send_date="01.08.2026",
        transmittal="AGCC-BCC-TRM-000010",
        row_index=10,
    )
    send_03 = _send(
        revision="03",
        send_date="15.08.2026",
        transmittal="AGCC-BCC-TRM-000011",
        row_index=11,
        note_raw="old note",
    )
    _ingest(database, (send_02, send_03))
    review_id = _annul_send(database, send_03, "03 снята")
    refreshed = _send(
        revision="03",
        send_date="15.08.2026",
        transmittal="AGCC-BCC-TRM-000011",
        row_index=44,
        note_raw="sheet note edited",
    )
    _ingest(database, (send_02, refreshed))
    row = database.get_issuance_review(review_id)
    assert row is not None
    assert row.decision == "annulled"
    assert row.comment == "03 снята"
    assert row.match_state == "matched"
    assert row.identity_fingerprint == send_identity_fingerprint(refreshed)
    assert row.evidence_fingerprint == send_evidence_fingerprint(refreshed)
    assert row.note == "sheet note edited"
    latest = latest_effective_issuance_kits(database, "1600", "SOS")
    assert latest[0].revision_text == "02"


def test_ladder_fills_trm(temp: Path) -> None:
    database = _open_db(temp / "ladder_trm")
    send_03 = _send(
        revision="03",
        send_date="15.08.2026",
        transmittal="",
        row_index=11,
    )
    _ingest(database, (send_03,))
    review_id = _annul_send(database, send_03, "03 снята")
    filled = _send(
        revision="03",
        send_date="15.08.2026",
        transmittal="AGCC-BCC-TRM-000099",
        row_index=12,
    )
    _ingest(database, (filled,))
    row = database.get_issuance_review(review_id)
    assert row is not None
    assert row.decision == "annulled"
    assert row.match_state == "matched"
    assert row.identity_fingerprint == send_identity_fingerprint(filled)
    assert row.send_transmittal == "AGCC-BCC-TRM-000099"
    assert latest_effective_issuance_kits(database, "1600", "SOS") == ()


def test_ambiguous_same_revision(temp: Path) -> None:
    database = _open_db(temp / "ambiguous")
    send_02 = _send(
        revision="02",
        send_date="01.08.2026",
        transmittal="AGCC-BCC-TRM-000010",
        row_index=10,
    )
    original = _send(
        revision="03",
        send_date="15.08.2026",
        transmittal="AGCC-BCC-TRM-OLD",
        row_index=11,
    )
    _ingest(database, (send_02, original))
    review_id = _annul_send(database, original, "03 снята")
    first = _send(
        revision="03",
        send_date="15.08.2026",
        transmittal="AGCC-BCC-TRM-000021",
        row_index=21,
    )
    second = _send(
        revision="03",
        send_date="15.08.2026",
        transmittal="AGCC-BCC-TRM-000022",
        row_index=22,
    )
    _ingest(database, (send_02, first, second))
    row = database.get_issuance_review(review_id)
    assert row is not None
    assert row.decision == "annulled"
    assert row.match_state == "ambiguous"
    assert row.identity_fingerprint == send_identity_fingerprint(original)
    latest = latest_effective_issuance_kits(database, "1600", "SOS")
    assert latest[0].revision_text == "03"
    assert latest[0].send_transmittal == "AGCC-BCC-TRM-000022"
    effective_revs = [item[1].revision_text for item in effective_issuance_sends(database, "1600", "SOS")]
    assert effective_revs == ["02", "03", "03"]


def test_legalize_orphan_does_not_beat_dated_02(temp: Path) -> None:
    database = _open_db(temp / "legalize")
    send_02 = _send(
        revision="02",
        send_date="01.08.2026",
        transmittal="AGCC-BCC-TRM-000010",
        row_index=10,
    )
    send_03 = _send(
        revision="03",
        send_date="15.08.2026",
        transmittal="AGCC-BCC-TRM-000011",
        row_index=11,
    )
    _ingest(database, (send_02, send_03))
    _annul_send(database, send_03, "03 снята")
    validate_issuance_decision(
        decision="legalized",
        revision_text="02-AN01",
        comment="",
    )
    database.upsert_issuance_review(
        "1600",
        "SOS",
        "orphan",
        "robot",
        orphan_identity_fingerprint("robot", "1600", "SOS", "02-AN01"),
        decision="legalized",
        revision_text="02-AN01",
    )
    effective = effective_issuance_sends(database, "1600", "SOS")
    texts = [item[1].revision_text for item in effective]
    assert "02" in texts
    assert "02-AN01" in texts
    assert all(send_id is None for send_id, send in effective if send.revision_text == "02-AN01")
    latest = latest_effective_issuance_kits(database, "1600", "SOS")
    assert latest[0].revision_text == "02"
    assert latest[0].send_date == "01.08.2026"


def test_validate_issuance_decision_rules() -> None:
    try:
        validate_issuance_decision(
            decision="legalized",
            revision_text="",
            comment="",
        )
    except ValueError:
        pass
    else:
        raise AssertionError("legalize without revision must raise")
    try:
        validate_issuance_decision(
            decision="annulled",
            revision_text="03",
            comment="",
        )
    except ValueError:
        pass
    else:
        raise AssertionError("annul without comment must raise")
    validate_issuance_decision(decision="active", revision_text="", comment="")
    validate_issuance_decision(decision="", revision_text="", comment="")


def test_journal_collapses_same_revision_across_files(temp: Path) -> None:
    database = _open_db(temp / "journal_collapse")
    kit = GoogleKit(
        title="1711",
        mark="POS",
        mark_raw="POS",
        title_system="1711-POS",
        sheet_revision="02",
        sheet_appendix=None,
        sheet_revision_text="02",
        status_sheet="",
        comment_raw="11.07.2024 код A рев.02",
        events=(
            KitEvent(
                raw="11.07.2024 код A рев.02",
                date="11.07.2024",
                stage="code_a",
                stage_label="код A",
                revision="02",
                appendix=None,
                transmittals=(),
                parsed=True,
            ),
        ),
        last_event=None,  # unused by journal; events tuple is the source
        row_index=1,
    )
    ingest_google_snapshot(
        database,
        (kit,),
        (),
        loaded_at=_LOADED_AT,
        source="test",
    )
    records = (
        _record(
            1,
            title="1711",
            mark="POS",
            revision="0",
            appendix=None,
            sequence=1,
            folder="01_рев.0_AGCC.287-1711-POS",
            mtime_ns=_mtime_ns(2023, 1, 1),
        ),
        _record(
            2,
            title="1711",
            mark="POS",
            revision="0",
            appendix=None,
            sequence=1,
            folder="01_рев.0_AGCC.287-1711-POS",
            mtime_ns=_mtime_ns(2023, 1, 2),
            file_kind="mto_xlsx",
        ),
        _record(
            3,
            title="1711",
            mark="POS",
            revision="02",
            appendix=None,
            sequence=3,
            folder="03_рев.02_AGCC.287-1711-POS",
            mtime_ns=_mtime_ns(2024, 7, 11),
        ),
        _record(
            4,
            title="1711",
            mark="POS",
            revision="02",
            appendix=None,
            sequence=3,
            folder="03_рев.02_AGCC.287-1711-POS",
            mtime_ns=_mtime_ns(2024, 7, 12),
            file_kind="mto_xlsx",
        ),
        _record(
            5,
            title="1711",
            mark="POS",
            revision="02",
            appendix=None,
            sequence=1,
            folder="robot",
            mtime_ns=_mtime_ns(2024, 7, 13),
            source=SourceKind.ROBOT,
            file_kind="mto_xlsx",
        ),
    )
    rows = [
        row
        for row in list_issuance_journal(database, records=records)
        if row.title == "1711" and row.mark == "POS"
    ]
    orphans = [row for row in rows if row.kind == "orphan"]
    by_rev = {row.revision_text: row for row in orphans}
    assert set(by_rev) == {"0", "02"}
    assert by_rev["0"].in_rd is True
    assert by_rev["0"].in_f is False
    assert by_rev["0"].source == "rd"
    assert by_rev["02"].in_f is True
    assert by_rev["02"].in_rd is True
    assert by_rev["02"].in_robot is True
    assert by_rev["02"].source == "rd"


def test_journal_includes_sheet_and_robot_orphan(temp: Path) -> None:
    database = _open_db(temp / "journal")
    send_02 = _send(
        revision="02",
        send_date="01.08.2026",
        transmittal="AGCC-BCC-TRM-000010",
        row_index=10,
    )
    send_03 = _send(
        revision="03",
        send_date="15.08.2026",
        transmittal="AGCC-BCC-TRM-000011",
        row_index=11,
    )
    _ingest(database, (send_02, send_03))
    robot = _record(
        7,
        title="1600",
        mark="SOS",
        revision="02",
        appendix="01",
        sequence=1,
        folder="01_рев.02-AN01_AGCC.287-1600-SOS",
        mtime_ns=_mtime_ns(2026, 8, 20),
        source=SourceKind.ROBOT,
        file_kind="mto_xlsx",
    )
    rows = list_issuance_journal(database, records=(robot,))
    sheet = [row for row in rows if row.kind == "send"]
    assert {row.revision_text for row in sheet} == {"02", "03"}
    orphans = [
        row
        for row in rows
        if row.kind == "orphan" and row.source == "robot"
    ]
    assert len(orphans) == 1
    assert orphans[0].revision_text == "02-AN01"
    assert orphans[0].decision == ""
    assert orphans[0].in_robot is True
    assert orphans[0].issuance_send_id is None


def test_apply_journal_decision_annuls_sheet(temp: Path) -> None:
    database = _open_db(temp / "apply_decision")
    send_02 = _send(
        revision="02",
        send_date="01.08.2026",
        transmittal="AGCC-BCC-TRM-000010",
        row_index=10,
    )
    send_03 = _send(
        revision="03",
        send_date="15.08.2026",
        transmittal="AGCC-BCC-TRM-000011",
        row_index=11,
    )
    _ingest(database, (send_02, send_03))
    rows = list_issuance_journal(database)
    sheet_03 = next(row for row in rows if row.revision_text == "03")
    assert sheet_03.review_id is None
    review_id = apply_journal_decision(
        database,
        sheet_03,
        decision="annulled",
        comment="03 снята, действует 02",
    )
    stored = database.get_issuance_review(review_id)
    assert stored is not None
    assert stored.kind == "send"
    assert stored.source == "issuance"
    assert stored.match_state == "matched"
    assert stored.identity_fingerprint == send_identity_fingerprint(send_03)
    latest = latest_effective_issuance_kits(database, "1600", "SOS")
    assert latest[0].revision_text == "02"


def test_note_annulled_is_implicit_and_skips_latest(temp: Path) -> None:
    database = _open_db(temp / "note_annul")
    send_02 = _send(
        revision="02",
        send_date="01.08.2026",
        transmittal="AGCC-BCC-TRM-000010",
        row_index=10,
    )
    send_03 = _send(
        revision="03",
        send_date="15.08.2026",
        transmittal="AGCC-BCC-TRM-000011",
        row_index=11,
        note_raw="TRM аннулирован",
    )
    _ingest(database, (send_02, send_03))
    rows = list_issuance_journal(database)
    sheet_02 = next(row for row in rows if row.revision_text == "02")
    sheet_03 = next(row for row in rows if row.revision_text == "03")
    assert sheet_02.decision == "active"
    assert sheet_03.decision == "annulled"
    assert sheet_03.review_id is None
    assert sheet_send_decision(send_03, None) == "annulled"
    latest = latest_effective_issuance_kits(database, "1600", "SOS")
    assert latest[0].revision_text == "02"
    review_id = apply_journal_decision(database, sheet_03, decision="annulled")
    stored = database.get_issuance_review(review_id)
    assert stored is not None
    assert stored.decision == "annulled"
    assert stored.comment == "TRM аннулирован"


def test_user_active_overrides_note_annulled(temp: Path) -> None:
    database = _open_db(temp / "note_override")
    send_03 = _send(
        revision="03",
        send_date="15.08.2026",
        transmittal="AGCC-BCC-TRM-000011",
        row_index=11,
        note_raw="ТРМ отменен",
    )
    _ingest(database, (send_03,))
    rows = list_issuance_journal(database)
    sheet_03 = next(row for row in rows if row.revision_text == "03")
    assert sheet_03.decision == "annulled"
    apply_journal_decision(database, sheet_03, decision="active")
    latest = latest_effective_issuance_kits(database, "1600", "SOS")
    assert latest[0].revision_text == "03"
    rows = list_issuance_journal(database)
    sheet_03 = next(row for row in rows if row.revision_text == "03")
    assert sheet_03.decision == "active"
    assert sheet_03.review_id is not None


def test_add_manual_journal_row_draft_and_legalize(temp: Path) -> None:
    database = _open_db(temp / "manual_row")
    draft_id = add_manual_journal_row(database, "1600", "SOS")
    draft = database.get_issuance_review(draft_id)
    assert draft is not None
    assert draft.kind == "manual"
    assert draft.source == "manual"
    assert draft.decision == ""
    assert draft.revision_text == ""
    assert draft.match_state == ""
    legal_id = add_manual_journal_row(
        database,
        "1600",
        "SOS",
        revision_text="02-AN01",
        decision="legalized",
    )
    legal = database.get_issuance_review(legal_id)
    assert legal is not None
    assert legal.decision == "legalized"
    assert legal.identity_fingerprint == orphan_identity_fingerprint(
        "manual", "1600", "SOS", "02-AN01", ""
    )
    try:
        add_manual_journal_row(
            database,
            "1600",
            "POS",
            decision="legalized",
        )
    except ValueError:
        pass
    else:
        raise AssertionError("legalize without revision must raise")


def test_pick_journal_row_for_effective_issuance(temp: Path) -> None:
    database = _open_db(temp / "pick_journal")
    send_01 = _send(
        revision="01",
        appendix="01",
        send_date="22.12.2025",
        transmittal="AGCC-BCC-TRM-000001",
        row_index=10,
    )
    send_02 = _send(
        revision="01",
        appendix="02",
        send_date="15.09.2026",
        transmittal="AGCC-BCC-TRM-000002",
        row_index=20,
    )
    _ingest(database, (send_01, send_02))
    latest = latest_effective_issuance_kits(database, "1600", "SOS")
    assert latest[0].revision_text == "01-AN02"
    rows = list_issuance_journal(database)
    picked = pick_journal_row_for_issuance(
        rows, title="1600", mark="SOS", issuance=latest[0]
    )
    assert picked is not None
    assert picked.revision_text == "01-AN02"
    assert journal_row_matches_issuance(picked, latest[0])
    older = next(row for row in rows if row.revision_text == "01-AN01")
    assert not journal_row_matches_issuance(older, latest[0])
    orphan = IssuanceJournalRow(
        title="1600",
        mark="SOS",
        kind="manual",
        source="manual",
        revision_text="01-AN02",
        send_date=send_02.send_date,
        send_date_sortable=send_02.send_date_sortable,
        send_transmittal=send_02.send_transmittal,
        incoming_control_date="",
        incoming_control_date_sortable="",
        confirm_transmittal="",
        sheet_status="",
        note="",
        decision="legalized",
        comment=None,
        match_state="",
        identity_fingerprint="not-the-send-fingerprint",
        evidence_fingerprint="",
        source_path="",
        path_key="",
        in_f=False,
        in_rd=False,
        in_robot=False,
        in_auto_mto=False,
        review_id=9,
        issuance_send_id=None,
    )
    assert journal_row_matches_issuance(orphan, send_02)


def main() -> None:
    """Run issuance-review fixtures against a temporary SQLite file."""

    test_validate_issuance_decision_rules()
    with tempfile.TemporaryDirectory(prefix="rd_catalog_issuance_review_") as raw:
        root = Path(raw)
        test_annul_03_latest_is_02(root)
        test_rematch_after_note_and_row_index(root)
        test_ladder_fills_trm(root)
        test_ambiguous_same_revision(root)
        test_legalize_orphan_does_not_beat_dated_02(root)
        test_journal_collapses_same_revision_across_files(root)
        test_journal_includes_sheet_and_robot_orphan(root)
        test_apply_journal_decision_annuls_sheet(root)
        test_note_annulled_is_implicit_and_skips_latest(root)
        test_user_active_overrides_note_annulled(root)
        test_add_manual_journal_row_draft_and_legalize(root)
        test_pick_journal_row_for_effective_issuance(root)
    print("RD catalog issuance review: OK")


if __name__ == "__main__":
    main()
