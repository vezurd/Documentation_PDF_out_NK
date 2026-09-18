"""Local checks for RD catalog pipeline matching, status, and liquidity."""

from __future__ import annotations

import json
import sys
import tempfile
from dataclasses import replace
from datetime import datetime
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from rd_catalog.db import CatalogDatabase, KitPackageRow, KitPipelineRow
from rd_catalog.issuance_review import (
    add_manual_journal_row,
    effective_issuance_sends,
    latest_effective_issuance_kits,
    orphan_identity_fingerprint,
    send_evidence_fingerprint,
    send_identity_fingerprint,
    validate_issuance_decision,
)
from rd_catalog.kits import (
    GoogleKit,
    IssuanceKit,
    KitEvent,
    KitSummary,
    build_kit_matrix,
    format_event_date_sortable,
    format_revision,
    kit_identity_key,
)
from rd_catalog.models import CollisionKind, FileRecord, ReviewState, SourceKind
from rd_catalog.monitor_views import official_rd_rev_text
from rd_catalog.parse import issued_package_dir
from rd_catalog.pipeline import (
    PIPELINE_STATUS_ALGORITHM_VERSION,
    KitPipelineStatus,
    get_kit_card,
    index_records_by_kit,
    ingest_google_snapshot,
    iter_mto_files_for_cells,
    kit_keys_under_folders,
    list_folder_tree_hints,
    list_kit_pipelines,
    list_mto_worklist,
    list_revision_matrix,
    official_detected_current_ids,
    patch_official_detected_current_ids,
    pick_official_rd_package,
    pipeline_algorithm_needs_rebuild,
    APPROVAL_REL_AHEAD,
    APPROVAL_REL_NO_REV,
    APPROVAL_REL_OTHER,
    APPROVAL_REL_PREVIOUS,
    PIPELINE_DISPLAY_V1,
    PIPELINE_DISPLAY_V2,
    pipeline_approval_label,
    pipeline_approval_relation,
    pipeline_review_label,
    pipeline_send_date_text,
    pipeline_tdo_passed_date_text,
    pipeline_status_label,
    rebuild_pipeline,
    records_in_contour,
    working_revision_for_display,
)


_LOADED_AT = "2026-09-04T12:00:00+00:00"


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


def _event(
    *,
    date: str,
    stage: str,
    stage_label: str,
    revision: str | None = None,
    appendix: str | None = None,
    transmittals: tuple[str, ...] = (),
    raw: str | None = None,
) -> KitEvent:
    rev = format_revision(revision, appendix)
    default = f"{date} {stage_label}"
    if rev:
        default += f" рев.{rev}"
    if transmittals:
        default += f" {transmittals[-1]}"
    return KitEvent(
        raw=raw or default,
        date=date,
        stage=stage,
        stage_label=stage_label,
        revision=revision,
        appendix=appendix,
        transmittals=transmittals,
        parsed=True,
    )


def _google(
    title: str,
    mark: str,
    events: tuple[KitEvent, ...],
    *,
    revision: str | None = "0",
    appendix: str | None = "02",
    row_index: int = 1,
    status_sheet: str = "",
) -> GoogleKit:
    return GoogleKit(
        title=title,
        mark=mark,
        mark_raw=mark,
        title_system=f"{title}-{mark}",
        sheet_revision=revision,
        sheet_appendix=appendix,
        sheet_revision_text=format_revision(revision, appendix),
        status_sheet=status_sheet,
        comment_raw="\n".join(event.raw for event in events),
        events=events,
        last_event=events[-1] if events else None,
        row_index=row_index,
    )


def _send(
    *,
    title: str,
    mark: str,
    revision: str | None,
    appendix: str | None,
    status: str,
    send_date: str,
    incoming: str = "",
    transmittal: str = "",
    row_index: int = 1,
    note_raw: str = "",
) -> IssuanceKit:
    return IssuanceKit(
        title=title,
        mark=mark,
        mark_raw=mark,
        title_system=f"{title}-{mark}",
        revision=revision,
        appendix=appendix,
        revision_text=format_revision(revision, appendix),
        status=status,
        send_date=send_date,
        send_date_sortable=format_event_date_sortable(send_date),
        send_transmittal=transmittal,
        incoming_control_date=incoming,
        incoming_control_date_sortable=format_event_date_sortable(incoming),
        confirm_transmittal="",
        note_raw=note_raw,
        row_index=row_index,
    )


def _open_db(temp: Path) -> CatalogDatabase:
    database = CatalogDatabase(temp / "rd_catalog.sqlite")
    database.initialize()
    return database


def _rebuild(
    database: CatalogDatabase,
    kits: tuple[GoogleKit, ...],
    sends: tuple[IssuanceKit, ...],
    records: list[FileRecord],
    current_ids: set[int],
) -> None:
    ingest_google_snapshot(
        database,
        kits,
        sends,
        loaded_at=_LOADED_AT,
        source="test",
    )
    rebuild_pipeline(
        database,
        records=records,
        detected_current_ids=current_ids,
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


def _pipeline(database: CatalogDatabase, title: str, mark: str):
    rows = database.list_kit_pipelines(title, mark)
    assert rows, f"missing pipeline for {title}-{mark}"
    return rows[0]


def test_9110_tdo_review_package_path(temp: Path) -> None:
    database = _open_db(temp / "c1")
    folder = "05_рев.0-AN02_AGCC.287-9110-KSB1"
    record = _record(
        1,
        title="9110",
        mark="KSB1",
        revision="0",
        appendix="02",
        sequence=5,
        folder=folder,
        mtime_ns=_mtime_ns(2026, 8, 5),
    )
    event = _event(
        date="05.08.2026",
        stage="tdo_passed",
        stage_label="прошла ТДО",
        revision="0",
        appendix="02",
        transmittals=("AGCC.287-PGS-PGS-TRM-20541",),
    )
    send = _send(
        title="9110",
        mark="KSB1",
        revision="0",
        appendix="02",
        status="Принят",
        send_date="01.08.2026",
        incoming="04.08.2026",
        transmittal="AGCC.287-PGS-PGS-TRM-20541",
    )
    _rebuild(
        database,
        (_google("9110", "KSB1", (event,)),),
        (send,),
        [record],
        {1},
    )
    row = _pipeline(database, "9110", "KSB1")
    assert row.status == KitPipelineStatus.TDO_REVIEW
    assert row.code is None
    assert row.algorithm_version == PIPELINE_STATUS_ALGORITHM_VERSION
    assert pipeline_status_label(row.status) == "Прошел ТДО"
    assert pipeline_review_label(row) == "Прошел ТДО (05.08.2026) · РД 0-AN02"
    assert pipeline_review_label(row, issuance=send) == (
        "Прошел ТДО (05.08.2026) · РД 0-AN02 · отпр. 01.08.2026"
    )
    assert row.tdo_date == "05.08.2026"
    packages = database.list_kit_packages("9110", "KSB1")
    assert len(packages) == 1
    package = packages[0]
    assert package.is_grey is False
    assert r"\PDF" not in package.package_path.upper()
    assert package.package_path.endswith(folder)
    card = get_kit_card(database, "9110", "KSB1")
    assert card is not None
    assert card.pipeline is not None
    assert card.google is not None
    assert len(card.cycles) == 1
    assert card.cycles[0].send_id is not None
    assert card.cycles[0].tdo_event_id is not None
    send_id, stored = database.list_issuance_sends_with_ids("9110", "KSB1")[0]
    event_id, stored_event = database.list_google_events_with_ids("9110", "KSB1")[0]
    assert card.cycles[0].send_id == send_id
    assert card.cycles[0].tdo_event_id == event_id
    assert stored_event.stage == "tdo_passed"
    assert stored.send_transmittal.endswith("TRM-20541")
    pipelines = list_kit_pipelines(database)
    assert any(item.title == "9110" and item.mark == "KSB1" for item in pipelines)
    hints = list_folder_tree_hints(database)
    hint = hints[(*kit_identity_key("9110", "KSB1"), folder.casefold())]
    assert hint.review_status == "ТДО"
    assert hint.review_full == "Прошел ТДО"
    assert hint.match_reason == "trm"
    assert "ТДО" in hint.f_label
    assert hint.is_working is False


def test_two_sends_one_package_two_cycles(temp: Path) -> None:
    database = _open_db(temp / "c2")
    folder = "03_рев.0-AN01_AGCC.287-1513-POS"
    record = _record(
        10,
        title="1513",
        mark="POS",
        revision="0",
        appendix="01",
        sequence=3,
        folder=folder,
        mtime_ns=_mtime_ns(2026, 3, 1),
    )
    events = (
        _event(
            date="02.03.2026",
            stage="tdo_sent",
            stage_label="отпр. на ТДО",
            revision="0",
            appendix="01",
            transmittals=("AGCC-BCC-TRM-000010",),
        ),
        _event(
            date="12.03.2026",
            stage="tdo_sent",
            stage_label="отпр. на ТДО",
            revision="0",
            appendix="01",
            transmittals=("AGCC-BCC-TRM-000011",),
        ),
    )
    sends = (
        _send(
            title="1513",
            mark="POS",
            revision="0",
            appendix="01",
            status="На рассмотрении вх.контроля",
            send_date="01.03.2026",
            transmittal="AGCC-BCC-TRM-000010",
            row_index=10,
        ),
        _send(
            title="1513",
            mark="POS",
            revision="0",
            appendix="01",
            status="На рассмотрении вх.контроля",
            send_date="11.03.2026",
            transmittal="AGCC-BCC-TRM-000011",
            row_index=11,
        ),
    )
    _rebuild(
        database,
        (_google("1513", "POS", events, revision="0", appendix="01"),),
        sends,
        [record],
        {10},
    )
    packages = [pkg for pkg in database.list_kit_packages("1513", "POS") if not pkg.is_grey]
    assert len(packages) == 1
    cycles = database.list_kit_cycles("1513", "POS")
    assert len(cycles) == 2
    assert {cycle.package_id for cycle in cycles} == {packages[0].id}
    assert {cycle.send_id for cycle in cycles} == {
        send_id for send_id, _send in database.list_issuance_sends_with_ids("1513", "POS")
    }
    assert {cycle.match_reason for cycle in cycles} == {"trm"}


def test_grey_issuance_without_rd(temp: Path) -> None:
    database = _open_db(temp / "c3")
    folder = "02_рев.01_AGCC.287-1715-SOT"
    record = _record(
        20,
        title="1715",
        mark="SOT",
        revision="01",
        appendix=None,
        sequence=2,
        folder=folder,
        mtime_ns=_mtime_ns(2026, 2, 1),
    )
    send = _send(
        title="1715",
        mark="SOT",
        revision="03",
        appendix=None,
        status="На рассмотрении вх.контроля",
        send_date="01.04.2026",
        transmittal="AGCC-BCC-TRM-000030",
    )
    _rebuild(
        database,
        (_google("1715", "SOT", (), revision="03", appendix=None),),
        (send,),
        [record],
        {20},
    )
    packages = database.list_kit_packages("1715", "SOT")
    grey = [pkg for pkg in packages if pkg.is_grey]
    rd = [pkg for pkg in packages if pkg.source == "rd"]
    assert len(rd) == 1
    assert len(grey) == 1
    assert grey[0].source == "issuance_grey"
    assert grey[0].revision_text == "03"
    assert grey[0].package_path == ""
    cycles = database.list_kit_cycles("1715", "SOT")
    assert len(cycles) == 1
    assert cycles[0].package_id == grey[0].id


def test_code_a_customer(temp: Path) -> None:
    database = _open_db(temp / "c4")
    folder = "04_рев.01-AN02_AGCC.287-2225-KSB"
    record = _record(
        30,
        title="2225",
        mark="KSB",
        revision="01",
        appendix="02",
        sequence=4,
        folder=folder,
        mtime_ns=_mtime_ns(2026, 5, 20),
    )
    events = (
        _event(
            date="11.04.2026",
            stage="tdo_sent",
            stage_label="отпр. на ТДО",
            revision="01",
            appendix="02",
            transmittals=("AGCC-BCC-TRM-000040",),
        ),
        _event(
            date="20.05.2026",
            stage="code_a",
            stage_label="код А",
            revision="01",
            appendix="02",
            transmittals=("AGCC-BCC-TRM-000041",),
        ),
    )
    send = _send(
        title="2225",
        mark="KSB",
        revision="01",
        appendix="02",
        status="Принят",
        send_date="10.04.2026",
        incoming="15.04.2026",
        transmittal="AGCC-BCC-TRM-000040",
    )
    _rebuild(
        database,
        (_google("2225", "KSB", events, revision="01", appendix="02"),),
        (send,),
        [record],
        {30},
    )
    row = _pipeline(database, "2225", "KSB")
    assert row.status == KitPipelineStatus.AGREED
    assert row.code == "A"
    assert row.code_stale is False
    assert row.code_origin == "customer"
    assert pipeline_status_label(row.status) == "Согласован"
    assert pipeline_approval_label(row, version=PIPELINE_DISPLAY_V1) == (
        "A · 01-AN02 · 20.05.2026"
    )
    assert pipeline_approval_label(row) == "—"


def test_code_b_origin_pi_and_customer(temp: Path) -> None:
    database = _open_db(temp / "c5")
    folder_pi = "06_рев.01_AGCC.287-1600-POS"
    folder_cu = "06_рев.01_AGCC.287-1600-KSB"
    records = [
        _record(
            40,
            title="1600",
            mark="POS",
            revision="01",
            appendix=None,
            sequence=6,
            folder=folder_pi,
            mtime_ns=_mtime_ns(2026, 6, 1),
        ),
        _record(
            41,
            title="1600",
            mark="KSB",
            revision="01",
            appendix=None,
            sequence=6,
            folder=folder_cu,
            mtime_ns=_mtime_ns(2026, 6, 1),
        ),
    ]
    pi_events = (
        _event(
            date="10.01.2026",
            stage="code_b",
            stage_label="код B",
            revision="01",
            appendix=None,
            transmittals=("AGCC-BCC-TRM-000050",),
        ),
    )
    cu_events = (
        _event(
            date="10.01.2026",
            stage="code_b",
            stage_label="код B",
            revision="01",
            appendix=None,
            transmittals=("AGCC-BCC-TRM-000060",),
        ),
    )
    sends = (
        _send(
            title="1600",
            mark="POS",
            revision="01",
            appendix=None,
            status="Принят",
            send_date="01.01.2026",
            incoming="05.01.2026",
            transmittal="AGCC-BCC-TRM-000050",
            row_index=1,
        ),
        _send(
            title="1600",
            mark="POS",
            revision="01",
            appendix=None,
            status="На рассмотрении вх.контроля",
            send_date="20.01.2026",
            transmittal="AGCC-BCC-TRM-000051",
            row_index=2,
        ),
        _send(
            title="1600",
            mark="KSB",
            revision="01",
            appendix=None,
            status="Принят",
            send_date="01.01.2026",
            incoming="05.01.2026",
            transmittal="AGCC-BCC-TRM-000060",
            row_index=3,
        ),
        _send(
            title="1600",
            mark="KSB",
            revision="02",
            appendix=None,
            status="На рассмотрении вх.контроля",
            send_date="20.01.2026",
            transmittal="AGCC-BCC-TRM-000061",
            row_index=4,
        ),
    )
    _rebuild(
        database,
        (
            _google("1600", "POS", pi_events, revision="01", appendix=None),
            _google("1600", "KSB", cu_events, revision="01", appendix=None),
        ),
        sends,
        records,
        {40, 41},
    )
    pi_row = _pipeline(database, "1600", "POS")
    assert pi_row.status == KitPipelineStatus.SENT_TDO
    assert pi_row.code == "B"
    assert pi_row.code_stale is True
    assert pi_row.code_origin == "pi"
    cu_row = _pipeline(database, "1600", "KSB")
    assert cu_row.status == KitPipelineStatus.TDO_REVIEW
    assert cu_row.code == "B"
    assert cu_row.code_stale is False
    assert cu_row.code_origin == "customer"


def test_liquidity_month_and_confirmed_ok(temp: Path) -> None:
    database = _open_db(temp / "c6")
    folder = "07_рев.02-AN01_AGCC.287-3330-PD"
    september_ns = _mtime_ns(2026, 9, 15)
    december_ns = _mtime_ns(2026, 12, 15)
    january_ns = _mtime_ns(2027, 1, 10)
    event = _event(
        date="01.09.2026",
        stage="tdo_passed",
        stage_label="прошла ТДО",
        revision="02",
        appendix="01",
        transmittals=("AGCC-BCC-TRM-000070",),
    )
    send = _send(
        title="3330",
        mark="PD",
        revision="02",
        appendix="01",
        status="Принят",
        send_date="20.08.2026",
        incoming="25.08.2026",
        transmittal="AGCC-BCC-TRM-000070",
    )
    kits = (_google("3330", "PD", (event,), revision="02", appendix="01"),)
    sends = (send,)

    sep_record = _record(
        50,
        title="3330",
        mark="PD",
        revision="02",
        appendix="01",
        sequence=7,
        folder=folder,
        mtime_ns=september_ns,
    )
    _rebuild(database, kits, sends, [sep_record], {50})
    assert _pipeline(database, "3330", "PD").suspicious is False

    dec_record = _record(
        50,
        title="3330",
        mark="PD",
        revision="02",
        appendix="01",
        sequence=7,
        folder=folder,
        mtime_ns=december_ns,
    )
    rebuild_pipeline(database, records=[dec_record], detected_current_ids={50})
    assert _pipeline(database, "3330", "PD").suspicious is True
    package = database.list_kit_packages("3330", "PD")[0]
    database.upsert_liquidity_review(
        "3330",
        "PD",
        package.transfer_name or folder,
        package.revision_text,
        decision="confirmed_ok",
        sequence=package.sequence,
        evidence_mtime_ns=december_ns,
    )
    rebuild_pipeline(database, records=[dec_record], detected_current_ids={50})
    assert _pipeline(database, "3330", "PD").suspicious is False
    assert database.list_liquidity_reviews("3330", "PD")[0].decision == "confirmed_ok"

    jan_record = _record(
        50,
        title="3330",
        mark="PD",
        revision="02",
        appendix="01",
        sequence=7,
        folder=folder,
        mtime_ns=january_ns,
    )
    rebuild_pipeline(database, records=[jan_record], detected_current_ids={50})
    assert _pipeline(database, "3330", "PD").suspicious is True
    card = get_kit_card(database, "3330", "PD")
    assert card is not None
    assert len(card.liquidity_reviews) == 1


def test_working_revision_not_suspicious(temp: Path) -> None:
    database = _open_db(temp / "c7")
    folder = "08_рев.03_AGCC.287-4440-SS30"
    record = _record(
        60,
        title="4440",
        mark="SS30",
        revision="03",
        appendix=None,
        sequence=8,
        folder=folder,
        mtime_ns=_mtime_ns(2026, 11, 1),
    )
    event = _event(
        date="01.06.2026",
        stage="tdo_passed",
        stage_label="прошла ТДО",
        revision="02",
        appendix="01",
        transmittals=("AGCC-BCC-TRM-000080",),
    )
    send = _send(
        title="4440",
        mark="SS30",
        revision="02",
        appendix="01",
        status="Принят",
        send_date="20.05.2026",
        incoming="25.05.2026",
        transmittal="AGCC-BCC-TRM-000080",
    )
    _rebuild(
        database,
        (_google("4440", "SS30", (event,), revision="02", appendix="01"),),
        (send,),
        [record],
        {60},
    )
    row = _pipeline(database, "4440", "SS30")
    assert row.working_revision_text == "03"
    assert row.official_revision_text == "02-AN01"
    assert row.working_as_build is False
    assert row.status == KitPipelineStatus.TDO_REVIEW
    assert pipeline_review_label(row) == "Прошел ТДО (01.06.2026) · РД 02-AN01"
    assert pipeline_review_label(row, issuance=send) == (
        "Прошел ТДО (01.06.2026) · РД 02-AN01 · отпр. 20.05.2026"
    )
    assert row.suspicious is False
    card = get_kit_card(database, "4440", "SS30")
    assert card is not None
    assert card.working_revision_text == "03"
    hints = list_folder_tree_hints(database)
    hint = hints[(*kit_identity_key("4440", "SS30"), folder.casefold())]
    assert hint.is_working is True
    assert hint.working_origin == "auto"
    assert hint.review_status == ""


def test_working_revision_for_display_only_when_ahead() -> None:
    assert working_revision_for_display("01-AN02", "01-AN01") == "01-AN02"
    assert working_revision_for_display("02", "02") == ""
    assert working_revision_for_display("01", "02") == ""
    assert working_revision_for_display("03", "") == "03"
    assert working_revision_for_display("", "01-AN01") == ""


def test_leftover_higher_rank_in_earlier_folder_is_not_working(temp: Path) -> None:
    """3860-SOT: leftover ``01`` in NN 09 must not beat current ``0-AN01``."""

    database = _open_db(temp / "c3860w")
    old_folder = "09_рев.01_AGCC.287-3860-SOT"
    new_folder = "10_рев.0-AN01_AGCC.287-3860-SOT"
    leftover = _record(
        38601,
        title="3860",
        mark="SOT",
        revision="01",
        appendix=None,
        sequence=9,
        folder=old_folder,
        mtime_ns=_mtime_ns(2025, 6, 1),
        file_kind="mto_xlsx",
    )
    current = _record(
        38602,
        title="3860",
        mark="SOT",
        revision="0",
        appendix="01",
        sequence=10,
        folder=new_folder,
        mtime_ns=_mtime_ns(2026, 3, 12),
        file_kind="mto_xlsx",
    )
    events = (
        _event(
            date="12.03.2026",
            stage="tdo_passed",
            stage_label="прошла ТДО",
            revision="0",
            appendix="01",
            transmittals=("TRM-A",),
        ),
        _event(
            date="14.04.2026",
            stage="code_a",
            stage_label="код А",
            revision="0",
            appendix="01",
            transmittals=("TRM-A",),
        ),
    )
    send = _send(
        title="3860",
        mark="SOT",
        revision="0",
        appendix="01",
        status="Принят",
        send_date="12.03.2026",
        incoming="12.03.2026",
        transmittal="TRM-A",
    )
    _rebuild(
        database,
        (_google("3860", "SOT", events, revision="0", appendix="01"),),
        (send,),
        [leftover, current],
        {current.id},
    )
    row = _pipeline(database, "3860", "SOT")
    assert row.working_revision_text == ""
    assert row.official_revision_text == "0-AN01"
    assert row.status == KitPipelineStatus.AGREED
    hints = list_folder_tree_hints(database)
    old_hint = hints[(*kit_identity_key("3860", "SOT"), old_folder.casefold())]
    new_hint = hints[(*kit_identity_key("3860", "SOT"), new_folder.casefold())]
    assert old_hint.is_working is False
    assert new_hint.is_working is False


def test_working_disk_keeps_issued_agreed_status(temp: Path) -> None:
    database = _open_db(temp / "c7b")
    folder = "06_рев.01-AN02_AGC.287-2000-KSB_as-build"
    record = _record(
        200,
        title="2000",
        mark="KSB",
        revision="01",
        appendix="02",
        sequence=6,
        folder=folder,
        mtime_ns=_mtime_ns(2025, 12, 1),
        as_build=True,
        file_kind="mto_xlsx",
    )
    events = (
        _event(
            date="18.11.2025",
            stage="tdo_passed",
            stage_label="прошла ТДО",
            revision="01",
            appendix="01",
            transmittals=("AGCC.287-BCC-PGS-TRM-000538",),
        ),
        _event(
            date="21.11.2025",
            stage="code_a",
            stage_label="код А",
            revision="01",
            appendix="01",
            transmittals=("PGS-BCC-TRM-000201",),
        ),
    )
    send = _send(
        title="2000",
        mark="KSB",
        revision="01",
        appendix="01",
        status="Принят вх.контр.",
        send_date="18.11.2025",
        incoming="18.11.2025",
        transmittal="AGCC.287-BCC-PGS-TRM-000538",
    )
    _rebuild(
        database,
        (_google("2000", "KSB", events, revision="01", appendix="01"),),
        (send,),
        [record],
        {200},
    )
    row = _pipeline(database, "2000", "KSB")
    assert row.working_revision_text == "01-AN02"
    assert row.official_revision_text == "01-AN01"
    assert row.status == KitPipelineStatus.AGREED
    assert row.code == "A"
    assert row.code_stale is False
    assert pipeline_review_label(row) == "Согласован (21.11.2025) · РД 01-AN01"
    assert pipeline_review_label(row, issuance=send) == (
        "Согласован (21.11.2025) · РД 01-AN01 · отпр. 18.11.2025"
    )
    assert pipeline_approval_label(row, version=PIPELINE_DISPLAY_V1) == (
        "A · 01-AN01 · 21.11.2025"
    )
    assert pipeline_approval_label(row) == "—"
    cells = list_revision_matrix(database)
    working_cells = [
        cell
        for cell in cells
        if cell.title == "2000"
        and cell.mark == "KSB"
        and cell.pipeline_status == "working"
    ]
    assert working_cells
    worklist = list_mto_worklist(database, records=[record])
    assert all(item.status != "working" for item in worklist)
    current_files = iter_mto_files_for_cells(
        database, records=[record], predicate="current"
    )
    assert current_files == ()


def test_iter_mto_current_uses_official_not_working(temp: Path) -> None:
    database = _open_db(temp / "c7d")
    issued = _record(
        210,
        title="2000",
        mark="KSB",
        revision="01",
        appendix="01",
        sequence=5,
        folder="05_рев.01-AN01_AGCC.287-2000-KSB",
        mtime_ns=_mtime_ns(2025, 11, 21),
        file_kind="mto_xlsx",
    )
    working = _record(
        211,
        title="2000",
        mark="KSB",
        revision="01",
        appendix="02",
        sequence=6,
        folder="06_рев.01-AN02_AGC.287-2000-KSB_as-build",
        mtime_ns=_mtime_ns(2025, 12, 1),
        as_build=True,
        file_kind="mto_xlsx",
    )
    records = [issued, working]
    events = (
        _event(
            date="18.11.2025",
            stage="tdo_passed",
            stage_label="прошла ТДО",
            revision="01",
            appendix="01",
            transmittals=("AGCC.287-BCC-PGS-TRM-000538",),
        ),
        _event(
            date="21.11.2025",
            stage="code_a",
            stage_label="код А",
            revision="01",
            appendix="01",
            transmittals=("PGS-BCC-TRM-000201",),
        ),
    )
    send = _send(
        title="2000",
        mark="KSB",
        revision="01",
        appendix="01",
        status="Принят вх.контр.",
        send_date="18.11.2025",
        incoming="18.11.2025",
        transmittal="AGCC.287-BCC-PGS-TRM-000538",
    )
    _rebuild(
        database,
        (_google("2000", "KSB", events, revision="01", appendix="01"),),
        (send,),
        records,
        {211},
    )
    current_files = iter_mto_files_for_cells(
        database, records=records, predicate="current"
    )
    assert [item.path_key for item in current_files] == [issued.path_key]
    cells = list_revision_matrix(database)
    working_cells = [
        cell
        for cell in cells
        if cell.title == "2000"
        and cell.mark == "KSB"
        and cell.revision_text == "01-AN02"
        and cell.pipeline_status == "working"
    ]
    assert working_cells
    worklist = list_mto_worklist(database, records=records)
    assert all(item.status != "working" for item in worklist)


def test_official_current_skips_leftover_files_in_working_package(
    temp: Path,
) -> None:
    """Leftover AN01 copies in the as-build NN must not win open-folder."""

    database = _open_db(temp / "c7e")
    issued_folder = "05_рев.01-AN01_AGCC.287-2000-KSB"
    working_folder = "07_рев.01-AN02_AGC.287-2000-KSB_as-build"
    issued_pdf = _record(
        200,
        title="2000",
        mark="KSB",
        revision="01",
        appendix="01",
        sequence=5,
        folder=issued_folder,
        mtime_ns=_mtime_ns(2025, 11, 21),
    )
    issued_mto = _record(
        210,
        title="2000",
        mark="KSB",
        revision="01",
        appendix="01",
        sequence=5,
        folder=issued_folder,
        mtime_ns=_mtime_ns(2025, 11, 21),
        file_kind="mto_xlsx",
    )
    working_pdf = _record(
        201,
        title="2000",
        mark="KSB",
        revision="01",
        appendix="02",
        sequence=7,
        folder=working_folder,
        mtime_ns=_mtime_ns(2025, 12, 1),
        as_build=True,
    )
    leftover_an01 = _record(
        202,
        title="2000",
        mark="KSB",
        revision="01",
        appendix="01",
        sequence=7,
        folder=working_folder,
        mtime_ns=_mtime_ns(2025, 12, 1),
        as_build=True,
    )
    leftover_empty = _record(
        203,
        title="2000",
        mark="KSB",
        revision="",
        appendix=None,
        sequence=7,
        folder=working_folder,
        mtime_ns=_mtime_ns(2025, 12, 1),
        as_build=True,
        file_kind="mto_xlsx",
    )
    working_mto = _record(
        211,
        title="2000",
        mark="KSB",
        revision="01",
        appendix="02",
        sequence=7,
        folder=working_folder,
        mtime_ns=_mtime_ns(2025, 12, 1),
        as_build=True,
        file_kind="mto_xlsx",
    )
    records = [
        issued_pdf,
        issued_mto,
        working_pdf,
        leftover_an01,
        leftover_empty,
        working_mto,
    ]
    overlay_ids = {201, 202, 203, 211}
    events = (
        _event(
            date="18.11.2025",
            stage="tdo_passed",
            stage_label="прошла ТДО",
            revision="01",
            appendix="01",
            transmittals=("AGCC.287-BCC-PGS-TRM-000538",),
        ),
        _event(
            date="21.11.2025",
            stage="code_a",
            stage_label="код А",
            revision="01",
            appendix="01",
            transmittals=("PGS-BCC-TRM-000201",),
        ),
    )
    send = _send(
        title="2000",
        mark="KSB",
        revision="01",
        appendix="01",
        status="Принят вх.контр.",
        send_date="18.11.2025",
        incoming="18.11.2025",
        transmittal="AGCC.287-BCC-PGS-TRM-000538",
    )
    google = (_google("2000", "KSB", events, revision="01", appendix="01"),)
    _rebuild(database, google, (send,), records, overlay_ids)
    pipelines = database.list_kit_pipelines()
    official_ids = official_detected_current_ids(
        records, overlay_ids, pipelines
    )
    assert official_ids == {issued_pdf.id, issued_mto.id}
    rows = build_kit_matrix(google, records, official_ids, issuance_kits=(send,))
    row = next(item for item in rows if item.title == "2000" and item.mark == "KSB")
    package = issued_package_dir(row.rd.paths[0])
    assert issued_folder in package
    assert working_folder not in package
    assert row.rd.revision_text == "01-AN01"
    worklist = list_mto_worklist(database, records=records)
    current_rows = [
        item
        for item in worklist
        if item.title == "2000" and item.mark == "KSB" and item.is_current
    ]
    assert current_rows
    assert issued_folder in current_rows[0].package_path
    assert working_folder not in current_rows[0].package_path
    assert all(
        working_folder not in item.package_path
        for item in worklist
        if item.title == "2000" and item.mark == "KSB"
    )


def test_send_an01_falls_back_to_disk_rev01_package(temp: Path) -> None:
    """2210-KSB: F/send 01-AN01, disk still 01 — РД · рев. and folder match."""

    database = _open_db(temp / "c2210")
    issued_folder = "06_рев.01_AGC.287-2210-KSB"
    working_folder = "07_рев.01-AN02_AGC.287-2210-KSB_as-build"
    issued = _record(
        2210,
        title="2210",
        mark="KSB",
        revision="01",
        appendix=None,
        sequence=6,
        folder=issued_folder,
        mtime_ns=_mtime_ns(2025, 12, 4),
    )
    working = _record(
        2211,
        title="2210",
        mark="KSB",
        revision="01",
        appendix="02",
        sequence=7,
        folder=working_folder,
        mtime_ns=_mtime_ns(2026, 1, 10),
        as_build=True,
    )
    records = [issued, working]
    overlay_ids = {working.id}
    events = (
        _event(
            date="18.11.2025",
            stage="tdo_passed",
            stage_label="прошла ТДО",
            revision="01",
            appendix="01",
            transmittals=("AGCC.287-BCC-PGS-TRM-000600",),
        ),
        _event(
            date="04.12.2025",
            stage="code_a",
            stage_label="код А",
            revision="01",
            appendix="01",
            transmittals=("PGS-BCC-TRM-000250",),
        ),
    )
    send = _send(
        title="2210",
        mark="KSB",
        revision="01",
        appendix="01",
        status="Принят вх.контр.",
        send_date="18.11.2025",
        incoming="18.11.2025",
        transmittal="AGCC.287-BCC-PGS-TRM-000600",
    )
    google = (_google("2210", "KSB", events, revision="01", appendix="01"),)
    _rebuild(database, google, (send,), records, overlay_ids)
    pipeline = _pipeline(database, "2210", "KSB")
    assert pipeline.official_revision_text == "01-AN01"
    assert pipeline.working_revision_text == "01-AN02"
    assert pipeline_review_label(pipeline) == "Согласован (04.12.2025) · РД 01-AN01"
    assert pipeline_review_label(pipeline, issuance=send) == (
        "Согласован (04.12.2025) · РД 01-AN01 · отпр. 18.11.2025"
    )
    official_ids = official_detected_current_ids(
        records, overlay_ids, (pipeline,)
    )
    assert official_ids == {issued.id}
    rows = build_kit_matrix(google, records, official_ids, issuance_kits=(send,))
    row = next(item for item in rows if item.title == "2210" and item.mark == "KSB")
    assert row.rd.present
    assert row.rd.revision_text == "01"
    assert official_rd_rev_text(row, pipeline) == (
        "01 · 01-AN01 Нет в РД"
    )
    assert row.summary == KitSummary.REV_MISMATCH
    package = issued_package_dir(row.rd.paths[0])
    assert issued_folder in package
    assert working_folder not in package
    current_packages = [
        pkg
        for pkg in database.list_kit_packages("2210", "KSB")
        if pkg.source == "rd" and pkg.is_current
    ]
    assert current_packages
    current_path = current_packages[0].package_path or ""
    assert issued_folder in current_path
    assert working_folder not in current_path


def test_manual_working_flag_hides_same_rev_from_official(temp: Path) -> None:
    database = _open_db(temp / "c7c")
    folder = "05_рев.01-AN01_AGCC.287-3140-KSB2"
    record = _record(
        314,
        title="3140",
        mark="KSB2",
        revision="01",
        appendix="01",
        sequence=5,
        folder=folder,
        mtime_ns=_mtime_ns(2026, 1, 30),
    )
    events = (
        _event(
            date="30.01.2026",
            stage="tdo_passed",
            stage_label="прошла ТДО",
            revision="01",
            appendix="01",
            transmittals=("TRM-A",),
        ),
        _event(
            date="10.02.2026",
            stage="code_a",
            stage_label="код А",
            revision="01",
            appendix="01",
            transmittals=("TRM-A",),
        ),
    )
    send = _send(
        title="3140",
        mark="KSB2",
        revision="01",
        appendix="01",
        status="Принят",
        send_date="30.01.2026",
        incoming="30.01.2026",
        transmittal="TRM-A",
    )
    kits = (_google("3140", "KSB2", events, revision="01", appendix="01"),)
    _rebuild(database, kits, (send,), [record], {314})
    row = _pipeline(database, "3140", "KSB2")
    assert row.status == KitPipelineStatus.AGREED
    assert row.working_revision_text == ""
    database.upsert_working_flag("3140", "KSB2", "01-AN01", transfer_name=folder)
    _rebuild(database, kits, (send,), [record], {314})
    row = _pipeline(database, "3140", "KSB2")
    assert row.working_revision_text == "01-AN01"
    assert row.official_revision_text == ""
    assert row.status == KitPipelineStatus.NOT_UPLOADED


def test_manual_working_flag_keeps_sibling_same_rev(temp: Path) -> None:
    """Flag NN 10; sibling NN 09 with the same filename rev stays official."""

    database = _open_db(temp / "c2612w")
    issued_folder = "09_рев.AN-02_AGC.287-2612-KSB_as-build"
    working_folder = "10_рев.AN-03_AGC.287-2612-KSB_as-build"
    issued = _record(
        26120,
        title="2612",
        mark="KSB",
        revision="01",
        appendix="02",
        sequence=9,
        folder=issued_folder,
        mtime_ns=_mtime_ns(2026, 4, 29),
        as_build=True,
    )
    working = _record(
        26121,
        title="2612",
        mark="KSB",
        revision="01",
        appendix="02",
        sequence=10,
        folder=working_folder,
        mtime_ns=_mtime_ns(2026, 8, 21),
        as_build=True,
    )
    events = (
        _event(
            date="19.11.2025",
            stage="tdo_passed",
            stage_label="прошла ТДО",
            revision="01",
            appendix="02",
            transmittals=("TRM-A",),
        ),
        _event(
            date="30.06.2026",
            stage="code_a",
            stage_label="код А",
            revision="01",
            appendix="02",
            transmittals=("TRM-A",),
        ),
    )
    send = _send(
        title="2612",
        mark="KSB",
        revision="01",
        appendix="02",
        status="Принят",
        send_date="18.11.2025",
        incoming="19.11.2025",
        transmittal="TRM-A",
    )
    records = [issued, working]
    google = (_google("2612", "KSB", events, revision="01", appendix="02"),)
    overlay_ids = {working.id}
    _rebuild(database, google, (send,), records, overlay_ids)
    database.upsert_working_flag(
        "2612",
        "KSB",
        "01-AN02",
        transfer_name=working_folder,
        sequence=10,
    )
    _rebuild(database, google, (send,), records, overlay_ids)
    row = _pipeline(database, "2612", "KSB")
    assert row.working_revision_text == "01-AN02"
    assert row.official_revision_text == "01-AN02"
    assert row.status == KitPipelineStatus.AGREED
    assert row.code == "A"
    assert row.code_stale is False
    assert 10 in row.working_sequences
    current_packages = [
        pkg
        for pkg in database.list_kit_packages("2612", "KSB")
        if pkg.source == "rd" and pkg.is_current
    ]
    assert current_packages
    current_path = current_packages[0].package_path or ""
    assert issued_folder in current_path
    assert working_folder not in current_path
    official_ids = official_detected_current_ids(records, overlay_ids, (row,))
    assert official_ids == {issued.id}
    blind = build_kit_matrix(
        google, records, official_ids, issuance_kits=(send,)
    )
    blind_row = next(
        item for item in blind if item.title == "2612" and item.mark == "KSB"
    )
    assert blind_row.summary is KitSummary.TRANSFER_REVIEW
    blind_notes = "\n".join(blind_row.transfer_review_notes)
    assert "Согласованная передача" in blind_notes
    assert "текущая" in blind_notes
    matrix = build_kit_matrix(
        google,
        records,
        official_ids,
        issuance_kits=(send,),
        working_folders_by_kit={
            kit_identity_key("2612", "KSB"): frozenset(
                {working_folder.casefold()}
            )
        },
    )
    kit_row = next(
        item for item in matrix if item.title == "2612" and item.mark == "KSB"
    )
    assert kit_row.rd.present is True
    assert kit_row.summary is not KitSummary.TRANSFER_REVIEW
    assert "Нет в РД" not in official_rd_rev_text(kit_row, row)
    hints = list_folder_tree_hints(database)
    issued_hint = hints[(*kit_identity_key("2612", "KSB"), issued_folder.casefold())]
    working_hint = hints[
        (*kit_identity_key("2612", "KSB"), working_folder.casefold())
    ]
    assert issued_hint.is_working is False
    assert issued_hint.working_origin == ""
    assert working_hint.is_working is True
    assert working_hint.working_origin == "manual"
    assert working_hint.mto_status == "code_a"


def _insert_overlay_collision(
    database: CatalogDatabase, path_key: str, kind: CollisionKind
) -> None:
    with database._connection() as connection, connection:
        connection.execute(
            """
            INSERT INTO scan_run(started_at, completed_at, status, is_baseline)
            VALUES ('2026-01-01T00:00:00+00:00', '2026-01-01T00:00:00+00:00',
                    'success', 1)
            """
        )
        run_id = int(connection.execute("SELECT last_insert_rowid()").fetchone()[0])
        connection.execute(
            """
            INSERT INTO current_collision(
                scope, source, kind, message, document_key,
                path_keys_json, scan_run_id
            ) VALUES (?, ?, ?, ?, ?, ?, ?)
            """,
            (
                "overlay_mto",
                "rd",
                kind.value,
                "fixture order conflict",
                "mto:fixture",
                json.dumps([path_key], ensure_ascii=False),
                run_id,
            ),
        )


def test_manual_annulled_flag_hides_same_rev_from_official(temp: Path) -> None:
    database = _open_db(temp / "c7c_an")
    folder = "05_рев.01-AN01_AGCC.287-3140-KSB2"
    record = _record(
        315,
        title="3140",
        mark="KSB2",
        revision="01",
        appendix="01",
        sequence=5,
        folder=folder,
        mtime_ns=_mtime_ns(2026, 1, 30),
        file_kind="mto_xlsx",
    )
    events = (
        _event(
            date="30.01.2026",
            stage="tdo_passed",
            stage_label="прошла ТДО",
            revision="01",
            appendix="01",
            transmittals=("TRM-A",),
        ),
        _event(
            date="10.02.2026",
            stage="code_a",
            stage_label="код А",
            revision="01",
            appendix="01",
            transmittals=("TRM-A",),
        ),
    )
    send = _send(
        title="3140",
        mark="KSB2",
        revision="01",
        appendix="01",
        status="Принят",
        send_date="30.01.2026",
        incoming="30.01.2026",
        transmittal="TRM-A",
    )
    kits = (_google("3140", "KSB2", events, revision="01", appendix="01"),)
    _rebuild(database, kits, (send,), [record], {315})
    database.upsert_annulled_flag("3140", "KSB2", "01-AN01", transfer_name=folder)
    _rebuild(database, kits, (send,), [record], {315})
    row = _pipeline(database, "3140", "KSB2")
    assert row.working_revision_text == ""
    assert row.official_revision_text == ""
    assert row.annulled_transfer_names
    assert row.status == KitPipelineStatus.NOT_UPLOADED
    hints = list_folder_tree_hints(database)
    hint = hints[(*kit_identity_key("3140", "KSB2"), folder.casefold())]
    assert hint.is_annulled is True
    assert hint.is_working is False
    cells = [
        cell
        for cell in list_revision_matrix(database)
        if cell.title == "3140" and cell.mark == "KSB2"
    ]
    assert cells
    assert all(cell.pipeline_status == "annulled" for cell in cells)
    worklist = list_mto_worklist(database, records=[record])
    assert not any(
        item.title == "3140" and item.mark == "KSB2" for item in worklist
    )


def test_void_folder_name_annuls_without_manual_flag(temp: Path) -> None:
    """NN folder with Void in the name is annulled on pipeline rebuild."""

    database = _open_db(temp / "c7c_void")
    folder = "04_рев.0-AN02_AGCC.287-7560-SKUD_Void"
    record = _record(
        75601,
        title="7560",
        mark="SKUD",
        revision="0",
        appendix="02",
        sequence=4,
        folder=folder,
        mtime_ns=_mtime_ns(2026, 8, 19),
        file_kind="mto_xlsx",
    )
    send = _send(
        title="7560",
        mark="SKUD",
        revision="0",
        appendix="02",
        status="Принят",
        send_date="19.08.2026",
        incoming="19.08.2026",
        transmittal="TRM-V",
    )
    kits = (_google("7560", "SKUD", (), revision="0", appendix="02"),)
    _rebuild(database, kits, (send,), [record], {75601})
    flags = database.list_annulled_flags("7560", "SKUD")
    assert any(flag.transfer_name == folder for flag in flags)
    row = _pipeline(database, "7560", "SKUD")
    assert any("void" in name.casefold() for name in row.annulled_transfer_names)
    hints = list_folder_tree_hints(database)
    hint = hints[(*kit_identity_key("7560", "SKUD"), folder.casefold())]
    assert hint.is_annulled is True
    assert hint.is_working is False


def test_manual_annulled_flag_keeps_sibling_same_rev(temp: Path) -> None:
    """Annul NN 10; sibling NN 09 with the same filename rev stays official."""

    database = _open_db(temp / "c2612a")
    issued_folder = "09_рев.AN-02_AGC.287-2612-KSB_as-build"
    leftover_folder = "10_рев.AN-03_AGC.287-2612-KSB_as-build"
    issued = _record(
        26130,
        title="2612",
        mark="KSB",
        revision="01",
        appendix="02",
        sequence=9,
        folder=issued_folder,
        mtime_ns=_mtime_ns(2026, 4, 29),
        as_build=True,
    )
    leftover = _record(
        26131,
        title="2612",
        mark="KSB",
        revision="01",
        appendix="02",
        sequence=10,
        folder=leftover_folder,
        mtime_ns=_mtime_ns(2026, 8, 21),
        as_build=True,
    )
    events = (
        _event(
            date="19.11.2025",
            stage="tdo_passed",
            stage_label="прошла ТДО",
            revision="01",
            appendix="02",
            transmittals=("TRM-A",),
        ),
        _event(
            date="30.06.2026",
            stage="code_a",
            stage_label="код А",
            revision="01",
            appendix="02",
            transmittals=("TRM-A",),
        ),
    )
    send = _send(
        title="2612",
        mark="KSB",
        revision="01",
        appendix="02",
        status="Принят",
        send_date="18.11.2025",
        incoming="19.11.2025",
        transmittal="TRM-A",
    )
    records = [issued, leftover]
    google = (_google("2612", "KSB", events, revision="01", appendix="02"),)
    overlay_ids = {leftover.id}
    _rebuild(database, google, (send,), records, overlay_ids)
    database.upsert_annulled_flag(
        "2612",
        "KSB",
        "01-AN02",
        transfer_name=leftover_folder,
        sequence=10,
    )
    _rebuild(database, google, (send,), records, overlay_ids)
    row = _pipeline(database, "2612", "KSB")
    assert row.working_revision_text == ""
    assert row.official_revision_text == "01-AN02"
    assert row.status == KitPipelineStatus.AGREED
    assert row.code == "A"
    current_packages = [
        pkg
        for pkg in database.list_kit_packages("2612", "KSB")
        if pkg.source == "rd" and pkg.is_current
    ]
    assert current_packages
    current_path = current_packages[0].package_path or ""
    assert issued_folder in current_path
    assert leftover_folder not in current_path
    official_ids = official_detected_current_ids(records, overlay_ids, (row,))
    assert official_ids == {issued.id}
    matrix = build_kit_matrix(
        google,
        records,
        official_ids,
        issuance_kits=(send,),
        annulled_folders_by_kit={
            kit_identity_key("2612", "KSB"): frozenset(
                {leftover_folder.casefold()}
            )
        },
    )
    kit_row = next(
        item for item in matrix if item.title == "2612" and item.mark == "KSB"
    )
    assert kit_row.rd.present is True
    assert kit_row.summary is not KitSummary.TRANSFER_REVIEW
    hints = list_folder_tree_hints(database)
    issued_hint = hints[(*kit_identity_key("2612", "KSB"), issued_folder.casefold())]
    leftover_hint = hints[
        (*kit_identity_key("2612", "KSB"), leftover_folder.casefold())
    ]
    assert leftover_hint.is_annulled is True
    assert leftover_hint.is_working is False
    assert issued_hint.is_annulled is False
    assert issued_hint.is_working is False


def test_annulled_overlay_head_is_not_auto_working(temp: Path) -> None:
    """Annul overlay-current NN 08 rev 01; last send 0 stays official."""

    database = _open_db(temp / "c3860a")
    issued_folder = "05_рев.0_AGCC.287-3860-SOT"
    leftover_folder = "08_рев.01_AGCC.287-3860-SOT"
    issued = _record(
        38600,
        title="3860",
        mark="SOT",
        revision="0",
        appendix=None,
        sequence=5,
        folder=issued_folder,
        mtime_ns=_mtime_ns(2025, 6, 1),
        file_kind="mto_xlsx",
    )
    leftover = _record(
        38601,
        title="3860",
        mark="SOT",
        revision="01",
        appendix=None,
        sequence=8,
        folder=leftover_folder,
        mtime_ns=_mtime_ns(2025, 11, 1),
        file_kind="mto_xlsx",
    )
    events = (
        _event(
            date="01.06.2025",
            stage="code_a",
            stage_label="код А",
            revision="0",
        ),
    )
    send = _send(
        title="3860",
        mark="SOT",
        revision="0",
        appendix=None,
        status="Принят",
        send_date="01.06.2025",
        incoming="02.06.2025",
        transmittal="TRM-A",
    )
    records = [issued, leftover]
    google = (_google("3860", "SOT", events, revision="0", appendix=None),)
    overlay_ids = {leftover.id}
    _rebuild(database, google, (send,), records, overlay_ids)
    database.upsert_annulled_flag(
        "3860", "SOT", "01", transfer_name=leftover_folder, sequence=8
    )
    _rebuild(database, google, (send,), records, overlay_ids)
    row = _pipeline(database, "3860", "SOT")
    assert row.working_revision_text == ""
    assert row.official_revision_text == "0"
    official_ids = official_detected_current_ids(records, overlay_ids, (row,))
    assert official_ids == {issued.id}
    leftover_hint = list_folder_tree_hints(database)[
        (*kit_identity_key("3860", "SOT"), leftover_folder.casefold())
    ]
    assert leftover_hint.is_annulled is True
    assert leftover_hint.is_working is False
    cells = {
        cell.revision_text: cell
        for cell in list_revision_matrix(database)
        if cell.title == "3860" and cell.mark == "SOT"
    }
    assert cells["01"].pipeline_status == "annulled"
    assert cells["0"].pipeline_status == "code_a"
    worklist = {
        item.revision_text: item
        for item in list_mto_worklist(database, records=records)
        if item.title == "3860" and item.mark == "SOT"
    }
    assert "01" not in worklist
    assert "0" in worklist


def test_annulled_flag_clears_working_flag(temp: Path) -> None:
    database = _open_db(temp / "c_mutex")
    folder = "08_рев.01_AGCC.287-3860-SOT"
    record = _record(
        38610,
        title="3860",
        mark="SOT",
        revision="01",
        appendix=None,
        sequence=8,
        folder=folder,
        mtime_ns=_mtime_ns(2025, 11, 1),
    )
    kits = (_google("3860", "SOT", (), revision="0", appendix=None),)
    send = _send(
        title="3860",
        mark="SOT",
        revision="0",
        appendix=None,
        status="Принят",
        send_date="01.06.2025",
        incoming="02.06.2025",
    )
    _rebuild(database, kits, (send,), [record], {38610})
    database.upsert_working_flag("3860", "SOT", "01", transfer_name=folder)
    _rebuild(database, kits, (send,), [record], {38610})
    assert _pipeline(database, "3860", "SOT").working_revision_text == "01"
    database.upsert_annulled_flag("3860", "SOT", "01", transfer_name=folder)
    _rebuild(database, kits, (send,), [record], {38610})
    row = _pipeline(database, "3860", "SOT")
    assert row.working_revision_text == ""
    assert row.annulled_transfer_names
    hint = list_folder_tree_hints(database)[
        (*kit_identity_key("3860", "SOT"), folder.casefold())
    ]
    assert hint.is_annulled is True
    assert hint.is_working is False


def test_annulled_folder_drops_overlay_collision_from_heatmap(temp: Path) -> None:
    database = _open_db(temp / "c_coll")
    issued_folder = "05_рев.01_AGCC.287-3860-SOT"
    leftover_folder = "08_рев.01_AGCC.287-3860-SOT"
    issued = _record(
        38620,
        title="3860",
        mark="SOT",
        revision="01",
        appendix=None,
        sequence=5,
        folder=issued_folder,
        mtime_ns=_mtime_ns(2025, 10, 1),
        file_kind="mto_xlsx",
    )
    leftover = _record(
        38621,
        title="3860",
        mark="SOT",
        revision="01",
        appendix=None,
        sequence=8,
        folder=leftover_folder,
        mtime_ns=_mtime_ns(2025, 11, 1),
        file_kind="mto_xlsx",
    )
    records = [issued, leftover]
    kits = (_google("3860", "SOT", (), revision="01", appendix=None),)
    send = _send(
        title="3860",
        mark="SOT",
        revision="01",
        appendix=None,
        status="Принят",
        send_date="01.10.2025",
        incoming="02.10.2025",
    )
    _rebuild(database, kits, (send,), records, {leftover.id})
    _insert_overlay_collision(
        database, leftover.path_key, CollisionKind.TRANSFER_ORDER_CONFLICT
    )
    rebuild_pipeline(database, records=records, detected_current_ids={leftover.id})
    before = next(
        cell
        for cell in list_revision_matrix(database)
        if cell.title == "3860" and cell.revision_text == "01"
    )
    assert CollisionKind.TRANSFER_ORDER_CONFLICT.value in json.loads(
        before.problem_kinds_json
    )
    database.upsert_annulled_flag(
        "3860", "SOT", "01", transfer_name=leftover_folder, sequence=8
    )
    rebuild_pipeline(database, records=records, detected_current_ids={leftover.id})
    after = next(
        cell
        for cell in list_revision_matrix(database)
        if cell.title == "3860" and cell.revision_text == "01"
    )
    assert CollisionKind.TRANSFER_ORDER_CONFLICT.value not in json.loads(
        after.problem_kinds_json
    )
    assert after.pipeline_status != "annulled"


def test_duplicate_nn_working_as_build_keeps_ifc_sibling(temp: Path) -> None:
    """2235-KSB: as-build NN 10 must not hide IFC ``10_SQ-…`` ``01-AN01``."""

    database = _open_db(temp / "c2235w")
    older_folder = "09_рев.01_AGC.287-2235-KSB"
    ifc_folder = "10_SQ-MFCU-AGCC-KSB-02364-0"
    as_built_folder = "10_рев.01-AN01_AGC.287-2235-KSB_as built"
    as_build_folder = "11_рев.01-AN01_AGC.287-2235-KSB_as build"
    older = _record(
        22350,
        title="2235",
        mark="KSB",
        revision="01",
        appendix=None,
        sequence=9,
        folder=older_folder,
        mtime_ns=_mtime_ns(2025, 5, 7),
    )
    ifc = _record(
        22351,
        title="2235",
        mark="KSB",
        revision="01",
        appendix="01",
        sequence=10,
        folder=ifc_folder,
        mtime_ns=_mtime_ns(2025, 11, 21),
    )
    leftover = _record(
        22352,
        title="2235",
        mark="KSB",
        revision="01",
        appendix="01",
        sequence=10,
        folder=as_built_folder,
        mtime_ns=_mtime_ns(2025, 12, 1),
        as_build=True,
    )
    as_built = _record(
        22353,
        title="2235",
        mark="KSB",
        revision="01",
        appendix="02",
        sequence=10,
        folder=as_built_folder,
        mtime_ns=_mtime_ns(2025, 12, 1),
        as_build=True,
    )
    head = _record(
        22354,
        title="2235",
        mark="KSB",
        revision="01",
        appendix="02",
        sequence=11,
        folder=as_build_folder,
        mtime_ns=_mtime_ns(2026, 6, 16),
        as_build=True,
    )
    records = [older, ifc, leftover, as_built, head]
    overlay_ids = {head.id}
    events = (
        _event(
            date="11.11.2025",
            stage="tdo_passed",
            stage_label="прошла ТДО",
            revision="01",
            appendix="01",
            transmittals=("AGCC.287-BCC-PGS-TRM-000530",),
        ),
        _event(
            date="29.11.2025",
            stage="code_a",
            stage_label="код А",
            revision="01",
            appendix="01",
            transmittals=("PGS-BCC-TRM-000208",),
        ),
    )
    send = _send(
        title="2235",
        mark="KSB",
        revision="01",
        appendix="01",
        status="Принят вх.контр.",
        send_date="21.11.2025",
        incoming="24.11.2025",
        transmittal="AGCC.287-BCC-PGS-TRM-000550",
    )
    google = (_google("2235", "KSB", events, revision="01", appendix="01"),)
    _rebuild(database, google, (send,), records, overlay_ids)
    pipeline = _pipeline(database, "2235", "KSB")
    assert pipeline.working_revision_text == "01-AN02"
    assert pipeline.official_revision_text == "01-AN01"
    assert pipeline.working_as_build is True
    assert 10 not in pipeline.working_sequences
    assert 11 in pipeline.working_sequences
    names = {name.casefold() for name in pipeline.working_transfer_names}
    assert as_built_folder.casefold() in names
    assert as_build_folder.casefold() in names
    assert ifc_folder.casefold() not in names
    official_ids = official_detected_current_ids(records, overlay_ids, (pipeline,))
    assert official_ids == {ifc.id}
    matrix = build_kit_matrix(
        google, records, official_ids, issuance_kits=(send,)
    )
    kit_row = next(
        item for item in matrix if item.title == "2235" and item.mark == "KSB"
    )
    assert kit_row.rd.present is True
    assert kit_row.rd.revision_text == "01-AN01"
    assert official_rd_rev_text(kit_row, pipeline) == "01-AN01"
    current_packages = [
        pkg
        for pkg in database.list_kit_packages("2235", "KSB")
        if pkg.source == "rd" and pkg.is_current
    ]
    assert current_packages
    current_path = current_packages[0].package_path or ""
    assert ifc_folder in current_path
    assert older_folder not in current_path
    assert as_built_folder not in current_path
    assert as_build_folder not in current_path
    hints = list_folder_tree_hints(database)
    key = kit_identity_key("2235", "KSB")
    assert hints[(*key, ifc_folder.casefold())].is_working is False
    assert hints[(*key, as_built_folder.casefold())].is_working is True
    assert hints[(*key, as_build_folder.casefold())].is_working is True


def test_unknown_kit_card_is_none(temp: Path) -> None:
    database = _open_db(temp / "c8")
    assert get_kit_card(database, "0000", "ZZ") is None


def test_9000_ksb_stale_code_b(temp: Path) -> None:
    database = _open_db(temp / "c9")
    folder = "05_рев.0_AGCC.287-9000-KSB"
    record = _record(
        90,
        title="9000",
        mark="KSB",
        revision="0",
        appendix=None,
        sequence=5,
        folder=folder,
        mtime_ns=_mtime_ns(2026, 8, 20),
    )
    events = (
        _event(
            date="20.08.2026",
            stage="code_b",
            stage_label="код B",
            revision="0",
            appendix=None,
            transmittals=("AGCC-BCC-TRM-000503",),
        ),
        _event(
            date="28.08.2026",
            stage="tdo_passed",
            stage_label="прошла ТДО",
            revision="0",
            appendix=None,
            transmittals=("AGCC-BCC-TRM-000934",),
        ),
    )
    sends = (
        _send(
            title="9000",
            mark="KSB",
            revision="0",
            appendix=None,
            status="Принят",
            send_date="15.08.2026",
            incoming="18.08.2026",
            transmittal="AGCC-BCC-TRM-000503",
            row_index=1,
        ),
        _send(
            title="9000",
            mark="KSB",
            revision="0",
            appendix=None,
            status="Принят",
            send_date="25.08.2026",
            incoming="27.08.2026",
            transmittal="AGCC-BCC-TRM-000934",
            row_index=2,
        ),
    )
    _rebuild(
        database,
        (_google("9000", "KSB", events, revision="0", appendix=None),),
        sends,
        [record],
        {90},
    )
    row = _pipeline(database, "9000", "KSB")
    assert row.status == KitPipelineStatus.TDO_REVIEW
    assert row.code == "B"
    assert row.code_stale is True
    assert "20.08" in row.code_date
    assert row.code_revision_text == "0"
    assert row.review_as_build is False
    approval = pipeline_approval_label(row)
    assert "B" in approval
    assert "20.08.2026" in approval
    assert approval != "B"
    assert pipeline_approval_relation(row) == APPROVAL_REL_PREVIOUS
    assert approval.endswith("прошлый цикл")
    assert pipeline_status_label(row.status) == "Прошел ТДО"
    assert pipeline_review_label(row) == "Прошел ТДО (28.08.2026) · РД 0"
    assert pipeline_review_label(row, issuance=sends[-1]) == (
        "Прошел ТДО (28.08.2026) · РД 0 · отпр. 25.08.2026"
    )
    events_with_ids = database.list_google_events_with_ids("9000", "KSB")
    code_b_id = next(
        event_id for event_id, event in events_with_ids if event.stage == "code_b"
    )
    cycles = database.list_kit_cycles("9000", "KSB")
    assert len(cycles) == 2
    last_cycle = max(cycles, key=lambda cycle: cycle.id or 0)
    assert last_cycle.code_event_id != code_b_id
    assert last_cycle.code_event_id is None


def test_as_build_after_a(temp: Path) -> None:
    database = _open_db(temp / "c10")
    ifc_folder = "04_рев.01-AN02_AGCC.287-5500-KSB"
    ab_folder = "06_рев.01-AN02_AB_AGCC.287-5500-KSB"
    ifc_record = _record(
        100,
        title="5500",
        mark="KSB",
        revision="01",
        appendix="02",
        sequence=4,
        folder=ifc_folder,
        mtime_ns=_mtime_ns(2026, 5, 20),
    )
    ab_record = _record(
        101,
        title="5500",
        mark="KSB",
        revision="01",
        appendix="02",
        sequence=6,
        folder=ab_folder,
        mtime_ns=_mtime_ns(2026, 9, 1),
        as_build=True,
    )
    events = (
        _event(
            date="20.05.2026",
            stage="code_a",
            stage_label="код A",
            revision="01",
            appendix="02",
            transmittals=("AGCC-BCC-TRM-000100",),
        ),
        _event(
            date="01.09.2026",
            stage="tdo_passed",
            stage_label="прошла ТДО",
            revision="01",
            appendix="02",
            transmittals=("AGCC-BCC-TRM-000101",),
        ),
    )
    sends = (
        _send(
            title="5500",
            mark="KSB",
            revision="01",
            appendix="02",
            status="Принят",
            send_date="10.04.2026",
            incoming="15.04.2026",
            transmittal="AGCC-BCC-TRM-000100",
            row_index=1,
        ),
        _send(
            title="5500",
            mark="KSB",
            revision="01",
            appendix="02",
            status="Принят",
            send_date="25.08.2026",
            incoming="28.08.2026",
            transmittal="AGCC-BCC-TRM-000101",
            row_index=2,
        ),
    )
    _rebuild(
        database,
        (_google("5500", "KSB", events, revision="01", appendix="02"),),
        sends,
        [ifc_record, ab_record],
        {101},
    )
    row = _pipeline(database, "5500", "KSB")
    assert row.review_as_build is True
    assert row.status == KitPipelineStatus.TDO_REVIEW
    assert row.code == "A"
    assert row.code_stale is True
    assert pipeline_review_label(row).endswith(" (AB)")
    assert pipeline_review_label(row, issuance=sends[-1]) == (
        "Прошел ТДО (01.09.2026) · РД 01-AN02 · отпр. 25.08.2026 (AB)"
    )


def test_current_b_no_later_cycle(temp: Path) -> None:
    database = _open_db(temp / "c11")
    folder = "03_рев.01_AGCC.287-7700-POS"
    record = _record(
        110,
        title="7700",
        mark="POS",
        revision="01",
        appendix=None,
        sequence=3,
        folder=folder,
        mtime_ns=_mtime_ns(2026, 6, 10),
    )
    events = (
        _event(
            date="10.06.2026",
            stage="tdo_passed",
            stage_label="прошла ТДО",
            revision="01",
            appendix=None,
            transmittals=("AGCC-BCC-TRM-000200",),
        ),
        _event(
            date="12.06.2026",
            stage="code_b",
            stage_label="код B",
            revision="01",
            appendix=None,
            transmittals=("AGCC-BCC-TRM-000200",),
        ),
    )
    send = _send(
        title="7700",
        mark="POS",
        revision="01",
        appendix=None,
        status="Принят",
        send_date="01.06.2026",
        incoming="05.06.2026",
        transmittal="AGCC-BCC-TRM-000200",
    )
    _rebuild(
        database,
        (_google("7700", "POS", events, revision="01", appendix=None),),
        (send,),
        [record],
        {110},
    )
    row = _pipeline(database, "7700", "POS")
    assert row.status == KitPipelineStatus.TDO_REVIEW
    assert row.code == "B"
    assert row.code_stale is False
    assert pipeline_approval_label(row, version=PIPELINE_DISPLAY_V1) == (
        "B · 01 · 12.06.2026"
    )
    assert pipeline_approval_label(row) == "—"
    assert pipeline_review_label(row, issuance=send) == (
        "Прошел ТДО (10.06.2026) · B · РД 01 · отпр. 01.06.2026"
    )
    assert pipeline_review_label(row, issuance=send, version=PIPELINE_DISPLAY_V1) == (
        "Прошел ТДО (10.06.2026) · РД 01 · отпр. 01.06.2026"
    )


def test_kit_keys_under_folders(temp: Path) -> None:
    pos = _record(
        110,
        title="7700",
        mark="POS",
        revision="01",
        appendix=None,
        sequence=3,
        folder="03_рев.01_AGCC.287-7700-POS",
        mtime_ns=_mtime_ns(2026, 6, 10),
    )
    ksb = _record(
        1,
        title="9110",
        mark="KSB1",
        revision="0",
        appendix="02",
        sequence=5,
        folder="05_рев.0-AN02_AGCC.287-9110-KSB1",
        mtime_ns=_mtime_ns(2026, 8, 5),
    )
    assert kit_keys_under_folders([pos, ksb], ()) == set()
    assert kit_keys_under_folders(
        [pos, ksb], (r"\\bcc\eng\PrDoc\РД\7700",)
    ) == {kit_identity_key("7700", "POS")}
    assert kit_keys_under_folders(
        [pos, ksb], (r"\\bcc\eng\PrDoc\РД",)
    ) == {
        kit_identity_key("7700", "POS"),
        kit_identity_key("9110", "KSB1"),
    }


def _pipe_sig(row) -> tuple:
    return (
        row.status,
        row.code,
        row.code_stale,
        row.working_revision_text,
        row.official_revision_text,
        row.suspicious,
        row.review_as_build,
    )


def _pkg_sig(packages) -> list[tuple]:
    return sorted(
        (
            package.source,
            package.sequence,
            package.package_path,
            package.revision_text,
            package.is_grey,
            package.is_current,
        )
        for package in packages
    )


def test_scoped_rebuild_leaves_other_kit_unchanged(temp: Path) -> None:
    database = _open_db(temp / "c_scope")
    ksb_folder = "05_рев.0-AN02_AGCC.287-9110-KSB1"
    ksb_record = _record(
        1,
        title="9110",
        mark="KSB1",
        revision="0",
        appendix="02",
        sequence=5,
        folder=ksb_folder,
        mtime_ns=_mtime_ns(2026, 8, 5),
    )
    ksb_event = _event(
        date="05.08.2026",
        stage="tdo_passed",
        stage_label="прошла ТДО",
        revision="0",
        appendix="02",
        transmittals=("AGCC.287-PGS-PGS-TRM-20541",),
    )
    ksb_send = _send(
        title="9110",
        mark="KSB1",
        revision="0",
        appendix="02",
        status="Принят",
        send_date="01.08.2026",
        incoming="04.08.2026",
        transmittal="AGCC.287-PGS-PGS-TRM-20541",
    )
    pos_folder = "03_рев.01_AGCC.287-7700-POS"
    pos_record = _record(
        110,
        title="7700",
        mark="POS",
        revision="01",
        appendix=None,
        sequence=3,
        folder=pos_folder,
        mtime_ns=_mtime_ns(2026, 6, 10),
    )
    pos_events = (
        _event(
            date="10.06.2026",
            stage="tdo_passed",
            stage_label="прошла ТДО",
            revision="01",
            appendix=None,
            transmittals=("AGCC-BCC-TRM-000200",),
        ),
        _event(
            date="12.06.2026",
            stage="code_b",
            stage_label="код B",
            revision="01",
            appendix=None,
            transmittals=("AGCC-BCC-TRM-000200",),
        ),
    )
    pos_send = _send(
        title="7700",
        mark="POS",
        revision="01",
        appendix=None,
        status="Принят",
        send_date="01.06.2026",
        incoming="05.06.2026",
        transmittal="AGCC-BCC-TRM-000200",
        row_index=2,
    )
    records = [ksb_record, pos_record]
    current_ids = {1, 110}
    _rebuild(
        database,
        (
            _google("9110", "KSB1", (ksb_event,)),
            _google("7700", "POS", pos_events, revision="01", appendix=None),
        ),
        (ksb_send, pos_send),
        records,
        current_ids,
    )
    ksb_pipe = _pipe_sig(_pipeline(database, "9110", "KSB1"))
    ksb_pkgs = database.list_kit_packages("9110", "KSB1")
    ksb_ids = [package.id for package in ksb_pkgs]
    ksb_cells = [
        (
            cell.revision_text,
            cell.pipeline_status,
            cell.package_ids_json,
        )
        for cell in database.list_revision_cells("9110", "KSB1")
    ]
    pos_later = _record(
        111,
        title="7700",
        mark="POS",
        revision="02",
        appendix=None,
        sequence=4,
        folder="04_рев.02_AGCC.287-7700-POS",
        mtime_ns=_mtime_ns(2026, 7, 10),
    )
    updated = [ksb_record, pos_record, pos_later]
    updated_ids = {1, 110, 111}
    rebuild_pipeline(
        database,
        records=updated,
        detected_current_ids=updated_ids,
        kit_keys={kit_identity_key("7700", "POS")},
    )
    assert _pipe_sig(_pipeline(database, "9110", "KSB1")) == ksb_pipe
    assert [package.id for package in database.list_kit_packages("9110", "KSB1")] == (
        ksb_ids
    )
    assert _pkg_sig(database.list_kit_packages("9110", "KSB1")) == _pkg_sig(ksb_pkgs)
    assert [
        (
            cell.revision_text,
            cell.pipeline_status,
            cell.package_ids_json,
        )
        for cell in database.list_revision_cells("9110", "KSB1")
    ] == ksb_cells
    pos_after_scoped = _pkg_sig(database.list_kit_packages("7700", "POS"))
    pos_pipe_scoped = _pipe_sig(_pipeline(database, "7700", "POS"))
    assert len(pos_after_scoped) == 2
    assert _pipeline(database, "7700", "POS").working_revision_text
    rebuild_pipeline(
        database,
        records=updated,
        detected_current_ids=updated_ids,
    )
    assert _pipe_sig(_pipeline(database, "9110", "KSB1")) == ksb_pipe
    assert _pkg_sig(database.list_kit_packages("9110", "KSB1")) == _pkg_sig(ksb_pkgs)
    assert _pkg_sig(database.list_kit_packages("7700", "POS")) == pos_after_scoped
    assert _pipe_sig(_pipeline(database, "7700", "POS")) == pos_pipe_scoped


def test_folder_hint_mto_on_older_revision(temp: Path) -> None:
    database = _open_db(temp / "mto_older")
    folder = "05_рев.02_AGCC.287-8800-KSB"
    pdf = _record(
        201,
        title="8800",
        mark="KSB",
        revision="02",
        appendix=None,
        sequence=5,
        folder=folder,
        mtime_ns=_mtime_ns(2026, 7, 1),
    )
    mto = _record(
        202,
        title="8800",
        mark="KSB",
        revision="01",
        appendix=None,
        sequence=5,
        folder=folder,
        mtime_ns=_mtime_ns(2026, 6, 1),
        file_kind="mto_xlsx",
    )
    _rebuild(
        database,
        (
            _google(
                "8800",
                "KSB",
                (
                    _event(
                        date="01.06.2026",
                        stage="code_a",
                        stage_label="код А",
                        revision="01",
                    ),
                ),
                revision="02",
                appendix=None,
            ),
        ),
        (
            _send(
                title="8800",
                mark="KSB",
                revision="02",
                appendix=None,
                status="Принят",
                send_date="20.06.2026",
                incoming="25.06.2026",
            ),
        ),
        [pdf, mto],
        {201, 202},
    )
    packages = [
        pkg
        for pkg in database.list_kit_packages("8800", "KSB")
        if pkg.source == "rd" and not pkg.is_grey
    ]
    assert len(packages) == 1
    assert packages[0].revision_text == "02"
    assert packages[0].mto_revision_text == "01"
    hints = list_folder_tree_hints(database)
    hint = hints[(*kit_identity_key("8800", "KSB"), folder.casefold())]
    assert hint.mto_status == "code_a"
    assert hint.mto_revision_text == "01"
    assert hint.has_mto_file is True


def test_pipeline_algorithm_needs_rebuild() -> None:
    current = KitPipelineRow(
        title="2000",
        mark="KSB",
        status="agreed",
        algorithm_version=PIPELINE_STATUS_ALGORITHM_VERSION,
        official_revision_text="01-AN01",
        working_revision_text="01-AN02",
    )
    stale = KitPipelineRow(
        title="2000",
        mark="KSB",
        status="tdo_review",
        algorithm_version=3,
        official_revision_text="01-AN02",
        working_revision_text="01-AN02",
    )
    assert pipeline_algorithm_needs_rebuild(()) is False
    assert pipeline_algorithm_needs_rebuild((current,)) is False
    assert pipeline_algorithm_needs_rebuild((stale,)) is True
    assert pipeline_algorithm_needs_rebuild((current, stale)) is True


def test_tdo_review_label_uses_passed_date() -> None:
    stored = KitPipelineRow(
        title="2000",
        mark="KSB",
        status=KitPipelineStatus.TDO_REVIEW.value,
        official_revision_text="02",
        tdo_date="13.09.2026",
    )
    assert pipeline_review_label(stored) == "Прошел ТДО (13.09.2026) · РД 02"
    assert pipeline_tdo_passed_date_text(stored) == "13.09.2026"
    empty = KitPipelineRow(
        title="2000",
        mark="KSB",
        status=KitPipelineStatus.TDO_REVIEW.value,
        official_revision_text="02",
    )
    assert pipeline_review_label(empty) == "Прошел ТДО · РД 02"
    event = _event(
        date="02.09.2026",
        stage="tdo_passed",
        stage_label="прошла ТДО",
        revision="02",
        appendix=None,
    )
    assert pipeline_review_label(empty, events=(event,)) == (
        "Прошел ТДО (02.09.2026) · РД 02"
    )
    assert pipeline_tdo_passed_date_text(empty, events=(event,)) == "02.09.2026"
    send = _send(
        title="2000",
        mark="KSB",
        revision="02",
        appendix=None,
        status="Принят",
        send_date="01.09.2026",
        transmittal="TRM-SEND",
    )
    assert pipeline_review_label(stored, issuance=send) == (
        "Прошел ТДО (13.09.2026) · РД 02 · отпр. 01.09.2026"
    )


def test_review_label_appends_send_date_for_all_statuses() -> None:
    send = _send(
        title="2000",
        mark="KSB",
        revision="01",
        appendix="01",
        status="Принят",
        send_date="18.11.2025",
        incoming="19.11.2025",
        transmittal="TRM-1",
    )
    agreed = KitPipelineRow(
        title="2000",
        mark="KSB",
        status=KitPipelineStatus.AGREED.value,
        official_revision_text="01-AN01",
    )
    assert pipeline_review_label(agreed) == "Согласован · РД 01-AN01"
    assert pipeline_review_label(agreed, issuance=send) == (
        "Согласован · РД 01-AN01 · отпр. 18.11.2025"
    )
    tdo = KitPipelineRow(
        title="2000",
        mark="KSB",
        status=KitPipelineStatus.TDO_REVIEW.value,
        official_revision_text="01-AN01",
        tdo_date="21.11.2025",
    )
    assert pipeline_review_label(tdo, issuance=send) == (
        "Прошел ТДО (21.11.2025) · РД 01-AN01 · отпр. 18.11.2025"
    )
    sent = KitPipelineRow(
        title="2000",
        mark="KSB",
        status=KitPipelineStatus.SENT_TDO.value,
        official_revision_text="01-AN01",
    )
    assert pipeline_review_label(sent, issuance=send) == (
        "Отправлен на ТДО · РД 01-AN01 · отпр. 18.11.2025"
    )
    missing = KitPipelineRow(
        title="2000",
        mark="KSB",
        status=KitPipelineStatus.NOT_UPLOADED.value,
        official_revision_text="01-AN01",
    )
    assert pipeline_review_label(missing, issuance=send) == (
        "Не загружен в СР · РД 01-AN01 · отпр. 18.11.2025"
    )
    as_build = replace(agreed, review_as_build=True)
    assert pipeline_review_label(as_build, issuance=send) == (
        "Согласован · РД 01-AN01 · отпр. 18.11.2025 (AB)"
    )
    other_rev = _send(
        title="2000",
        mark="KSB",
        revision="02",
        appendix=None,
        status="Принят",
        send_date="01.01.2026",
        transmittal="TRM-2",
    )
    f_sent = _event(
        date="03.08.2026",
        stage="tdo_sent",
        stage_label="отпр. на ТДО",
        revision="01",
        appendix="01",
    )
    assert pipeline_review_label(
        agreed, events=(f_sent,), issuance=other_rev, version=PIPELINE_DISPLAY_V2
    ) == "Согласован · РД 01-AN01 · отпр. 03.08.2026"
    v3_ahead = pipeline_review_label(
        agreed, events=(f_sent,), issuance=other_rev
    )
    assert "выдача 02" in v3_ahead
    assert "РД 01-AN01" not in v3_ahead
    assert pipeline_send_date_text(
        agreed, events=(f_sent,), version=PIPELINE_DISPLAY_V2
    ) == "03.08.2026"


def test_pipeline_approval_relation_suffix() -> None:
    current = KitPipelineRow(
        title="2000",
        mark="KSB",
        status=KitPipelineStatus.AGREED.value,
        code="A",
        code_revision_text="01-AN01",
        code_date="04.03.2026",
        official_revision_text="01-AN01",
    )
    assert pipeline_approval_relation(current) == ""
    assert pipeline_approval_label(current, version=PIPELINE_DISPLAY_V1) == (
        "A · 01-AN01 · 04.03.2026"
    )
    assert pipeline_approval_label(current) == "—"
    ahead = replace(
        current,
        code_stale=True,
        official_revision_text="01",
        code_date="29.05.2026",
    )
    assert pipeline_approval_relation(ahead) == APPROVAL_REL_AHEAD
    assert pipeline_approval_label(ahead) == (
        "A · 01-AN01 · 29.05.2026 · новее диска"
    )
    previous_same = replace(
        current,
        status=KitPipelineStatus.TDO_REVIEW.value,
        code="B",
        code_stale=True,
        code_revision_text="0",
        official_revision_text="0",
        code_date="20.08.2026",
    )
    assert pipeline_approval_relation(previous_same) == APPROVAL_REL_PREVIOUS
    assert pipeline_approval_label(previous_same).endswith("прошлый цикл")
    previous_lower = replace(
        current,
        code_stale=True,
        code_revision_text="01",
        official_revision_text="02",
    )
    assert pipeline_approval_relation(previous_lower) == APPROVAL_REL_PREVIOUS
    no_rev = replace(current, code_stale=True, code_revision_text="")
    assert pipeline_approval_relation(no_rev) == APPROVAL_REL_NO_REV
    assert pipeline_approval_label(no_rev).endswith("без рев. в F")
    other = replace(current, code_stale=True, official_revision_text="")
    assert pipeline_approval_relation(other) == APPROVAL_REL_OTHER
    assert pipeline_approval_label(other).endswith("не этого цикла")


def test_pipeline_display_v2_current_bc_on_review() -> None:
    tdo_b = KitPipelineRow(
        title="7700",
        mark="POS",
        status=KitPipelineStatus.TDO_REVIEW.value,
        code="B",
        code_revision_text="01",
        code_date="12.06.2026",
        official_revision_text="01",
        tdo_date="10.06.2026",
    )
    assert pipeline_review_label(tdo_b) == "Прошел ТДО (10.06.2026) · B · РД 01"
    assert pipeline_approval_label(tdo_b) == "—"
    assert pipeline_review_label(tdo_b, version=PIPELINE_DISPLAY_V1) == (
        "Прошел ТДО (10.06.2026) · РД 01"
    )
    assert pipeline_approval_label(tdo_b, version=PIPELINE_DISPLAY_V1) == (
        "B · 01 · 12.06.2026"
    )
    stale_ahead = replace(
        tdo_b,
        code_stale=True,
        code_revision_text="0-AN02",
        official_revision_text="0-AN01",
    )
    assert pipeline_review_label(stale_ahead) == (
        "Прошел ТДО (10.06.2026) · РД 0-AN01"
    )
    assert pipeline_approval_label(stale_ahead).endswith("новее диска")
    agreed_a = KitPipelineRow(
        title="2000",
        mark="KSB",
        status=KitPipelineStatus.AGREED.value,
        code="A",
        code_revision_text="01-AN01",
        code_date="21.11.2025",
        official_revision_text="01-AN01",
    )
    assert " · A" not in pipeline_review_label(agreed_a)
    assert pipeline_review_label(agreed_a) == (
        "Согласован (21.11.2025) · РД 01-AN01"
    )
    assert pipeline_approval_label(agreed_a) == "—"


def test_pipeline_review_label_agreed_date_not_send_date() -> None:
    row = KitPipelineRow(
        title="8950",
        mark="SOT1",
        status=KitPipelineStatus.AGREED.value,
        code="A",
        code_revision_text="03-AN02",
        code_date="21.03.2026",
        official_revision_text="03-AN02",
    )
    send = _send(
        title="8950",
        mark="SOT1",
        revision="03",
        appendix="02",
        status="Принят",
        send_date="18.03.2026",
        transmittal="TRM-8950",
    )
    assert pipeline_review_label(row, issuance=send) == (
        "Согласован (21.03.2026) · РД 03-AN02 · отпр. 18.03.2026"
    )
    later_a = _event(
        date="22.03.2026",
        stage="code_a",
        stage_label="код А",
        revision="03",
        appendix="02",
    )
    assert pipeline_review_label(row, events=(later_a,), issuance=send) == (
        "Согласован (22.03.2026) · РД 03-AN02 · отпр. 18.03.2026"
    )


def test_pipeline_display_v3_google_face_ahead_of_disk() -> None:
    events = (
        _event(
            date="01.06.2026",
            stage="tdo_sent",
            stage_label="отправлена на ТДО",
            revision="0",
            appendix="02",
            transmittals=("TRM-300",),
        ),
        _event(
            date="10.06.2026",
            stage="tdo_passed",
            stage_label="прошла ТДО",
            revision="0",
            appendix="02",
            transmittals=("TRM-300",),
        ),
        _event(
            date="12.06.2026",
            stage="code_b",
            stage_label="код B",
            revision="0",
            appendix="02",
            transmittals=("TRM-300",),
        ),
    )
    google = _google("6550", "SKUD", events, revision="0", appendix="02")
    send = _send(
        title="6550",
        mark="SKUD",
        revision="0",
        appendix="02",
        status="На рассмотрении",
        send_date="01.06.2026",
        incoming="10.06.2026",
        transmittal="TRM-300",
    )
    row = KitPipelineRow(
        title="6550",
        mark="SKUD",
        status=KitPipelineStatus.TDO_REVIEW.value,
        official_revision_text="0-AN01",
        tdo_date="09.10.2023",
        code="B",
        code_stale=True,
        code_revision_text="0-AN02",
        code_date="12.06.2026",
    )
    label = pipeline_review_label(row, google=google, issuance=send)
    assert label == (
        "Прошел ТДО (10.06.2026) · B · выдача 0-AN02 · отпр. 01.06.2026"
    )
    assert "0-AN01" not in label
    assert "09.10.2023" not in label
    assert pipeline_approval_label(row, google=google, issuance=send) == "—"
    v2_review = pipeline_review_label(
        row, google=google, issuance=send, version=PIPELINE_DISPLAY_V2
    )
    assert v2_review == "Прошел ТДО (09.10.2023) · РД 0-AN01"
    assert pipeline_approval_label(
        row, google=google, issuance=send, version=PIPELINE_DISPLAY_V2
    ).endswith("новее диска")
    v1_review = pipeline_review_label(
        row, google=google, issuance=send, version=PIPELINE_DISPLAY_V1
    )
    assert " · B" not in v1_review
    assert "РД 0-AN01" in v1_review


def test_pipeline_display_v3_de_caption_does_not_force_agreed() -> None:
    google = _google(
        "7417",
        "SOS",
        (),
        revision="0",
        appendix="02",
        status_sheet="РД Согласовано",
    )
    row = KitPipelineRow(
        title="7417",
        mark="SOS",
        status=KitPipelineStatus.TDO_REVIEW.value,
        official_revision_text="0-AN01",
        tdo_date="09.10.2023",
        code="A",
        code_stale=True,
        code_revision_text="0-AN02",
        code_date="29.05.2026",
    )
    label = pipeline_review_label(row, google=google)
    assert label.startswith("Не загружен в СР")
    assert "D/E 0-AN02" in label
    assert not label.startswith("Согласован")
    assert pipeline_approval_label(row, google=google).endswith("новее диска")


def test_index_records_by_kit_groups_and_omits() -> None:
    first = _record(
        1,
        title="2245",
        mark="KSB",
        revision="01",
        appendix=None,
        sequence=1,
        folder="01_рев.01_2245-KSB",
        mtime_ns=_mtime_ns(2026, 8, 1),
    )
    second = _record(
        2,
        title="3240",
        mark="KSB1",
        revision="01",
        appendix=None,
        sequence=1,
        folder="01_рев.01_3240-KSB1",
        mtime_ns=_mtime_ns(2026, 8, 1),
    )
    first_again = _record(
        3,
        title="2245",
        mark="KSB",
        revision="02",
        appendix=None,
        sequence=2,
        folder="02_рев.02_2245-KSB",
        mtime_ns=_mtime_ns(2026, 8, 2),
    )
    omitted = _record(
        4,
        title="12",
        mark="KSB",
        revision="01",
        appendix=None,
        sequence=1,
        folder="01_рев.01_12-KSB",
        mtime_ns=_mtime_ns(2026, 8, 1),
    )
    grouped = index_records_by_kit([first, second, first_again, omitted])
    key_a = kit_identity_key("2245", "KSB")
    key_b = kit_identity_key("3240", "KSB1")
    assert set(grouped) == {key_a, key_b}
    assert grouped[key_a] == [first, first_again]
    assert grouped[key_b] == [second]


def test_patch_official_ids_scopes_to_one_kit() -> None:
    """Patching one working kit must not drop another kit's official ids."""

    first = _record(
        11,
        title="2245",
        mark="KSB",
        revision="01",
        appendix=None,
        sequence=1,
        folder="01_рев.01_2245-KSB",
        mtime_ns=_mtime_ns(2026, 8, 1),
    )
    second = _record(
        22,
        title="3240",
        mark="KSB1",
        revision="01",
        appendix=None,
        sequence=1,
        folder="01_рев.01_3240-KSB1",
        mtime_ns=_mtime_ns(2026, 8, 1),
    )
    idle = KitPipelineRow(
        title="2245",
        mark="KSB",
        status="agreed",
        official_revision_text="01",
    )
    working = KitPipelineRow(
        title="3240",
        mark="KSB1",
        status="working",
        official_revision_text="",
        working_revision_text="01",
    )
    detected = {11, 22}
    before = official_detected_current_ids([first, second], detected, (idle,))
    assert before == {11, 22}
    after = patch_official_detected_current_ids(
        before,
        records=[first, second],
        detected_current_ids=detected,
        pipelines=(idle, working),
        kit_keys={kit_identity_key("3240", "KSB1")},
    )
    full = official_detected_current_ids(
        [first, second], detected, (idle, working)
    )
    assert after == full
    assert 11 in after
    assert 22 not in after


def test_folder_hint_package_without_mto(temp: Path) -> None:
    database = _open_db(temp / "no_mto_hint")
    folder = "04_рев.02_AGCC.287-8810-KSB"
    pdf = _record(
        211,
        title="8810",
        mark="KSB",
        revision="02",
        appendix=None,
        sequence=4,
        folder=folder,
        mtime_ns=_mtime_ns(2026, 7, 15),
    )
    _rebuild(
        database,
        (
            _google(
                "8810",
                "KSB",
                (
                    _event(
                        date="15.07.2026",
                        stage="code_a",
                        stage_label="код А",
                        revision="02",
                    ),
                ),
                revision="02",
                appendix=None,
            ),
        ),
        (
            _send(
                title="8810",
                mark="KSB",
                revision="02",
                appendix=None,
                status="Принят",
                send_date="01.07.2026",
                incoming="05.07.2026",
            ),
        ),
        [pdf],
        {211},
    )
    packages = [
        pkg
        for pkg in database.list_kit_packages("8810", "KSB")
        if pkg.source == "rd" and not pkg.is_grey
    ]
    assert len(packages) == 1
    assert packages[0].revision_text == "02"
    assert packages[0].mto_revision_text == ""
    hints = list_folder_tree_hints(database)
    hint = hints[(*kit_identity_key("8810", "KSB"), folder.casefold())]
    assert hint.has_mto_file is False
    assert hint.mto_revision_text == "02"
    assert hint.mto_status == "code_a"


def test_1600_sos_annul_03_uses_effective_sends(temp: Path) -> None:
    database = _open_db(temp / "sos_annul")
    folder = "02_рев.02_AGCC.287-1600-SOS"
    record = _record(
        1,
        title="1600",
        mark="SOS",
        revision="02",
        appendix=None,
        sequence=2,
        folder=folder,
        mtime_ns=_mtime_ns(2026, 8, 1),
    )
    send_02 = _send(
        title="1600",
        mark="SOS",
        revision="02",
        appendix=None,
        status="Отправлен",
        send_date="01.08.2026",
        transmittal="AGCC-BCC-TRM-000010",
        row_index=10,
    )
    send_03 = _send(
        title="1600",
        mark="SOS",
        revision="03",
        appendix=None,
        status="Отправлен",
        send_date="15.08.2026",
        transmittal="AGCC-BCC-TRM-000011",
        row_index=11,
        note_raw="old note",
    )
    _rebuild(database, (), (send_02, send_03), [record], {1})
    grey_03 = [
        pkg
        for pkg in database.list_kit_packages("1600", "SOS")
        if pkg.is_grey and pkg.revision_text == "03"
    ]
    assert grey_03
    review_id = _annul_send(database, send_03, "03 снята, действует 02")
    rebuild_pipeline(
        database,
        records=[record],
        detected_current_ids={1},
    )
    latest = latest_effective_issuance_kits(database, "1600", "SOS")
    assert len(latest) == 1
    assert latest[0].revision_text == "02"
    card = get_kit_card(database, "1600", "SOS")
    assert card is not None
    assert card.issuance is not None
    assert card.issuance.revision_text == "02"
    assert not [
        pkg
        for pkg in database.list_kit_packages("1600", "SOS")
        if pkg.is_grey and pkg.revision_text == "03"
    ]
    raw = database.list_issuance_sends("1600", "SOS")
    assert {item.revision_text for item in raw} == {"02", "03"}

    refreshed = _send(
        title="1600",
        mark="SOS",
        revision="03",
        appendix=None,
        status="Отправлен",
        send_date="15.08.2026",
        transmittal="AGCC-BCC-TRM-000011",
        row_index=44,
        note_raw="sheet note edited",
    )
    ingest_google_snapshot(
        database,
        (),
        (send_02, refreshed),
        loaded_at=_LOADED_AT,
        source="test",
    )
    row = database.get_issuance_review(review_id)
    assert row is not None
    assert row.decision == "annulled"
    assert row.comment == "03 снята, действует 02"
    rebuild_pipeline(
        database,
        records=[record],
        detected_current_ids={1},
    )
    latest = latest_effective_issuance_kits(database, "1600", "SOS")
    assert latest[0].revision_text == "02"
    still = database.get_issuance_review(review_id)
    assert still is not None
    assert still.decision == "annulled"


def test_1711_sot_code_a_without_sheet_send_stays_agreed(temp: Path) -> None:
    """F code A on overlay 02 is current even without «Выдача РД ПД».

    Legalize with the RD file date after the letter must not invent a
    later TRM-less cycle that stale-marks A and drops review to sent_tdo.
    """

    database = _open_db(temp / "sot_legalize")
    folder = "05_рев.02_AGCC.287-1711-SOT"
    record = _record(
        1711,
        title="1711",
        mark="SOT",
        revision="02",
        appendix=None,
        sequence=5,
        folder=folder,
        mtime_ns=_mtime_ns(2025, 3, 19),
    )
    events = (
        _event(
            date="03.07.2023",
            stage="dup",
            stage_label="ДУП",
        ),
        _event(
            date="24.03.2024",
            stage="code_a",
            stage_label="код А",
            revision="02",
            appendix=None,
            transmittals=("AGCC-BCC-TRM-000293",),
        ),
    )
    _rebuild(
        database,
        (_google("1711", "SOT", events, revision="02", appendix=None),),
        (),
        [record],
        {1711},
    )
    row = _pipeline(database, "1711", "SOT")
    assert row.status == KitPipelineStatus.AGREED
    assert row.code == "A"
    assert row.code_stale is False
    assert row.code_revision_text == "02"
    assert "24.03" in row.code_date

    add_manual_journal_row(
        database,
        "1711",
        "SOT",
        revision_text="02",
        send_date="19.03.2025",
        note="легализация РД",
        decision="legalized",
    )
    rebuild_pipeline(
        database,
        records=[record],
        detected_current_ids={1711},
    )
    after = _pipeline(database, "1711", "SOT")
    assert after.status == KitPipelineStatus.AGREED
    assert after.code == "A"
    assert after.code_stale is False
    legalized = [
        cycle
        for cycle in database.list_kit_cycles("1711", "SOT")
        if cycle.match_reason == "legalized"
    ]
    assert legalized
    assert legalized[0].send_id is None
    assert legalized[0].revision_text == "02"


def test_1600_sos_legalize_orphan_in_heatmap(temp: Path) -> None:
    database = _open_db(temp / "sos_legalize")
    folder = "02_рев.02_AGCC.287-1600-SOS"
    record = _record(
        1,
        title="1600",
        mark="SOS",
        revision="02",
        appendix=None,
        sequence=2,
        folder=folder,
        mtime_ns=_mtime_ns(2026, 8, 1),
    )
    send_02 = _send(
        title="1600",
        mark="SOS",
        revision="02",
        appendix=None,
        status="Отправлен",
        send_date="01.08.2026",
        transmittal="AGCC-BCC-TRM-000010",
        row_index=10,
    )
    send_03 = _send(
        title="1600",
        mark="SOS",
        revision="03",
        appendix=None,
        status="Отправлен",
        send_date="15.08.2026",
        transmittal="AGCC-BCC-TRM-000011",
        row_index=11,
    )
    _rebuild(database, (), (send_02, send_03), [record], {1})
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
    rebuild_pipeline(
        database,
        records=[record],
        detected_current_ids={1},
    )
    effective = effective_issuance_sends(database, "1600", "SOS")
    texts = [send.revision_text for _send_id, send in effective]
    assert "02" in texts
    assert "02-AN01" in texts
    assert latest_effective_issuance_kits(database, "1600", "SOS")[0].revision_text == (
        "02"
    )
    legalized = [
        cycle
        for cycle in database.list_kit_cycles("1600", "SOS")
        if cycle.revision_text == "02-AN01"
    ]
    assert legalized
    assert legalized[0].send_id is None
    assert legalized[0].match_reason == "legalized"
    cells = list_revision_matrix(database)
    assert any(
        cell.title == "1600"
        and cell.mark == "SOS"
        and cell.revision_text == "02-AN01"
        for cell in cells
    )
    worklist = list_mto_worklist(database, records=[record])
    an01 = [
        row
        for row in worklist
        if row.title == "1600"
        and row.mark == "SOS"
        and row.revision_text == "02-AN01"
    ]
    assert an01
    assert an01[0].in_issuance is True


def test_pick_official_rd_package_skips_delta_between_agreed_and_working() -> None:
    """8950-SOO1: open-folder is NN 11 (03-AN01), not is_current NN 13."""

    base = r"\\bcc\eng\PrDoc\РД\8950\14_SOO1\Для передачи"
    packages = (
        KitPackageRow(
            title="8950",
            mark="SOO1",
            source="rd",
            sequence=11,
            transfer_name="11_рев.03-AN01_от_2026.01.22",
            package_path=rf"{base}\11_рев.03-AN01_от_2026.01.22",
            revision_text="03-AN01",
            is_current=False,
            id=11,
        ),
        KitPackageRow(
            title="8950",
            mark="SOO1",
            source="rd",
            sequence=13,
            transfer_name="13_рев.03_AN02_AGCC.287-8950-SOO1",
            package_path=rf"{base}\13_рев.03_AN02_AGCC.287-8950-SOO1",
            revision_text="03-AN02",
            is_current=True,
            id=13,
        ),
        KitPackageRow(
            title="8950",
            mark="SOO1",
            source="rd",
            sequence=14,
            transfer_name="14_рев.03_AN03_AGCC.287-8950-SOO1",
            package_path=rf"{base}\14_рев.03_AN03_AGCC.287-8950-SOO1",
            revision_text="03-AN03",
            is_current=False,
            id=14,
        ),
    )
    pipeline = KitPipelineRow(
        title="8950",
        mark="SOO1",
        status="agreed",
        code="A",
        official_revision_text="03-AN01",
        working_revision_text="03-AN03",
        working_transfer_names=("14_рев.03_AN03_AGCC.287-8950-SOO1",),
        working_sequences=(14,),
    )
    picked = pick_official_rd_package(packages, pipeline)
    assert picked is not None
    assert picked.sequence == 11
    assert picked.revision_text == "03-AN01"


def test_pick_official_rd_package_falls_back_when_official_rev_missing() -> None:
    """Send 01-AN01, disk still 01 → newest non-working package."""

    base = r"\\bcc\eng\PrDoc\РД\2000\KSB\Для передачи"
    packages = (
        KitPackageRow(
            title="2000",
            mark="KSB",
            source="rd",
            sequence=8,
            transfer_name="08_рев.01_AGCC.287-2000-KSB",
            package_path=rf"{base}\08_рев.01_AGCC.287-2000-KSB",
            revision_text="01",
            is_current=True,
            id=8,
        ),
        KitPackageRow(
            title="2000",
            mark="KSB",
            source="rd",
            sequence=9,
            transfer_name="09_рев.01-AN02_working",
            package_path=rf"{base}\09_рев.01-AN02_working",
            revision_text="01-AN02",
            is_current=False,
            id=9,
        ),
    )
    pipeline = KitPipelineRow(
        title="2000",
        mark="KSB",
        status="agreed",
        code="A",
        official_revision_text="01-AN01",
        working_revision_text="01-AN02",
        working_transfer_names=("09_рев.01-AN02_working",),
        working_sequences=(9,),
    )
    picked = pick_official_rd_package(packages, pipeline)
    assert picked is not None
    assert picked.sequence == 8
    assert picked.revision_text == "01"


def main() -> None:
    """Run pipeline fixtures against a temporary SQLite file."""

    test_index_records_by_kit_groups_and_omits()
    test_patch_official_ids_scopes_to_one_kit()
    test_pipeline_algorithm_needs_rebuild()
    test_tdo_review_label_uses_passed_date()
    test_review_label_appends_send_date_for_all_statuses()
    test_pipeline_approval_relation_suffix()
    test_pipeline_display_v2_current_bc_on_review()
    test_pipeline_review_label_agreed_date_not_send_date()
    test_pipeline_display_v3_google_face_ahead_of_disk()
    test_pipeline_display_v3_de_caption_does_not_force_agreed()
    test_pick_official_rd_package_skips_delta_between_agreed_and_working()
    test_pick_official_rd_package_falls_back_when_official_rev_missing()
    with tempfile.TemporaryDirectory(prefix="rd_catalog_pipeline_") as raw:
        root = Path(raw)
        test_9110_tdo_review_package_path(root)
        test_two_sends_one_package_two_cycles(root)
        test_grey_issuance_without_rd(root)
        test_code_a_customer(root)
        test_code_b_origin_pi_and_customer(root)
        test_liquidity_month_and_confirmed_ok(root)
        test_working_revision_not_suspicious(root)
        test_working_revision_for_display_only_when_ahead()
        test_leftover_higher_rank_in_earlier_folder_is_not_working(root)
        test_working_disk_keeps_issued_agreed_status(root)
        test_iter_mto_current_uses_official_not_working(root)
        test_official_current_skips_leftover_files_in_working_package(root)
        test_send_an01_falls_back_to_disk_rev01_package(root)
        test_manual_working_flag_hides_same_rev_from_official(root)
        test_manual_working_flag_keeps_sibling_same_rev(root)
        test_manual_annulled_flag_hides_same_rev_from_official(root)
        test_void_folder_name_annuls_without_manual_flag(root)
        test_manual_annulled_flag_keeps_sibling_same_rev(root)
        test_annulled_overlay_head_is_not_auto_working(root)
        test_annulled_flag_clears_working_flag(root)
        test_annulled_folder_drops_overlay_collision_from_heatmap(root)
        test_duplicate_nn_working_as_build_keeps_ifc_sibling(root)
        test_unknown_kit_card_is_none(root)
        test_9000_ksb_stale_code_b(root)
        test_as_build_after_a(root)
        test_current_b_no_later_cycle(root)
        test_kit_keys_under_folders(root)
        test_scoped_rebuild_leaves_other_kit_unchanged(root)
        test_folder_hint_mto_on_older_revision(root)
        test_folder_hint_package_without_mto(root)
        test_1600_sos_annul_03_uses_effective_sends(root)
        test_1711_sot_code_a_without_sheet_send_stays_agreed(root)
        test_1600_sos_legalize_orphan_in_heatmap(root)
    test_records_in_contour_keeps_sq_and_drops_loose_rd()
    print("RD catalog pipeline: OK")


def test_records_in_contour_keeps_sq_and_drops_loose_rd() -> None:
    """SQ stays; only canonical issued RD paths remain in the contour."""

    rd_root = r"\\bcc\eng\PrDoc\РД"
    canonical = _record(
        1,
        title="2235",
        mark="KSB",
        revision="01",
        appendix="AN01",
        sequence=1,
        folder="01_рев.01-AN01",
        mtime_ns=_mtime_ns(2026, 1, 1),
    )
    loose = replace(
        canonical,
        id=2,
        path=rf"{rd_root}\2235\KSB\loose.pdf",
        path_key="rd/2235-KSB/loose",
    )
    sq = replace(
        canonical,
        id=3,
        source=SourceKind.SQ,
        path=r"\\bcc\eng\SQ\2235\KSB\file.xlsx",
        path_key="sq/2235-KSB/3",
    )
    kept = list(records_in_contour([canonical, loose, sq], rd_root))
    assert canonical in kept
    assert sq in kept
    assert loose not in kept
    reused = list(
        records_in_contour(
            [canonical, loose, sq],
            rd_root,
            previous_records=[canonical, loose, sq],
            previous_contour=kept,
        )
    )
    assert reused == kept
    moved = replace(
        loose,
        path=(
            rf"{rd_root}\2235\06_KSB\Для передачи\01_рев.01-AN01"
            rf"\PDF\AGCC.287-2235-KSB.OD-0001_01-AN01_RU.pdf"
        ),
    )
    after_move = list(
        records_in_contour(
            [canonical, moved, sq],
            rd_root,
            previous_records=[canonical, loose, sq],
            previous_contour=kept,
        )
    )
    assert canonical in after_move
    assert moved in after_move
    assert sq in after_move


if __name__ == "__main__":
    main()
