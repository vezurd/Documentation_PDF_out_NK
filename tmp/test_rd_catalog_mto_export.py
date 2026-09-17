"""Local checks for RD catalog MTO export rules, pins, and plans."""

from __future__ import annotations

import json
import sys
import tempfile
from dataclasses import replace
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
sys.path.insert(0, str(Path(__file__).resolve().parent))

from rd_catalog.config import CatalogConfig
from rd_catalog.kits import kit_identity_key
from rd_catalog.models import FileRecord, ReviewState, SourceKind
from rd_catalog.mto_export import (
    DEFAULT_EXPORT_RULE,
    DEFAULT_ROBOT_TARGET_NAME,
    MTO_EXPORT_ALGORITHM_VERSION,
    PIN_COLUMN_HEADER,
    PIN_STALE_SUFFIX,
    ExportPin,
    ExportSelection,
    ExportTarget,
    build_export_plan,
    export_pin_evidence_for_kit,
    export_pin_view,
    list_export_pin_candidates,
    load_export_pins,
    load_export_targets,
    pin_evidence,
    pin_is_current,
    reconcile_export_pin,
    remove_export_pin,
    resolve_export_selections,
    save_export_pins,
    save_export_targets,
    scan_export_target,
    target_files_from_records,
    upsert_export_pin,
)
from rd_catalog.pipeline import ingest_google_snapshot, rebuild_pipeline, resolve_approved_contour
from test_rd_catalog_pipeline import (
    _LOADED_AT,
    _event,
    _google,
    _mtime_ns,
    _open_db,
    _record,
    _send,
)

RD_ROOT = r"\\bcc\eng\PrDoc\РД"
_TITLE = "3240"
_MARK = "KSB1"
_KEY = kit_identity_key(_TITLE, _MARK)


@pytest.fixture
def temp(tmp_path: Path) -> Path:
    return tmp_path


def _config(temp: Path, *, robot_flat: bool = False) -> CatalogConfig:
    return CatalogConfig(
        rd_root=Path(RD_ROOT),
        sq_root=temp / "sq",
        robot_root=temp / "robot",
        runtime_dir=temp,
        db_path=temp / "rd_catalog.sqlite",
        robot_flat_structure=robot_flat,
        skip_dirs=(),
    )


def _mto(
    file_id: int,
    *,
    title: str = _TITLE,
    mark: str = _MARK,
    revision: str,
    appendix: str | None,
    sequence: int,
    folder: str,
    mtime_ns: int,
    source: SourceKind = SourceKind.RD,
    as_build: bool = False,
    size: int | None = None,
    present: bool = True,
) -> FileRecord:
    record = _record(
        file_id,
        title=title,
        mark=mark,
        revision=revision,
        appendix=appendix,
        sequence=sequence,
        folder=folder,
        mtime_ns=mtime_ns,
        source=source,
        file_kind="mto_xlsx",
        present=present,
        as_build=as_build,
    )
    data = dict(record.data)
    data["size"] = 1000 + file_id if size is None else size
    return replace(record, data=data)


def _robot_file(
    file_id: int,
    *,
    path: str,
    title: str,
    mark: str,
    revision: str,
    appendix: str | None,
    size: int,
    mtime_ns: int,
    name: str,
) -> FileRecord:
    return FileRecord(
        id=file_id,
        path=path,
        path_key=f"robot/{title}-{mark}/{file_id}",
        source=SourceKind.ROBOT,
        present=True,
        review_state=ReviewState.ACKNOWLEDGED,
        first_seen_run_id=1,
        last_seen_run_id=1,
        data={
            "file_kind": "mto_xlsx",
            "name": name,
            "parse_status": "parsed",
            "title": title,
            "mark": mark,
            "revision": revision,
            "appendix": appendix,
            "mtime_ns": mtime_ns,
            "size": size,
            "title_system": f"{title}-{mark}",
            "discipline_block": "MTO-0001",
        },
    )


def _rebuild(
    database,
    kits,
    sends,
    records,
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
        rd_root=RD_ROOT,
    )


def _3240_records() -> dict[str, FileRecord]:
    mtime = _mtime_ns(2026, 6, 1)
    nn07 = _mto(
        7,
        revision="01",
        appendix="01",
        sequence=7,
        folder="07_рев.01-AN01_AGCC.287-3240-KSB1",
        mtime_ns=mtime,
    )
    nn09 = _mto(
        9,
        revision="01",
        appendix="02",
        sequence=9,
        folder="09_рев.01-AN02_AGCC.287-3240-KSB1",
        mtime_ns=mtime + 1,
    )
    nn11 = _mto(
        11,
        revision="01",
        appendix="03",
        sequence=11,
        folder="11_рев.01-AN03_AGCC.287-3240-KSB1",
        mtime_ns=mtime + 2,
    )
    nn12 = _mto(
        12,
        revision="02",
        appendix=None,
        sequence=12,
        folder="12_рев.02_AGCC.287-3240-KSB1",
        mtime_ns=mtime + 3,
        as_build=True,
    )
    return {"nn07": nn07, "nn09": nn09, "nn11": nn11, "nn12": nn12}


def _3240_google():
    trm_a = "AGCC.287-PGS-PGS-TRM-32407"
    trm_b = "AGCC.287-PGS-PGS-TRM-32409"
    return (
        _google(
            _TITLE,
            _MARK,
            (
                _event(
                    date="01.05.2026",
                    stage="code_a",
                    stage_label="код А",
                    revision="01",
                    appendix="01",
                    transmittals=(trm_a,),
                ),
                _event(
                    date="10.06.2026",
                    stage="tdo_passed",
                    stage_label="прошла ТДО",
                    revision="01",
                    appendix="02",
                    transmittals=(trm_b,),
                ),
            ),
            revision="01",
            appendix="03",
        ),
        (
            _send(
                title=_TITLE,
                mark=_MARK,
                revision="01",
                appendix="01",
                status="Принят",
                send_date="20.04.2026",
                incoming="25.04.2026",
                transmittal=trm_a,
            ),
            _send(
                title=_TITLE,
                mark=_MARK,
                revision="01",
                appendix="02",
                status="Принят",
                send_date="01.06.2026",
                incoming="05.06.2026",
                transmittal=trm_b,
                row_index=2,
            ),
        ),
        trm_a,
        trm_b,
    )


def _load_3240(temp: Path):
    files = _3240_records()
    records = [files["nn07"], files["nn09"], files["nn11"], files["nn12"]]
    google, sends, _trm_a, _trm_b = _3240_google()
    database = _open_db(temp / "3240")
    current_ids = {files["nn12"].id}
    _rebuild(database, (google,), sends, records, current_ids)
    return database, records, current_ids, files


def _target(temp: Path, *, rule: str = "latest_issued") -> ExportTarget:
    return ExportTarget(
        name="test-target",
        root=str(temp / "robot"),
        flat_structure=False,
        is_default_robot=False,
        rule=rule,
        filter_text="",
    )


def _pick(
    database,
    records,
    current_ids,
    *,
    rule: str,
    target: ExportTarget,
    target_files=None,
    pins=(),
    verdicts=None,
):
    rows = resolve_export_selections(
        database,
        records=records,
        detected_current_ids=current_ids,
        rule=rule,
        target=target,
        target_files=target_files or {},
        pins=pins,
        rd_root=RD_ROOT,
        verdicts=verdicts,
    )
    found = [row for row in rows if kit_identity_key(row.title, row.mark) == _KEY]
    assert found, f"missing 3240-KSB1 among {[(item.title, item.mark) for item in rows]}"
    return found[0]


def test_four_rules_pick_different_files(temp: Path) -> None:
    database, records, current_ids, files = _load_3240(temp)
    target = _target(temp)
    approved = _pick(
        database, records, current_ids, rule="approved", target=target
    )
    latest_tdo = _pick(
        database, records, current_ids, rule="latest_tdo", target=target
    )
    latest_issued = _pick(
        database, records, current_ids, rule="latest_issued", target=target
    )
    latest_ifc = _pick(
        database, records, current_ids, rule="latest_no_as_build", target=target
    )
    assert approved.source_path == files["nn07"].path
    assert latest_tdo.source_path == files["nn09"].path
    assert latest_issued.source_path == files["nn12"].path
    assert latest_ifc.source_path == files["nn11"].path
    assert len({approved.source_path, latest_tdo.source_path, latest_issued.source_path, latest_ifc.source_path}) == 4
    contour = resolve_approved_contour(
        database,
        records=records,
        detected_current_ids=current_ids,
        title=_TITLE,
        mark=_MARK,
        rd_root=RD_ROOT,
    )
    assert contour is not None
    assert contour.mto_path == files["nn07"].path
    assert latest_tdo.source_path != approved.source_path


def _one(rows, title: str, mark: str):
    found = [row for row in rows if row.title == title and row.mark == mark]
    assert found, f"missing {title}-{mark}"
    return found[0]


def test_latest_tdo_cycle_link_present(temp: Path) -> None:
    database, records, current_ids, files = _load_3240(temp)
    row = _pick(
        database, records, current_ids, rule="latest_tdo", target=_target(temp)
    )
    assert row.source_path == files["nn09"].path
    assert row.rule_path == files["nn09"].path
    assert row.state == "add"


def test_latest_tdo_folder_rev_without_cycle(temp: Path) -> None:
    """No issuance send → no kit_cycle. Folder ``рев.`` matches the TDO event."""

    database = _open_db(temp / "tdo_folder")
    mto = _mto(
        19,
        title="1900",
        mark="KSB",
        revision="01",
        appendix="02",
        sequence=9,
        folder="09_рев.01-AN02_AGCC.287-1900-KSB",
        mtime_ns=_mtime_ns(2026, 4, 1),
    )
    google = _google(
        "1900",
        "KSB",
        (
            _event(
                date="10.04.2026",
                stage="tdo_passed",
                stage_label="прошла ТДО",
                revision="01",
                appendix="02",
            ),
        ),
        revision="01",
        appendix="02",
    )
    _rebuild(database, (google,), (), [mto], {19})
    rows = resolve_export_selections(
        database,
        records=[mto],
        detected_current_ids={19},
        rule="latest_tdo",
        target=_target(temp),
        target_files={},
        pins=(),
        rd_root=RD_ROOT,
    )
    row = _one(rows, "1900", "KSB")
    assert row.source_path == mto.path
    assert row.state == "add"


def test_latest_tdo_inherits_earlier_mto(temp: Path) -> None:
    """TDO package is PDF-only; overlay as-of that NN inherits the earlier MTO."""

    database = _open_db(temp / "tdo_inherit")
    mtime = _mtime_ns(2026, 5, 1)
    earlier = _mto(
        5,
        title="1715",
        mark="POS",
        revision="03",
        appendix=None,
        sequence=5,
        folder="05_рев.03_AGCC.287-1715-POS",
        mtime_ns=mtime,
    )
    tdo_pdf = _record(
        6,
        title="1715",
        mark="POS",
        revision="04",
        appendix=None,
        sequence=6,
        folder="06_рев.04_AGCC.287-1715-POS",
        mtime_ns=mtime + 1,
    )
    records = [earlier, tdo_pdf]
    trm = "AGCC.287-PGS-PGS-TRM-17151"
    google = _google(
        "1715",
        "POS",
        (
            _event(
                date="01.05.2026",
                stage="tdo_passed",
                stage_label="прошла ТДО",
                revision="04",
                transmittals=(trm,),
            ),
        ),
        revision="04",
        appendix=None,
    )
    send = _send(
        title="1715",
        mark="POS",
        revision="04",
        appendix=None,
        status="Принят",
        send_date="20.04.2026",
        incoming="25.04.2026",
        transmittal=trm,
    )
    _rebuild(database, (google,), (send,), records, {6})
    rows = resolve_export_selections(
        database,
        records=records,
        detected_current_ids={6},
        rule="latest_tdo",
        target=_target(temp),
        target_files={},
        pins=(),
        rd_root=RD_ROOT,
    )
    row = _one(rows, "1715", "POS")
    assert row.source_path == earlier.path
    assert row.state == "add"


def test_latest_tdo_later_mto_after_tdo_package(temp: Path) -> None:
    """TDO cycle binds an early PDF-only NN; MTO lives only in a later transfer.

    Overlay as-of the TDO sequence is empty. The approved last_package /
    file_rev chain must still find the later MTO.
    """

    database = _open_db(temp / "tdo_later")
    mtime = _mtime_ns(2026, 3, 1)
    tdo_pdf = _record(
        3,
        title="1800",
        mark="KSB",
        revision="01",
        appendix=None,
        sequence=3,
        folder="03_рев.01_AGCC.287-1800-KSB",
        mtime_ns=mtime,
    )
    later_mto = _mto(
        8,
        title="1800",
        mark="KSB",
        revision="02",
        appendix=None,
        sequence=8,
        folder="08_рев.02_AGCC.287-1800-KSB",
        mtime_ns=mtime + 1,
    )
    records = [tdo_pdf, later_mto]
    trm = "AGCC.287-PGS-PGS-TRM-18003"
    google = _google(
        "1800",
        "KSB",
        (
            _event(
                date="01.03.2026",
                stage="tdo_passed",
                stage_label="прошла ТДО",
                revision="01",
                transmittals=(trm,),
            ),
        ),
        revision="01",
        appendix=None,
    )
    send = _send(
        title="1800",
        mark="KSB",
        revision="01",
        appendix=None,
        status="Принят",
        send_date="20.02.2026",
        incoming="25.02.2026",
        transmittal=trm,
    )
    _rebuild(database, (google,), (send,), records, {8})
    rows = resolve_export_selections(
        database,
        records=records,
        detected_current_ids={8},
        rule="latest_tdo",
        target=_target(temp),
        target_files={},
        pins=(),
        rd_root=RD_ROOT,
    )
    row = _one(rows, "1800", "KSB")
    assert row.source_path == later_mto.path
    assert row.state == "add"


def test_latest_tdo_without_passed_event_stays_no_source(temp: Path) -> None:
    """RD MTO is present but F never passed ТДО → latest_tdo is no_source."""

    database = _open_db(temp / "tdo_none")
    mto = _mto(
        21,
        title="2100",
        mark="KSB",
        revision="01",
        appendix=None,
        sequence=1,
        folder="01_рев.01_AGCC.287-2100-KSB",
        mtime_ns=_mtime_ns(2026, 1, 1),
    )
    google = _google(
        "2100",
        "KSB",
        (
            _event(
                date="01.01.2026",
                stage="tdo_sent",
                stage_label="отпр на ТДО",
                revision="01",
            ),
        ),
        revision="01",
        appendix=None,
    )
    _rebuild(database, (google,), (), [mto], {21})
    rows = resolve_export_selections(
        database,
        records=[mto],
        detected_current_ids={21},
        rule="latest_tdo",
        target=_target(temp),
        target_files={},
        pins=(),
        rd_root=RD_ROOT,
    )
    tdo_row = _one(rows, "2100", "KSB")
    assert tdo_row.state == "no_source"
    issued = resolve_export_selections(
        database,
        records=[mto],
        detected_current_ids={21},
        rule="latest_issued",
        target=_target(temp),
        target_files={},
        pins=(),
        rd_root=RD_ROOT,
    )
    assert _one(issued, "2100", "KSB").source_path == mto.path


def test_google_only_kit_is_no_source(temp: Path) -> None:
    database = _open_db(temp / "google_only")
    google = _google(
        "8888",
        "KSB",
        (
            _event(
                date="01.01.2026",
                stage="sr_upload",
                stage_label="загрузка в СР",
                revision="01",
            ),
        ),
        revision="01",
        appendix=None,
    )
    _rebuild(database, (google,), (), [], set())
    rows = resolve_export_selections(
        database,
        records=[],
        detected_current_ids=set(),
        rule="latest_issued",
        target=_target(temp),
        target_files={},
        pins=(),
        rd_root=RD_ROOT,
    )
    found = [row for row in rows if row.title == "8888" and row.mark == "KSB"]
    assert len(found) == 1
    assert found[0].state == "no_source"
    assert found[0].source_path == ""
    assert found[0].rule_path == ""
    assert found[0].destination_path == ""
    assert found[0].origin == "rule"


def test_states_add_replace_same(temp: Path) -> None:
    database, records, current_ids, files = _load_3240(temp)
    target = _target(temp)
    source = files["nn12"]
    dest_dir = Path(target.root) / _TITLE / _MARK
    dest_same = str(dest_dir / Path(source.path).name)
    dest_other = str(dest_dir / "AGCC.287-3240-KSB1.MTO-0001_01_RU.xlsx")
    add_row = _pick(
        database, records, current_ids, rule="latest_issued", target=target
    )
    assert add_row.state == "add"
    assert add_row.existing_target_path == ""
    robot_same = _robot_file(
        90,
        path=dest_same,
        title=_TITLE,
        mark=_MARK,
        revision="02",
        appendix=None,
        size=int(source.data["size"]),
        mtime_ns=int(source.data["mtime_ns"]),
        name=Path(source.path).name,
    )
    same_row = _pick(
        database,
        records + [robot_same],
        current_ids,
        rule="latest_issued",
        target=target,
        target_files={_KEY: dest_same},
    )
    assert same_row.state == "unknown"
    assert same_row.existing_target_path == dest_same
    robot_other = _robot_file(
        91,
        path=dest_other,
        title=_TITLE,
        mark=_MARK,
        revision="01",
        appendix=None,
        size=50,
        mtime_ns=1,
        name=Path(dest_other).name,
    )
    replace_row = _pick(
        database,
        records + [robot_other],
        current_ids,
        rule="latest_issued",
        target=target,
        target_files={_KEY: dest_other},
    )
    assert replace_row.state == "unknown"
    assert replace_row.existing_target_path == dest_other
    robot_same_name_diff = _robot_file(
        92,
        path=dest_same,
        title=_TITLE,
        mark=_MARK,
        revision="02",
        appendix=None,
        size=int(source.data["size"]) + 100000,
        mtime_ns=int(source.data["mtime_ns"]),
        name=Path(source.path).name,
    )
    content_row = _pick(
        database,
        records + [robot_same_name_diff],
        current_ids,
        rule="latest_issued",
        target=target,
        target_files={_KEY: dest_same},
    )
    assert content_row.state == "unknown"
    assert content_row.existing_target_path == dest_same
    blob = " ".join(content_row.warnings)
    assert "другим содержимым" in blob
    assert str(int(source.data["size"])) in blob
    assert str(int(robot_same_name_diff.data["size"])) in blob


def test_pin_honoured_refresh_and_stale(temp: Path) -> None:
    database, records, current_ids, files = _load_3240(temp)
    target = _target(temp)
    events = database.list_google_events(_TITLE, _MARK)
    mto_files = [
        record
        for record in records
        if record.source is SourceKind.RD
    ]
    evidence = pin_evidence(events, mto_files)
    pin = ExportPin(
        title=_TITLE,
        mark=_MARK,
        package_path=str(Path(files["nn07"].path).parent.parent),
        file_path=files["nn07"].path,
        revision_text="01-AN01",
        evidence=evidence,
        created_at="2026-09-10T00:00:00+00:00",
    )
    honoured = _pick(
        database,
        records,
        current_ids,
        rule="latest_issued",
        target=target,
        pins=(pin,),
    )
    assert honoured.origin == "pin"
    assert honoured.source_path == files["nn07"].path
    assert honoured.rule_path == files["nn12"].path
    assert honoured.state == "add"

    current_pin = ExportPin(
        title=_TITLE,
        mark=_MARK,
        package_path="",
        file_path=files["nn12"].path,
        revision_text="02",
        evidence=evidence,
        created_at="2026-09-10T00:00:00+00:00",
    )
    bumped = replace(
        files["nn07"],
        data={**files["nn07"].data, "mtime_ns": int(files["nn07"].data["mtime_ns"]) + 99},
    )
    bumped_records = [bumped if row.id == 7 else row for row in records]
    refreshed = _pick(
        database,
        bumped_records,
        current_ids,
        rule="latest_issued",
        target=target,
        pins=(current_pin,),
    )
    assert refreshed.origin == "pin"
    assert refreshed.source_path == files["nn12"].path
    assert not pin_is_current(
        current_pin, pin_evidence(events, [row for row in bumped_records if row.source is SourceKind.RD])
    )

    stale = _pick(
        database,
        bumped_records,
        current_ids,
        rule="latest_issued",
        target=target,
        pins=(pin,),
    )
    assert stale.origin == "pin_stale"
    assert stale.state == "pin_stale"
    assert stale.source_path == files["nn07"].path
    assert stale.rule_path == files["nn12"].path
    assert stale.source_path != stale.rule_path
    assert files["nn12"].path in " ".join(stale.warnings)
    assert files["nn07"].path in " ".join(stale.warnings)


def test_raw_journal_text_does_not_invalidate_pin() -> None:
    shared = dict(
        date="01.05.2026",
        stage="code_a",
        stage_label="код А",
        revision="01",
        appendix="01",
        transmittals=("TRM-1",),
    )
    first = _event(**shared, raw="01.05.2026 код А рев.01-AN01 TRM-1  typo")
    second = _event(**shared, raw="01.05.2026 код А рев.01-AN01 TRM-1 FIXED")
    assert first.raw != second.raw
    assert pin_evidence((first,), ()) == pin_evidence((second,), ())
    digest = pin_evidence((first,), ())
    assert len(digest) == 64
    assert MTO_EXPORT_ALGORITHM_VERSION == 1


def test_latest_issued_falls_back_to_sq(temp: Path) -> None:
    database = _open_db(temp / "sq_only")
    sq = _mto(
        50,
        title="5000",
        mark="KSB",
        revision="03",
        appendix=None,
        sequence=1,
        folder="01_рев.03_AGCC.287-5000-KSB",
        mtime_ns=_mtime_ns(2026, 7, 1),
        source=SourceKind.SQ,
    )
    google = _google(
        "5000",
        "KSB",
        (
            _event(
                date="01.07.2026",
                stage="sr_upload",
                stage_label="загрузка в СР",
                revision="03",
            ),
        ),
        revision="03",
        appendix=None,
    )
    _rebuild(database, (google,), (), [sq], set())
    rows = resolve_export_selections(
        database,
        records=[sq],
        detected_current_ids=set(),
        rule="latest_issued",
        target=_target(temp),
        target_files={},
        pins=(),
        rd_root=RD_ROOT,
    )
    found = [row for row in rows if row.title == "5000" and row.mark == "KSB"]
    assert found and found[0].source_path == sq.path
    assert found[0].state == "add"


def test_default_robot_target_synthesized_and_not_deleted(temp: Path) -> None:
    config = _config(temp)
    loaded = load_export_targets(config)
    assert len(loaded) == 1
    assert loaded[0].is_default_robot
    assert loaded[0].name == DEFAULT_ROBOT_TARGET_NAME
    assert loaded[0].root == str(config.robot_root)
    assert loaded[0].rule == DEFAULT_EXPORT_RULE
    assert not (Path(config.runtime_dir) / "mto_export_targets.json").exists()

    custom = ExportTarget(
        name="Архив",
        root=str(temp / "archive"),
        flat_structure=True,
        is_default_robot=False,
        rule="approved",
        filter_text="KSB",
    )
    saved = save_export_targets(config, (custom,))
    assert saved[0].is_default_robot
    assert saved[0].root == str(config.robot_root)
    assert saved[0].flat_structure == config.robot_flat_structure
    assert saved[1].name == "Архив"

    hijack = replace(
        saved[0],
        root=r"C:\not-the-robot",
        flat_structure=True,
        rule="approved",
        filter_text="x",
    )
    saved_again = save_export_targets(config, (hijack, saved[1]))
    assert saved_again[0].root == str(config.robot_root)
    assert saved_again[0].flat_structure is False
    assert saved_again[0].rule == "approved"
    assert saved_again[0].filter_text == "x"

    corrupt = Path(config.runtime_dir) / "mto_export_targets.json"
    corrupt.write_text("{not json", encoding="utf-8")
    healed = load_export_targets(config)
    assert len(healed) == 1 and healed[0].is_default_robot


def test_build_export_plan_raises_on_escape(temp: Path) -> None:
    target = _target(temp)
    escaped = ExportSelection(
        title=_TITLE,
        mark=_MARK,
        rule="latest_issued",
        source_path=r"C:\rd\file.xlsx",
        source_revision_text="02",
        package_path="",
        origin="rule",
        state="add",
        destination_path=r"C:\Windows\file.xlsx",
        existing_target_path="",
        confidence="high",
        warnings=(),
    )
    with pytest.raises(ValueError, match="escapes target root"):
        build_export_plan((escaped,), target=target)

    database, records, current_ids, files = _load_3240(temp)
    row = _pick(
        database, records, current_ids, rule="latest_issued", target=target
    )
    plan = build_export_plan((row,), target=target)
    assert len(plan.items) == 1
    assert plan.items[0].selection.state == "add"
    same = replace(
        row,
        state="same_data",
        existing_target_path=row.destination_path,
        destination_path=row.destination_path,
    )
    skipped = build_export_plan((same,), target=target)
    assert plan.items[0].archive_dir == ""
    assert skipped.items == ()
    assert skipped.skipped[0].state == "same_data"


def test_verdicts_same_data_replace_unknown(temp: Path) -> None:
    database, records, current_ids, files = _load_3240(temp)
    target = _target(temp)
    source = files["nn12"]
    dest_dir = Path(target.root) / _TITLE / _MARK
    dest_same = str(dest_dir / Path(source.path).name)
    dest_other = str(dest_dir / "AGCC.287-3240-KSB1.MTO-0001_01_RU.xlsx")
    robot_same = _robot_file(
        90,
        path=dest_same,
        title=_TITLE,
        mark=_MARK,
        revision="02",
        appendix=None,
        size=int(source.data["size"]),
        mtime_ns=int(source.data["mtime_ns"]),
        name=Path(source.path).name,
    )
    none_row = _pick(
        database,
        records + [robot_same],
        current_ids,
        rule="latest_issued",
        target=target,
        target_files={_KEY: dest_same},
        verdicts=None,
    )
    assert none_row.state == "unknown"
    empty_map = _pick(
        database,
        records + [robot_same],
        current_ids,
        rule="latest_issued",
        target=target,
        target_files={_KEY: dest_same},
        verdicts={},
    )
    assert empty_map.state == "unknown"
    equal_row = _pick(
        database,
        records + [robot_same],
        current_ids,
        rule="latest_issued",
        target=target,
        target_files={_KEY: dest_same},
        verdicts={(source.id, 90): "content_equal"},
    )
    assert equal_row.state == "same_data"
    equal_plan = build_export_plan((equal_row,), target=target)
    assert equal_plan.items == ()
    assert equal_plan.skipped[0].state == "same_data"
    swapped_key = _pick(
        database,
        records + [robot_same],
        current_ids,
        rule="latest_issued",
        target=target,
        target_files={_KEY: dest_same},
        verdicts={(90, source.id): "content_equal"},
    )
    assert swapped_key.state == "same_data"
    diff_row = _pick(
        database,
        records + [robot_same],
        current_ids,
        rule="latest_issued",
        target=target,
        target_files={_KEY: dest_same},
        verdicts={(source.id, 90): "content_diff"},
    )
    assert diff_row.state == "replace"
    diff_plan = build_export_plan((diff_row,), target=target)
    assert len(diff_plan.items) == 1
    unknown_plan = build_export_plan((none_row,), target=target)
    assert unknown_plan.items == ()
    assert unknown_plan.skipped[0].state == "unknown"
    robot_other = _robot_file(
        91,
        path=dest_other,
        title=_TITLE,
        mark=_MARK,
        revision="01",
        appendix=None,
        size=50,
        mtime_ns=1,
        name=Path(dest_other).name,
    )
    name_diff = _pick(
        database,
        records + [robot_other],
        current_ids,
        rule="latest_issued",
        target=target,
        target_files={_KEY: dest_other},
        verdicts=None,
    )
    assert name_diff.state == "unknown"
    name_diff_verdict = _pick(
        database,
        records + [robot_other],
        current_ids,
        rule="latest_issued",
        target=target,
        target_files={_KEY: dest_other},
        verdicts={(source.id, 91): "content_diff"},
    )
    assert name_diff_verdict.state == "replace"


def test_scan_export_target_missing_and_present(temp: Path) -> None:
    missing = ExportTarget(
        name="gone",
        root=str(temp / "does-not-exist"),
        flat_structure=False,
        is_default_robot=False,
        rule="latest_issued",
        filter_text="",
    )
    assert scan_export_target(missing) == {}
    folder = temp / "robot" / _TITLE / _MARK
    folder.mkdir(parents=True)
    name = "AGCC.287-3240-KSB1.MTO-0001_02_RU.xlsx"
    (folder / name).write_bytes(b"fixture")
    found = scan_export_target(_target(temp))
    assert found[_KEY].endswith(name)
    robot_record = _robot_file(
        5,
        path=found[_KEY],
        title=_TITLE,
        mark=_MARK,
        revision="02",
        appendix=None,
        size=7,
        mtime_ns=1,
        name=name,
    )
    from_records = target_files_from_records([robot_record])
    assert from_records[_KEY] == found[_KEY]


def test_pins_roundtrip(temp: Path) -> None:
    config = _config(temp)
    assert load_export_pins(config) == ()
    pin = ExportPin(
        title=_TITLE,
        mark=_MARK,
        package_path="pkg",
        file_path=r"C:\a.xlsx",
        revision_text="01",
        evidence="abc",
        created_at="2026-09-10T00:00:00+00:00",
    )
    saved = save_export_pins(config, (pin,))
    assert saved == (pin,)
    reloaded = load_export_pins(config)
    assert reloaded == (pin,)
    payload = json.loads(
        (Path(config.runtime_dir) / "mto_export_pins.json").read_text(encoding="utf-8")
    )
    assert payload["version"] == 1
    corrupt = Path(config.runtime_dir) / "mto_export_pins.json"
    corrupt.write_text("{nope", encoding="utf-8")
    assert load_export_pins(config) == ()


def test_pin_candidates_and_view(temp: Path) -> None:
    database, records, current_ids, files = _load_3240(temp)
    target = _target(temp)
    rule_row = _pick(
        database,
        records,
        current_ids,
        rule="latest_issued",
        target=target,
    )
    candidates = list_export_pin_candidates(
        records,
        title=_TITLE,
        mark=_MARK,
        rd_root=RD_ROOT,
        rule_path=rule_row.rule_path,
    )
    assert [item.sequence for item in candidates] == [12, 11, 9, 7]
    assert candidates[0].file_path == files["nn12"].path
    assert candidates[0].is_rule_choice
    assert candidates[0].size == files["nn12"].data["size"]
    assert candidates[0].package_name == files["nn12"].data["transfer_name"]
    assert candidates[-1].file_path == files["nn07"].path
    assert not candidates[-1].is_rule_choice
    empty = export_pin_view(None)
    assert empty.text == ""
    assert empty.tooltip == ""
    assert not empty.is_stale
    pin = ExportPin(
        title=_TITLE,
        mark=_MARK,
        package_path=candidates[-1].package_path,
        file_path=candidates[-1].file_path,
        revision_text=candidates[-1].revision_text,
        evidence="abc",
        created_at="2026-09-10T00:00:00+00:00",
    )
    current = export_pin_view(pin, origin="pin", rule_path=rule_row.rule_path)
    assert current.text == Path(pin.file_path).name
    assert PIN_STALE_SUFFIX not in current.text
    assert pin.package_path in current.tooltip
    assert pin.file_path in current.tooltip
    assert pin.created_at in current.tooltip
    assert "файл по правилу" not in current.tooltip
    stale = export_pin_view(pin, origin="pin_stale", rule_path=rule_row.rule_path)
    assert stale.text.endswith(PIN_STALE_SUFFIX.strip())
    assert Path(pin.file_path).name in stale.text
    assert stale.is_stale
    assert rule_row.rule_path in stale.tooltip
    assert PIN_COLUMN_HEADER == "Ручной выбор MTO"


def test_silent_refresh_persists_via_sink(temp: Path) -> None:
    database, records, current_ids, files = _load_3240(temp)
    target = _target(temp)
    pin = ExportPin(
        title=_TITLE,
        mark=_MARK,
        package_path="",
        file_path=files["nn12"].path,
        revision_text="02",
        evidence="deadbeef",
        created_at="2026-09-01T00:00:00+00:00",
    )
    refreshed: list[ExportPin] = []
    rows = resolve_export_selections(
        database,
        records=records,
        detected_current_ids=current_ids,
        rule="latest_issued",
        target=target,
        target_files={},
        pins=(pin,),
        rd_root=RD_ROOT,
        refreshed_pins=refreshed,
    )
    found = [row for row in rows if row.title == _TITLE and row.mark == _MARK]
    assert found and found[0].origin == "pin"
    assert found[0].state != "pin_stale"
    assert refreshed and refreshed[0].evidence != "deadbeef"
    expected = export_pin_evidence_for_kit(
        database,
        records,
        title=_TITLE,
        mark=_MARK,
        rd_root=RD_ROOT,
    )
    assert refreshed[0].evidence == expected
    origin, same = reconcile_export_pin(
        pin, evidence="other-digest", rule_path=files["nn07"].path
    )
    assert origin == "pin_stale"
    assert same is pin


def test_upsert_and_remove_pin(temp: Path) -> None:
    config = _config(temp)
    first = ExportPin(
        title=_TITLE,
        mark=_MARK,
        package_path="a",
        file_path=r"C:\a.xlsx",
        revision_text="01",
        evidence="one",
        created_at="2026-09-01T00:00:00+00:00",
    )
    second = ExportPin(
        title=_TITLE,
        mark=_MARK,
        package_path="b",
        file_path=r"C:\b.xlsx",
        revision_text="02",
        evidence="two",
        created_at="2026-09-02T00:00:00+00:00",
    )
    other = ExportPin(
        title="1111",
        mark="KSB",
        package_path="c",
        file_path=r"C:\c.xlsx",
        revision_text="01",
        evidence="oth",
        created_at="2026-09-03T00:00:00+00:00",
    )
    upsert_export_pin(config, first)
    upsert_export_pin(config, other)
    saved = upsert_export_pin(config, second)
    by_key = {kit_identity_key(item.title, item.mark): item for item in saved}
    assert by_key[_KEY].file_path == r"C:\b.xlsx"
    assert kit_identity_key("1111", "KSB") in by_key
    remaining = remove_export_pin(config, _TITLE, _MARK)
    left = {kit_identity_key(item.title, item.mark) for item in remaining}
    assert _KEY not in left
    assert kit_identity_key("1111", "KSB") in left


def main() -> None:
    """Run MTO export fixtures against a temporary directory."""

    test_raw_journal_text_does_not_invalidate_pin()
    with tempfile.TemporaryDirectory(prefix="rd_catalog_mto_export_") as raw:
        root = Path(raw)
        test_four_rules_pick_different_files(root / "rules")
        test_latest_tdo_cycle_link_present(root / "tdo_cycle")
        test_latest_tdo_folder_rev_without_cycle(root / "tdo_folder")
        test_latest_tdo_inherits_earlier_mto(root / "tdo_inherit")
        test_latest_tdo_later_mto_after_tdo_package(root / "tdo_later")
        test_latest_tdo_without_passed_event_stays_no_source(root / "tdo_none")
        test_google_only_kit_is_no_source(root / "google")
        test_states_add_replace_same(root / "states")
        test_verdicts_same_data_replace_unknown(root / "verdicts")
        test_pin_honoured_refresh_and_stale(root / "pins")
        test_latest_issued_falls_back_to_sq(root / "sq")
        test_default_robot_target_synthesized_and_not_deleted(root / "targets")
        test_build_export_plan_raises_on_escape(root / "plan")
        test_scan_export_target_missing_and_present(root / "scan")
        test_pins_roundtrip(root / "pins_io")
        test_pin_candidates_and_view(root / "pin_view")
        test_silent_refresh_persists_via_sink(root / "pin_sink")
        test_upsert_and_remove_pin(root / "pin_upsert")
    print("RD catalog MTO export: OK")


if __name__ == "__main__":
    main()
