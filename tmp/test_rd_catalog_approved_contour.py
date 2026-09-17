"""Local checks for approved-contour derivation and bounded overlay."""

from __future__ import annotations

import json
import sys
import tempfile
from dataclasses import replace
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
sys.path.insert(0, str(Path(__file__).resolve().parent))

from rd_catalog.models import CollisionKind, FileKind, FileRecord, SourceKind
from rd_catalog.overlay import build_rd_overlays, build_rd_overlays_as_of
from rd_catalog.parse import parse_catalog_file, parse_transfer_folder
from rd_catalog.pipeline import (
    resolve_all_approved_contours,
    resolve_approved_contour,
)
from test_rd_catalog_pipeline import (
    _event,
    _google,
    _mtime_ns,
    _open_db,
    _rebuild,
    _record,
    _send,
)


@pytest.fixture
def temp(tmp_path: Path) -> Path:
    return tmp_path


def _with_folder_rev(
    record: FileRecord, revision: str | None, appendix: str | None = None
) -> FileRecord:
    data = dict(record.data)
    data["transfer_revision"] = revision
    data["transfer_appendix"] = appendix
    return replace(record, data=data)


def _insert_overlay_mto_collision(
    database,
    path_keys: list[str],
    kind: CollisionKind,
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
                "fixture overlay mto collision",
                "mto:fixture|mto-0001",
                json.dumps(path_keys, ensure_ascii=False),
                run_id,
            ),
        )


def _resolve(database, records, current_ids, title: str, mark: str):
    return resolve_approved_contour(
        database,
        records=records,
        detected_current_ids=current_ids,
        title=title,
        mark=mark,
    )


def test_1600_sot_later_package_with_older_files_does_not_win(temp: Path) -> None:
    database = _open_db(temp / "sot1600")
    mtime = _mtime_ns(2026, 6, 1)
    nn19 = _with_folder_rev(
        _record(
            19,
            title="1600",
            mark="SOT",
            revision="03",
            appendix=None,
            sequence=19,
            folder="19_рев.04_AGCC.287-1600-SOT",
            mtime_ns=mtime,
            file_kind="mto_xlsx",
        ),
        "04",
    )
    nn20 = _with_folder_rev(
        _record(
            20,
            title="1600",
            mark="SOT",
            revision="03",
            appendix=None,
            sequence=20,
            folder="20_рев.04 по протоколу",
            mtime_ns=mtime + 1,
            file_kind="mto_xlsx",
        ),
        "04",
    )
    nn21 = _with_folder_rev(
        _record(
            21,
            title="1600",
            mark="SOT",
            revision="02",
            appendix=None,
            sequence=21,
            folder="21_рев.03_AGCC.287-1600-SOT",
            mtime_ns=mtime + 2,
            file_kind="mto_xlsx",
        ),
        "03",
    )
    records = [nn19, nn20, nn21]
    trm = "AGCC.287-PGS-PGS-TRM-16001"
    _rebuild(
        database,
        (
            _google(
                "1600",
                "SOT",
                (
                    _event(
                        date="01.06.2026",
                        stage="code_a",
                        stage_label="код А",
                        revision="04",
                        transmittals=(trm,),
                    ),
                ),
                revision="04",
                appendix=None,
            ),
        ),
        (
            _send(
                title="1600",
                mark="SOT",
                revision="03",
                appendix=None,
                status="Принят",
                send_date="20.05.2026",
                incoming="25.05.2026",
                transmittal=trm,
            ),
        ),
        records,
        {21},
    )
    contour = _resolve(database, records, {21}, "1600", "SOT")
    assert contour is not None
    assert contour.match_reason == "cycle_trm"
    assert contour.package_sequence == 20
    assert contour.mto_path == nn20.path
    assert nn21.path not in contour.mto_path
    assert contour.mto_source == "package"
    assert "no_mto_in_contour" not in contour.ambiguity


def test_6400_sot_filename_lags_approved_revision(temp: Path) -> None:
    database = _open_db(temp / "sot6400")
    mtime = _mtime_ns(2026, 7, 1)
    nn09 = _with_folder_rev(
        _record(
            9,
            title="6400",
            mark="SOT",
            revision="05",
            appendix=None,
            sequence=9,
            folder="09_рев.06_AGCC.287-6400-SOT",
            mtime_ns=mtime,
            file_kind="mto_xlsx",
        ),
        "06",
    )
    nn10 = _with_folder_rev(
        _record(
            10,
            title="6400",
            mark="SOT",
            revision="05",
            appendix="01",
            sequence=10,
            folder="10_рев.06-AN01_AGCC.287-6400-SOT",
            mtime_ns=mtime + 1,
            file_kind="mto_xlsx",
        ),
        "06",
        "01",
    )
    records = [nn09, nn10]
    trm = "AGCC.287-PGS-PGS-TRM-64001"
    _rebuild(
        database,
        (
            _google(
                "6400",
                "SOT",
                (
                    _event(
                        date="10.07.2026",
                        stage="code_a",
                        stage_label="код А",
                        revision="06",
                        appendix="01",
                        transmittals=(trm,),
                    ),
                ),
                revision="06",
                appendix="01",
            ),
        ),
        (
            _send(
                title="6400",
                mark="SOT",
                revision="05",
                appendix="01",
                status="Принят",
                send_date="01.07.2026",
                incoming="05.07.2026",
                transmittal=trm,
            ),
        ),
        records,
        {10},
    )
    contour = _resolve(database, records, {10}, "6400", "SOT")
    assert contour is not None
    assert contour.mto_path == nn10.path
    assert contour.mto_source == "package"
    assert "no_mto_in_contour" not in contour.ambiguity
    assert contour.approved_revision_text == "06-AN01"


def test_inherited_mto_from_earlier_package(temp: Path) -> None:
    database = _open_db(temp / "inherited")
    mtime = _mtime_ns(2026, 5, 1)
    earlier = _record(
        5,
        title="1710",
        mark="POS",
        revision="03",
        appendix=None,
        sequence=5,
        folder="05_рев.03_AGCC.287-1710-POS",
        mtime_ns=mtime,
        file_kind="mto_xlsx",
    )
    approved_pdf = _record(
        6,
        title="1710",
        mark="POS",
        revision="04",
        appendix=None,
        sequence=6,
        folder="06_рев.04_AGCC.287-1710-POS",
        mtime_ns=mtime + 1,
    )
    records = [earlier, approved_pdf]
    trm = "AGCC.287-PGS-PGS-TRM-17101"
    _rebuild(
        database,
        (
            _google(
                "1710",
                "POS",
                (
                    _event(
                        date="01.05.2026",
                        stage="code_a",
                        stage_label="код А",
                        revision="04",
                        transmittals=(trm,),
                    ),
                ),
                revision="04",
                appendix=None,
            ),
        ),
        (
            _send(
                title="1710",
                mark="POS",
                revision="04",
                appendix=None,
                status="Принят",
                send_date="20.04.2026",
                incoming="25.04.2026",
                transmittal=trm,
            ),
        ),
        records,
        {6},
    )
    contour = _resolve(database, records, {6}, "1710", "POS")
    assert contour is not None
    assert contour.package_sequence == 6
    assert contour.mto_path == earlier.path
    assert contour.mto_source == "inherited"
    assert contour.confidence == "medium"
    assert "no_mto_in_contour" not in contour.ambiguity


def test_exact_full_revision_wins_over_base_only(temp: Path) -> None:
    database = _open_db(temp / "exact_over_base")
    mtime = _mtime_ns(2026, 8, 1)
    nn09 = _with_folder_rev(
        _record(
            9,
            title="6100",
            mark="SOS",
            revision="04",
            appendix=None,
            sequence=9,
            folder="09_рев.04_AGCC.287-6100-SOS",
            mtime_ns=mtime,
            file_kind="mto_xlsx",
        ),
        "04",
    )
    nn10 = _with_folder_rev(
        _record(
            10,
            title="6100",
            mark="SOS",
            revision="04",
            appendix="01",
            sequence=10,
            folder="10_рев.AN01_AGCC.287-6100-SOS",
            mtime_ns=mtime + 1,
            file_kind="mto_xlsx",
        ),
        None,
    )
    records = [nn09, nn10]
    trm = "AGCC.287-PGS-PGS-TRM-61001"
    _rebuild(
        database,
        (
            _google(
                "6100",
                "SOS",
                (
                    _event(
                        date="10.08.2026",
                        stage="code_a",
                        stage_label="код А",
                        revision="04",
                        appendix="01",
                        transmittals=(trm,),
                    ),
                ),
                revision="04",
                appendix="01",
            ),
        ),
        (
            _send(
                title="6100",
                mark="SOS",
                revision="04",
                appendix=None,
                status="Принят",
                send_date="01.08.2026",
                incoming="05.08.2026",
                transmittal=trm,
            ),
        ),
        records,
        {10},
    )
    contour = _resolve(database, records, {10}, "6100", "SOS")
    assert contour is not None
    assert contour.approved_revision_text == "04-AN01"
    assert contour.package_sequence == 10
    assert contour.mto_path == nn10.path
    assert contour.match_reason.startswith("cycle_")
    assert contour.confidence == "high"
    assert not any("Согласована ревизия" in note for note in contour.warnings)


def test_base_only_fallback_lowers_confidence(temp: Path) -> None:
    database = _open_db(temp / "base_only")
    nn09 = _with_folder_rev(
        _record(
            9,
            title="6110",
            mark="SOS",
            revision="04",
            appendix=None,
            sequence=9,
            folder="09_рев.04_AGCC.287-6110-SOS",
            mtime_ns=_mtime_ns(2026, 8, 1),
            file_kind="mto_xlsx",
        ),
        "04",
    )
    records = [nn09]
    trm = "AGCC.287-PGS-PGS-TRM-61101"
    _rebuild(
        database,
        (
            _google(
                "6110",
                "SOS",
                (
                    _event(
                        date="10.08.2026",
                        stage="code_a",
                        stage_label="код А",
                        revision="04",
                        appendix="01",
                        transmittals=(trm,),
                    ),
                ),
                revision="04",
                appendix="01",
            ),
        ),
        (
            _send(
                title="6110",
                mark="SOS",
                revision="04",
                appendix=None,
                status="Принят",
                send_date="01.08.2026",
                incoming="05.08.2026",
                transmittal=trm,
            ),
        ),
        records,
        {9},
    )
    contour = _resolve(database, records, {9}, "6110", "SOS")
    assert contour is not None
    assert contour.package_sequence == 9
    assert contour.mto_path == nn09.path
    assert contour.confidence == "medium"
    assert any(
        "Согласована ревизия 04-AN01, а выбрана папка с ревизией 04." in note
        for note in contour.warnings
    )


def test_folder_an01_matches_approved_04_an01(temp: Path) -> None:
    database = _open_db(temp / "an01_folder")
    record = _with_folder_rev(
        _record(
            10,
            title="6120",
            mark="SOS",
            revision="03",
            appendix=None,
            sequence=10,
            folder="10_рев.AN01_AGCC.287-6120-SOS",
            mtime_ns=_mtime_ns(2026, 8, 1),
            file_kind="mto_xlsx",
        ),
        None,
    )
    records = [record]
    _rebuild(
        database,
        (
            _google(
                "6120",
                "SOS",
                (
                    _event(
                        date="01.08.2026",
                        stage="code_a",
                        stage_label="код А",
                        revision="04",
                        appendix="01",
                    ),
                ),
                revision="04",
                appendix="01",
            ),
        ),
        (),
        records,
        {10},
    )
    contour = _resolve(database, records, {10}, "6120", "SOS")
    assert contour is not None
    assert contour.package_sequence == 10
    assert contour.approved_revision_text == "04-AN01"
    assert contour.match_reason == "folder_rev"
    assert contour.mto_path == record.path


def test_exact_match_keeps_high_confidence(temp: Path) -> None:
    database = _open_db(temp / "exact_high")
    record = _with_folder_rev(
        _record(
            10,
            title="6130",
            mark="SOS",
            revision="04",
            appendix="01",
            sequence=10,
            folder="10_рев.04-AN01_AGCC.287-6130-SOS",
            mtime_ns=_mtime_ns(2026, 8, 1),
            file_kind="mto_xlsx",
        ),
        "04",
        "01",
    )
    records = [record]
    trm = "AGCC.287-PGS-PGS-TRM-61301"
    _rebuild(
        database,
        (
            _google(
                "6130",
                "SOS",
                (
                    _event(
                        date="10.08.2026",
                        stage="code_a",
                        stage_label="код А",
                        revision="04",
                        appendix="01",
                        transmittals=(trm,),
                    ),
                ),
                revision="04",
                appendix="01",
            ),
        ),
        (
            _send(
                title="6130",
                mark="SOS",
                revision="04",
                appendix="01",
                status="Принят",
                send_date="01.08.2026",
                incoming="05.08.2026",
                transmittal=trm,
            ),
        ),
        records,
        {10},
    )
    contour = _resolve(database, records, {10}, "6130", "SOS")
    assert contour is not None
    assert contour.package_sequence == 10
    assert contour.mto_path == record.path
    assert contour.confidence == "high"
    assert not any("Согласована ревизия" in note for note in contour.warnings)


def test_later_folder_same_label_does_not_steal_exact(temp: Path) -> None:
    """A later NN that keeps folder ``рев.01-AN01`` but files at ``01-AN02``
    must not beat the package whose filename revision is the approved text.
    """

    database = _open_db(temp / "folder_label_lag")
    mtime = _mtime_ns(2026, 8, 1)
    nn05 = _with_folder_rev(
        _record(
            5,
            title="3150",
            mark="KSB",
            revision="01",
            appendix="01",
            sequence=5,
            folder="05_рев.01-AN01_AGCC.287-3150-KSB",
            mtime_ns=mtime,
            file_kind="mto_xlsx",
        ),
        "01",
        "01",
    )
    nn06 = _with_folder_rev(
        _record(
            6,
            title="3150",
            mark="KSB",
            revision="01",
            appendix="02",
            sequence=6,
            folder="06_рев.01-AN01_AGCC.287-3150-KSB",
            mtime_ns=mtime + 1,
            file_kind="mto_xlsx",
        ),
        "01",
        "01",
    )
    records = [nn05, nn06]
    trm = "AGCC.287-PGS-PGS-TRM-31501"
    _rebuild(
        database,
        (
            _google(
                "3150",
                "KSB",
                (
                    _event(
                        date="10.08.2026",
                        stage="code_a",
                        stage_label="код А",
                        revision="01",
                        appendix="01",
                        transmittals=(trm,),
                    ),
                ),
                revision="01",
                appendix="01",
            ),
        ),
        (
            _send(
                title="3150",
                mark="KSB",
                revision="01",
                appendix="01",
                status="Принят",
                send_date="01.08.2026",
                incoming="05.08.2026",
                transmittal=trm,
            ),
        ),
        records,
        {6},
    )
    contour = _resolve(database, records, {6}, "3150", "KSB")
    assert contour is not None
    assert contour.approved_revision_text == "01-AN01"
    assert contour.package_sequence == 5
    assert contour.mto_path == nn05.path
    assert contour.confidence == "high"


def test_folder_rev_fallback_2225_ksb(temp: Path) -> None:
    database = _open_db(temp / "ksb2225")
    record = _with_folder_rev(
        _record(
            10,
            title="2225",
            mark="KSB",
            revision="01",
            appendix="02",
            sequence=10,
            folder="10_рев.01-AN01_AGCC.287-2225-KSB_as-built",
            mtime_ns=_mtime_ns(2026, 8, 1),
            file_kind="mto_xlsx",
        ),
        "01",
        "01",
    )
    records = [record]
    _rebuild(
        database,
        (
            _google(
                "2225",
                "KSB",
                (
                    _event(
                        date="01.08.2026",
                        stage="code_a",
                        stage_label="код А",
                        revision="01",
                        appendix="01",
                    ),
                ),
                revision="01",
                appendix="01",
            ),
        ),
        (),
        records,
        {10},
    )
    contour = _resolve(database, records, {10}, "2225", "KSB")
    assert contour is not None
    assert contour.match_reason == "folder_rev"
    assert contour.package_sequence == 10
    assert contour.approved_revision_text == "01-AN01"
    assert contour.mto_path == record.path


def test_dup_same_revision_in_contour_is_low_confidence(temp: Path) -> None:
    database = _open_db(temp / "dup")
    mtime = _mtime_ns(2026, 4, 1)
    folder = "01_рев.01_AGCC.287-2245-KSB"
    first = _record(
        51,
        title="2245",
        mark="KSB",
        revision="01",
        appendix=None,
        sequence=1,
        folder=folder,
        mtime_ns=mtime,
        file_kind="mto_xlsx",
    )
    second = _record(
        52,
        title="2245",
        mark="KSB",
        revision="01",
        appendix=None,
        sequence=1,
        folder=folder,
        mtime_ns=mtime + 10,
        file_kind="mto_xlsx",
    )
    records = [first, second]
    _rebuild(
        database,
        (
            _google(
                "2245",
                "KSB",
                (
                    _event(
                        date="01.04.2026",
                        stage="code_a",
                        stage_label="код А",
                        revision="01",
                    ),
                ),
                revision="01",
                appendix=None,
            ),
        ),
        (
            _send(
                title="2245",
                mark="KSB",
                revision="01",
                appendix=None,
                status="Принят",
                send_date="20.03.2026",
                incoming="25.03.2026",
            ),
        ),
        records,
        {52},
    )
    _insert_overlay_mto_collision(
        database,
        [first.path_key, second.path_key],
        CollisionKind.DUP_SAME_REVISION,
    )
    contour = _resolve(database, records, {52}, "2245", "KSB")
    assert contour is not None
    assert contour.confidence == "low"
    assert "dup_same_revision" in contour.ambiguity


def _mto_parsed(name: str, transfer_name: str, mtime_ns: int):
    transfer = parse_transfer_folder(transfer_name, under_gate=True)
    return parse_catalog_file(
        f"C:/fixture/{transfer_name}/{name}",
        SourceKind.RD,
        size=100,
        mtime_ns=mtime_ns,
        transfer=transfer,
    )


def test_build_rd_overlays_as_of_bound_and_unbound() -> None:
    older = _mto_parsed(
        "AGCC.287-1600-SOT.MTO-0001_03_RU.xlsx",
        "20_рев.04_AGCC.287-1600-SOT",
        20,
    )
    newer = _mto_parsed(
        "AGCC.287-1600-SOT.MTO-0001_02_RU.xlsx",
        "21_рев.03_AGCC.287-1600-SOT",
        21,
    )
    files = [older, newer]
    unbounded = build_rd_overlays_as_of(files, max_sequence=None)
    baseline = build_rd_overlays(files)
    assert unbounded[0].current == baseline[0].current
    assert unbounded[1].current == baseline[1].current
    assert unbounded[0].entries == baseline[0].entries
    assert unbounded[1].entries == baseline[1].entries
    assert unbounded[0].collisions == baseline[0].collisions
    assert unbounded[1].collisions == baseline[1].collisions

    bounded = build_rd_overlays_as_of(files, max_sequence=20)
    assert bounded[1].file_kind is FileKind.MTO_XLSX
    assert list(bounded[1].current.values()) == [older]
    assert newer not in bounded[1].current.values()
    assert list(unbounded[1].current.values()) == [newer]


def test_resolve_unknown_kit_is_none(temp: Path) -> None:
    database = _open_db(temp / "unknown")
    database.initialize()
    assert (
        _resolve(database, [], set(), "9999", "KSB") is None
    )
    assert resolve_all_approved_contours(
        database, records=[], detected_current_ids=set()
    ) == ()


def main() -> None:
    """Run approved-contour fixtures against a temporary SQLite file."""

    test_build_rd_overlays_as_of_bound_and_unbound()
    with tempfile.TemporaryDirectory(prefix="rd_catalog_approved_contour_") as raw:
        root = Path(raw)
        test_1600_sot_later_package_with_older_files_does_not_win(root)
        test_6400_sot_filename_lags_approved_revision(root)
        test_inherited_mto_from_earlier_package(root)
        test_exact_full_revision_wins_over_base_only(root)
        test_base_only_fallback_lowers_confidence(root)
        test_folder_an01_matches_approved_04_an01(root)
        test_exact_match_keeps_high_confidence(root)
        test_later_folder_same_label_does_not_steal_exact(root)
        test_folder_rev_fallback_2225_ksb(root)
        test_dup_same_revision_in_contour_is_low_confidence(root)
        test_resolve_unknown_kit_is_none(root)
    print("RD catalog approved contour: OK")


if __name__ == "__main__":
    main()
