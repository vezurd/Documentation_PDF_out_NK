"""Local checks for RD catalog per-revision MTO heatmap cells."""

from __future__ import annotations

import json
import sys
import tempfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
sys.path.insert(0, str(Path(__file__).resolve().parent))

from rd_catalog.models import CollisionKind
from rd_catalog.pipeline import (
    REVISION_MATRIX_ALGORITHM_VERSION,
    iter_mto_files_for_cells,
    list_revision_columns,
    list_revision_matrix,
    rebuild_pipeline,
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


def _insert_collision(database, path_key: str, kind: CollisionKind) -> None:
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
                "mto:2245-ksb|mto-0001",
                json.dumps([path_key], ensure_ascii=False),
                run_id,
            ),
        )


def _cell_map(cells):
    return {cell.revision_text: cell for cell in cells}


def test_revision_matrix_codes_current_and_collision(temp: Path) -> None:
    database = _open_db(temp / "matrix")
    mtime = _mtime_ns(2026, 8, 5)
    rec_01 = _record(
        11,
        title="2245",
        mark="KSB",
        revision="01",
        appendix=None,
        sequence=1,
        folder="01_рев.01_2245-KSB",
        mtime_ns=mtime,
        file_kind="mto_xlsx",
    )
    rec_02 = _record(
        12,
        title="2245",
        mark="KSB",
        revision="02",
        appendix=None,
        sequence=2,
        folder="02_рев.02_2245-KSB",
        mtime_ns=mtime,
        file_kind="mto_xlsx",
    )
    rec_03 = _record(
        13,
        title="2245",
        mark="KSB",
        revision="03",
        appendix=None,
        sequence=8,
        folder="08_рев.03_2245-KSB_as-built",
        mtime_ns=mtime,
        file_kind="mto_xlsx",
        as_build=True,
    )
    records = [rec_01, rec_02, rec_03]
    events = (
        _event(
            date="01.06.2026",
            stage="code_a",
            stage_label="код А",
            revision="01",
        ),
        _event(
            date="10.07.2026",
            stage="tdo_passed",
            stage_label="прошла ТДО",
            revision="02",
        ),
        _event(
            date="01.08.2026",
            stage="us_build",
            stage_label="US-BUILD",
            revision="03",
        ),
    )
    _rebuild(
        database,
        (_google("2245", "KSB", events, revision="03", appendix=None),),
        (
            _send(
                title="2245",
                mark="KSB",
                revision="01",
                appendix=None,
                status="Принят",
                send_date="20.05.2026",
                incoming="25.05.2026",
                row_index=1,
            ),
            _send(
                title="2245",
                mark="KSB",
                revision="02",
                appendix=None,
                status="Принят",
                send_date="01.07.2026",
                incoming="05.07.2026",
                row_index=2,
            ),
        ),
        records,
        {13},
    )
    _insert_collision(
        database, rec_01.path_key, CollisionKind.TRANSFER_ORDER_CONFLICT
    )
    rebuild_pipeline(
        database,
        records=records,
        detected_current_ids={13},
    )

    cells = [
        cell
        for cell in list_revision_matrix(database)
        if cell.title == "2245" and cell.mark == "KSB"
    ]
    by_rev = _cell_map(cells)
    assert set(by_rev) == {"01", "02", "03"}
    assert list_revision_columns(cells) == ("01", "02", "03")

    cell_01 = by_rev["01"]
    assert cell_01.pipeline_status == "code_a"
    assert cell_01.letters == "A"
    assert cell_01.has_mto is True
    assert cell_01.is_current is False
    assert cell_01.is_current_ifc is False
    assert cell_01.is_as_build is False
    assert cell_01.algorithm_version == REVISION_MATRIX_ALGORITHM_VERSION
    kinds = json.loads(cell_01.problem_kinds_json)
    assert CollisionKind.TRANSFER_ORDER_CONFLICT.value in kinds
    package_ids = json.loads(cell_01.package_ids_json)
    assert package_ids
    stored_ids = {pkg.id for pkg in database.list_kit_packages("2245", "KSB")}
    assert set(package_ids) <= stored_ids

    cell_02 = by_rev["02"]
    assert cell_02.pipeline_status == "tdo_review"
    assert cell_02.letters == "T"
    assert cell_02.is_current is False
    assert cell_02.is_current_ifc is True
    assert cell_02.is_as_build is False

    cell_03 = by_rev["03"]
    assert cell_03.has_mto is True
    assert cell_03.is_as_build is True
    assert cell_03.is_current is True
    assert cell_03.is_current_ifc is False
    assert cell_03.pipeline_status == "working"
    assert "AB" in cell_03.letters.split("·")
    packages = database.list_kit_packages("2245", "KSB")
    as_build_packages = [pkg for pkg in packages if pkg.is_as_build]
    assert as_build_packages
    assert all(pkg.revision_text == "03" for pkg in as_build_packages)

    code_a_files = iter_mto_files_for_cells(
        database, records=records, predicate="code_a"
    )
    assert [item.path_key for item in code_a_files] == [rec_01.path_key]
    assert rec_01.path in {item.path for item in code_a_files}

    ifc_files = iter_mto_files_for_cells(
        database, records=records, predicate="current_ifc"
    )
    assert [item.path_key for item in ifc_files] == [rec_02.path_key]

    current_files = iter_mto_files_for_cells(
        database, records=records, predicate="current"
    )
    assert [item.path_key for item in current_files] == [rec_02.path_key]

    without_ab = iter_mto_files_for_cells(
        database, records=records, predicate="exclude_as_build"
    )
    assert rec_03.path_key not in {item.path_key for item in without_ab}
    assert rec_01.path_key in {item.path_key for item in without_ab}
    assert rec_02.path_key in {item.path_key for item in without_ab}

    try:
        iter_mto_files_for_cells(database, records=records, predicate="nope")
    except ValueError as exc:
        assert "predicate" in str(exc).casefold()
    else:
        raise AssertionError("expected ValueError for unknown predicate")


def main() -> None:
    """Run heatmap fixtures against a temporary SQLite file."""

    with tempfile.TemporaryDirectory(prefix="rd_catalog_rev_matrix_") as raw:
        test_revision_matrix_codes_current_and_collision(Path(raw))
    print("RD catalog revision matrix: OK")


if __name__ == "__main__":
    main()
