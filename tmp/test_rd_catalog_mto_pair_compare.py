"""Local checks for RD catalog symmetric MTO pair comparison."""

from __future__ import annotations

import json
import os
import sqlite3
import subprocess
import sys
import tempfile
from dataclasses import replace
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
sys.path.insert(0, str(Path(__file__).resolve().parent))

from rd_catalog.config import CatalogConfig
from rd_catalog.db import SCHEMA_VERSION, CatalogDatabase, canonical_mto_pair_ids
from rd_catalog.models import (
    FileKind,
    FileRecord,
    ParseStatus,
    ParsedFile,
    ScanSummary,
    SourceKind,
    SourceScanResult,
    make_path_key,
)
from rd_catalog.mto_diff import MTO_COMPARE_ALGORITHM_VERSION
from rd_catalog.mto_export import ExportSelection
from rd_catalog.mto_pair_compare import (
    MtoFilePair,
    PairCompareOutcome,
    build_pair_compare_plan,
    compare_pair,
    execute_pair_compare,
    pair_pool_status,
    pairs_from_export_selections,
)
from rd_catalog.mto_pair_compare_job import (
    config_to_pair_job_dict,
    execute_mto_pair_compare,
    session_verdicts_from_payload,
)
from rd_catalog.overlay import build_overlay

REPO_ROOT = Path(__file__).resolve().parents[1]
COPY_DB = REPO_ROOT / "tmp" / "rd_catalog_copy.sqlite3"

LEFT_PATH = r"C:\rd\AGCC.287-3240-KSB1.MTO-0001_01_RU.xlsx"
RIGHT_PATH = r"C:\robot\AGCC.287-3240-KSB1.MTO-0001_01_RU.xlsx"
CUSTOM_RIGHT_PATH = r"C:\archive\custom\AGCC.287-3240-KSB1.MTO-0001_01_RU.xlsx"

BASE_ROW = {
    "row_type": "position_row",
    "CODE": "BCC0001",
    "UNITS": "шт",
    "VALUES": 1,
    "TAGS": ["3240-SH-01"],
    "NAME": "Прибор",
    "VENDOR": "Vendor",
    "TYPE_MARK": "Type A",
}
ROW_B = {**BASE_ROW, "CODE": "BCC0002", "NAME": "Кабель"}
HEADER_ROW = {
    "row_type": "header_row",
    "CODE": "SHOULD_IGNORE",
    "UNITS": "шт",
    "VALUES": 99,
    "TAGS": ["nope"],
    "NAME": "Заголовок",
    "VENDOR": "X",
    "TYPE_MARK": "H",
}


@pytest.fixture
def temp(tmp_path: Path) -> Path:
    return tmp_path


def _parsed(
    path: str,
    source: SourceKind,
    *,
    size: int,
    mtime_ns: int,
) -> ParsedFile:
    return ParsedFile(
        path=path,
        path_key=make_path_key(path),
        name=Path(path).name,
        source=source,
        file_kind=FileKind.MTO_XLSX,
        size=size,
        mtime_ns=mtime_ns,
        parse_status=ParseStatus.PARSED,
        title="3240",
        mark="KSB1",
        title_system="3240-KSB1",
        discipline_block="MTO-0001",
        revision="01",
        language="RU",
        extension="xlsx",
    )


def _store_pair(
    db_path: Path,
    *,
    left_size: int = 10,
    right_size: int = 20,
    left_mtime: int = 100,
    right_mtime: int = 200,
) -> tuple[CatalogDatabase, FileRecord, FileRecord]:
    database = CatalogDatabase(db_path)
    database.initialize()
    left = _parsed(LEFT_PATH, SourceKind.RD, size=left_size, mtime_ns=left_mtime)
    right = _parsed(RIGHT_PATH, SourceKind.ROBOT, size=right_size, mtime_ns=right_mtime)
    summary = ScanSummary(
        sources={
            SourceKind.RD: SourceScanResult(
                source=SourceKind.RD,
                root=r"C:\rd",
                files=[left],
            ),
            SourceKind.ROBOT: SourceScanResult(
                source=SourceKind.ROBOT,
                root=r"C:\robot",
                files=[right],
            ),
        },
        rd_mto_overlay=build_overlay([left], FileKind.MTO_XLSX),
    )
    database.store_scan(summary)
    records = {record.path_key: record for record in database.list_files()}
    return database, records[left.path_key], records[right.path_key]


def _store_left_only(db_path: Path) -> tuple[CatalogDatabase, FileRecord]:
    database = CatalogDatabase(db_path)
    database.initialize()
    left = _parsed(LEFT_PATH, SourceKind.RD, size=10, mtime_ns=100)
    summary = ScanSummary(
        sources={
            SourceKind.RD: SourceScanResult(
                source=SourceKind.RD,
                root=r"C:\rd",
                files=[left],
            ),
        },
        rd_mto_overlay=build_overlay([left], FileKind.MTO_XLSX),
    )
    database.store_scan(summary)
    records = {record.path_key: record for record in database.list_files()}
    return database, records[left.path_key]


def _session_map(
    outcome: PairCompareOutcome,
) -> dict[tuple[str, str], str]:
    return {
        key: result.content_status
        for key, result in outcome.session_verdicts.items()
    }


def _pair_row_count(database: CatalogDatabase) -> int:
    with database._connection() as connection:
        return int(
            connection.execute(
                "SELECT COUNT(*) FROM mto_pair_comparison"
            ).fetchone()[0]
        )


def _catalog_config(root: Path, db_path: Path) -> CatalogConfig:
    return CatalogConfig(
        rd_root=root / "rd",
        sq_root=root / "sq",
        robot_root=root / "robot",
        runtime_dir=root / "runtime",
        db_path=db_path,
        robot_flat_structure=True,
        skip_dirs=(),
    )


def _done_payload(stdout: str) -> dict:
    done = None
    for line in stdout.splitlines():
        text = line.strip()
        if not text:
            continue
        payload = json.loads(text)
        if payload.get("t") == "done":
            done = payload
    assert done is not None, stdout
    return done


def _loader(mapping: dict[str, list[dict[str, object]]], calls: list[str] | None = None):
    def load(path: str | Path) -> list[dict[str, object]]:
        key = str(path)
        if calls is not None:
            calls.append(key)
        if key not in mapping:
            raise FileNotFoundError(key)
        return mapping[key]

    return load


def _pair() -> MtoFilePair:
    return MtoFilePair(LEFT_PATH, RIGHT_PATH, "3240", "KSB1")


def _selection(
    *,
    state: str,
    source_path: str = "",
    existing_target_path: str = "",
) -> ExportSelection:
    return ExportSelection(
        title="3240",
        mark="KSB1",
        rule="latest_issued",
        source_path=source_path,
        source_revision_text="01",
        package_path="",
        origin="rule",
        state=state,
        destination_path="",
        existing_target_path=existing_target_path,
        confidence="high",
        warnings=(),
    )


def _sqlite_counts(path: Path) -> tuple[int, int, bool, int]:
    connection = sqlite3.connect(path)
    try:
        version_row = connection.execute(
            "SELECT value FROM schema_meta WHERE key = 'schema_version'"
        ).fetchone()
        version = int(version_row[0]) if version_row else 0
        mto_n = int(
            connection.execute("SELECT COUNT(*) FROM mto_comparison").fetchone()[0]
        )
        tables = {
            str(row[0])
            for row in connection.execute(
                "SELECT name FROM sqlite_master WHERE type = 'table'"
            )
        }
        has_pair = "mto_pair_comparison" in tables
        pair_n = 0
        if has_pair:
            pair_n = int(
                connection.execute(
                    "SELECT COUNT(*) FROM mto_pair_comparison"
                ).fetchone()[0]
            )
    finally:
        connection.close()
    return version, mto_n, has_pair, pair_n


def test_content_equal_ignores_format_and_order() -> None:
    left_rows = [BASE_ROW, ROW_B, HEADER_ROW]
    right_rows = [
        {**ROW_B, "extra_format": "bold"},
        {**BASE_ROW, "VALUES": "1.0"},
        HEADER_ROW,
        {"row_type": "note_row", "NAME": "примечание"},
    ]
    result = compare_pair(
        LEFT_PATH,
        RIGHT_PATH,
        loader=_loader({LEFT_PATH: left_rows, RIGHT_PATH: right_rows}),
    )
    assert result.content_status == "content_equal"
    assert result.error is None
    assert result.added == 0
    assert result.removed == 0
    assert result.changed == 0
    assert result.left_rows == 2
    assert result.right_rows == 2


def test_values_change_is_content_diff() -> None:
    changed = {**BASE_ROW, "VALUES": 2}
    result = compare_pair(
        LEFT_PATH,
        RIGHT_PATH,
        loader=_loader({LEFT_PATH: [BASE_ROW], RIGHT_PATH: [changed]}),
    )
    assert result.content_status == "content_diff"
    assert result.changed == 1
    assert result.added == 0
    assert result.removed == 0


def test_non_position_rows_are_ignored() -> None:
    result = compare_pair(
        LEFT_PATH,
        RIGHT_PATH,
        loader=_loader(
            {
                LEFT_PATH: [BASE_ROW, HEADER_ROW],
                RIGHT_PATH: [BASE_ROW],
            }
        ),
    )
    assert result.content_status == "content_equal"


def test_loader_error_is_not_compared() -> None:
    def boom(_path: str | Path) -> list[dict[str, object]]:
        raise RuntimeError("xlsx boom")

    result = compare_pair(LEFT_PATH, RIGHT_PATH, loader=boom)
    assert result.content_status == "not_compared"
    assert result.error is not None
    assert "xlsx boom" in result.error


def test_canonical_order_one_row_one_computation(temp: Path) -> None:
    database, left, right = _store_pair(temp / "canon.sqlite")
    records = [left, right]
    pair_ab = MtoFilePair(left.path, right.path, "3240", "KSB1")
    pair_ba = MtoFilePair(right.path, left.path, "3240", "KSB1")
    mapping = {left.path: [BASE_ROW], right.path: [BASE_ROW]}
    calls: list[str] = []
    loader = _loader(mapping, calls)
    plan_ab = build_pair_compare_plan(
        database, pairs=(pair_ab,), records=records
    )
    outcome = execute_pair_compare(
        database, plan=plan_ab, records=records, loader=loader
    )
    assert outcome.written == 1
    assert len(calls) == 2
    row_ab = database.get_mto_pair_comparison(
        left.id, right.id, MTO_COMPARE_ALGORITHM_VERSION
    )
    row_ba = database.get_mto_pair_comparison(
        right.id, left.id, MTO_COMPARE_ALGORITHM_VERSION
    )
    assert row_ab is not None
    assert row_ba is not None
    assert row_ab["id"] == row_ba["id"]
    listed = database.list_mto_pair_comparisons(
        [(left.id, right.id), (right.id, left.id)]
    )
    assert len(listed) == 1
    assert canonical_mto_pair_ids(left.id, right.id) in listed
    plan_ba = build_pair_compare_plan(
        database, pairs=(pair_ba,), records=records
    )
    assert plan_ba.pending == ()
    assert len(plan_ba.cached) == 1
    outcome_again = execute_pair_compare(
        database, plan=plan_ba, records=records, loader=loader
    )
    assert outcome_again.written == 0
    assert len(calls) == 2
    both_plan = build_pair_compare_plan(
        database, pairs=(pair_ab, pair_ba), records=records
    )
    assert len(both_plan.pending) + len(both_plan.cached) == 1


def test_cache_invalidation(temp: Path) -> None:
    database, left, right = _store_pair(temp / "cache.sqlite")
    records = [left, right]
    loader = _loader({left.path: [BASE_ROW], right.path: [BASE_ROW]})
    pair = MtoFilePair(left.path, right.path, "3240", "KSB1")
    execute_pair_compare(
        database,
        plan=build_pair_compare_plan(database, pairs=(pair,), records=records),
        records=records,
        loader=loader,
    )
    assert database.mto_pair_comparison_valid(left=left, right=right)
    assert database.mto_pair_comparison_valid(left=right, right=left)
    disk_mtime = int(left.data.get("disk_mtime_ns") or left.data["mtime_ns"])
    bumped = replace(
        left,
        data={
            **left.data,
            "mtime_ns": disk_mtime + 1,
            "disk_mtime_ns": disk_mtime + 1,
        },
    )
    assert not database.mto_pair_comparison_valid(left=bumped, right=right)
    with database._connection() as connection, connection:
        connection.execute(
            "UPDATE file_entry SET mtime_ns = mtime_ns + 5 WHERE id = ?",
            (left.id,),
        )
    refreshed = {record.id: record for record in database.list_files()}
    new_left = refreshed[left.id]
    new_right = refreshed[right.id]
    assert not database.mto_pair_comparison_valid(left=new_left, right=new_right)
    plan = build_pair_compare_plan(
        database, pairs=(pair,), records=[new_left, new_right]
    )
    assert plan.pending == (pair,)


def test_failed_pair_is_failed_not_pending(temp: Path) -> None:
    database, left, right = _store_pair(temp / "fail.sqlite")
    records = [left, right]
    pair = MtoFilePair(left.path, right.path, "3240", "KSB1")

    def boom(_path: str | Path) -> list[dict[str, object]]:
        raise PermissionError("workbook is open")

    pending_status = pair_pool_status(database, pairs=(pair,), records=records)
    assert pending_status.pending == 1
    assert pending_status.is_drained is False
    outcome = execute_pair_compare(
        database,
        plan=build_pair_compare_plan(database, pairs=(pair,), records=records),
        records=records,
        loader=boom,
    )
    assert outcome.written == 1
    stored = database.get_mto_pair_comparison(
        left.id, right.id, MTO_COMPARE_ALGORITHM_VERSION
    )
    assert stored is not None
    assert stored["content_status"] == "not_compared"
    assert stored["error"]
    assert "PermissionError" in str(stored["error"])
    assert not database.mto_pair_comparison_valid(left=left, right=right)
    plan = build_pair_compare_plan(database, pairs=(pair,), records=records)
    assert plan.pending == (pair,)
    assert plan.cached == ()
    without_session = pair_pool_status(database, pairs=(pair,), records=records)
    assert without_session.pending == 1
    assert without_session.failed == 0
    assert without_session.is_drained is False
    with_session = pair_pool_status(
        database,
        pairs=(pair,),
        records=records,
        session_verdicts=_session_map(outcome),
    )
    assert with_session.total == 1
    assert with_session.pending == 0
    assert with_session.failed == 1
    assert with_session.compared == 0
    assert with_session.unpersisted == 0
    assert with_session.is_drained is True

    success = execute_pair_compare(
        database,
        plan=plan,
        records=records,
        loader=_loader({left.path: [BASE_ROW], right.path: [BASE_ROW]}),
    )
    assert success.written == 1
    overwritten = database.get_mto_pair_comparison(
        left.id, right.id, MTO_COMPARE_ALGORITHM_VERSION
    )
    assert overwritten is not None
    assert overwritten["content_status"] == "content_equal"
    assert database.mto_pair_comparison_valid(left=left, right=right)
    cached_plan = build_pair_compare_plan(
        database, pairs=(pair,), records=records
    )
    assert cached_plan.pending == ()
    assert cached_plan.cached == (pair,)


def test_unpersisted_custom_target_pair(temp: Path) -> None:
    database, left = _store_left_only(temp / "custom.sqlite")
    records = [left]
    pair = MtoFilePair(left.path, CUSTOM_RIGHT_PATH, "3240", "KSB1")
    loader = _loader({left.path: [BASE_ROW], CUSTOM_RIGHT_PATH: [BASE_ROW]})
    plan = build_pair_compare_plan(database, pairs=(pair,), records=records)
    assert plan.pending == (pair,)
    assert plan.cached == ()
    before = _pair_row_count(database)
    outcome = execute_pair_compare(
        database, plan=plan, records=records, loader=loader
    )
    assert outcome.written == 0
    assert _pair_row_count(database) == before
    assert len(outcome.session_verdicts) == 1
    session_key, result = next(iter(outcome.session_verdicts.items()))
    assert session_key == tuple(
        sorted((make_path_key(left.path), make_path_key(CUSTOM_RIGHT_PATH)))
    )
    assert result.content_status == "content_equal"
    with_session = pair_pool_status(
        database,
        pairs=(pair,),
        records=records,
        session_verdicts=_session_map(outcome),
    )
    assert with_session.total == 1
    assert with_session.pending == 0
    assert with_session.compared == 1
    assert with_session.failed == 0
    assert with_session.unpersisted == 1
    assert with_session.is_drained is True
    without_session = pair_pool_status(
        database, pairs=(pair,), records=records
    )
    assert without_session.pending == 1
    assert without_session.compared == 0
    assert without_session.unpersisted == 1
    assert without_session.is_drained is False


def test_json_line_session_verdicts(temp: Path) -> None:
    database, left = _store_left_only(temp / "cli.sqlite")
    pair = MtoFilePair(left.path, CUSTOM_RIGHT_PATH, "3240", "KSB1")
    loader = _loader({left.path: [BASE_ROW], CUSTOM_RIGHT_PATH: [ROW_B]})
    config = _catalog_config(temp, database.path)
    in_process = execute_mto_pair_compare(
        config, pairs=(pair,), loader=loader
    )
    assert in_process.failure is None
    assert in_process.comparison_count == 1
    assert _pair_row_count(database) == 0
    assert len(in_process.session_verdicts) == 1
    in_key, in_result = next(iter(in_process.session_verdicts.items()))
    assert in_result.content_status == "content_diff"
    roundtrip = session_verdicts_from_payload(
        [
            {
                "left_path_key": in_key[0],
                "right_path_key": in_key[1],
                "content_status": in_result.content_status,
                "left_fingerprint": in_result.left_fingerprint,
                "right_fingerprint": in_result.right_fingerprint,
                "left_rows": in_result.left_rows,
                "right_rows": in_result.right_rows,
                "added": in_result.added,
                "removed": in_result.removed,
                "changed": in_result.changed,
                "error": in_result.error,
            }
        ]
    )
    assert roundtrip[in_key].content_status == "content_diff"

    runtime = temp / "runtime"
    runtime.mkdir(parents=True, exist_ok=True)
    job_path = runtime / "mto_pair_job_test.json"
    job_path.write_text(
        json.dumps(
            config_to_pair_job_dict(config, pairs=(pair,), batch_size=1),
            ensure_ascii=False,
        ),
        encoding="utf-8",
    )
    env = os.environ.copy()
    env["PYTHONPATH"] = str(REPO_ROOT)
    env["PYTHONUTF8"] = "1"
    cli_result = subprocess.run(
        [
            sys.executable,
            "-m",
            "rd_catalog.mto_pair_compare_cli",
            "--job",
            str(job_path),
        ],
        cwd=REPO_ROOT,
        env=env,
        capture_output=True,
        text=True,
        timeout=60,
        check=False,
    )
    assert cli_result.returncode == 0, cli_result.stderr
    done = _done_payload(cli_result.stdout)
    assert "session_verdicts" in done
    decoded = session_verdicts_from_payload(done["session_verdicts"])
    assert len(decoded) == 1
    cli_result_row = next(iter(decoded.values()))
    assert cli_result_row.content_status == "not_compared"
    assert cli_result_row.error
    assert _pair_row_count(database) == 0


def test_pairs_from_export_selections_skips_add_and_no_source() -> None:
    add = _selection(state="add", source_path=LEFT_PATH, existing_target_path="")
    missing = _selection(state="no_source")
    comparable = _selection(
        state="unknown",
        source_path=LEFT_PATH,
        existing_target_path=RIGHT_PATH,
    )
    pin_stale = _selection(
        state="pin_stale",
        source_path=LEFT_PATH,
        existing_target_path=RIGHT_PATH,
    )
    swapped = _selection(
        state="replace",
        source_path=RIGHT_PATH,
        existing_target_path=LEFT_PATH,
    )
    found = pairs_from_export_selections(
        (add, missing, comparable, pin_stale, swapped)
    )
    assert len(found) == 1
    assert found[0].left_path == LEFT_PATH
    assert found[0].right_path == RIGHT_PATH


def test_pool_drained_after_attempts(temp: Path) -> None:
    database, left, right = _store_pair(temp / "pool.sqlite")
    records = [left, right]
    pair = MtoFilePair(left.path, right.path, "3240", "KSB1")
    empty = pair_pool_status(database, pairs=(pair,), records=records)
    assert empty.is_drained is False
    execute_pair_compare(
        database,
        plan=build_pair_compare_plan(database, pairs=(pair,), records=records),
        records=records,
        loader=_loader({left.path: [BASE_ROW], right.path: [BASE_ROW]}),
    )
    done = pair_pool_status(database, pairs=(pair,), records=records)
    assert done.compared == 1
    assert done.pending == 0
    assert done.failed == 0
    assert done.unpersisted == 0
    assert done.is_drained is True


def test_schema_5_migrates_keeping_mto_comparison(temp: Path) -> None:
    path = temp / "v5.sqlite"
    path.parent.mkdir(parents=True, exist_ok=True)
    connection = sqlite3.connect(path)
    try:
        connection.execute(
            "CREATE TABLE schema_meta "
            "(key TEXT PRIMARY KEY, value TEXT NOT NULL)"
        )
        CatalogDatabase._migrate_0_to_1(connection)
        CatalogDatabase._migrate_1_to_2(connection)
        CatalogDatabase._migrate_2_to_3(connection)
        CatalogDatabase._migrate_3_to_4(connection)
        CatalogDatabase._migrate_4_to_5(connection)
        connection.execute(
            "INSERT INTO scan_run(started_at, status, is_baseline) "
            "VALUES ('t', 'success', 1)"
        )
        run_id = int(connection.execute("SELECT last_insert_rowid()").fetchone()[0])
        connection.execute(
            """
            INSERT INTO file_entry(
                id, path, path_key, source, file_kind, name, size, mtime_ns,
                first_seen_run_id, last_seen_run_id, present, review_state,
                parse_status, transfer_is_as_build
            ) VALUES (11, 'C:\\rd\\a.xlsx', 'c:\\rd\\a.xlsx', 'rd', 'mto_xlsx',
                      'a.xlsx', 1, 1, ?, ?, 1, 'acknowledged', 'parsed', 0)
            """,
            (run_id, run_id),
        )
        connection.execute(
            """
            INSERT INTO mto_comparison(
                rd_file_id, robot_file_id, rd_fingerprint, robot_fingerprint,
                algorithm_version, status, stats_json, diff_json, compared_at
            ) VALUES (11, 11, 'rd', 'robot', 1, 'content_equal', '{}', '{}', 't')
            """
        )
        connection.commit()
    finally:
        connection.close()
    version, mto_before, has_pair, _pair_n = _sqlite_counts(path)
    assert version == 5
    assert mto_before == 1
    assert has_pair is False
    database = CatalogDatabase(path)
    database.initialize()
    assert database.schema_version() == SCHEMA_VERSION
    assert SCHEMA_VERSION == 14
    version_after, mto_after, has_pair_after, pair_n = _sqlite_counts(path)
    assert version_after == SCHEMA_VERSION
    assert mto_after == mto_before == 1
    assert has_pair_after is True
    assert pair_n == 0
    cached = database.get_mto_comparison(11, 11, 1)
    assert cached is not None
    assert cached["rd_fingerprint"] == "rd"


def test_real_copy_migration() -> None:
    if not COPY_DB.is_file():
        pytest.skip(f"missing copy database: {COPY_DB}")
    before = _sqlite_counts(COPY_DB)
    database = CatalogDatabase(COPY_DB)
    database.initialize()
    after = _sqlite_counts(COPY_DB)
    assert after[0] == 7
    assert after[1] == before[1]
    assert after[2] is True
    print(
        "copy migration: "
        f"mto_comparison {before[1]} -> {after[1]}; "
        f"schema {before[0]} -> {after[0]}; "
        f"mto_pair_comparison rows={after[3]}"
    )


def main() -> None:
    """Run pair-compare fixtures against a temporary directory."""

    test_content_equal_ignores_format_and_order()
    test_values_change_is_content_diff()
    test_non_position_rows_are_ignored()
    test_loader_error_is_not_compared()
    with tempfile.TemporaryDirectory(prefix="rd_catalog_mto_pair_") as raw:
        root = Path(raw)
        test_canonical_order_one_row_one_computation(root / "canon")
        test_cache_invalidation(root / "cache")
        test_failed_pair_is_failed_not_pending(root / "fail")
        test_unpersisted_custom_target_pair(root / "custom")
        test_json_line_session_verdicts(root / "cli")
        test_pairs_from_export_selections_skips_add_and_no_source()
        test_pool_drained_after_attempts(root / "pool")
        test_schema_5_migrates_keeping_mto_comparison(root / "migrate")
    if COPY_DB.is_file():
        test_real_copy_migration()
    else:
        print(f"skip real copy: {COPY_DB} is missing")
    print("RD catalog MTO pair compare: OK")


if __name__ == "__main__":
    main()
