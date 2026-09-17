"""Unit tests for MTO compare plan, job, and CLI."""

from __future__ import annotations

import json
import subprocess
import sys
import tempfile
from pathlib import Path

import openpyxl

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from rd_catalog.config import CatalogConfig
from rd_catalog.db import CatalogDatabase
from rd_catalog.models import (
    FileKind,
    ScanSummary,
    SourceKind,
    SourceScanResult,
)
from rd_catalog.mto_compare_job import (
    config_to_mto_job_dict,
    execute_mto_compare,
)
from rd_catalog.mto_compare_queue import build_mto_compare_plan
from rd_catalog.overlay import build_overlay
from rd_catalog.parse import parse_catalog_file, parse_transfer_folder

REPO_ROOT = Path(__file__).resolve().parents[1]

BASE_ROW = {
    "CODE": "BCC0001",
    "UNITS": "шт",
    "VALUES": 1,
    "TAGS": ["2225-SH-01-S-UZ-0001"],
    "NAME": "Прибор",
    "VENDOR": "Vendor",
    "TYPE_MARK": "Type A",
}


def _write_mto(path: Path, rows: list[dict[str, object]]) -> None:
    workbook = openpyxl.Workbook()
    worksheet = workbook.active
    worksheet.title = "Спецификация"
    for row in rows:
        worksheet.append(
            [
                ", ".join(row["TAGS"]) if isinstance(row["TAGS"], list) else row["TAGS"],
                "",
                row["NAME"],
                row["TYPE_MARK"],
                row["CODE"],
                row["VENDOR"],
                row["UNITS"],
                row["VALUES"],
            ]
        )
    workbook.save(path)
    workbook.close()


def _parsed(
    path: Path,
    source: SourceKind,
    *,
    mtime_ns: int = 100,
):
    transfer = (
        parse_transfer_folder("01_рев.01_2225-KSB")
        if source is SourceKind.RD
        else None
    )
    return parse_catalog_file(
        path,
        source,
        size=path.stat().st_size,
        mtime_ns=mtime_ns,
        transfer=transfer,
    )


def _loader(data: dict[str, list[dict[str, object]]]):
    return lambda path: data[str(path)]


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


def _store_scan(
    database: CatalogDatabase,
    rd_files: list,
    robot_files: list,
    overlay,
) -> None:
    summary = ScanSummary(
        sources={
            SourceKind.RD: SourceScanResult(
                source=SourceKind.RD,
                root=str(Path(rd_files[0].path).parent.parent),
                files=rd_files,
            ),
            SourceKind.ROBOT: SourceScanResult(
                source=SourceKind.ROBOT,
                root=str(Path(robot_files[0].path).parent),
                files=robot_files,
            ),
        },
        rd_mto_overlay=overlay,
    )
    database.store_scan(summary)


def main() -> None:
    """Exercise build_mto_compare_plan cache filtering and ordering."""

    with tempfile.TemporaryDirectory(prefix="rd_catalog_mto_compare_job_") as temp:
        root = Path(temp)
        rd_dir = root / "rd" / "01_рев.01_2225-KSB"
        robot_dir = root / "robot"
        rd_dir.mkdir(parents=True)
        robot_dir.mkdir()

        name_one = "AGCC.287-2225-KSB.MTO-0001_01_RU.xlsx"
        name_two = "AGCC.287-2225-KSB.MTO-0002_01_RU.xlsx"
        rd_path_one = rd_dir / name_one
        rd_path_two = rd_dir / name_two
        robot_path_one = robot_dir / name_one
        robot_path_two = robot_dir / name_two
        _write_mto(rd_path_one, [BASE_ROW])
        _write_mto(rd_path_two, [BASE_ROW])
        _write_mto(robot_path_one, [BASE_ROW])
        _write_mto(robot_path_two, [BASE_ROW])

        rd_file_one = _parsed(rd_path_one, SourceKind.RD)
        rd_file_two = _parsed(rd_path_two, SourceKind.RD)
        robot_file_one = _parsed(robot_path_one, SourceKind.ROBOT)
        robot_file_two = _parsed(robot_path_two, SourceKind.ROBOT)
        key_one = ("2225-ksb", "mto-0001")
        key_two = ("2225-ksb", "mto-0002")

        database = CatalogDatabase(root / "runtime" / "catalog.sqlite")
        database.initialize()

        empty_overlay = build_overlay([], FileKind.MTO_XLSX)
        empty_plan = build_mto_compare_plan(
            database,
            empty_overlay,
            [robot_file_one],
        )
        assert empty_plan.pairs == ()
        assert empty_plan.skipped_cache_hits == 0
        assert not empty_plan.scope_limited

        empty_progress: list[tuple[int, int]] = []
        empty_job = execute_mto_compare(
            _catalog_config(root, database.path),
            progress=lambda completed, total: empty_progress.append(
                (completed, total)
            ),
        )
        assert empty_job.failure is None
        assert empty_job.planned == 0
        assert empty_job.comparison_count == 0
        assert empty_progress == [(0, 0)]

        overlay_one = build_overlay([rd_file_one], FileKind.MTO_XLSX)
        _store_scan(database, [rd_file_one], [robot_file_one], overlay_one)
        uncached_plan = build_mto_compare_plan(
            database,
            overlay_one,
            [robot_file_one],
        )
        assert len(uncached_plan.pairs) == 1
        assert uncached_plan.pairs[0].key == key_one
        assert uncached_plan.skipped_cache_hits == 0

        config = _catalog_config(root, database.path)
        assert database.list_mto_comparisons() == []
        compare_outcome = execute_mto_compare(
            config,
            loader=_loader(
                {
                    str(rd_path_one): [BASE_ROW],
                    str(robot_path_one): [BASE_ROW],
                }
            ),
        )
        assert compare_outcome.failure is None
        assert compare_outcome.planned == 1
        assert compare_outcome.comparison_count == 1
        assert len(database.list_mto_comparisons()) == 1

        cached_data = {
            str(rd_path_one): [BASE_ROW],
            str(robot_path_one): [BASE_ROW],
        }
        database.compare_and_store_mto(
            overlay_one,
            [robot_file_one],
            loader=_loader(cached_data),
        )
        cached_plan = build_mto_compare_plan(
            database,
            overlay_one,
            [robot_file_one],
        )
        assert cached_plan.pairs == ()
        assert cached_plan.skipped_cache_hits == 1

        overlay_two = build_overlay([rd_file_one, rd_file_two], FileKind.MTO_XLSX)
        priority_db = CatalogDatabase(root / "runtime_priority" / "catalog.sqlite")
        priority_db.initialize()
        _store_scan(
            priority_db,
            [rd_file_one, rd_file_two],
            [robot_file_one, robot_file_two],
            overlay_two,
        )
        priority_plan = build_mto_compare_plan(
            priority_db,
            overlay_two,
            [robot_file_one, robot_file_two],
            priority_keys=(key_two,),
        )
        assert len(priority_plan.pairs) == 2
        assert priority_plan.pairs[0].key == key_two
        assert priority_plan.pairs[1].key == key_one

        scoped_plan = build_mto_compare_plan(
            priority_db,
            overlay_two,
            [robot_file_one, robot_file_two],
            scope_keys=(key_one,),
        )
        assert len(scoped_plan.pairs) == 1
        assert scoped_plan.pairs[0].key == key_one
        assert scoped_plan.scope_limited

        changed_rd = _parsed(rd_path_one, SourceKind.RD, mtime_ns=999)
        changed_overlay = build_overlay([changed_rd, rd_file_two], FileKind.MTO_XLSX)
        stat_mismatch_plan = build_mto_compare_plan(
            database,
            changed_overlay,
            [robot_file_one, robot_file_two],
        )
        assert key_one in stat_mismatch_plan.document_keys
        assert key_two in stat_mismatch_plan.document_keys
        assert stat_mismatch_plan.skipped_cache_hits == 0

        cancel_db = CatalogDatabase(root / "runtime_cancel" / "catalog.sqlite")
        cancel_db.initialize()
        cancel_overlay = build_overlay(
            [rd_file_one, rd_file_two],
            FileKind.MTO_XLSX,
        )
        _store_scan(
            cancel_db,
            [rd_file_one, rd_file_two],
            [robot_file_one, robot_file_two],
            cancel_overlay,
        )
        cancel_config = _catalog_config(root, cancel_db.path)
        immediate_cancel = execute_mto_compare(
            cancel_config,
            is_cancelled=lambda: True,
            loader=_loader(
                {
                    str(rd_path_one): [BASE_ROW],
                    str(robot_path_one): [BASE_ROW],
                    str(rd_path_two): [BASE_ROW],
                    str(robot_path_two): [BASE_ROW],
                }
            ),
        )
        assert immediate_cancel.cancelled
        assert immediate_cancel.comparison_count == 0
        assert len(immediate_cancel.remaining_keys) == 2
        assert cancel_db.list_mto_comparisons() == []

        compare_calls = {"count": 0}

        def cancel_after_first() -> bool:
            return compare_calls["count"] >= 1

        def track_progress(completed: int, total: int) -> None:
            compare_calls["count"] = completed

        mid_cancel = execute_mto_compare(
            cancel_config,
            batch_size=1,
            is_cancelled=cancel_after_first,
            progress=track_progress,
            loader=_loader(
                {
                    str(rd_path_one): [BASE_ROW],
                    str(robot_path_one): [BASE_ROW],
                    str(rd_path_two): [BASE_ROW],
                    str(robot_path_two): [BASE_ROW],
                }
            ),
        )
        assert mid_cancel.cancelled
        assert mid_cancel.comparison_count >= 1
        assert len(mid_cancel.remaining_keys) >= 1
        assert len(cancel_db.list_mto_comparisons()) == 1

        scoped_config = _catalog_config(root, priority_db.path)
        scoped_outcome = execute_mto_compare(
            scoped_config,
            document_keys={key_one},
            loader=_loader(
                {
                    str(rd_path_one): [BASE_ROW],
                    str(robot_path_one): [BASE_ROW],
                    str(rd_path_two): [BASE_ROW],
                    str(robot_path_two): [BASE_ROW],
                }
            ),
        )
        assert scoped_outcome.failure is None
        assert scoped_outcome.planned == 1
        assert scoped_outcome.comparison_count == 1
        scoped_rows = priority_db.list_mto_comparisons()
        assert len(scoped_rows) == 1
        assert key_one[0] in scoped_rows[0]["rd_path"].lower()

        cli_db = CatalogDatabase(root / "runtime_cli" / "catalog.sqlite")
        cli_db.initialize()
        _store_scan(
            cli_db,
            [rd_file_one],
            [robot_file_one],
            overlay_one,
        )
        cli_config = _catalog_config(root, cli_db.path)
        job_path = root / "runtime_cli" / "mto_job_test.json"
        job_path.write_text(
            json.dumps(
                config_to_mto_job_dict(cli_config, batch_size=1),
                ensure_ascii=False,
            ),
            encoding="utf-8",
        )
        import os

        env = os.environ.copy()
        env["PYTHONPATH"] = str(REPO_ROOT)
        env["PYTHONUTF8"] = "1"
        cli_result = subprocess.run(
            [
                sys.executable,
                "-m",
                "rd_catalog.mto_compare_cli",
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
        assert '"t": "done"' in cli_result.stdout or '"t":"done"' in cli_result.stdout

    print("RD catalog MTO compare job: OK")


if __name__ == "__main__":
    main()
