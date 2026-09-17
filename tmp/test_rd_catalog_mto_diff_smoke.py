"""Local smoke checks for side-effect-free RD Catalog MTO comparison."""

from __future__ import annotations

import sys
import tempfile
from pathlib import Path

import openpyxl

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from rd_catalog.db import CatalogDatabase
from rd_catalog.config import CatalogConfig
from rd_catalog.models import (
    FileKind,
    MtoContentStatus,
    MtoIssueKind,
    MtoReadinessStatus,
    ScanSummary,
    SourceKind,
    SourceScanResult,
)
from rd_catalog.mto_diff import (
    CANONICAL_MTO_FIELDS,
    MtoPair,
    compare_mto_pair,
    load_canonical_mto,
    pair_current_mto,
)
from rd_catalog.overlay import build_overlay
from rd_catalog.parse import parse_catalog_file, parse_transfer_folder
from rd_catalog.mto_compare_job import execute_mto_compare
from rd_catalog.scan_thread import ScanThread


BASE_ROW = {
    "CODE": "BCC0001",
    "UNITS": "шт",
    "VALUES": 1,
    "TAGS": ["2225-SH-01-S-UZ-0001", "2225-SH-01-S-UZ-0002"],
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


def _snapshot(directory: Path) -> set[str]:
    return {
        str(path.relative_to(directory))
        for path in directory.rglob("*")
        if path.is_file()
    }


def main() -> None:
    """Exercise normalization, diff structure, pairing, and DB cache reuse."""

    with tempfile.TemporaryDirectory(prefix="rd_catalog_mto_diff_") as temp:
        root = Path(temp)
        rd_dir = root / "rd" / "01_рев.01_2225-KSB"
        robot_dir = root / "robot"
        rd_dir.mkdir(parents=True)
        robot_dir.mkdir()
        name = "AGCC.287-2225-KSB.MTO-0001_01_RU.xlsx"
        rd_path = rd_dir / name
        robot_path = robot_dir / name
        robot_extra_path = (
            robot_dir / "AGCC.287-2225-KSB.MTO-0002_01_RU.xlsx"
        )
        _write_mto(rd_path, [BASE_ROW])
        _write_mto(robot_path, [BASE_ROW])
        _write_mto(robot_extra_path, [BASE_ROW])
        rd_file = _parsed(rd_path, SourceKind.RD)
        robot_file = _parsed(robot_path, SourceKind.ROBOT)
        robot_extra_file = _parsed(robot_extra_path, SourceKind.ROBOT)
        pair = MtoPair(
            ("2225-ksb", "mto-0001"),
            rd_file,
            (robot_file,),
        )
        before = _snapshot(root)

        # The production adapter uses get_mto_std_from_file and keeps one
        # material row without importing report-writing comparison modules.
        real_document = load_canonical_mto(rd_path)
        assert len(real_document.rows) == 1
        assert real_document.rows[0].as_dict() == {
            "CODE": "BCC0001",
            "UNITS": "шт",
            "VALUES": "1",
            "TAGS": ["2225-SH-01-S-UZ-0001", "2225-SH-01-S-UZ-0002"],
            "NAME": "Прибор",
            "VENDOR": "Vendor",
            "TYPE_MARK": "Type A",
        }
        assert "RFQ.Value_Compare" not in sys.modules

        reordered_data = {
            str(rd_path): [BASE_ROW, {**BASE_ROW, "CODE": "BCC0002"}],
            str(robot_path): [{**BASE_ROW, "CODE": "BCC0002"}, BASE_ROW],
        }
        equal = compare_mto_pair(pair, loader=_loader(reordered_data))
        assert equal.readiness is MtoReadinessStatus.READY
        assert equal.content_status is MtoContentStatus.EQUAL
        assert equal.diff.is_empty

        replacements = {
            "CODE": "BCC9999",
            "UNITS": "компл.",
            "VALUES": 2,
            "TAGS": ["2225-SH-01-S-UZ-9999"],
            "NAME": "Другой прибор",
            "VENDOR": "Other vendor",
            "TYPE_MARK": "Type B",
        }
        for field in CANONICAL_MTO_FIELDS:
            changed_row = {**BASE_ROW, field: replacements[field]}
            result = compare_mto_pair(
                pair,
                loader=_loader(
                    {
                        str(rd_path): [BASE_ROW],
                        str(robot_path): [changed_row],
                    }
                ),
            )
            assert result.readiness is MtoReadinessStatus.WARNING, field
            assert result.content_status is MtoContentStatus.DIFF, field
            assert result.diff.changed, field
            assert result.diff.changed[0].fields[0].field == field, field

        duplicate = compare_mto_pair(
            pair,
            loader=_loader(
                {
                    str(rd_path): [BASE_ROW, BASE_ROW],
                    str(robot_path): [BASE_ROW],
                }
            ),
        )
        assert duplicate.content_status is MtoContentStatus.DIFF
        assert len(duplicate.diff.removed) == 1

        numeric = compare_mto_pair(
            pair,
            loader=_loader(
                {
                    str(rd_path): [{**BASE_ROW, "VALUES": 1}],
                    str(robot_path): [{**BASE_ROW, "VALUES": "1,000"}],
                }
            ),
        )
        assert numeric.content_status is MtoContentStatus.EQUAL

        overlay = build_overlay([rd_file], FileKind.MTO_XLSX)
        pairing_with_extra = pair_current_mto(
            overlay, (robot_file, robot_extra_file)
        )
        assert pairing_with_extra.robot_extras == (robot_extra_file,)
        missing_pair = pair_current_mto(overlay, ()).pairs[0]
        missing = compare_mto_pair(missing_pair, loader=_loader({}))
        assert missing.readiness is MtoReadinessStatus.BLOCKED
        assert missing.issues == (MtoIssueKind.ROBOT_MISSING,)

        duplicate_robot = _parsed(robot_path, SourceKind.ROBOT, mtime_ns=101)
        duplicate_pair = pair_current_mto(
            overlay, (robot_file, duplicate_robot)
        ).pairs[0]
        blocked_duplicate = compare_mto_pair(
            duplicate_pair,
            loader=_loader({}),
        )
        assert blocked_duplicate.readiness is MtoReadinessStatus.BLOCKED
        assert blocked_duplicate.issues == (MtoIssueKind.ROBOT_DUPLICATE,)

        summary = ScanSummary(
            sources={
                SourceKind.RD: SourceScanResult(
                    source=SourceKind.RD,
                    root=str(rd_dir.parent),
                    files=[rd_file],
                ),
                SourceKind.ROBOT: SourceScanResult(
                    source=SourceKind.ROBOT,
                    root=str(robot_dir),
                    files=[robot_file, robot_extra_file],
                ),
            },
            rd_mto_overlay=overlay,
        )
        database = CatalogDatabase(root / "runtime" / "catalog.sqlite")
        database.initialize()
        database.store_scan(summary)
        cached_data = {
            str(rd_path): [BASE_ROW],
            str(robot_path): [BASE_ROW],
        }
        first = database.compare_and_store_mto(
            overlay,
            [robot_file],
            loader=_loader(cached_data),
        )[0]
        assert not first.cache_hit

        def fail_loader(path: str | Path):
            raise AssertionError(f"cache failed for {path}")

        second = database.compare_and_store_mto(
            overlay,
            [robot_file],
            loader=fail_loader,
        )[0]
        assert second.cache_hit
        assert second.readiness is MtoReadinessStatus.READY
        assert database.list_mto_comparisons()[0]["status"] == "ready"
        robot_extras = database.list_robot_extra_mto()
        assert len(robot_extras) == 1
        assert robot_extras[0]["robot_path"] == str(robot_extra_path)
        assert robot_extras[0]["status"] == "warning"
        assert robot_extras[0]["diff"]["issues"] == [
            MtoIssueKind.ROBOT_EXTRA.value
        ]
        cancelled = database.compare_and_store_mto(
            overlay,
            [robot_file],
            loader=fail_loader,
            is_cancelled=lambda: True,
        )
        assert cancelled == []

        incremental_db = root / "runtime_incremental" / "catalog.sqlite"
        config = CatalogConfig(
            rd_root=rd_dir.parent,
            sq_root=root / "sq",
            robot_root=robot_dir,
            runtime_dir=incremental_db.parent,
            db_path=incremental_db,
            robot_flat_structure=True,
            skip_dirs=("old", "tmp"),
        )
        incremental_database = CatalogDatabase(incremental_db)
        incremental_database.initialize()
        rd_only = ScanThread(config, (SourceKind.RD,))
        rd_only.run()
        assert rd_only.failure is None
        assert len(
            incremental_database.reconstruct_current_rd_mto_overlay().current
        ) == 1
        assert incremental_database.list_mto_comparisons() == []
        assert rd_only.comparison_count == 0
        assert rd_only.mto_compare_skipped is True
        assert rd_only.compare_enqueue_scope == "full"
        rd_compare = execute_mto_compare(config)
        assert rd_compare.failure is None
        assert incremental_database.list_mto_comparisons()[0]["status"] == "blocked"

        robot_only = ScanThread(config, (SourceKind.ROBOT,))
        robot_only.run()
        assert robot_only.failure is None
        assert incremental_database.list_mto_comparisons()[0]["status"] == "blocked"
        assert len(incremental_database.list_present_robot_mto_files()) == 2
        assert robot_only.comparison_count == 0
        assert robot_only.mto_compare_skipped is True
        assert robot_only.compare_enqueue_scope == "full"
        robot_compare = execute_mto_compare(config)
        assert robot_compare.failure is None
        comparison_rows = incremental_database.list_mto_comparisons()
        assert comparison_rows[0]["status"] == "ready"

        after = _snapshot(root)
        expected_new = {
            str(Path("runtime", "catalog.sqlite")),
            str(Path("runtime_incremental", "catalog.sqlite")),
        }
        assert after - before == expected_new, after - before
    print("RD catalog MTO diff smoke: OK")


if __name__ == "__main__":
    main()
