"""Local smoke checks for incremental scan and baseline semantics."""

from __future__ import annotations

import tempfile
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from rd_catalog.config import CatalogConfig
from rd_catalog.db import CatalogDatabase
from rd_catalog.models import ReviewState, ScanRunStatus, SourceError, SourceKind
from rd_catalog.scan import scan_catalog, scan_robot_source


def _config(root: Path, db_path: Path) -> CatalogConfig:
    return CatalogConfig(
        rd_root=root,
        sq_root=root / "Ответы на SQ запросы",
        robot_root=root / "robot",
        runtime_dir=db_path.parent,
        db_path=db_path,
        robot_flat_structure=True,
        skip_dirs=("old", "tmp", "архив"),
    )


def main() -> None:
    """Run two local scans and verify add/change/missing transitions."""

    with tempfile.TemporaryDirectory(prefix="rd_catalog_incremental_") as temp:
        root = Path(temp, "rd")
        transfer = (
            root / "2225" / "KSB" / "Для передачи" / "01_рев.01_2225-KSB_as-built"
        )
        transfer.mkdir(parents=True)
        file_a = transfer / "AGCC.287-2225-KSB.OD-0001_01_RU.pdf"
        file_b = transfer / "AGCC.287-2225-KSB.OD-0002_01_RU.pdf"
        file_a.write_bytes(b"a")
        file_b.write_bytes(b"b")
        sq_dir = root / "Ответы на SQ запросы"
        sq_dir.mkdir()
        (sq_dir / "AGCC.287-2225-KSB.OD-0099_01_RU.pdf").write_bytes(b"sq")

        config = _config(root, Path(temp, "runtime", "catalog.sqlite"))
        database = CatalogDatabase(config.db_path)
        database.initialize()

        first = scan_catalog(config, sources=(SourceKind.RD,))
        assert len(first.files) == 2
        assert all(file.transfer and file.transfer.is_as_build for file in first.files)
        first_run = database.store_scan(first)
        assert first_run == 1
        assert all(
            row.review_state is ReviewState.ACKNOWLEDGED
            for row in database.list_files()
        )

        file_a.write_bytes(b"changed-size")
        file_b.unlink()
        file_c = transfer / "AGCC.287-2225-KSB.OD-0003_01_RU.pdf"
        file_c.write_bytes(b"c")

        second = scan_catalog(config, sources=(SourceKind.RD,))
        database.store_scan(second)
        records = {record.path: record for record in database.list_files()}
        assert records[str(file_a)].review_state is ReviewState.PENDING
        assert records[str(file_a)].present
        assert records[str(file_b)].review_state is ReviewState.PENDING
        assert not records[str(file_b)].present
        assert records[str(file_c)].review_state is ReviewState.PENDING
        assert records[str(file_c)].present

        overlay_before_partial = database.current_overlay()
        unavailable = _config(
            Path(temp, "unavailable_rd"),
            config.db_path,
        )
        partial = scan_catalog(unavailable, sources=(SourceKind.RD,))
        assert partial.status is ScanRunStatus.PARTIAL
        database.store_scan(partial)
        after_partial = {record.path: record for record in database.list_files()}
        assert after_partial[str(file_a)].present
        assert after_partial[str(file_c)].present
        assert database.current_overlay() == overlay_before_partial

        broken = Path(temp, "rd", "broken_unreadable")
        broken.mkdir()
        file_c.unlink()
        nested = scan_catalog(config, sources=(SourceKind.RD,))
        nested.sources[SourceKind.RD].errors.append(
            SourceError(SourceKind.RD, str(broken), "access denied")
        )
        nested.status = ScanRunStatus.PARTIAL
        database.store_scan(nested)
        nested_records = {record.path: record for record in database.list_files()}
        assert nested_records[str(file_a)].present
        assert not nested_records[str(file_c)].present
        assert nested_records[str(file_c)].review_state is ReviewState.PENDING

        robot_root = Path(temp, "robot")
        valid_robot_dir = robot_root / "2225" / "KSB"
        invalid_robot_dir = robot_root / "wrong"
        valid_robot_dir.mkdir(parents=True)
        invalid_robot_dir.mkdir()
        robot_name = "AGCC.287-2225-KSB.MTO-0001_01_RU.xlsx"
        (valid_robot_dir / robot_name).write_bytes(b"robot")
        (invalid_robot_dir / robot_name).write_bytes(b"excluded")
        robot_result = scan_robot_source(robot_root, flat_structure=False)
        assert len(robot_result.files) == 1
        assert robot_result.excluded_files == 1

        other_dir = robot_root / "1513" / "POS"
        other_dir.mkdir(parents=True)
        other_name = "AGCC.287-1513-POS.MTO-0001_01_RU.xlsx"
        other_file = other_dir / other_name
        other_file.write_bytes(b"other")
        robot_config = CatalogConfig(
            rd_root=root,
            sq_root=root / "Ответы на SQ запросы",
            robot_root=robot_root,
            runtime_dir=config.runtime_dir,
            db_path=config.db_path,
            robot_flat_structure=False,
            skip_dirs=("old", "tmp", "архив"),
        )
        full_robot = scan_catalog(robot_config, sources=(SourceKind.ROBOT,))
        assert len(full_robot.files) == 2
        database.store_scan(full_robot)
        old_ksb = valid_robot_dir / robot_name
        old_ksb.unlink()
        new_ksb = valid_robot_dir / "AGCC.287-2225-KSB.MTO-0001_02_RU.xlsx"
        new_ksb.write_bytes(b"new-rev")
        scoped = scan_catalog(
            robot_config,
            sources=(SourceKind.ROBOT,),
            robot_subtree=valid_robot_dir,
        )
        assert scoped.sources[SourceKind.ROBOT].subtree
        assert {Path(file.path).name for file in scoped.files} == {new_ksb.name}
        database.store_scan(scoped)
        robot_records = {
            record.path: record
            for record in database.list_files(source=SourceKind.ROBOT)
        }
        assert not robot_records[str(old_ksb)].present
        assert robot_records[str(new_ksb)].present
        assert robot_records[str(other_file)].present

        from rd_catalog.scan_job import execute_scan

        job = execute_scan(
            robot_config,
            (SourceKind.ROBOT,),
            robot_subtree=valid_robot_dir,
        )
        assert job.failure is None
        assert job.summary is not None
        assert job.summary.sources[SourceKind.ROBOT].subtree
        assert job.comparison_count == 0
        assert job.mto_compare_skipped is True
        assert job.compare_enqueue_scope == "subtree"

        mixed_root = Path(temp, "mixed_rd")
        own_mark = mixed_root / "2630" / "KSB"
        stray_mark = mixed_root / "2612" / "KSB"
        own_file = (
            own_mark
            / "Для передачи"
            / "08_рев.01"
            / "AGCC.287-2630-KSB.OD-0001_01_RU.pdf"
        )
        stray_file = (
            stray_mark
            / "Для передачи"
            / "09_рев.01"
            / "AGCC.287-2630-KSB.WIR-0010_01_RU.pdf"
        )
        own_file.parent.mkdir(parents=True)
        stray_file.parent.mkdir(parents=True)
        own_file.write_bytes(b"own")
        stray_file.write_bytes(b"stray")
        mixed_config = CatalogConfig(
            rd_root=mixed_root,
            sq_root=mixed_root / "Ответы на SQ запросы",
            robot_root=mixed_root / "robot",
            runtime_dir=config.runtime_dir,
            db_path=Path(temp, "runtime", "mixed.sqlite"),
            robot_flat_structure=True,
            skip_dirs=("old", "tmp", "архив"),
        )
        mixed_db = CatalogDatabase(mixed_config.db_path)
        mixed_db.initialize()
        (mixed_config.sq_root).mkdir()
        seeded = scan_catalog(
            mixed_config,
            sources=(SourceKind.RD,),
            rd_subtree=(str(own_mark), str(stray_mark)),
        )
        assert seeded.sources[SourceKind.RD].subtrees
        assert len(seeded.sources[SourceKind.RD].subtrees) == 2
        mixed_db.store_scan(seeded)
        stray_file.unlink()
        own_only = scan_catalog(
            mixed_config,
            sources=(SourceKind.RD,),
            rd_subtree=str(own_mark),
        )
        mixed_db.store_scan(own_only)
        own_only_rows = {
            record.path: record
            for record in mixed_db.list_files(source=SourceKind.RD)
        }
        assert own_only_rows[str(own_file)].present
        assert own_only_rows[str(stray_file)].present
        both = scan_catalog(
            mixed_config,
            sources=(SourceKind.RD,),
            rd_subtree=(str(own_mark), str(stray_mark)),
        )
        mixed_db.store_scan(both)
        both_rows = {
            record.path: record
            for record in mixed_db.list_files(source=SourceKind.RD)
        }
        assert both_rows[str(own_file)].present
        assert not both_rows[str(stray_file)].present

    print("RD catalog incremental smoke: OK")


if __name__ == "__main__":
    main()
