"""Local checks for moving an SQ kit folder into a new RD transfer."""

from __future__ import annotations

import os
import sys
import tempfile
from datetime import datetime
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from rd_catalog.config import CatalogConfig
from rd_catalog.db import CatalogDatabase
from rd_catalog.kits import (
    KitMatrixRow,
    KitSummary,
    SourceKitSnapshot,
    summary_tooltip,
)
from rd_catalog.models import FileKind, FileRecord, ReviewState, SourceKind
from rd_catalog.parse import is_transfer_folder_name, parse_transfer_folder
from rd_catalog.scan import scan_catalog, scan_document_source
from rd_catalog.sq_to_rd import (
    SqToRdError,
    execute_sq_to_rd_transfer,
    folder_matches_mark,
    next_transfer_sequence,
    plan_sq_to_rd_transfer,
    transfer_folder_name,
)


def _record(
    file_id: int,
    *,
    path: str,
    source: SourceKind,
    title: str = "8950",
    mark: str = "POS1",
    revision: str = "04",
    appendix: str | None = "04",
    mtime_ns: int = 1,
) -> FileRecord:
    name = Path(path).name
    return FileRecord(
        id=file_id,
        path=path,
        path_key=path.casefold(),
        source=source,
        present=True,
        review_state=ReviewState.ACKNOWLEDGED,
        first_seen_run_id=1,
        last_seen_run_id=1,
        data={
            "file_kind": FileKind.PDF.value,
            "name": name,
            "parse_status": "parsed",
            "title": title,
            "mark": mark,
            "title_system": f"{title}-{mark}",
            "discipline_block": "OD-0001",
            "revision": revision,
            "appendix": appendix,
            "mtime_ns": mtime_ns,
        },
    )


def _write(path: Path, payload: bytes = b"sq") -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_bytes(payload)


def _utime(path: Path, when: datetime) -> int:
    stamp_ns = int(when.timestamp() * 1_000_000_000)
    os.utime(path, ns=(stamp_ns, stamp_ns))
    return stamp_ns


def main() -> None:
    """Plan and execute SQ→RD moves on a local tree; never touch UNC."""

    assert folder_matches_mark("POS1", "POS1")
    assert folder_matches_mark("12_POS1", "POS1")
    assert folder_matches_mark("04-SOT", "SOT")
    assert not folder_matches_mark("12_POS1_old", "POS1")
    assert not folder_matches_mark("12_POS2", "POS1")
    assert next_transfer_sequence(()) == 1
    assert next_transfer_sequence((10, 9)) == 11
    when = datetime(2026, 8, 15, 14, 30)
    name = transfer_folder_name(11, "04-AN04", when)
    assert name == "11_рев.04-AN04_от_2026.08.15"
    assert is_transfer_folder_name(name, under_gate=True)
    parsed = parse_transfer_folder(name, under_gate=True)
    assert parsed.sequence == 11
    assert parsed.revision == "04"
    assert parsed.appendix == "04"

    aligned = KitMatrixRow(
        title="8950",
        mark="POS1",
        title_system="8950-POS1",
        rd=SourceKitSnapshot(present=True, revision_text="04-AN01"),
        robot=SourceKitSnapshot(),
        sq=SourceKitSnapshot(present=True, revision_text="04-AN04"),
        google=None,
        issuance=None,
        flags=(),
        summary=KitSummary.ALIGNED,
    )
    assert "ревизии Google↔РД совпадают" in summary_tooltip(aligned)
    review = KitMatrixRow(
        title="2869",
        mark="SKUD",
        title_system="2869-SKUD",
        rd=SourceKitSnapshot(present=True),
        robot=SourceKitSnapshot(),
        sq=SourceKitSnapshot(),
        google=None,
        issuance=None,
        flags=(),
        summary=KitSummary.TRANSFER_REVIEW,
        transfer_review_notes=("10_рев.04 vs 09_рев.04-AN01",),
    )
    tip = summary_tooltip(review)
    assert "Текущий состав при этом не откатывается" in tip
    assert "10_рев.04 vs 09_рев.04-AN01" in tip

    with tempfile.TemporaryDirectory(prefix="rd_catalog_sq_to_rd_") as temp:
        root = Path(temp)
        rd_root = root / "rd"
        sq_root = rd_root / "Ответы на SQ запросы"
        gate = rd_root / "8950" / "12_POS1" / "Для передачи"
        old_transfer = gate / "10_рев.04-AN01_AGCC.287-8950-POS1" / "PDF"
        rd_pdf = old_transfer / "AGCC.287-8950-POS1.OD-0001_04-AN01_RU.pdf"
        _write(rd_pdf, b"rd")
        sq_pdf = (
            sq_root
            / "8950"
            / "POS1"
            / "PDF"
            / "AGCC.287-8950-POS1.OD-0001_04-AN04_RU.pdf"
        )
        sq_dwg = (
            sq_root
            / "8950"
            / "POS1"
            / "DWG"
            / "AGCC.287-8950-POS1.OD-0001_04-AN04_RU.dwg"
        )
        _write(sq_pdf, b"sq-pdf")
        _write(sq_dwg, b"sq-dwg")
        stamp = _utime(sq_pdf, when)
        _utime(sq_dwg, datetime(2026, 8, 14, 10, 0))

        records = [
            _record(
                1,
                path=str(rd_pdf),
                source=SourceKind.RD,
                revision="04",
                appendix="01",
            ),
            _record(2, path=str(sq_pdf), source=SourceKind.SQ),
        ]
        plan = plan_sq_to_rd_transfer(
            title="8950",
            mark="POS1",
            records=records,
            sq_paths=(str(sq_pdf),),
            rd_paths=(str(rd_pdf),),
            sq_revision_text="04-AN04",
            sq_max_mtime_ns=stamp,
            rd_root=rd_root,
            sq_root=sq_root,
        )
        assert plan.sequence == 11
        assert plan.transfer_name == "11_рев.04-AN04_от_2026.08.15"
        assert plan.created_gate is False
        assert Path(plan.mark_folder).name == "12_POS1"
        assert set(plan.child_names) == {"PDF", "DWG"}

        result = execute_sq_to_rd_transfer(plan)
        dest = Path(result.destination_folder)
        assert dest.is_dir()
        assert (dest / "PDF" / sq_pdf.name).is_file()
        assert (dest / "DWG" / sq_dwg.name).is_file()
        assert not sq_pdf.exists()
        assert result.source_removed
        assert not Path(plan.source_folder).exists()
        assert rd_pdf.is_file()

        try:
            plan_sq_to_rd_transfer(
                title="8950",
                mark="POS1",
                records=records,
                sq_paths=(str(sq_pdf),),
                rd_paths=(str(rd_pdf),),
                sq_revision_text="04-AN04",
                rd_root=rd_root,
                sq_root=sq_root,
            )
            raise AssertionError("expected missing SQ folder")
        except SqToRdError as exc:
            assert "Нет файлов SQ" in str(exc) or "не найдена" in str(exc)

        other_sq = (
            sq_root
            / "8950"
            / "mixed"
            / "AGCC.287-8950-POS1.OD-0002_04-AN04_RU.pdf"
        )
        foreign = (
            sq_root
            / "8950"
            / "mixed"
            / "AGCC.287-8950-SKUD.OD-0001_01_RU.pdf"
        )
        _write(other_sq, b"pos1")
        _write(foreign, b"skud")
        mixed_records = [
            _record(10, path=str(other_sq), source=SourceKind.SQ),
            _record(
                11,
                path=str(foreign),
                source=SourceKind.SQ,
                mark="SKUD",
                revision="01",
                appendix=None,
            ),
        ]
        try:
            plan_sq_to_rd_transfer(
                title="8950",
                mark="POS1",
                records=mixed_records,
                sq_paths=(str(other_sq),),
                rd_paths=(),
                sq_revision_text="04-AN04",
                rd_root=rd_root,
                sq_root=sq_root,
            )
            raise AssertionError("expected foreign-kit refusal")
        except SqToRdError as exc:
            assert "другого комплекта" in str(exc)

        loose = sq_root / "AGCC.287-9110-KSB.OD-0001_03_RU.pdf"
        _write(loose, b"loose")
        try:
            plan_sq_to_rd_transfer(
                title="9110",
                mark="KSB",
                records=[
                    _record(
                        20,
                        path=str(loose),
                        source=SourceKind.SQ,
                        title="9110",
                        mark="KSB",
                        revision="03",
                        appendix=None,
                    )
                ],
                sq_paths=(str(loose),),
                rd_paths=(),
                sq_revision_text="03",
                rd_root=rd_root,
                sq_root=sq_root,
            )
            raise AssertionError("expected sq_root refusal")
        except SqToRdError as exc:
            assert "весь каталог SQ" in str(exc)

        fresh_sq = (
            sq_root
            / "9192"
            / "SKUD"
            / "AGCC.287-9192-SKUD.OD-0001_01-AN02_RU.pdf"
        )
        _write(fresh_sq, b"new-title")
        _utime(fresh_sq, datetime(2026, 1, 2, 8, 0))
        created = plan_sq_to_rd_transfer(
            title="9192",
            mark="SKUD",
            records=[
                _record(
                    30,
                    path=str(fresh_sq),
                    source=SourceKind.SQ,
                    title="9192",
                    mark="SKUD",
                    revision="01",
                    appendix="02",
                )
            ],
            sq_paths=(str(fresh_sq),),
            rd_paths=(),
            sq_revision_text="01-AN02",
            rd_root=rd_root,
            sq_root=sq_root,
        )
        assert created.sequence == 1
        assert created.created_gate
        assert created.created_mark
        assert created.transfer_name == "01_рев.01-AN02_от_2026.01.02"
        executed = execute_sq_to_rd_transfer(created)
        new_dest = Path(executed.destination_folder)
        assert new_dest.parent.name == "Для передачи"
        assert new_dest.parent.parent.name == "SKUD"
        assert (new_dest / fresh_sq.name).is_file()

    with tempfile.TemporaryDirectory(prefix="rd_catalog_sq_to_rd_scan_") as temp:
        root = Path(temp)
        rd_root = root / "rd"
        sq_root = rd_root / "Ответы на SQ запросы"
        robot_root = root / "robot"
        runtime = root / "runtime"
        runtime.mkdir()
        other_pdf = (
            rd_root
            / "2225"
            / "10_KSB"
            / "Для передачи"
            / "01_рев.01"
            / "AGCC.287-2225-KSB.OD-0001_01_RU.pdf"
        )
        old_pdf = (
            rd_root
            / "8950"
            / "12_POS1"
            / "Для передачи"
            / "10_рев.04"
            / "PDF"
            / "AGCC.287-8950-POS1.OD-0001_04_RU.pdf"
        )
        new_pdf = (
            rd_root
            / "8950"
            / "12_POS1"
            / "Для передачи"
            / "11_рев.04-AN04_от_2026.08.15"
            / "PDF"
            / "AGCC.287-8950-POS1.OD-0001_04-AN04_RU.pdf"
        )
        sq_keep = (
            sq_root / "2225" / "KSB" / "AGCC.287-2225-KSB.OD-0002_01_RU.pdf"
        )
        sq_moved = (
            sq_root
            / "8950"
            / "POS1"
            / "AGCC.287-8950-POS1.OD-0001_04-AN04_RU.pdf"
        )
        _write(other_pdf, b"other-rd")
        _write(old_pdf, b"old-rd")
        _write(sq_keep, b"sq-keep")
        _write(sq_moved, b"sq-moved")
        config = CatalogConfig(
            rd_root=rd_root,
            sq_root=sq_root,
            robot_root=robot_root,
            runtime_dir=runtime,
            db_path=runtime / "catalog.sqlite",
            robot_flat_structure=True,
            skip_dirs=("old", "tmp"),
        )
        database = CatalogDatabase(config.db_path)
        database.initialize()
        first = scan_catalog(config, sources=(SourceKind.RD, SourceKind.SQ))
        database.store_scan(first)
        overlay_before = {
            row["document_key"]: row["path"]
            for row in database.current_overlay()
        }
        other_key = next(
            key for key in overlay_before if "2225" in key.casefold()
        )
        pos_key = next(
            key for key in overlay_before if "8950" in key.casefold()
        )
        assert Path(overlay_before[pos_key]).name.endswith("_04_RU.pdf")

        scoped_rd = scan_document_source(
            rd_root,
            SourceKind.RD,
            skip_dirs=config.skip_dirs,
            sq_root=sq_root,
            subtree=old_pdf.parent.parent,
        )
        assert scoped_rd.subtree
        assert len(scoped_rd.files) == 1
        assert Path(scoped_rd.files[0].path) == old_pdf

        _write(new_pdf, b"new-rd")
        sq_moved.unlink()
        Path(sq_moved.parent).rmdir()
        gate = new_pdf.parent.parent.parent
        assert gate.name == "Для передачи"
        scoped = scan_catalog(
            config,
            sources=(SourceKind.RD, SourceKind.SQ),
            rd_subtree=gate,
            sq_subtree=sq_moved.parent,
        )
        assert scoped.sources[SourceKind.RD].subtree
        assert scoped.sources[SourceKind.SQ].subtree
        assert {Path(item.path) for item in scoped.files} == {old_pdf, new_pdf}
        database.store_scan(scoped)
        overlay_after = {
            row["document_key"]: row["path"]
            for row in database.current_overlay()
        }
        assert overlay_after[other_key] == overlay_before[other_key]
        assert Path(overlay_after[pos_key]) == new_pdf
        present = {
            record.path: record.present
            for record in database.list_files(source=SourceKind.SQ)
        }
        assert present[str(sq_keep)]
        assert not present[str(sq_moved)]
        assert database.list_files(source=SourceKind.RD, present_only=True)
        other_still = next(
            record
            for record in database.list_files(source=SourceKind.RD)
            if record.path == str(other_pdf)
        )
        assert other_still.present

    print("RD catalog SQ to RD transfer: OK")


if __name__ == "__main__":
    main()
