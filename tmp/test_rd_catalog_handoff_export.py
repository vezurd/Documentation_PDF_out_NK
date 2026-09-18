"""Local checks for official-kit manager dump (MTO + BBB)."""

from __future__ import annotations

import sys
import tempfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from rd_catalog.db import KitPackageRow, KitPipelineRow
from rd_catalog.handoff_export import (
    HANDOFF_LAYOUT_FLAT,
    HANDOFF_LAYOUT_TITLE_MARK,
    HandoffKitInput,
    build_handoff_copy_plan,
    build_handoff_rows,
    dest_matches_source_fingerprint,
    execute_handoff_copy,
    is_bbb_media_path,
    kit_handoff_eligible,
    load_handoff_destinations,
    remember_handoff_destination,
    validate_handoff_dest,
)
from rd_catalog.kits import format_revision
from rd_catalog.models import FileRecord, ReviewState, SourceKind


def _pipeline(**kwargs) -> KitPipelineRow:
    values = {
        "title": "8950",
        "mark": "SOO1",
        "status": "agreed",
        "code": "A",
        "official_revision_text": "03-AN01",
        "working_revision_text": "03-AN03",
        "working_transfer_names": ("14_рев.03_AN03_AGCC.287-8950-SOO1",),
        "working_sequences": (14,),
    }
    values.update(kwargs)
    return KitPipelineRow(**values)


def _package(*, sequence: int, revision: str, current: bool, folder: str) -> KitPackageRow:
    base = r"\\bcc\eng\PrDoc\РД\8950\14_SOO1\Для передачи"
    return KitPackageRow(
        title="8950",
        mark="SOO1",
        source="rd",
        sequence=sequence,
        transfer_name=folder,
        package_path=rf"{base}\{folder}",
        revision_text=revision,
        is_current=current,
        id=sequence,
    )


def _file(
    file_id: int,
    *,
    folder: str,
    media: str,
    discipline: str,
    revision: str,
    appendix: str | None = None,
    file_kind: str = "pdf",
    size: int = 10,
    mtime_ns: int = 100,
) -> FileRecord:
    rev_text = format_revision(revision, appendix)
    suffix = "xlsx" if file_kind == "mto_xlsx" else (
        "dwg" if file_kind == "source_editable" else "pdf"
    )
    filename = f"AGCC.287-8950-SOO1.{discipline}_{rev_text}_RU.{suffix}"
    path = (
        rf"\\bcc\eng\PrDoc\РД\8950\14_SOO1\Для передачи\{folder}"
        rf"\{media}\{filename}"
    )
    return FileRecord(
        id=file_id,
        path=path,
        path_key=f"rd/{file_id}",
        source=SourceKind.RD,
        present=True,
        review_state=ReviewState.ACKNOWLEDGED,
        first_seen_run_id=1,
        last_seen_run_id=1,
        data={
            "file_kind": file_kind,
            "name": filename,
            "parse_status": "parsed",
            "title": "8950",
            "mark": "SOO1",
            "revision": revision,
            "appendix": appendix,
            "mtime_ns": mtime_ns,
            "disk_mtime_ns": mtime_ns,
            "size": size,
            "transfer_sequence": int(folder.split("_", 1)[0]),
            "transfer_name": folder,
            "discipline_block": discipline,
            "title_system": "8950-SOO1",
        },
    )


def test_eligibility() -> None:
    code_a = HandoffKitInput(
        title="8950",
        mark="SOO1",
        official_revision_text="03-AN01",
        status="agreed",
        code_a=True,
    )
    tdo = HandoffKitInput(
        title="1000",
        mark="KSB",
        official_revision_text="01",
        status="sent_tdo",
        code_a=False,
    )
    other = HandoffKitInput(
        title="2000",
        mark="KSB",
        official_revision_text="01",
        status="not_uploaded",
        code_a=False,
    )
    assert kit_handoff_eligible(code_a, include_tdo=False)
    assert not kit_handoff_eligible(tdo, include_tdo=False)
    assert kit_handoff_eligible(tdo, include_tdo=True)
    assert not kit_handoff_eligible(other, include_tdo=True)


def test_bbb_media_path() -> None:
    assert is_bbb_media_path(r"\\bcc\x\11_рев\BBB\file.pdf")
    assert not is_bbb_media_path(r"\\bcc\x\11_рев\PDF\file.pdf")
    assert not is_bbb_media_path(r"\\bcc\x\11_рев\DWG\file.dwg")


def test_rows_use_official_package_not_is_current() -> None:
    folder11 = "11_рев.03-AN01_от_2026.01.22"
    folder13 = "13_рев.03_AN02_AGCC.287-8950-SOO1"
    folder14 = "14_рев.03_AN03_AGCC.287-8950-SOO1"
    pipeline = _pipeline()
    packages = (
        _package(sequence=11, revision="03-AN01", current=False, folder=folder11),
        _package(sequence=13, revision="03-AN02", current=True, folder=folder13),
        _package(sequence=14, revision="03-AN03", current=False, folder=folder14),
    )
    records = (
        _file(1, folder=folder11, media="DWG", discipline="MTO-0001",
              revision="03", appendix="01", file_kind="mto_xlsx"),
        _file(2, folder=folder11, media="PDF", discipline="BOE-0001",
              revision="03", appendix="01"),
        _file(3, folder=folder11, media="PDF", discipline="BOM-0001",
              revision="03", appendix="01"),
        _file(4, folder=folder11, media="PDF", discipline="BOQ-0001",
              revision="03", appendix="01"),
        _file(5, folder=folder11, media="BBB", discipline="WIR-0001",
              revision="03", appendix="01"),
        _file(6, folder=folder13, media="PDF", discipline="LAY-0003",
              revision="03", appendix="02"),
    )
    kit = HandoffKitInput(
        title="8950",
        mark="SOO1",
        official_revision_text="03-AN01",
        status="agreed",
        code_a=True,
        pipeline=pipeline,
    )
    rows = build_handoff_rows((kit,), packages, records)
    assert len(rows) == 1
    row = rows[0]
    assert folder11 in row.package_path
    assert folder13 not in row.package_path
    assert row.mto.text == "03-AN01"
    assert row.boe.text == "03-AN01"
    assert row.bom.present and row.boq.present
    assert row.bbb_count == 1
    assert row.has_problem is False
    assert any(item[0].endswith(".xlsx") for item in row.copy_files)
    assert any("BBB" in item[0] for item in row.copy_files)
    assert not any("LAY-0003" in item[0] for item in row.copy_files)


def test_missing_boe_is_a_note() -> None:
    folder11 = "11_рев.03-AN01_от_2026.01.22"
    pipeline = _pipeline()
    packages = (
        _package(sequence=11, revision="03-AN01", current=True, folder=folder11),
    )
    records = (
        _file(1, folder=folder11, media="DWG", discipline="MTO-0001",
              revision="03", appendix="01", file_kind="mto_xlsx"),
    )
    kit = HandoffKitInput(
        title="8950",
        mark="SOO1",
        official_revision_text="03-AN01",
        status="agreed",
        code_a=True,
        pipeline=pipeline,
    )
    row = build_handoff_rows((kit,), packages, records)[0]
    assert row.boe.text == "нет"
    assert "нет BOE" in row.notes
    assert row.has_problem is True


def test_dest_guard_and_fingerprint(tmp: Path) -> None:
    rd_root = r"\\bcc\eng\PrDoc\РД"
    try:
        validate_handoff_dest(rf"{rd_root}\8950", rd_root=rd_root)
        raise AssertionError("expected reject")
    except ValueError:
        pass
    dest = tmp / "out"
    dest.mkdir()
    src = tmp / "src.xlsx"
    src.write_bytes(b"abc")
    dest_file = dest / "src.xlsx"
    dest_file.write_bytes(b"abc")
    stat = src.stat()
    dest_file.touch()
    # copy2-equivalent: match size+mtime of source
    import os
    os.utime(dest_file, ns=(stat.st_atime_ns, stat.st_mtime_ns))
    assert dest_matches_source_fingerprint(
        dest_file, size=stat.st_size, mtime_ns=stat.st_mtime_ns
    )
    dest_file.write_bytes(b"abcd")
    assert not dest_matches_source_fingerprint(
        dest_file, size=stat.st_size, mtime_ns=stat.st_mtime_ns
    )


def test_copy_skips_matching_fingerprint(tmp: Path) -> None:
    folder11 = "11_рев.03-AN01_от_2026.01.22"
    package_dir = tmp / folder11
    mto_name = "AGCC.287-8950-SOO1.MTO-0001_03-AN01_RU.xlsx"
    src = package_dir / "DWG" / mto_name
    src.parent.mkdir(parents=True)
    src.write_bytes(b"payload")
    dest = tmp / "dump"
    dest.mkdir()
    pipeline = _pipeline()
    packages = (
        KitPackageRow(
            title="8950",
            mark="SOO1",
            source="rd",
            sequence=11,
            transfer_name=folder11,
            package_path=str(package_dir),
            revision_text="03-AN01",
            is_current=True,
            id=11,
        ),
    )
    stat = src.stat()
    record = _file(
        1,
        folder=folder11,
        media="DWG",
        discipline="MTO-0001",
        revision="03",
        appendix="01",
        file_kind="mto_xlsx",
        size=stat.st_size,
        mtime_ns=stat.st_mtime_ns,
    )
    record = FileRecord(
        id=record.id,
        path=str(src),
        path_key=record.path_key,
        source=record.source,
        present=True,
        review_state=record.review_state,
        first_seen_run_id=1,
        last_seen_run_id=1,
        data=dict(record.data),
    )
    kit = HandoffKitInput(
        title="8950",
        mark="SOO1",
        official_revision_text="03-AN01",
        status="agreed",
        code_a=True,
        pipeline=pipeline,
    )
    rows = build_handoff_rows((kit,), packages, (record,))
    assert len(rows) == 1
    assert len(rows[0].copy_files) == 1
    plan = build_handoff_copy_plan(
        rows, dest_root=str(dest), layout=HANDOFF_LAYOUT_FLAT
    )
    first = execute_handoff_copy(plan)
    assert first.copied == 1, first.failed
    assert first.skipped == 0
    assert first.list_path
    second = execute_handoff_copy(plan)
    assert second.copied == 0, second.failed
    assert second.skipped == 1
    title_plan = build_handoff_copy_plan(
        rows, dest_root=str(dest), layout=HANDOFF_LAYOUT_TITLE_MARK
    )
    assert "8950" in title_plan.items[0].dest_path
    assert "SOO1" in title_plan.items[0].dest_path


def test_destinations_json(tmp: Path) -> None:
    first = remember_handoff_destination(tmp, r"D:\out")
    assert first.last == r"D:\out"
    second = remember_handoff_destination(tmp, r"E:\mgr")
    assert second.paths[0] == r"E:\mgr"
    assert r"D:\out" in second.paths
    loaded = load_handoff_destinations(tmp)
    assert loaded.last == r"E:\mgr"


def main() -> None:
    test_eligibility()
    test_bbb_media_path()
    test_rows_use_official_package_not_is_current()
    test_missing_boe_is_a_note()
    with tempfile.TemporaryDirectory(prefix="handoff_export_") as raw:
        root = Path(raw)
        test_dest_guard_and_fingerprint(root)
        test_copy_skips_matching_fingerprint(root)
        test_destinations_json(root)
    print("RD catalog handoff export: OK")


if __name__ == "__main__":
    main()
