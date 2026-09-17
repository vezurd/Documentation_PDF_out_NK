"""Lexical and wiring checks for the strict RD issued-transfer layout filter."""

from __future__ import annotations

import sys
import tempfile
from dataclasses import replace
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
sys.path.insert(0, str(Path(__file__).resolve().parent))

from rd_catalog.db import CatalogDatabase
from rd_catalog.kits import GoogleKit, KitSummary, build_kit_matrix
from rd_catalog.models import FileKind, FileRecord, ReviewState, SourceKind
from rd_catalog.overlay import build_rd_overlays
from rd_catalog.parse import (
    has_canonical_rd_issued_path,
    list_layout_violations,
    parse_catalog_file,
    parse_transfer_folder,
    record_has_canonical_layout,
)
from rd_catalog.pipeline import rebuild_pipeline
from rd_catalog.scan import scan_document_source
from test_rd_catalog_pipeline import _mtime_ns, _record


RD_ROOT = r"C:\rd"
PIPELINE_RD_ROOT = r"\\bcc\eng\PrDoc\РД"
_NO_GATE_FOLDER = "01_рев.0_AGCC.287-1600-SOT"


def _layout_record(
    path: str,
    *,
    source: SourceKind = SourceKind.RD,
    title: str = "",
    mark: str = "",
    present: bool = True,
) -> FileRecord:
    return FileRecord(
        id=1,
        path=path,
        path_key=path.casefold(),
        source=source,
        present=present,
        review_state=ReviewState.ACKNOWLEDGED,
        first_seen_run_id=1,
        last_seen_run_id=1,
        data={
            "title": title,
            "mark": mark,
            "parse_status": "parsed",
            "file_kind": FileKind.PDF.value,
        },
    )


def _violation_reason(path: str, *, title: str = "", mark: str = "") -> str:
    rows = list_layout_violations(
        records=[_layout_record(path, title=title, mark=mark)],
        rd_root=RD_ROOT,
    )
    assert len(rows) == 1, path
    return rows[0].reason


def _overlay_files() -> tuple[object, object]:
    transfer = parse_transfer_folder(_NO_GATE_FOLDER, under_gate=True)
    canonical = parse_catalog_file(
        rf"{RD_ROOT}\1600\06_SOT\Для передачи\{_NO_GATE_FOLDER}"
        r"\PDF\AGCC.287-1600-SOT.OD-0001_0_RU.pdf",
        SourceKind.RD,
        size=10,
        mtime_ns=2,
        transfer=transfer,
    )
    no_gate = parse_catalog_file(
        rf"{RD_ROOT}\1600\06_SOT\{_NO_GATE_FOLDER}"
        r"\PDF\AGCC.287-1600-SOT.OD-0001_0_RU.pdf",
        SourceKind.RD,
        size=10,
        mtime_ns=1,
        transfer=transfer,
    )
    return canonical, no_gate


def _overlay_seen(overlay) -> set:
    return set(overlay.current.values()) | {entry.file for entry in overlay.entries}


def _no_gate_pipeline_record() -> FileRecord:
    canonical = _record(
        7,
        title="1600",
        mark="SOT",
        revision="0",
        appendix=None,
        sequence=1,
        folder=_NO_GATE_FOLDER,
        mtime_ns=_mtime_ns(2026, 6, 1),
    )
    return replace(
        canonical,
        path=canonical.path.replace(r"\Для передачи", ""),
    )


def _google_1600_sot() -> GoogleKit:
    return GoogleKit(
        title="1600",
        mark="SOT",
        mark_raw="SOT",
        title_system="1600-SOT",
        sheet_revision="0",
        sheet_appendix=None,
        sheet_revision_text="0",
        status_sheet="",
        comment_raw="",
        events=(),
        last_event=None,
        row_index=1,
    )


def test_canonical_package_root_accepted() -> None:
    path = (
        r"C:\rd\9110\KSB\Для передачи\05_рев.0_AGCC.287-9110-KSB"
        r"\AGCC.287-9110-KSB.OD-0001_0_RU.pdf"
    )
    assert has_canonical_rd_issued_path(path, RD_ROOT)
    assert record_has_canonical_layout(_layout_record(path), RD_ROOT)
    assert (
        list_layout_violations(records=[_layout_record(path)], rd_root=RD_ROOT) == ()
    )


def test_canonical_media_folders_accepted() -> None:
    package = r"C:\rd\9110\KSB\Для передачи\05_рев.0_AGCC.287-9110-KSB"
    for media, name in (
        ("PDF", "file.pdf"),
        ("DWG", "file.dwg"),
        ("BBB", "file.xlsx"),
    ):
        path = rf"{package}\{media}\{name}"
        assert has_canonical_rd_issued_path(path, RD_ROOT), path
        assert record_has_canonical_layout(_layout_record(path), RD_ROOT), path


def test_no_gate_rejected() -> None:
    path = (
        r"C:\rd\9000\KSB\01_рев.0_AGCC.287-9000-KSB\PDF"
        r"\AGCC.287-9000-KSB.OD-0001_0_RU.pdf"
    )
    assert not has_canonical_rd_issued_path(path, RD_ROOT)
    assert not record_has_canonical_layout(
        _layout_record(path, title="9000", mark="KSB"), RD_ROOT
    )
    row = list_layout_violations(
        records=[_layout_record(path, title="9000", mark="KSB")],
        rd_root=RD_ROOT,
    )[0]
    assert row.reason == "no_gate"
    assert row.reason_label == (
        "Создайте папку «Для передачи» между маркой и пакетом NN"
    )
    assert row.package_hint == "01_рев.0_AGCC.287-9000-KSB"


def test_loose_in_mark_rejected() -> None:
    path = r"C:\rd\9000\KSB\AGCC.287-9000-KSB.OD-0001_0_RU.pdf"
    assert _violation_reason(path, title="9000", mark="KSB") == "loose_in_mark"
    assert not record_has_canonical_layout(
        _layout_record(path, title="9000", mark="KSB"), RD_ROOT
    )


def test_extra_subfolder_rejected() -> None:
    comma_folder = (
        r"C:\rd\9110\KSB\Для передачи\05_рев.0_AGCC.287-9110-KSB"
        r"\DWG, DOC, XLS\file.xlsx"
    )
    nested = (
        r"C:\rd\9110\KSB\Для передачи\03_рев.01_AGCC.287-9110-KSB"
        r"\DWG\test\file.xlsx"
    )
    assert _violation_reason(comma_folder, title="9110", mark="KSB") == (
        "extra_subfolder"
    )
    assert _violation_reason(nested, title="9110", mark="KSB") == "extra_subfolder"


def test_not_a_package_rejected() -> None:
    path = r"C:\rd\9110\KSB\Для передачи\Корректировка для АН\file.pdf"
    row = list_layout_violations(
        records=[_layout_record(path, title="9110", mark="KSB")],
        rd_root=RD_ROOT,
    )[0]
    assert row.reason == "not_a_package"
    assert row.package_hint == ""


def test_sq_and_robot_never_rejected() -> None:
    loose = r"C:\rd\9000\KSB\file.pdf"
    sq = _layout_record(loose, source=SourceKind.SQ, title="9000", mark="KSB")
    robot = _layout_record(
        r"C:\robot\9000\KSB\file.xlsx",
        source=SourceKind.ROBOT,
        title="9000",
        mark="KSB",
    )
    assert record_has_canonical_layout(sq, RD_ROOT)
    assert record_has_canonical_layout(robot, RD_ROOT)
    assert list_layout_violations(records=[sq, robot], rd_root=RD_ROOT) == ()


def test_digit_mark_folder_still_accepted() -> None:
    path = (
        r"C:\rd\1715\11_SOT\Для передачи\04_рев.0_AGCC.287-1715-SOT"
        r"\AGCC.287-1715-SOT.OD-0001_0_RU.pdf"
    )
    assert has_canonical_rd_issued_path(path, RD_ROOT)
    assert record_has_canonical_layout(
        _layout_record(path, title="1715", mark="SOT"), RD_ROOT
    )
    assert (
        list_layout_violations(
            records=[_layout_record(path, title="1715", mark="SOT")],
            rd_root=RD_ROOT,
        )
        == ()
    )


def test_overlay_wiring_excludes_no_gate_when_rd_root_set() -> None:
    canonical, no_gate = _overlay_files()
    filtered_pdf, _filtered_mto = build_rd_overlays(
        [canonical, no_gate], rd_root=RD_ROOT
    )
    seen = _overlay_seen(filtered_pdf)
    assert canonical in seen
    assert no_gate not in seen
    assert list(filtered_pdf.current.values()) == [canonical]


def test_overlay_wiring_keeps_no_gate_without_rd_root() -> None:
    canonical, no_gate = _overlay_files()
    open_pdf, _open_mto = build_rd_overlays([canonical, no_gate])
    assert no_gate in _overlay_seen(open_pdf)


def test_rebuild_pipeline_wiring_drops_no_gate_when_rd_root_set() -> None:
    record = _no_gate_pipeline_record()
    with tempfile.TemporaryDirectory(prefix="rd_layout_pipe_") as raw:
        database = CatalogDatabase(Path(raw) / "filtered.sqlite")
        database.initialize()
        rebuild_pipeline(
            database,
            records=[record],
            detected_current_ids={record.id},
            rd_root=PIPELINE_RD_ROOT,
        )
        assert database.list_kit_packages("1600", "SOT") == []


def test_rebuild_pipeline_wiring_keeps_no_gate_without_rd_root() -> None:
    record = _no_gate_pipeline_record()
    with tempfile.TemporaryDirectory(prefix="rd_layout_pipe_open_") as raw:
        database = CatalogDatabase(Path(raw) / "open.sqlite")
        database.initialize()
        rebuild_pipeline(
            database,
            records=[record],
            detected_current_ids={record.id},
        )
        packages = database.list_kit_packages("1600", "SOT")
        assert packages
        assert all(pkg.source == "rd" for pkg in packages)


def test_kit_matrix_wiring_hides_no_gate_when_rd_root_set() -> None:
    record = _no_gate_pipeline_record()
    google = _google_1600_sot()
    filtered = build_kit_matrix(
        (google,),
        [record],
        {record.id},
        rd_root=PIPELINE_RD_ROOT,
    )
    assert len(filtered) == 1
    assert not filtered[0].rd.present
    assert filtered[0].summary is KitSummary.GOOGLE_ONLY


def test_kit_matrix_wiring_keeps_no_gate_without_rd_root() -> None:
    record = _no_gate_pipeline_record()
    google = _google_1600_sot()
    open_rows = build_kit_matrix((google,), [record], {record.id})
    assert len(open_rows) == 1
    assert open_rows[0].rd.present


def test_scan_rd_stores_only_canonical_issued_paths() -> None:
    """RD walk stores BBB-under-NN; skips loose title dumps. SQ stays open."""

    with tempfile.TemporaryDirectory(prefix="rd_layout_scan_") as raw:
        rd = Path(raw, "rd")
        package = rd / "9110" / "KSB" / "Для передачи" / "05_рев.0_AGCC.287-9110-KSB"
        bbb = package / "BBB"
        bbb.mkdir(parents=True)
        canonical_dwg = bbb / "AGCC.287-9110-KSB.OD-0001_0_RU.dwg"
        canonical_dwg.write_bytes(b"bbb-dwg")
        loose_dir = rd / "2000"
        loose_dir.mkdir(parents=True)
        loose_pdf = loose_dir / "AGCC.287-2000-KSB.OD-0001_0_RU.pdf"
        loose_pdf.write_bytes(b"loose-pdf")
        scanned = scan_document_source(rd, SourceKind.RD, skip_dirs=())
        paths = {item.path for item in scanned.files}
        assert str(canonical_dwg) in paths
        assert str(loose_pdf) not in paths

        sq = Path(raw, "sq")
        sq.mkdir()
        sq_loose = sq / "AGCC.287-2000-KSB.OD-0001_0_RU.pdf"
        sq_loose.write_bytes(b"sq-loose")
        sq_scanned = scan_document_source(sq, SourceKind.SQ, skip_dirs=())
        assert any(item.path == str(sq_loose) for item in sq_scanned.files)


def main() -> None:
    """Run layout predicate and wiring assertions without UNC or a live database."""

    test_canonical_package_root_accepted()
    test_canonical_media_folders_accepted()
    test_no_gate_rejected()
    test_loose_in_mark_rejected()
    test_extra_subfolder_rejected()
    test_not_a_package_rejected()
    test_sq_and_robot_never_rejected()
    test_digit_mark_folder_still_accepted()
    test_overlay_wiring_excludes_no_gate_when_rd_root_set()
    test_overlay_wiring_keeps_no_gate_without_rd_root()
    test_rebuild_pipeline_wiring_drops_no_gate_when_rd_root_set()
    test_rebuild_pipeline_wiring_keeps_no_gate_without_rd_root()
    test_kit_matrix_wiring_hides_no_gate_when_rd_root_set()
    test_kit_matrix_wiring_keeps_no_gate_without_rd_root()
    test_scan_rd_stores_only_canonical_issued_paths()
    print("RD catalog strict layout: OK")


if __name__ == "__main__":
    main()
