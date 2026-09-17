"""Local checks for RD source/editable file kinds and PDF pairing."""

from __future__ import annotations

import sys
import tempfile
from datetime import datetime
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from rd_catalog.doc_bundle import (
    DocumentTreeLabelOptions,
    ANNULLED_MARKER,
    ANNULLED_TOOLTIP,
    WORKING_TOOLTIP_MANUAL,
    annulled_folder_tooltip,
    bundle_documents,
    editable_pairs_with_pdf,
    folder_latest_save_date,
    folder_mean_override,
    folder_revision_label,
    folder_tree_label,
    latest_save_mtime_ns,
    mtime_override_tooltip,
    same_document_path_keys,
    working_folder_tooltip,
)
from rd_catalog.models import (
    FileKind,
    FileRecord,
    ParseStatus,
    ReviewState,
    SourceKind,
)
from rd_catalog.overlay import build_rd_overlays
from rd_catalog.parse import _file_kind_from_name, parse_catalog_file
from rd_catalog.scan import _candidate_kind, scan_document_source


def _record(
    file_id: int,
    *,
    path: str,
    file_kind: str,
    title: str = "1513",
    mark: str = "POS",
    revision: str | None = "01",
    appendix: str | None = None,
    core_stem: str = "AGCC.287-1513-POS.OD-0001",
    discipline_block: str = "OD-0001",
    transfer_name: str | None = None,
    source: SourceKind = SourceKind.RD,
    mtime_ns: int = 1,
    present: bool = True,
) -> FileRecord:
    name = Path(path).name
    return FileRecord(
        id=file_id,
        path=path,
        path_key=path.casefold(),
        source=source,
        present=present,
        review_state=ReviewState.ACKNOWLEDGED,
        first_seen_run_id=1,
        last_seen_run_id=1,
        data={
            "file_kind": file_kind,
            "name": name,
            "parse_status": "parsed",
            "title": title,
            "mark": mark,
            "revision": revision,
            "appendix": appendix,
            "core_stem": core_stem,
            "discipline_block": discipline_block,
            "mtime_ns": mtime_ns,
            "transfer_name": transfer_name,
        },
    )


def main() -> None:
    """Run source-file classification and pairing assertions."""

    assert _candidate_kind("AGCC.287-1513-POS.OD-0001_01_RU.pdf") is FileKind.PDF
    assert _candidate_kind("x.pdf", sq_only=True) is FileKind.PDF
    assert (
        _candidate_kind(
            "AGCC.287-2225-KSB.MTO-0001_01_RU.xlsx", sq_only=True
        )
        is FileKind.MTO_XLSX
    )
    assert (
        _candidate_kind("AGCC.MTO.xlsx", allow_source=False) is FileKind.MTO_XLSX
    )
    assert (
        _candidate_kind("AGCC.287-мто.xlsx", allow_source=False) is FileKind.MTO_XLSX
    )
    assert _candidate_kind("foo.dwg", allow_source=False) is None
    assert _candidate_kind("foo.dwg", allow_source=True) is FileKind.SOURCE_EDITABLE
    assert _candidate_kind("foo.dwg", sq_only=True, allow_source=True) is None
    assert _candidate_kind("BOQ.xlsx", allow_source=True) is FileKind.SOURCE_EDITABLE

    assert _file_kind_from_name("BOQ.xlsx") is FileKind.SOURCE_EDITABLE
    assert _file_kind_from_name("MTO.xlsx") is FileKind.MTO_XLSX

    dwg = parse_catalog_file(
        "AGCC.287-2225-KSB.OD-0001_01_RU.dwg",
        SourceKind.RD,
        size=10,
        mtime_ns=20,
    )
    assert dwg.file_kind is FileKind.SOURCE_EDITABLE
    assert dwg.parse_status is ParseStatus.PARSED
    assert dwg.extension == "dwg"
    assert dwg.document_key == "agcc.287-2225-ksb.od-0001"
    assert dwg.parse_error is None

    boq_xlsx = parse_catalog_file(
        "AGCC.287-2225-KSB.OD-0001_01_RU.xlsx",
        SourceKind.RD,
        size=10,
        mtime_ns=20,
    )
    assert boq_xlsx.file_kind is FileKind.SOURCE_EDITABLE
    assert boq_xlsx.parse_status is ParseStatus.PARSED
    assert boq_xlsx.parse_error is None

    mto = parse_catalog_file(
        "AGCC.287-2225-KSB.MTO-0001_01-AN02_RU.xlsx",
        SourceKind.RD,
        size=10,
        mtime_ns=20,
    )
    assert mto.file_kind is FileKind.MTO_XLSX
    assert mto.parse_status is ParseStatus.PARSED

    pdf_dir = r"C:\rd\1513\01_рев.01\PDF"
    assert editable_pairs_with_pdf(pdf_dir, pdf_dir) is True
    assert editable_pairs_with_pdf(pdf_dir, str(Path(pdf_dir).parent / "DWG")) is True
    assert editable_pairs_with_pdf(pdf_dir, str(Path(pdf_dir) / "DWG")) is True
    assert editable_pairs_with_pdf(pdf_dir, r"C:\other\DWG") is False

    same_folder_pdf = _record(
        1,
        path=r"C:\rd\kit\AGCC.287-1513-POS.OD-0001_01_RU.pdf",
        file_kind=FileKind.PDF.value,
    )
    same_folder_dwg = _record(
        2,
        path=r"C:\rd\kit\AGCC.287-1513-POS.OD-0001_01_RU.dwg",
        file_kind=FileKind.SOURCE_EDITABLE.value,
    )
    glued = bundle_documents(
        [same_folder_pdf, same_folder_dwg],
        detected_current_ids={1},
    )
    assert len(glued) == 1
    assert glued[0].pdf is same_folder_pdf
    assert glued[0].editables == (same_folder_dwg,)
    assert glued[0].is_current is True

    mask_miss_pdf = _record(
        4,
        path=r"C:\rd\kit\SQ-PP-AGCC-KSB-00051-0_signed.pdf",
        file_kind=FileKind.PDF.value,
        title="",
        mark="",
        core_stem="",
    )
    mask_miss_xlsx = _record(
        5,
        path=r"C:\rd\kit\SQ-PP-AGCC-KSB-00051-0.xlsx",
        file_kind=FileKind.SOURCE_EDITABLE.value,
        title="",
        mark="",
        core_stem="",
    )
    hidden = bundle_documents(
        [same_folder_pdf, mask_miss_pdf, mask_miss_xlsx],
        detected_current_ids={1},
    )
    assert len(hidden) == 1
    assert hidden[0].pdf is same_folder_pdf

    foreign_dwg = _record(
        3,
        path=r"C:\foreign\AGCC.287-1513-POS.OD-0001_01_RU.dwg",
        file_kind=FileKind.SOURCE_EDITABLE.value,
    )
    split = bundle_documents([same_folder_pdf, foreign_dwg])
    assert len(split) == 2
    by_pdf = {bundle.pdf is not None: bundle for bundle in split}
    assert by_pdf[True].pdf is same_folder_pdf
    assert by_pdf[True].editables == ()
    assert by_pdf[False].pdf is None
    assert by_pdf[False].editables == (foreign_dwg,)

    od_pdf = _record(
        10,
        path=r"C:\rd\06_рев.02\AGCC.287-2245-KSB.OD-0001_01-AN02_RU.pdf",
        file_kind=FileKind.PDF.value,
        title="2245",
        mark="KSB",
        revision="01",
        appendix="02",
        core_stem="AGCC.287-2245-KSB.OD-0001",
        discipline_block="OD-0001",
        transfer_name="06_рев.02_AGCC.287-2245-KSB",
    )
    wir_pdf = _record(
        11,
        path=r"C:\rd\06_рев.02\AGCC.287-2245-KSB.WIR-0001.pdf",
        file_kind=FileKind.PDF.value,
        title="2245",
        mark="KSB",
        revision=None,
        core_stem="AGCC.287-2245-KSB.WIR-0001",
        discipline_block="WIR-0001",
        transfer_name="06_рев.02_AGCC.287-2245-KSB",
    )
    folder_bundles = bundle_documents([od_pdf, wir_pdf], detected_current_ids={10})
    assert len(folder_bundles) == 2
    assert {bundle.folder_key for bundle in folder_bundles} == {
        "06_рев.02_agcc.287-2245-ksb"
    }
    folder_files = [od_pdf, wir_pdf]
    assert folder_revision_label(folder_files) == "01-AN02"

    ns_old = int(datetime(2024, 1, 2, 12, 0).timestamp() * 1_000_000_000)
    ns_new = int(datetime(2024, 8, 16, 18, 0).timestamp() * 1_000_000_000)
    dated_pdf = _record(
        12,
        path=r"C:\rd\06_рев.02\PDF\AGCC.287-2245-KSB.OD-0001_01-AN02_RU.pdf",
        file_kind=FileKind.PDF.value,
        title="2245",
        mark="KSB",
        revision="01",
        appendix="02",
        core_stem="AGCC.287-2245-KSB.OD-0001",
        discipline_block="OD-0001",
        transfer_name="06_рев.02_AGCC.287-2245-KSB",
        mtime_ns=ns_old,
    )
    dated_dwg = _record(
        13,
        path=r"C:\rd\06_рев.02\DWG\AGCC.287-2245-KSB.OD-0001_01-AN02_RU.dwg",
        file_kind=FileKind.SOURCE_EDITABLE.value,
        title="2245",
        mark="KSB",
        revision="01",
        appendix="02",
        core_stem="AGCC.287-2245-KSB.OD-0001",
        discipline_block="OD-0001",
        transfer_name="06_рев.02_AGCC.287-2245-KSB",
        mtime_ns=ns_new,
    )
    expected_date = datetime.fromtimestamp(ns_new / 1_000_000_000).strftime(
        "%Y.%m.%d"
    )
    assert latest_save_mtime_ns([dated_pdf, dated_dwg]) == ns_new
    assert folder_latest_save_date([dated_pdf, dated_dwg]) == expected_date
    assert folder_tree_label([dated_pdf, dated_dwg]) == f"01-AN02 ({expected_date})"
    assert (
        folder_tree_label([dated_pdf, dated_dwg], is_working=True)
        == f"01-AN02 ({expected_date}) · рабочая"
    )
    assert working_folder_tooltip(is_working=False) == ""
    assert (
        working_folder_tooltip(is_working=True, origin="manual")
        == WORKING_TOOLTIP_MANUAL
    )
    assert (
        folder_tree_label([dated_pdf, dated_dwg], is_annulled=True)
        == f"01-AN02 ({expected_date}) · {ANNULLED_MARKER}"
    )
    assert (
        folder_tree_label(
            [dated_pdf, dated_dwg], is_working=True, is_annulled=True
        )
        == f"01-AN02 ({expected_date}) · {ANNULLED_MARKER}"
    )
    assert annulled_folder_tooltip(is_annulled=False) == ""
    assert annulled_folder_tooltip(is_annulled=True) == ANNULLED_TOOLTIP
    assert (
        folder_tree_label(
            [dated_pdf, dated_dwg],
            options=DocumentTreeLabelOptions(show_date=False, show_folder=True),
            folder_name="06_рев.02_AGCC.287-2245-KSB",
        )
        == "01-AN02 · 06_рев.02_AGCC.287-2245-KSB"
    )
    assert (
        folder_tree_label(
            [dated_pdf, dated_dwg],
            options=DocumentTreeLabelOptions(
                show_date=True,
                show_folder=True,
                show_review=True,
                show_current=True,
                show_mto=True,
                show_working=True,
                show_as_build=True,
            ),
            folder_name="06_рев.02_AGCC.287-2245-KSB",
            review_status="код A",
            is_current=True,
            has_mto=False,
            is_working=True,
            has_as_build=True,
        )
        == (
            f"01-AN02 ({expected_date}) · 06_рев.02_AGCC.287-2245-KSB"
            " · код A · текущая · рабочая · AB"
        )
    )
    undated = _record(
        14,
        path=r"C:\rd\06_рев.02\PDF\AGCC.287-2245-KSB.WIR-0001.pdf",
        file_kind=FileKind.PDF.value,
        title="2245",
        mark="KSB",
        revision=None,
        core_stem="AGCC.287-2245-KSB.WIR-0001",
        discipline_block="WIR-0001",
        transfer_name="06_рев.02_AGCC.287-2245-KSB",
        mtime_ns=0,
    )
    assert folder_tree_label([undated]) == "Без ревизии"

    pdf_parsed = parse_catalog_file(
        r"C:\rd\01_рев.01\AGCC.287-2225-KSB.OD-0001_01_RU.pdf",
        SourceKind.RD,
        size=10,
        mtime_ns=20,
    )
    overlay_pdf, overlay_mto = build_rd_overlays([pdf_parsed, dwg, mto])
    assert all(entry.file.file_kind is FileKind.PDF for entry in overlay_pdf.entries)
    assert all(
        entry.file.file_kind is FileKind.MTO_XLSX for entry in overlay_mto.entries
    )

    with tempfile.TemporaryDirectory(prefix="rd_catalog_source_") as temp:
        rd = Path(temp, "rd")
        transfer = rd / "1513" / "11_POS" / "Для передачи" / "01_рев.01"
        pdf_folder = transfer / "PDF"
        dwg_folder = transfer / "DWG"
        other_folder = transfer / "other"
        pdf_folder.mkdir(parents=True)
        dwg_folder.mkdir(parents=True)
        other_folder.mkdir(parents=True)
        pdf_name = "AGCC.287-1513-POS.OD-0001_01_RU.pdf"
        dwg_name = "AGCC.287-1513-POS.OD-0001_01_RU.dwg"
        mto_name = "AGCC.287-1513-POS.MTO-0001_01_RU.xlsx"
        (pdf_folder / pdf_name).write_bytes(b"pdf")
        (pdf_folder / dwg_name).write_bytes(b"dwg-next-to-pdf")
        (dwg_folder / dwg_name).write_bytes(b"dwg-dir")
        (other_folder / dwg_name).write_bytes(b"stray-dwg")
        (transfer / mto_name).write_bytes(b"mto")
        bbb_folder = transfer / "BBB"
        bbb_folder.mkdir(parents=True)
        (bbb_folder / dwg_name).write_bytes(b"bbb-only-dwg")
        mark_dwg_dir = rd / "2225" / "10_KSB" / "DWG"
        mark_dwg_dir.mkdir(parents=True)
        (mark_dwg_dir / dwg_name).write_bytes(b"mark-level-dwg")
        loose_dir = rd / "2000"
        loose_dir.mkdir(parents=True)
        loose_pdf = loose_dir / "AGCC.287-2000-KSB.OD-0001_0_RU.pdf"
        loose_dwg = loose_dir / "AGCC.287-2000-KSB.CAE-0002.1_0_RU.dwg"
        loose_pdf.write_bytes(b"loose-pdf")
        loose_dwg.write_bytes(b"loose-dwg")
        scanned = scan_document_source(rd, SourceKind.RD, skip_dirs=())
        scanned_paths = {item.path for item in scanned.files}
        names_by_kind: dict[FileKind, set[str]] = {}
        for item in scanned.files:
            names_by_kind.setdefault(item.file_kind, set()).add(item.path)
        pdf_paths = names_by_kind.get(FileKind.PDF, set())
        editable_paths = names_by_kind.get(FileKind.SOURCE_EDITABLE, set())
        mto_paths = names_by_kind.get(FileKind.MTO_XLSX, set())
        assert any(path.endswith(pdf_name) for path in pdf_paths)
        assert any(
            Path(path).parent.name == "PDF" and path.endswith(dwg_name)
            for path in editable_paths
        )
        assert any(
            Path(path).parent.name == "DWG" and path.endswith(dwg_name)
            for path in editable_paths
        )
        assert any(
            Path(path).parent.name == "BBB" and path.endswith(dwg_name)
            for path in editable_paths
        )
        assert not any(
            Path(path).parent.name == "other" and path.endswith(dwg_name)
            for path in editable_paths
        )
        assert any(path.endswith(mto_name) for path in mto_paths)
        assert str(loose_pdf) not in scanned_paths
        assert str(loose_dwg) not in scanned_paths
        assert str(mark_dwg_dir / dwg_name) not in scanned_paths

        sq = Path(temp, "sq")
        sq.mkdir()
        sq_loose = sq / "AGCC.287-2000-KSB.OD-0001_0_RU.pdf"
        sq_loose.write_bytes(b"sq-loose")
        sq_scanned = scan_document_source(sq, SourceKind.SQ, skip_dirs=())
        assert any(item.path == str(sq_loose) for item in sq_scanned.files)

    rec = _record(
        90,
        path=r"\\x\AGCC.287-1513-POS.MTO-0001_01_RU.xlsx",
        file_kind=FileKind.MTO_XLSX.value,
    )
    rec.data["mtime_override_applied"] = True
    rec.data["mtime_override_date"] = "15.02.2026"
    rec.data["mtime_override_reason"] = "code_b"
    rec.data["disk_mtime_ns"] = 0
    assert "код B" in mtime_override_tooltip(rec)
    rec.data["mtime_override_reason"] = "code_c"
    assert "код C" in mtime_override_tooltip(rec)
    rec.data["mtime_override_reason"] = "manual"
    assert "вручную" in mtime_override_tooltip(rec)
    rec.data["mtime_override_reason"] = "folder_mean"
    assert "средняя по папке" in mtime_override_tooltip(rec)

    def _noon_ns(year: int, month: int, day: int) -> int:
        return int(datetime(year, month, day, 12, 0, 0).timestamp() * 1_000_000_000)

    mto_xlsx = _record(
        91,
        path=r"\\x\pkg\DWG\AGCC.287-7360-SKUD.MTO-0001_0_RU.xlsx",
        file_kind=FileKind.MTO_XLSX.value,
        title="7360",
        mark="SKUD",
        revision="0",
        core_stem="AGCC.287-7360-SKUD.MTO-0001",
        discipline_block="MTO-0001",
        mtime_ns=_noon_ns(2025, 8, 22),
    )
    mto_pdf = _record(
        92,
        path=r"\\x\pkg\PDF\AGCC.287-7360-SKUD.MTO-0001_0_RU.pdf",
        file_kind=FileKind.PDF.value,
        title="7360",
        mark="SKUD",
        revision="0",
        core_stem="AGCC.287-7360-SKUD.MTO-0001",
        discipline_block="MTO-0001",
        mtime_ns=_noon_ns(2025, 8, 22),
    )
    od_files = [
        _record(
            93 + index,
            path=rf"\\x\pkg\PDF\AGCC.287-7360-SKUD.OD-000{index}_0_RU.pdf",
            file_kind=FileKind.PDF.value,
            title="7360",
            mark="SKUD",
            revision="0",
            core_stem=f"AGCC.287-7360-SKUD.OD-000{index}",
            discipline_block=f"OD-000{index}",
            mtime_ns=_noon_ns(2023, 7, 7),
        )
        for index in range(1, 4)
    ]
    folder = [*od_files, mto_pdf, mto_xlsx]
    exclude = same_document_path_keys(mto_xlsx, folder)
    assert mto_pdf.path_key.casefold() in exclude
    mean = folder_mean_override(folder, exclude_path_keys=exclude)
    assert mean is not None
    assert mean.override_date == "07.07.2023"
    assert mean.used_count == 3
    assert mean.excluded_count == 2
    assert folder_mean_override((mto_pdf, mto_xlsx), exclude_path_keys=exclude) is None

    print("RD catalog source files: OK")


if __name__ == "__main__":
    main()
