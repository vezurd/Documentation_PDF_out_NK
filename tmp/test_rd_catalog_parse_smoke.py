"""Local smoke checks for RD catalog AGCC parsing."""

from __future__ import annotations

import sys
import tempfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from rd_catalog.config import load_config
from rd_catalog.doc_bundle import folder_transfer_sequence, folder_tree_sort_key
from rd_catalog.models import FileKind, FileRecord, ParseStatus, ReviewState, SourceKind
from rd_catalog.parse import (
    constructed_kit_rd_mark_folder,
    folder_hosts_filename_mark,
    folder_matches_mark,
    has_canonical_rd_issued_path,
    is_transfer_folder_name,
    is_transfer_gate_folder_name,
    issued_package_dir,
    issued_path_mark_folder,
    issued_path_title_folder,
    path_relative_to_transfer_gate,
    kit_rd_mark_folder_from_path,
    matches_agcc_filename,
    normalize_unicode_dashes,
    parse_catalog_file,
    parse_transfer_folder,
    path_is_as_build,
    transfer_name_is_void,
    constructed_kit_rd_title_folder,
    unique_kit_rd_mark_folders,
)
from utils.file_name_converts import AgccFilenamePatterns
from rd_catalog.scan import _find_transfer, _is_skipped_dir, scan_document_source


def main() -> None:
    """Run parsing assertions without touching network paths."""

    transfer = parse_transfer_folder(
        "07_рев.02-AN03_2225\u2010KSB_as-built"
    )
    assert transfer.parse_status is ParseStatus.PARSED
    assert transfer.sequence == 7
    assert transfer.revision == "02"
    assert transfer.appendix == "03"
    assert transfer.title == "2225"
    assert transfer.mark == "KSB"
    assert transfer.is_as_build
    void_folder = parse_transfer_folder(
        "04_рев.0-AN02_AGCC.287-7560-SKUD_Void",
        under_gate=True,
    )
    assert void_folder.parse_status is ParseStatus.PARSED
    assert void_folder.sequence == 4
    assert void_folder.is_void
    assert transfer_name_is_void("04_рев.0-AN02_AGCC.287-7560-SKUD_Void")
    assert transfer_name_is_void("04_Void_рев.0-AN02_AGCC.287-7560-SKUD")
    assert not transfer_name_is_void("04_рев.0-AN02_AGCC.287-7560-SKUD")
    assert is_transfer_folder_name(
        "04_рев.0-AN02_AGCC.287-7560-SKUD_Void", under_gate=True
    )
    assert path_is_as_build(
        r"\\server\share\РД\as-build\MTO as-build\2245\KSB\file.pdf"
    )
    assert path_is_as_build(
        r"\\host\share\РД\2225\KSB\Для передачи"
        r"\07_рев.02_2225-KSB_as-built\PDF"
        r"\AGCC.287-2225-KSB.OD-0001_02_RU.pdf"
    )
    assert not path_is_as_build(
        r"\\server\share\РД\2245\KSB\Для передачи\05_рев.01"
        r"\document-as-build.pdf"
    )
    assert not path_is_as_build("as-build.pdf")

    bare = parse_transfer_folder("01_рев.01")
    assert bare.parse_status is ParseStatus.PARSED
    assert bare.sequence == 1
    assert bare.revision == "01"
    assert bare.title_system is None

    numbered_only = parse_transfer_folder("03")
    assert numbered_only.parse_status is ParseStatus.PARSED
    assert numbered_only.sequence == 3

    assert not is_transfer_folder_name("Для передачи")
    assert not is_transfer_folder_name("1513")
    assert not is_transfer_folder_name("11_SPP")
    assert not is_transfer_folder_name("12_POS1")
    assert not is_transfer_folder_name("06_4130-KSB1")
    assert not is_transfer_folder_name("03_Для передачи")
    assert not is_transfer_folder_name("10_рев.AN01_AGCC.287-8950-POS1")
    assert is_transfer_folder_name("01_рев.01")
    assert is_transfer_folder_name("02_рев.01-AN02_1513-SPP")
    assert is_transfer_folder_name(
        "10_рев.AN01_AGCC.287-8950-POS1", under_gate=True
    )
    assert is_transfer_gate_folder_name("Для передачи")
    assert is_transfer_gate_folder_name("На_отправку")
    assert is_transfer_gate_folder_name("03_Для передачи")
    assert is_transfer_gate_folder_name("2025.12.08_Для передачи")

    assert parse_transfer_folder("11_SPP").parse_status is ParseStatus.UNPARSED_FOLDER
    assert parse_transfer_folder("1513").parse_status is ParseStatus.UNPARSED_FOLDER
    an_only = parse_transfer_folder(
        "10_рев.AN01_AGCC.287-8950-POS1", under_gate=True
    )
    assert an_only.parse_status is ParseStatus.PARSED
    assert an_only.sequence == 10

    found = _find_transfer(
        r"\\host\share\РД\1513\11_SPP\Для передачи\01_рев.01",
        r"\\host\share\РД",
    )
    assert found.parse_status is ParseStatus.PARSED
    assert found.sequence == 1
    assert found.revision == "01"

    an01_under_gate = _find_transfer(
        r"\\host\share\РД\8950\12_POS1\Для передачи\10_рев.AN01_AGCC.287-8950-POS1\PDF",
        r"\\host\share\РД",
    )
    assert an01_under_gate.parse_status is ParseStatus.PARSED
    assert an01_under_gate.sequence == 10

    mark_not_transfer = _find_transfer(
        r"\\host\share\РД\8950\12_POS1",
        r"\\host\share\РД",
    )
    assert mark_not_transfer.parse_status is ParseStatus.UNPARSED_FOLDER

    skipped_mark = _find_transfer(
        r"\\host\share\РД\1513\11_SPP\Для передачи",
        r"\\host\share\РД",
    )
    assert skipped_mark.parse_status is ParseStatus.UNPARSED_FOLDER

    parsed = parse_catalog_file(
        "AGCC.287\u20102225-KSB.MTO-0001_01-AN02_RU.xlsx",
        SourceKind.RD,
        size=10,
        mtime_ns=20,
        transfer=transfer,
    )
    assert parsed.parse_status is ParseStatus.PARSED
    assert parsed.title_system == "2225-KSB"
    assert parsed.discipline_block == "MTO-0001"
    assert parsed.revision == "01"
    assert parsed.appendix == "02"
    assert parsed.document_key == ("2225-ksb", "mto-0001")

    unparsed = parse_catalog_file(
        "not_an_agcc_document.pdf",
        SourceKind.RD,
        size=0,
        mtime_ns=0,
    )
    assert unparsed.parse_status is ParseStatus.UNPARSED_FILE
    assert matches_agcc_filename("AGCC.287-7421-SKUD.1.OD-0001_01_RU.pdf")
    assert matches_agcc_filename("AGCC.287\u20102225-KSB.MTO-0001_01-AN02_RU.xlsx")
    assert matches_agcc_filename("AGCC.287-1600-SS30.MTO-0001_0_RU.xlsx")
    assert matches_agcc_filename("AGCC.287-6400-SS30.MTO-0001_01-AN01_RU.xlsx")
    assert matches_agcc_filename("AGCC.287-8150-PD21.MTO-0001_01_RU.xlsx")
    ss30 = parse_catalog_file(
        "AGCC.287-1600-SS30.MTO-0001_0_RU.xlsx",
        SourceKind.RD,
        size=10,
        mtime_ns=20,
    )
    assert ss30.parse_status is ParseStatus.PARSED
    assert ss30.title_system == "1600-SS30"
    assert ss30.mark == "SS30"
    assert ss30.revision == "0"
    assert not matches_agcc_filename("SQ-PP-AGCC-KSB-00051-0_signed.pdf")
    assert not matches_agcc_filename("SQ-PP-AGCC-KSB-00051-0.xlsx")
    assert not matches_agcc_filename("not_an_agcc_document.pdf")
    assert normalize_unicode_dashes("A\u2010B\u2014C") == "A-B-C"
    dash_name = (
        "AGCC.287\u20109130-KSB.MTO-0001_0\u2013AN02_RU.xlsx"
    )
    dash_parsed = parse_catalog_file(
        dash_name,
        SourceKind.RD,
        size=10,
        mtime_ns=20,
    )
    assert dash_parsed.parse_status is ParseStatus.PARSED
    assert dash_parsed.name == dash_name
    assert dash_parsed.title_system == "9130-KSB"
    assert dash_parsed.revision == "0"
    assert dash_parsed.appendix == "02"
    assert AgccFilenamePatterns.scan_title_system(
        "see 7180\u2010SOT in text"
    ) == "7180-SOT"
    assert AgccFilenamePatterns.parse_strict(
        "AGCC.287-7180-SOT.OD-0001_02\u2010AN02_RU.pdf"
    ) is not None

    issued_pdf = (
        r"\\bcc\eng\PrDoc\РД\9110\06_KSB_21\Для передачи"
        r"\05_рев.0-AN02_AGCC.287-9110-KSB1\PDF"
        r"\AGCC.287-9110-KSB1.OD-0001_0-AN02_RU.pdf"
    )
    issued_folder = (
        r"\\bcc\eng\PrDoc\РД\9110\06_KSB_21\Для передачи"
        r"\05_рев.0-AN02_AGCC.287-9110-KSB1"
    )
    assert issued_package_dir(issued_pdf) == issued_folder
    assert issued_path_title_folder(issued_pdf) == "9110"
    assert issued_path_title_folder(issued_folder) == "9110"
    assert issued_path_mark_folder(issued_pdf) == "06_KSB_21"
    assert issued_path_mark_folder(issued_folder) == "06_KSB_21"
    assert folder_matches_mark("POS1", "POS1")
    assert folder_matches_mark("12_POS1", "POS1")
    assert folder_matches_mark("04-SOT", "SOT")
    assert not folder_matches_mark("12_POS2", "POS1")
    assert folder_hosts_filename_mark("SOS", "SOS")
    assert folder_hosts_filename_mark("06_KSB_21", "KSB1")
    assert folder_hosts_filename_mark("06_KSB_21", "KSB")
    assert folder_hosts_filename_mark("05_KSB", "KSB1")
    assert folder_hosts_filename_mark("22_POS2", "POS2")
    assert not folder_hosts_filename_mark("SOS", "SOT")
    assert not folder_hosts_filename_mark("12_POS1", "POS2")
    assert r"\PDF" not in issued_package_dir(issued_pdf)
    assert path_relative_to_transfer_gate(issued_pdf) == (
        r"05_рев.0-AN02_AGCC.287-9110-KSB1\PDF"
        r"\AGCC.287-9110-KSB1.OD-0001_0-AN02_RU.pdf"
    )
    assert path_relative_to_transfer_gate(r"\\stub\rd\file.pdf") == "file.pdf"
    user_pdf = (
        r"\\bcc\eng\PrDoc\377_НИПИГАЗ\АГХК\КСБ\РД\8950\22_POS2\Для передачи"
        r"\12_рев.AN01_AGCC.287-8950-POS2\PDF"
        r"\AGCC.287-8950-POS2.OD-0001_03-AN01_RU.pdf"
    )
    user_kit = (
        r"\\bcc\eng\PrDoc\377_НИПИГАЗ\АГХК\КСБ\РД\8950\22_POS2\Для передачи"
        r"\12_рев.AN01_AGCC.287-8950-POS2"
    )
    assert issued_package_dir(user_pdf) == user_kit
    assert issued_path_title_folder(user_pdf) == "8950"
    stray_wir = (
        r"\\bcc\eng\PrDoc\РД\2612\05_KSB\Для передачи"
        r"\09_рев.AN-01_AGC.287-2612-KSB_as-build\PDF"
        r"\AGCC.287-2630-KSB.WIR-0010_01-AN01_RU.pdf"
    )
    assert issued_path_title_folder(stray_wir) == "2612"
    assert issued_path_title_folder(r"\\stub\rd\file.pdf") == ""
    rd_root = r"\\bcc\eng\PrDoc\РД"
    assert has_canonical_rd_issued_path(issued_pdf, rd_root)
    issued_no_media = (
        issued_folder + r"\AGCC.287-9110-KSB1.OD-0001_0-AN02_RU.pdf"
    )
    assert path_relative_to_transfer_gate(issued_no_media) == (
        r"05_рев.0-AN02_AGCC.287-9110-KSB1"
        r"\AGCC.287-9110-KSB1.OD-0001_0-AN02_RU.pdf"
    )
    assert has_canonical_rd_issued_path(issued_no_media, rd_root)
    assert has_canonical_rd_issued_path(
        issued_pdf.replace("Для передачи", "На_отправку"), rd_root
    )
    assert has_canonical_rd_issued_path(
        issued_pdf.replace("Для передачи", "Для_передачи"), rd_root
    )
    assert not has_canonical_rd_issued_path(
        r"\\bcc\eng\PrDoc\РД\9110\06_KSB_21\05_рев.01\PDF"
        r"\AGCC.287-9110-KSB1.OD-0001_01_RU.pdf",
        rd_root,
    )
    assert not has_canonical_rd_issued_path(
        r"\\bcc\eng\PrDoc\РД\9110\06_KSB_21"
        r"\AGCC.287-9110-KSB1.OD-0001_01_RU.pdf",
        rd_root,
    )
    assert not has_canonical_rd_issued_path(
        r"\\bcc\eng\PrDoc\РД\9110\06_KSB_21\Для передачи"
        r"\AGCC.287-9110-KSB1.OD-0001_01_RU.pdf",
        rd_root,
    )
    assert not has_canonical_rd_issued_path(
        r"\\bcc\eng\PrDoc\РД\9110\06_KSB_21\Для передачи"
        r"\05_рев.01\extra\PDF\AGCC.287-9110-KSB1.OD-0001_01_RU.pdf",
        rd_root,
    )
    issued_dwg = issued_pdf.replace("\\PDF\\", "\\DWG\\").replace(".pdf", ".dwg")
    assert issued_package_dir(issued_dwg) == issued_folder
    issued_bbb = issued_pdf.replace("\\PDF\\", "\\BBB\\")
    assert issued_package_dir(issued_bbb) == issued_folder
    assert has_canonical_rd_issued_path(issued_bbb, rd_root)
    mark_folder = r"\\bcc\eng\PrDoc\РД\9110\06_KSB_21"
    assert kit_rd_mark_folder_from_path(issued_pdf, rd_root, title="9110") == (
        mark_folder
    )
    assert kit_rd_mark_folder_from_path(
        issued_no_media, rd_root, title="9110"
    ) == mark_folder
    assert kit_rd_mark_folder_from_path(issued_pdf, rd_root, title="8950") == ""
    assert unique_kit_rd_mark_folders(
        (issued_pdf, issued_no_media),
        rd_root,
        title="9110",
    ) == (mark_folder,)
    second_mark = issued_pdf.replace("\\06_KSB_21\\", "\\KSB1\\")
    assert unique_kit_rd_mark_folders(
        (issued_pdf, second_mark),
        rd_root,
        title="9110",
    ) == (mark_folder, r"\\bcc\eng\PrDoc\РД\9110\KSB1")
    assert constructed_kit_rd_mark_folder(rd_root, "9110", "KSB1") == (
        r"\\bcc\eng\PrDoc\РД\9110\KSB1"
    )
    gate_under_title = (
        r"\\bcc\eng\PrDoc\377_НИПИГАЗ\АГХК\КСБ\РД\9000\03_Для передачи"
        r"\05_рев.01\PDF\AGCC.287-9000-KSB.OD-0001_01_RU.pdf"
    )
    user_rd_root = r"\\bcc\eng\PrDoc\377_НИПИГАЗ\АГХК\КСБ\РД"
    assert kit_rd_mark_folder_from_path(
        gate_under_title, user_rd_root, title="9000"
    ) == ""
    assert unique_kit_rd_mark_folders(
        (gate_under_title,),
        user_rd_root,
        title="9000",
    ) == ()
    assert constructed_kit_rd_title_folder(user_rd_root, "9000") == (
        r"\\bcc\eng\PrDoc\377_НИПИГАЗ\АГХК\КСБ\РД\9000"
    )
    robot_xlsx = (
        r"\\bcc\eng\PrDoc\RFP_MTO_VO\_ГОТОВЫЕ_ДЛЯ_РОБОТА\9110\KSB1"
        r"\AGCC.287-9110-KSB1.MTO-0001_0-AN02_RU.xlsx"
    )
    robot_folder = (
        r"\\bcc\eng\PrDoc\RFP_MTO_VO\_ГОТОВЫЕ_ДЛЯ_РОБОТА\9110\KSB1"
    )
    assert issued_package_dir(robot_xlsx) == robot_folder
    assert not is_transfer_folder_name("KSB1")

    skips = load_config().skip_dirs
    assert "пример" in skips
    assert "Замечания" in skips
    assert _is_skipped_dir("пример", skips)
    assert _is_skipped_dir("Пример ОФОРМЛЕНИЯ AN1", skips)
    assert _is_skipped_dir("замечания", skips)
    assert _is_skipped_dir("Замечания", skips)
    assert _is_skipped_dir("Для сравнения", skips)
    assert _is_skipped_dir("_Сравнение МТО МТО", skips)
    assert _is_skipped_dir("05_ePlan", skips)
    assert _is_skipped_dir("05_Eplan", skips)
    assert _is_skipped_dir("02_МДЗ", skips)
    assert _is_skipped_dir("03_Интерфейсы", skips)
    assert _is_skipped_dir("01_Таблица оснащения", skips)
    assert _is_skipped_dir("00_3D", skips)
    assert _is_skipped_dir("05_Выгрузка nanoCAD", skips)
    assert _is_skipped_dir("Объекты в закупку", skips)
    assert _is_skipped_dir("99_Замечания для исполнителей", skips)
    assert _is_skipped_dir("_WORK", skips)
    assert _is_skipped_dir("02_FROM", skips)
    assert _is_skipped_dir("2xxx", skips)
    assert _is_skipped_dir("99_Вспомогательные", skips)
    assert _is_skipped_dir("Прочее", skips)
    assert _is_skipped_dir("ЗИП", skips)
    assert _is_skipped_dir("99_Наработки", skips)
    assert _is_skipped_dir("Загрузка СР", skips)
    assert _is_skipped_dir("Первая величина", skips)
    assert _is_skipped_dir("03_SENT", skips)
    assert _is_skipped_dir("Цесис", skips)
    assert _is_skipped_dir("черновики", skips)
    assert _is_skipped_dir("Оld", skips)
    assert _is_skipped_dir("Comment Attachments", skips)
    assert not _is_skipped_dir("Для передачи", skips)
    assert not _is_skipped_dir("Для_передачи", skips)
    assert not _is_skipped_dir("11_SKUD", skips)
    assert not _is_skipped_dir("05_рев.0_AGCC.287-3152-KSB", skips)
    assert not _is_skipped_dir(
        "08_рев.04_AGCC.287-8441-SKUD_в работе", skips
    )

    with tempfile.TemporaryDirectory(prefix="rd_catalog_config_") as temp:
        config = load_config(environment={"LOCALAPPDATA": temp})
        assert config.runtime_dir == Path(
            temp, "Documentation_PDF_out_NK", "rd_catalog"
        )
        assert not config.runtime_dir.exists()

        rd = Path(temp, "rd")
        kept = (
            rd
            / "7421"
            / "11_SKUD"
            / "Для передачи"
            / "04_рев.01_7421-SKUD.1"
            / "PDF"
        )
        sample = (
            rd
            / "3240"
            / "10_KSB1"
            / "Для передачи"
            / "08_рев.01-AN02_AGCC.287-3240-KSB1"
            / "DWG"
            / "Пример ОФОРМЛЕНИЯ AN1"
            / "Для сравнения"
        )
        remarks = (
            rd
            / "8950"
            / "12_POS1"
            / "замечания"
            / "01"
            / "Comment Attachments"
            / "AGCC.287-8950-POS1.OD-0001"
        )
        kept.mkdir(parents=True)
        sample.mkdir(parents=True)
        remarks.mkdir(parents=True)
        kept_name = "AGCC.287-7421-SKUD.1.OD-0001_01_RU.pdf"
        skipped_name = "AGCC.287-7421-SKUD.1.MTO-0001_01-AN01_RU.xlsx"
        remarks_name = "AGCC.287-8950-POS1.OD-0001_0_RU_НК.pdf"
        alien_pdf = "SQ-PP-AGCC-KSB-00051-0_signed.pdf"
        alien_xlsx = "SQ-PP-AGCC-KSB-00051-0.xlsx"
        (kept / kept_name).write_bytes(b"ok")
        (kept / alien_pdf).write_bytes(b"sq-pdf")
        (kept / alien_xlsx).write_bytes(b"sq-xlsx")
        (sample / skipped_name).write_bytes(b"sample")
        (remarks / remarks_name).write_bytes(b"comment")
        scanned = scan_document_source(rd, SourceKind.RD, skip_dirs=skips)
        names = {Path(item.path).name for item in scanned.files}
        assert kept_name in names
        assert skipped_name not in names
        assert remarks_name not in names
        assert alien_pdf not in names
        assert alien_xlsx not in names
        assert scanned.excluded_files >= 2

    def _folder_record(
        path: str,
        *,
        transfer_sequence: int | None = None,
        transfer_name: str | None = None,
    ) -> FileRecord:
        data: dict[str, object] = {
            "file_kind": FileKind.PDF.value,
            "name": Path(path).name,
            "parse_status": "parsed",
            "title": "8950",
            "mark": "POS4",
            "revision": "0",
        }
        if transfer_sequence is not None:
            data["transfer_sequence"] = transfer_sequence
        if transfer_name:
            data["transfer_name"] = transfer_name
        return FileRecord(
            id=1,
            path=path,
            path_key=path.casefold(),
            source=SourceKind.RD,
            present=True,
            review_state=ReviewState.ACKNOWLEDGED,
            first_seen_run_id=1,
            last_seen_run_id=1,
            data=data,
        )

    rec_nn4 = _folder_record(
        r"C:\rd\04\file.pdf",
        transfer_sequence=4,
        transfer_name="04_Рев.0_AGCC.287-8950-POS4",
    )
    rec_nn2 = _folder_record(
        r"C:\rd\02\file.pdf",
        transfer_sequence=2,
        transfer_name="02_Рев.0_AGCC.287-8950-POS4",
    )
    rec_from_name = _folder_record(
        r"C:\rd\01\file.pdf",
        transfer_name="01_AGCC.287-8950-POS4.MTO-0001_0_RU",
    )
    rec_loose = _folder_record(r"C:\rd\POS4\DWG\file.dwg")
    assert folder_transfer_sequence([rec_nn4]) == 4
    assert folder_transfer_sequence([rec_from_name]) == 1
    assert folder_transfer_sequence([rec_loose]) is None
    ordered = sorted(
        (
            ([rec_nn4], "04_рев.0_agcc.287-8950-pos4"),
            ([rec_nn2], "02_рев.0_agcc.287-8950-pos4"),
            ([rec_from_name], "01_agcc.287-8950-pos4.mto-0001_0_ru"),
            ([rec_loose], r"c:\rd\pos4\dwg"),
        ),
        key=lambda item: folder_tree_sort_key(item[0], folder_key=item[1]),
    )
    assert [item[1] for item in ordered] == [
        "01_agcc.287-8950-pos4.mto-0001_0_ru",
        "02_рев.0_agcc.287-8950-pos4",
        "04_рев.0_agcc.287-8950-pos4",
        r"c:\rd\pos4\dwg",
    ]
    print("RD catalog parse smoke: OK")


if __name__ == "__main__":
    main()
