"""Local checks for Google kits parsing and the presence matrix."""

from __future__ import annotations

import json
import sys
import tempfile
from datetime import datetime
from pathlib import Path
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from rd_catalog.config import CatalogConfig, load_config
from rd_catalog.customer_pi_auto_mto import compare_mto_pair
from rd_catalog.google_kits import (
    fetch_google_kits,
    issuance_cache_paths,
    load_cached_google_kits,
)
from rd_catalog.kits import (
    IssuanceKit,
    KitFlag,
    KitSummary,
    build_kit_matrix,
    code_a_date_for_revision,
    last_code_letter_for_revision,
    format_event_date_sortable,
    format_revision,
    extract_confirm_transmittal,
    is_rd_kit_mark,
    kit_identity_key,
    kit_revision_match_flags,
    kit_robot_origin,
    last_event_parts,
    last_event_text,
    mixed_title_open_folders,
    mixed_title_rescan_folders,
    mto_content_equal_by_kit,
    outlook_od_search_query,
    normalize_revision_digits,
    parse_google_matrix,
    parse_history_comment,
    parse_history_line,
    parse_issuance_matrix,
    parse_issuance_sends,
    parse_sheet_revision,
    strip_mark_suffix,
)
from rd_catalog.models import FileRecord, ReviewState, SourceKind
from rd_catalog.transfer_review_compare import (
    cache_entry_key,
    labels_from_cache,
    result_from_entry,
    result_to_entry,
    save_transfer_review_compare_cache,
)


def _record(
    file_id: int,
    *,
    source: SourceKind,
    title: str,
    mark: str,
    revision: str = "01",
    appendix: str | None = "02",
    transfer_revision: str | None = None,
    transfer_sequence: int = 3,
    transfer_name: str | None = None,
    path: str | None = None,
    file_kind: str = "pdf",
    present: bool = True,
    core_stem: str | None = None,
    discipline_block: str | None = None,
    mtime_ns: int = 1_700_000_000_000_000_000,
) -> FileRecord:
    resolved_path = path or rf"\\stub\{source.value}\{title}-{mark}.pdf"
    if discipline_block is None:
        discipline_block = "OD-0001" if file_kind == "pdf" else "MTO-0001"
    return FileRecord(
        id=file_id,
        path=resolved_path,
        path_key=f"{source.value}/{title}-{mark}/{file_id}",
        source=source,
        present=present,
        review_state=ReviewState.ACKNOWLEDGED,
        first_seen_run_id=1,
        last_seen_run_id=1,
        data={
            "file_kind": file_kind,
            "name": f"{title}-{mark}.pdf",
            "parse_status": "parsed",
            "title": title,
            "mark": mark,
            "revision": revision,
            "appendix": appendix,
            "transfer_revision": transfer_revision or revision,
            "transfer_appendix": appendix,
            "mtime_ns": mtime_ns,
            "transfer_is_as_build": 0,
            "transfer_sequence": transfer_sequence,
            "transfer_name": transfer_name or f"{transfer_sequence:02d}_рев.{revision}",
            "core_stem": core_stem or f"agcc.287-{title}-{mark}.{file_id}",
            "discipline_block": discipline_block,
            "title_system": f"{title}-{mark}",
        },
    )


def main() -> None:
    """Run parser/filter/matrix assertions without UNC or a live Google fetch."""

    assert strip_mark_suffix("POS_IFC") == "POS"
    assert strip_mark_suffix("KSB_IFR") == "KSB"
    assert is_rd_kit_mark("POS")
    assert is_rd_kit_mark("KSB1")
    assert is_rd_kit_mark("EPLAN")
    assert not is_rd_kit_mark("ТО")
    assert not is_rd_kit_mark("МДЗ")
    assert not is_rd_kit_mark("ИД")
    assert not is_rd_kit_mark("Интерфейс")
    assert not is_rd_kit_mark("NANOCAD")

    assert outlook_od_search_query("1600", "SOS") == '"1600-SOS.OD"'
    assert outlook_od_search_query("1600", "SOT") == '"1600-SOT.OD"'
    assert outlook_od_search_query("", "SOT") == ""
    assert outlook_od_search_query("1600", "  ") == ""

    assert is_rd_kit_mark("SS30")
    assert not is_rd_kit_mark("3D")
    assert not is_rd_kit_mark("3D 90")
    assert normalize_revision_digits("1") == "01"
    assert normalize_revision_digits("0") == "0"
    assert parse_sheet_revision("1") == ("01", None)
    assert parse_sheet_revision("0-AN02") == ("0", "02")

    assert parse_sheet_revision("Рев. 01-AN02") == ("01", "02")
    assert parse_sheet_revision("Рев. 0") == ("0", None)
    assert parse_sheet_revision("01-АН01") == ("01", "01")
    assert parse_sheet_revision("A") == ("A", None)
    assert parse_sheet_revision("Рев. A") == ("A", None)
    assert parse_sheet_revision("IFC") == (None, None)
    assert parse_sheet_revision("02\u2013AN02") == ("02", "02")
    assert parse_sheet_revision("02\u2010AN02") == ("02", "02")
    assert extract_confirm_transmittal(
        "PGS\u2013BCC\u2013TRM\u2013000334 extra"
    ) == "PGS-BCC-TRM-000334"

    comment = (
        "31.03.2024 AGCC-BCC-TRM-000311\n"
        "11.04.2025 отправлен на ТДО AGCC.287-PGS-PGS-TRM-20541\n"
        "20.05.2025 код А AGCC-BCC-TRM-000697\n"
        "17.12.2025 прошла ТДО рев.01-AN02 AGCC.287-BCC-PGS-TRM-000612"
    )
    events = parse_history_comment(comment)
    assert len(events) == 4
    assert events[0].date == "31.03.2024"
    assert events[0].transmittals == ("AGCC-BCC-TRM-000311",)
    assert events[1].stage == "tdo_sent"
    assert events[2].stage == "code_a"
    assert events[3].stage == "tdo_passed"
    assert events[3].revision == "01"
    assert events[3].appendix == "02"

    code_c_cyr = parse_history_line(
        "14.10.2024 код С CRS-PGS-AGCC-TRM-007715"
    )
    assert code_c_cyr.stage == "code_c"
    assert code_c_cyr.stage_label == "код C"
    assert code_c_cyr.date == "14.10.2024"
    code_c_lat = parse_history_line("код C ...")
    assert code_c_lat.stage == "code_c"
    assert code_c_lat.stage_label == "код C"
    assert parse_history_line("20.05.2025 код B AGCC-BCC-TRM-000001").stage == "code_b"
    sr_upload = parse_history_line("2025.06.03  - Отправлено на загрузку в СР")
    assert sr_upload.stage == "sr_upload"
    assert sr_upload.stage_label == "загрузка в СР"
    assert parse_history_line("отпр. на загрузку").stage == "sr_upload"
    assert parse_history_line("отпр на загр").stage == "sr_upload"
    assert parse_history_line("корр. в раб. пор").stage == "other"
    us_build = parse_history_line("US-BUILD")
    assert us_build.stage == "us_build"
    assert us_build.stage_label == "US-BUILD"
    assert parse_history_line("us build").stage == "us_build"
    assert parse_history_line("as-built").stage == "us_build"
    assert parse_history_line("выпустить до").stage == "other"
    en_dash_a = parse_history_line(
        "16.03.2026 код А не рев. 02\u2013AN02  PGS-BCC-TRM-000334"
    )
    assert en_dash_a.stage == "code_a"
    assert en_dash_a.revision == "02"
    assert en_dash_a.appendix == "02"
    assert "\u2013" in en_dash_a.raw
    hyphen_tdo = parse_history_line(
        "19.12.2025 прошла ТДО рев. 02\u2010AN02 AGCC.287-BCC-PGS-TRM-000637"
    )
    assert hyphen_tdo.stage == "tdo_passed"
    assert hyphen_tdo.revision == "02"
    assert hyphen_tdo.appendix == "02"
    dash_trm = parse_history_line(
        "16.03.2026 код А на рев. 02-AN02 PGS\u2013BCC\u2013TRM\u2013000334"
    )
    assert dash_trm.transmittals == ("PGS-BCC-TRM-000334",)
    assert parse_history_line("us\u2013build").stage == "us_build"

    rows = [
        ["титул", "ИД/ТО/МДЗ/РД", "имя", "Тип/ Стадия", "Статус", "Комментарий"],
        ["1513", "POS_IFC", "", "Рев. 01-AN02", "Прошла входной контроль", comment],
        ["1513", "Интерфейс", "", "неактивно", "ТО_Согласовано", "06.05.2024 RE"],
        ["1600", "ТО", "", "", "ТО_Согласовано", ""],
        ["2225", "KSB_IFC", "", "Рев. 01", "РД Согласовано", "10.08.2024 код А"],
    ]
    kits, stats = parse_google_matrix(rows)
    assert stats.kept == 2
    assert stats.skipped >= 2
    by_mark = {kit.mark: kit for kit in kits}
    assert by_mark["POS"].sheet_revision == "01"
    assert by_mark["POS"].sheet_appendix == "02"
    assert format_event_date_sortable("17.12.2025") == "2025.12.17"
    date, stage, rev, trm = last_event_parts(by_mark["POS"])
    assert date == "2025.12.17"
    assert stage == "прошла ТДО"
    assert rev == "01-AN02"
    assert "TRM-000612" in trm
    assert last_event_text(by_mark["POS"]).startswith("2025.12.17")

    pos_kit = by_mark["POS"]
    ksb_kit = by_mark["KSB"]
    google_only = build_kit_matrix((pos_kit,), (), set())
    assert google_only[0].summary is KitSummary.GOOGLE_ONLY
    assert KitFlag.GAP_RD in google_only[0].flags

    extra_rd = build_kit_matrix(
        (),
        [_record(1, source=SourceKind.RD, title="1715", mark="SOT")],
        {1},
    )
    assert extra_rd[0].summary is KitSummary.EXTRA_RD

    aligned = build_kit_matrix(
        (ksb_kit,),
        [
            _record(
                2,
                source=SourceKind.RD,
                title="2225",
                mark="KSB",
                revision="01",
                appendix=None,
                transfer_revision="01",
            ),
            _record(3, source=SourceKind.ROBOT, title="2225", mark="KSB"),
            _record(4, source=SourceKind.SQ, title="2225", mark="KSB"),
        ],
        {2},
    )
    assert aligned[0].rd.present and aligned[0].robot.present and aligned[0].sq.present
    assert aligned[0].summary is KitSummary.ALIGNED

    aligned_without_sq = build_kit_matrix(
        (ksb_kit,),
        [
            _record(
                20,
                source=SourceKind.RD,
                title="2225",
                mark="KSB",
                revision="01",
                appendix=None,
                transfer_revision="01",
            ),
            _record(21, source=SourceKind.ROBOT, title="2225", mark="KSB"),
        ],
        {20},
    )
    assert aligned_without_sq[0].summary is KitSummary.ALIGNED
    assert KitFlag.GAP_SQ not in aligned_without_sq[0].flags

    matching_revs = build_kit_matrix(
        (ksb_kit,),
        [
            _record(
                30,
                source=SourceKind.RD,
                title="2225",
                mark="KSB",
                revision="01",
                appendix=None,
                transfer_revision="01",
            ),
            _record(
                31,
                source=SourceKind.ROBOT,
                title="2225",
                mark="KSB",
                revision="1",
                appendix=None,
                transfer_revision="1",
            ),
        ],
        {30},
    )[0]
    match_flags = kit_revision_match_flags(matching_revs)
    assert match_flags["rd"] is True
    assert match_flags["google"] is True
    assert match_flags["robot"] is True
    assert match_flags["sq"] is None
    assert match_flags["issuance"] is None

    mismatched_revs = build_kit_matrix(
        (ksb_kit,),
        [
            _record(
                32,
                source=SourceKind.RD,
                title="2225",
                mark="KSB",
                revision="0",
                appendix="02",
                transfer_revision="0",
            ),
            _record(
                33,
                source=SourceKind.ROBOT,
                title="2225",
                mark="KSB",
                revision="0",
                appendix="01",
                transfer_revision="0",
            ),
        ],
        {32},
    )[0]
    mismatch_flags = kit_revision_match_flags(mismatched_revs)
    assert mismatch_flags["rd"] is True
    assert mismatch_flags["robot"] is False
    assert mismatch_flags["google"] is False

    mto_baseline = build_kit_matrix(
        (),
        [
            _record(
                34,
                source=SourceKind.RD,
                title="1600",
                mark="SOO",
                revision="03",
                appendix=None,
                transfer_sequence=10,
                core_stem="agcc.287-1600-soo.od-0001",
                discipline_block="OD-0001",
            ),
            _record(
                35,
                source=SourceKind.RD,
                title="1600",
                mark="SOO",
                revision="02",
                appendix=None,
                transfer_sequence=10,
                file_kind="mto_xlsx",
                core_stem="agcc.287-1600-soo.mto-0001",
                discipline_block="MTO-0001",
            ),
            _record(
                36,
                source=SourceKind.ROBOT,
                title="1600",
                mark="SOO",
                revision="02",
                appendix=None,
            ),
        ],
        {34, 35},
    )[0]
    mto_flags = kit_revision_match_flags(mto_baseline)
    assert mto_baseline.rd.revision_text == "03"
    assert mto_baseline.rd.mto_revision_text == "02"
    assert mto_flags["rd"] is True
    assert mto_flags["rd_mto"] is True
    assert mto_flags["robot"] is True

    filename_over_folder = build_kit_matrix(
        (),
        [
            _record(
                40,
                source=SourceKind.RD,
                title="9130",
                mark="KSB",
                revision="0",
                appendix="02",
                transfer_revision="0",
                transfer_sequence=3,
                transfer_name="03_рев.0-\u0410N02_AGCC.287\u20109130-KSB",
                path=(
                    r"\\stub\rd\9130\03_рев.0-AN02\PDF"
                    r"\AGCC.287-9130-KSB.MTO-0001_0-AN02_RU.pdf"
                ),
            ),
        ],
        {40},
    )[0]
    assert filename_over_folder.rd.revision_text == "0-AN02"

    older_9192 = (
        r"\\bcc\eng\rd\9192\04_SOT\Для передачи"
        r"\02_рев.0-AN01_AGCC.287-9192-SOT\PDF"
        r"\AGCC.287-9192-SOT.OD-0001_0-AN01_RU.pdf"
    )
    newer_9192 = (
        r"\\bcc\eng\rd\9192\04_SOT\Для передачи"
        r"\03_рев.0-AN02_AGCC.287-9192-SOT\PDF"
        r"\AGCC.287-9192-SOT.OD-0001_0-AN02_RU.pdf"
    )
    mixed_transfers = build_kit_matrix(
        (),
        [
            _record(
                41,
                source=SourceKind.RD,
                title="9192",
                mark="SOT",
                revision="0",
                appendix="01",
                transfer_revision="0",
                transfer_sequence=2,
                transfer_name="02_рев.0-AN01_AGCC.287-9192-SOT",
                path=older_9192,
            ),
            _record(
                42,
                source=SourceKind.RD,
                title="9192",
                mark="SOT",
                revision="0",
                appendix="02",
                transfer_revision="0",
                transfer_sequence=3,
                transfer_name="03_рев.0-AN02_AGCC.287-9192-SOT",
                path=newer_9192,
            ),
            _record(
                43,
                source=SourceKind.RD,
                title="9192",
                mark="SOT",
                revision="0",
                appendix="02",
                transfer_revision="0",
                transfer_sequence=3,
                transfer_name="03_рев.0-AN02_AGCC.287-9192-SOT",
                path=(
                    r"\\bcc\eng\rd\9192\04_SOT\Для передачи"
                    r"\03_рев.0-AN02_AGCC.287-9192-SOT\DWG"
                    r"\AGCC.287-9192-SOT.MTO-0001_0-AN02_RU.xlsx"
                ),
                file_kind="mto_xlsx",
            ),
        ],
        {41, 42, 43},
    )[0]
    assert mixed_transfers.rd.revision_text == "0-AN02"
    assert mixed_transfers.rd.paths[0] == newer_9192
    assert mixed_transfers.rd.transfer_name == "03_рев.0-AN02_AGCC.287-9192-SOT"

    same_rev_older = (
        r"\\bcc\eng\rd\9192\02_рев.0-AN01\PDF"
        r"\AGCC.287-9192-SOT.CAE-0002_0-AN01_RU.pdf"
    )
    same_rev_newer = (
        r"\\bcc\eng\rd\9192\03_рев.0-AN01\PDF"
        r"\AGCC.287-9192-SOT.CAE-0002_0-AN01_RU.pdf"
    )
    same_rev_reissue = build_kit_matrix(
        (),
        [
            _record(
                44,
                source=SourceKind.RD,
                title="9192",
                mark="SOT",
                revision="0",
                appendix="01",
                transfer_sequence=2,
                transfer_name="02_рев.0-AN01_AGCC.287-9192-SOT",
                path=same_rev_older,
            ),
            _record(
                45,
                source=SourceKind.RD,
                title="9192",
                mark="SOT",
                revision="0",
                appendix="01",
                transfer_sequence=3,
                transfer_name="03_рев.0-AN01_AGCC.287-9192-SOT",
                path=same_rev_newer,
            ),
        ],
        {44, 45},
    )[0]
    assert same_rev_reissue.rd.revision_text == "0-AN01"
    assert same_rev_reissue.rd.paths[0] == same_rev_newer

    cancelled_od = (
        r"\\stub\rd\9000\03_рев.0-AN02\PDF"
        r"\AGCC.287-9000-KSB.OD-0001_V_RU.pdf"
    )
    working_mto = (
        r"\\stub\rd\9000\03_рев.0-AN02"
        r"\AGCC.287-9000-KSB.MTO-0001_0-AN02_RU.xlsx"
    )
    letter_below_numeric = build_kit_matrix(
        (),
        [
            _record(
                50,
                source=SourceKind.RD,
                title="9000",
                mark="KSB",
                revision="V",
                appendix=None,
                transfer_sequence=3,
                path=cancelled_od,
            ),
            _record(
                51,
                source=SourceKind.RD,
                title="9000",
                mark="KSB",
                revision="0",
                appendix="02",
                transfer_sequence=3,
                path=working_mto,
                file_kind="mto_xlsx",
            ),
        ],
        {50, 51},
    )[0]
    assert letter_below_numeric.rd.revision_text == "V"
    assert letter_below_numeric.rd.mto_revision_text == "0-AN02"
    assert letter_below_numeric.rd.paths[0] == working_mto

    cancelled_only = build_kit_matrix(
        (),
        [
            _record(
                52,
                source=SourceKind.RD,
                title="9001",
                mark="KSB",
                revision="V",
                appendix=None,
                transfer_sequence=3,
                path=cancelled_od.replace("9000", "9001"),
            ),
        ],
        {52},
    )[0]
    assert cancelled_only.rd.revision_text == "V"

    mismatch = build_kit_matrix(
        (pos_kit,),
        [
            _record(
                5,
                source=SourceKind.RD,
                title="1513",
                mark="POS",
                revision="02",
                appendix=None,
                transfer_revision="02",
            )
        ],
        {5},
    )
    assert mismatch[0].summary is KitSummary.REV_MISMATCH

    stem = "agcc.287-8950-pos1.od-0001"
    gate = r"\\bcc\eng\PrDoc\РД\8950\12_POS1\Для передачи"
    inverted_kit = build_kit_matrix(
        (),
        [
            _record(
                60,
                source=SourceKind.RD,
                title="8950",
                mark="POS1",
                revision="04",
                appendix=None,
                transfer_sequence=10,
                core_stem=stem,
                mtime_ns=10,
                path=(
                    gate + r"\10_рев.04_AGCC.287-8950-POS1\PDF"
                    r"\AGCC.287-8950-POS1.OD-0001_04_RU.pdf"
                ),
                transfer_name="10_рев.04_AGCC.287-8950-POS1",
            ),
            _record(
                61,
                source=SourceKind.RD,
                title="8950",
                mark="POS1",
                revision="04",
                appendix="01",
                transfer_sequence=9,
                core_stem=stem,
                mtime_ns=20,
                path=(
                    gate + r"\09_рев.04-AN01_AGCC.287-8950-POS1\PDF"
                    r"\AGCC.287-8950-POS1.OD-0001_04-AN01_RU.pdf"
                ),
                transfer_name="09_рев.04-AN01_AGCC.287-8950-POS1",
            ),
        ],
        {60},
    )[0]
    assert inverted_kit.rd.revision_text == "04"
    assert inverted_kit.summary is KitSummary.TRANSFER_REVIEW
    assert KitFlag.TRANSFER_REVIEW in inverted_kit.flags
    assert inverted_kit.transfer_review_notes
    notes = "\n".join(inverted_kit.transfer_review_notes)
    assert "04-AN01" in notes
    assert "Что смущает:" in notes
    assert "ревизия ниже" in notes
    assert "файл старше" in notes
    assert "NN 10 текущая" in notes
    assert "Пути от папки передачи:" in notes
    assert "к отпр." in notes
    assert r"10_рев.04_AGCC.287-8950-POS1\PDF" in notes
    assert r"09_рев.04-AN01_AGCC.287-8950-POS1\PDF" in notes
    assert "текущая" in notes
    assert "спорная" in notes

    inverted_working = build_kit_matrix(
        (),
        [
            _record(
                60,
                source=SourceKind.RD,
                title="8950",
                mark="POS1",
                revision="04",
                appendix=None,
                transfer_sequence=10,
                core_stem=stem,
                mtime_ns=10,
                path=(
                    gate + r"\10_рев.04_AGCC.287-8950-POS1\PDF"
                    r"\AGCC.287-8950-POS1.OD-0001_04_RU.pdf"
                ),
                transfer_name="10_рев.04_AGCC.287-8950-POS1",
            ),
            _record(
                61,
                source=SourceKind.RD,
                title="8950",
                mark="POS1",
                revision="04",
                appendix="01",
                transfer_sequence=9,
                core_stem=stem,
                mtime_ns=20,
                path=(
                    gate + r"\09_рев.04-AN01_AGCC.287-8950-POS1\PDF"
                    r"\AGCC.287-8950-POS1.OD-0001_04-AN01_RU.pdf"
                ),
                transfer_name="09_рев.04-AN01_AGCC.287-8950-POS1",
            ),
        ],
        {60},
        working_folders_by_kit={
            kit_identity_key("8950", "POS1"): frozenset(
                {"09_рев.04-an01_agcc.287-8950-pos1"}
            )
        },
    )[0]
    assert inverted_working.summary is not KitSummary.TRANSFER_REVIEW
    assert not inverted_working.transfer_review_notes

    inverted_annulled = build_kit_matrix(
        (),
        [
            _record(
                60,
                source=SourceKind.RD,
                title="8950",
                mark="POS1",
                revision="04",
                appendix=None,
                transfer_sequence=10,
                core_stem=stem,
                mtime_ns=10,
                path=(
                    gate + r"\10_рев.04_AGCC.287-8950-POS1\PDF"
                    r"\AGCC.287-8950-POS1.OD-0001_04_RU.pdf"
                ),
                transfer_name="10_рев.04_AGCC.287-8950-POS1",
            ),
            _record(
                61,
                source=SourceKind.RD,
                title="8950",
                mark="POS1",
                revision="04",
                appendix="01",
                transfer_sequence=9,
                core_stem=stem,
                mtime_ns=20,
                path=(
                    gate + r"\09_рев.04-AN01_AGCC.287-8950-POS1\PDF"
                    r"\AGCC.287-8950-POS1.OD-0001_04-AN01_RU.pdf"
                ),
                transfer_name="09_рев.04-AN01_AGCC.287-8950-POS1",
            ),
        ],
        {60},
        annulled_folders_by_kit={
            kit_identity_key("8950", "POS1"): frozenset(
                {"09_рев.04-an01_agcc.287-8950-pos1"}
            )
        },
    )[0]
    assert inverted_annulled.summary is not KitSummary.TRANSFER_REVIEW
    assert not inverted_annulled.transfer_review_notes

    def _ns(year: int, month: int, day: int, hour: int, minute: int) -> int:
        return int(datetime(year, month, day, hour, minute).timestamp() * 1_000_000_000)

    def _issuance_row(
        title: str,
        mark: str,
        *,
        revision: str,
        appendix: str | None,
        send_date: str,
    ) -> IssuanceKit:
        return IssuanceKit(
            title=title,
            mark=mark,
            mark_raw=mark,
            title_system=f"{title}-{mark}",
            revision=revision,
            appendix=appendix,
            revision_text=format_revision(revision, appendix),
            status="",
            send_date=send_date,
            send_date_sortable=format_event_date_sortable(send_date),
            send_transmittal="",
            incoming_control_date="",
            incoming_control_date_sortable="",
            confirm_transmittal="",
            note_raw="",
            row_index=1,
        )

    sot_gate = r"\\bcc\eng\PrDoc\РД\2869\SOT\Для передачи"
    sot_stem = "agcc.287-2869-sot.od-0001"
    sot_mto_stem = "agcc.287-2869-sot.mto-0001"
    send_close = build_kit_matrix(
        (),
        [
            _record(
                64,
                source=SourceKind.RD,
                title="2869",
                mark="SOT",
                revision="0",
                appendix="02",
                transfer_sequence=10,
                core_stem=sot_stem,
                mtime_ns=_ns(2025, 3, 1, 14, 22),
                path=(
                    sot_gate + r"\10_рев.01_AN01_AGCC.287-2869-SOT_as-build\PDF"
                    r"\AGCC.287-2869-SOT-OD-0001_0-AN02_RU.pdf"
                ),
                transfer_name="10_рев.01_AN01_AGCC.287-2869-SOT_as-build",
            ),
            _record(
                65,
                source=SourceKind.RD,
                title="2869",
                mark="SOT",
                revision="01",
                appendix="02",
                transfer_sequence=9,
                core_stem=sot_stem,
                mtime_ns=_ns(2025, 4, 12, 9, 10),
                path=(
                    sot_gate + r"\09_рев.01_AN-02_AGCC.287-2869-SOT_as-build\PDF"
                    r"\AGCC.287-2869-SOT-OD-0001_01-AN02_RU.pdf"
                ),
                transfer_name="09_рев.01_AN-02_AGCC.287-2869-SOT_as-build",
            ),
            _record(
                66,
                source=SourceKind.RD,
                title="2869",
                mark="SOT",
                revision="01",
                appendix="02",
                transfer_sequence=9,
                file_kind="mto_xlsx",
                core_stem=sot_mto_stem,
                discipline_block="MTO-0001",
                mtime_ns=_ns(2025, 4, 12, 9, 12),
                path=(
                    sot_gate + r"\09_рев.01_AN-02_AGCC.287-2869-SOT_as-build"
                    r"\AGCC.287-2869-SOT.MTO-0001_01-AN02_RU.xlsx"
                ),
                transfer_name="09_рев.01_AN-02_AGCC.287-2869-SOT_as-build",
            ),
        ],
        {64},
        issuance_kits=(
            _issuance_row(
                "2869",
                "SOT",
                revision="01",
                appendix="02",
                send_date="12.04.2025",
            ),
        ),
    )[0]
    send_notes = "\n".join(send_close.transfer_review_notes)
    assert send_close.summary is KitSummary.TRANSFER_REVIEW
    assert "MTO — нет в NN 10, есть в NN 09" in send_notes
    assert "MTO-0001" in send_notes
    assert "нет в пакете" in send_notes
    assert "Близость к отправке: Выдача · 01-AN02 · 12.04.2025" in send_notes
    assert "лучше" in send_notes
    assert "к отправке ближе спорная NN 09" in send_notes
    assert "сверка" in send_notes
    assert "нет пары" in send_notes

    agreed_comment = (
        "24.03.2026 отпр на ТДО рев. 0-AN01 AGCC.287-BCC-PGS-TRM-000776\n"
        "24.03.2026 прошла ТДО рев. 0-AN01 AGCC.287-BCC-PGS-TRM-000776\n"
        "27.04.2026 код А на рев. 0-AN01 PGS-BCC-TRM-000393"
    )
    agreed_kits, _agreed_stats = parse_google_matrix(
        [
            ["титул", "ИД/ТО/МДЗ/РД", "имя", "Тип/ Стадия", "Статус", "Комментарий"],
            ["3000", "KSB", "", "Рев. 0-AN01", "РД Согласовано", agreed_comment],
        ]
    )
    kit_3000 = agreed_kits[0]
    assert code_a_date_for_revision(kit_3000, "0-AN01") == "27.04.2026"
    assert code_a_date_for_revision(kit_3000, "02") == ""
    assert last_code_letter_for_revision(kit_3000, "0-AN01") == (
        "27.04.2026",
        "code_a",
    )
    assert last_code_letter_for_revision(kit_3000, "02") == ("", "")

    letter_comment = (
        "01.01.2026 код А на рев. 0-AN01 AGCC-BCC-TRM-000001\n"
        "15.02.2026 код B на рев. 0-AN01 AGCC-BCC-TRM-000002\n"
        "20.03.2026 код C на рев. 02 AGCC-BCC-TRM-000003"
    )
    letter_kits, _letter_stats = parse_google_matrix(
        [
            ["титул", "ИД/ТО/МДЗ/РД", "имя", "Тип/ Стадия", "Статус", "Комментарий"],
            ["6550", "SKUD", "", "Рев. 0-AN01", "На рассмотрении ТДО", letter_comment],
        ]
    )
    kit_letters = letter_kits[0]
    assert last_code_letter_for_revision(kit_letters, "0-AN01") == (
        "15.02.2026",
        "code_b",
    )
    assert last_code_letter_for_revision(kit_letters, "02") == (
        "20.03.2026",
        "code_c",
    )
    assert code_a_date_for_revision(kit_letters, "0-AN01") == "01.01.2026"
    tdo_only_kits, _tdo_stats = parse_google_matrix(
        [
            ["титул", "ИД/ТО/МДЗ/РД", "имя", "Тип/ Стадия", "Статус", "Комментарий"],
            [
                "6550",
                "SKUD",
                "",
                "Рев. 0-AN01",
                "На рассмотрении ТДО",
                (
                    "16.09.2026 отпр на ТДО рев. 0-AN01 "
                    "AGCC.287-BCC-PGS-TRM-000776\n"
                    "17.09.2026 прошла ТДО рев. 0-AN01 "
                    "AGCC.287-BCC-PGS-TRM-000776"
                ),
            ],
        ]
    )
    assert last_code_letter_for_revision(tdo_only_kits[0], "0-AN01") == ("", "")
    issuance_3000 = _issuance_row(
        "3000",
        "KSB",
        revision="0",
        appendix="01",
        send_date="24.03.2026",
    )
    gate_3000 = r"\\bcc\eng\PrDoc\РД\3000\KSB\Для передачи"
    stem_3000_od = "agcc.287-3000-ksb.od-0001"
    stem_3000_mto = "agcc.287-3000-ksb.mto-0001"

    def _pack_3000(
        file_id: int,
        *,
        seq: int,
        folder: str,
        revision: str,
        appendix: str | None,
        file_kind: str,
        stem: str,
        mtime_ns: int,
        name: str,
    ) -> FileRecord:
        media = "DWG" if file_kind == "mto_xlsx" else "PDF"
        return _record(
            file_id,
            source=SourceKind.RD,
            title="3000",
            mark="KSB",
            revision=revision,
            appendix=appendix,
            transfer_sequence=seq,
            file_kind=file_kind,
            core_stem=stem,
            mtime_ns=mtime_ns,
            path=rf"{gate_3000}\{folder}\{media}\{name}",
            transfer_name=folder,
        )

    late_mto = build_kit_matrix(
        (kit_3000,),
        [
            _pack_3000(
                301,
                seq=8,
                folder="08_рев.0-AN01_AGCC.287-3000-KSB",
                revision="0",
                appendix="01",
                file_kind="mto_xlsx",
                stem=stem_3000_mto,
                mtime_ns=_ns(2026, 2, 26, 17, 3),
                name="AGCC.287-3000-KSB.MTO-0001_0-AN01_RU.xlsx",
            ),
            _pack_3000(
                302,
                seq=8,
                folder="08_рев.0-AN01_AGCC.287-3000-KSB",
                revision="0",
                appendix="01",
                file_kind="pdf",
                stem=stem_3000_od,
                mtime_ns=_ns(2026, 2, 26, 15, 21),
                name="AGCC.287-3000-KSB.OD-0001_0-AN01_RU.pdf",
            ),
            _pack_3000(
                303,
                seq=9,
                folder="09_рев.0-AN01_AGCC.287-3000-KSB",
                revision="0",
                appendix="01",
                file_kind="mto_xlsx",
                stem=stem_3000_mto,
                mtime_ns=_ns(2026, 5, 19, 12, 39),
                name="AGCC.287-3000-KSB.MTO-0001_0-AN01_RU.xlsx",
            ),
            _pack_3000(
                304,
                seq=9,
                folder="09_рев.0-AN01_AGCC.287-3000-KSB",
                revision="0",
                appendix="01",
                file_kind="pdf",
                stem=stem_3000_od,
                mtime_ns=_ns(2026, 3, 23, 17, 14),
                name="AGCC.287-3000-KSB.OD-0001_0-AN01_RU.pdf",
            ),
            _pack_3000(
                305,
                seq=11,
                folder="11_рев.0-AN02_AGCC.287-3000-KSB",
                revision="0",
                appendix="02",
                file_kind="mto_xlsx",
                stem=stem_3000_mto,
                mtime_ns=_ns(2026, 5, 19, 13, 24),
                name="AGCC.287-3000-KSB.MTO-0001_0-AN02_RU.xlsx",
            ),
            _record(
                306,
                source=SourceKind.ROBOT,
                title="3000",
                mark="KSB",
                revision="0",
                appendix="01",
                file_kind="mto_xlsx",
                core_stem=stem_3000_mto,
                mtime_ns=_ns(2026, 3, 20, 17, 6),
            ),
        ],
        {303, 304},
        issuance_kits=(issuance_3000,),
    )[0]
    late_notes = "\n".join(late_mto.transfer_review_notes)
    assert late_mto.summary is KitSummary.TRANSFER_REVIEW
    assert "Согласованная передача" in late_notes
    assert "позже письма A" in late_notes
    assert "27.04.2026" in late_notes
    assert "19.05.2026" in late_notes
    assert "NN 09" in late_notes

    timely_mto = build_kit_matrix(
        (kit_3000,),
        [
            _pack_3000(
                311,
                seq=9,
                folder="09_рев.0-AN01_AGCC.287-3000-KSB",
                revision="0",
                appendix="01",
                file_kind="mto_xlsx",
                stem=stem_3000_mto,
                mtime_ns=_ns(2026, 4, 20, 12, 0),
                name="AGCC.287-3000-KSB.MTO-0001_0-AN01_RU.xlsx",
            ),
            _pack_3000(
                312,
                seq=9,
                folder="09_рев.0-AN01_AGCC.287-3000-KSB",
                revision="0",
                appendix="01",
                file_kind="pdf",
                stem=stem_3000_od,
                mtime_ns=_ns(2026, 3, 23, 17, 14),
                name="AGCC.287-3000-KSB.OD-0001_0-AN01_RU.pdf",
            ),
            _record(
                313,
                source=SourceKind.ROBOT,
                title="3000",
                mark="KSB",
                revision="0",
                appendix="01",
                file_kind="mto_xlsx",
                core_stem=stem_3000_mto,
                mtime_ns=_ns(2026, 4, 20, 12, 0),
            ),
        ],
        {311, 312},
        issuance_kits=(issuance_3000,),
    )[0]
    assert timely_mto.summary is KitSummary.ALIGNED
    assert not timely_mto.transfer_review_notes

    working_folder_10 = "10_рев.0-AN01_AGCC.287-3000-KSB_as-build"
    official_folder_09 = "09_рев.0-AN01_AGCC.287-3000-KSB"
    working_pack = [
        _pack_3000(
            321,
            seq=9,
            folder=official_folder_09,
            revision="0",
            appendix="01",
            file_kind="mto_xlsx",
            stem=stem_3000_mto,
            mtime_ns=_ns(2026, 4, 20, 12, 0),
            name="AGCC.287-3000-KSB.MTO-0001_0-AN01_RU.xlsx",
        ),
        _pack_3000(
            322,
            seq=9,
            folder=official_folder_09,
            revision="0",
            appendix="01",
            file_kind="pdf",
            stem=stem_3000_od,
            mtime_ns=_ns(2026, 3, 23, 17, 14),
            name="AGCC.287-3000-KSB.OD-0001_0-AN01_RU.pdf",
        ),
        _pack_3000(
            323,
            seq=10,
            folder=working_folder_10,
            revision="0",
            appendix="01",
            file_kind="mto_xlsx",
            stem=stem_3000_mto,
            mtime_ns=_ns(2026, 5, 19, 13, 24),
            name="AGCC.287-3000-KSB.MTO-0001_0-AN01_RU.xlsx",
        ),
        _pack_3000(
            324,
            seq=10,
            folder=working_folder_10,
            revision="0",
            appendix="01",
            file_kind="pdf",
            stem=stem_3000_od,
            mtime_ns=_ns(2026, 5, 19, 13, 0),
            name="AGCC.287-3000-KSB.OD-0001_0-AN01_RU.pdf",
        ),
        _record(
            325,
            source=SourceKind.ROBOT,
            title="3000",
            mark="KSB",
            revision="0",
            appendix="01",
            file_kind="mto_xlsx",
            core_stem=stem_3000_mto,
            mtime_ns=_ns(2026, 4, 20, 12, 0),
        ),
    ]
    working_blind = build_kit_matrix(
        (kit_3000,),
        working_pack,
        {321, 322},
        issuance_kits=(issuance_3000,),
    )[0]
    assert working_blind.summary is KitSummary.TRANSFER_REVIEW
    working_notes = "\n".join(working_blind.transfer_review_notes)
    assert "текущая" in working_notes
    working_marked = build_kit_matrix(
        (kit_3000,),
        working_pack,
        {321, 322},
        issuance_kits=(issuance_3000,),
        working_folders_by_kit={
            kit_identity_key("3000", "KSB"): frozenset(
                {working_folder_10.casefold()}
            )
        },
    )[0]
    assert working_marked.summary is KitSummary.ALIGNED
    assert not working_marked.transfer_review_notes

    def _mto_cells(code: str, qty: str) -> list[dict[str, object]]:
        return [
            {
                "CODE": code,
                "UNITS": "шт",
                "VALUES": qty,
                "TAGS": [],
                "NAME": "",
                "VENDOR": "",
                "TYPE_MARK": "",
            }
        ]

    mto_payload = {
        "nn10": _mto_cells("BCC1", "2"),
        "nn09": _mto_cells("BCC1", "2"),
        "diff": _mto_cells("BCC1", "1"),
    }

    def _loader(path: str | Path):
        name = Path(path).name
        if "DIFF" in name:
            return mto_payload["diff"]
        if "_0-AN02_" in name:
            return mto_payload["nn10"]
        return mto_payload["nn09"]

    def _grade(left: str, right: str) -> str:
        return compare_mto_pair(left, right, loader=_loader).paren_label

    mto10_path = (
        sot_gate + r"\10_рев.01_AN01_AGCC.287-2869-SOT_as-build"
        r"\AGCC.287-2869-SOT.MTO-0001_0-AN02_RU.xlsx"
    )
    mto09_path = (
        sot_gate + r"\09_рев.01_AN-02_AGCC.287-2869-SOT_as-build"
        r"\AGCC.287-2869-SOT.MTO-0001_01-AN02_RU.xlsx"
    )
    both_mtos = build_kit_matrix(
        (),
        [
            _record(
                67,
                source=SourceKind.RD,
                title="2869",
                mark="SOT",
                revision="0",
                appendix="02",
                transfer_sequence=10,
                core_stem=sot_stem,
                mtime_ns=_ns(2025, 3, 1, 14, 22),
                path=(
                    sot_gate + r"\10_рев.01_AN01_AGCC.287-2869-SOT_as-build\PDF"
                    r"\AGCC.287-2869-SOT-OD-0001_0-AN02_RU.pdf"
                ),
                transfer_name="10_рев.01_AN01_AGCC.287-2869-SOT_as-build",
            ),
            _record(
                68,
                source=SourceKind.RD,
                title="2869",
                mark="SOT",
                revision="01",
                appendix="02",
                transfer_sequence=9,
                core_stem=sot_stem,
                mtime_ns=_ns(2025, 4, 12, 9, 10),
                path=(
                    sot_gate + r"\09_рев.01_AN-02_AGCC.287-2869-SOT_as-build\PDF"
                    r"\AGCC.287-2869-SOT-OD-0001_01-AN02_RU.pdf"
                ),
                transfer_name="09_рев.01_AN-02_AGCC.287-2869-SOT_as-build",
            ),
            _record(
                69,
                source=SourceKind.RD,
                title="2869",
                mark="SOT",
                revision="0",
                appendix="02",
                transfer_sequence=10,
                file_kind="mto_xlsx",
                core_stem=sot_mto_stem,
                discipline_block="MTO-0001",
                mtime_ns=_ns(2025, 3, 1, 14, 20),
                path=mto10_path,
                transfer_name="10_рев.01_AN01_AGCC.287-2869-SOT_as-build",
            ),
            _record(
                80,
                source=SourceKind.RD,
                title="2869",
                mark="SOT",
                revision="01",
                appendix="02",
                transfer_sequence=9,
                file_kind="mto_xlsx",
                core_stem=sot_mto_stem,
                discipline_block="MTO-0001",
                mtime_ns=_ns(2025, 4, 12, 9, 12),
                path=mto09_path,
                transfer_name="09_рев.01_AN-02_AGCC.287-2869-SOT_as-build",
            ),
        ],
        {67, 69},
        mto_compare=_grade,
    )[0]
    both_notes = "\n".join(both_mtos.transfer_review_notes)
    assert "четкое" in both_notes
    assert both_notes.count("четкое") >= 2
    assert both_mtos.transfer_review_mto_pairs
    pair = both_mtos.transfer_review_mto_pairs[0]
    assert {pair.left_path, pair.right_path} == {mto10_path, mto09_path}

    pdf_copies = build_kit_matrix(
        (),
        [
            _record(
                81,
                source=SourceKind.RD,
                title="2869",
                mark="SOT",
                revision="0",
                appendix="02",
                transfer_sequence=10,
                file_kind="mto_xlsx",
                core_stem=sot_mto_stem,
                discipline_block="MTO-0001",
                mtime_ns=_ns(2025, 3, 1, 14, 20),
                path=mto10_path,
                transfer_name="10_рев.01_AN01_AGCC.287-2869-SOT_as-build",
            ),
            _record(
                82,
                source=SourceKind.RD,
                title="2869",
                mark="SOT",
                revision="0",
                appendix="02",
                transfer_sequence=10,
                file_kind="pdf",
                core_stem=sot_mto_stem,
                discipline_block="MTO-0001",
                mtime_ns=_ns(2025, 3, 1, 14, 20),
                path=(
                    sot_gate + r"\10_рев.01_AN01_AGCC.287-2869-SOT_as-build\PDF"
                    r"\AGCC.287-2869-SOT.MTO-0001_0-AN02_RU.pdf"
                ),
                transfer_name="10_рев.01_AN01_AGCC.287-2869-SOT_as-build",
            ),
            _record(
                83,
                source=SourceKind.RD,
                title="2869",
                mark="SOT",
                revision="01",
                appendix="02",
                transfer_sequence=9,
                file_kind="mto_xlsx",
                core_stem=sot_mto_stem,
                discipline_block="MTO-0001",
                mtime_ns=_ns(2025, 4, 12, 9, 12),
                path=mto09_path,
                transfer_name="09_рев.01_AN-02_AGCC.287-2869-SOT_as-build",
            ),
            _record(
                84,
                source=SourceKind.RD,
                title="2869",
                mark="SOT",
                revision="01",
                appendix="02",
                transfer_sequence=9,
                file_kind="pdf",
                core_stem=sot_mto_stem,
                discipline_block="MTO-0001",
                mtime_ns=_ns(2025, 4, 12, 9, 12),
                path=(
                    sot_gate + r"\09_рев.01_AN-02_AGCC.287-2869-SOT_as-build\PDF"
                    r"\AGCC.287-2869-SOT.MTO-0001_01-AN02_RU.pdf"
                ),
                transfer_name="09_рев.01_AN-02_AGCC.287-2869-SOT_as-build",
            ),
        ],
        {81, 82},
        mto_compare=_grade,
    )[0]
    pdf_notes = "\n".join(pdf_copies.transfer_review_notes)
    assert "MTO-0001_0-AN02_RU.xlsx" in pdf_notes
    assert "MTO-0001_01-AN02_RU.xlsx" in pdf_notes
    assert "MTO-0001_0-AN02_RU.pdf" not in pdf_notes
    assert "MTO-0001_01-AN02_RU.pdf" not in pdf_notes
    assert "ошибка" not in pdf_notes
    assert pdf_copies.transfer_review_mto_pairs
    pdf_pair = pdf_copies.transfer_review_mto_pairs[0]
    assert {pdf_pair.left_path, pdf_pair.right_path} == {mto10_path, mto09_path}

    key_ab = cache_entry_key(mto10_path, 1, mto09_path, 2)
    key_ba = cache_entry_key(mto09_path, 2, mto10_path, 1)
    assert key_ab == key_ba
    with tempfile.TemporaryDirectory(prefix="rd_tr_cmp_") as raw:
        runtime = Path(raw)
        from rd_catalog.customer_pi_auto_mto import MtoPairCompareResult

        matched = MtoPairCompareResult(kind="matched", grade="exact")
        save_transfer_review_compare_cache(
            runtime, {key_ab: result_to_entry(matched)}
        )
        labels = labels_from_cache(runtime)
        assert labels[key_ab] == "четкое"
        assert result_from_entry(result_to_entry(matched)).paren_label == "четкое"

    wir_stem = "agcc.287-1600-sot.wir-0011"
    od_stem = "agcc.287-1600-sot.od-0001"
    mto_stem = "agcc.287-1600-sot.mto-0001"
    wir_only_inversion = build_kit_matrix(
        (),
        [
            _record(
                70,
                source=SourceKind.RD,
                title="1600",
                mark="SOT",
                revision="04",
                appendix=None,
                transfer_sequence=19,
                core_stem=od_stem,
                discipline_block="OD-0001",
                mtime_ns=19,
            ),
            _record(
                71,
                source=SourceKind.RD,
                title="1600",
                mark="SOT",
                revision="03",
                appendix=None,
                transfer_sequence=18,
                core_stem=od_stem,
                discipline_block="OD-0001",
                mtime_ns=18,
            ),
            _record(
                72,
                source=SourceKind.RD,
                title="1600",
                mark="SOT",
                revision="03",
                appendix=None,
                transfer_sequence=19,
                core_stem=wir_stem,
                discipline_block="WIR-0011",
                mtime_ns=19,
            ),
            _record(
                73,
                source=SourceKind.RD,
                title="1600",
                mark="SOT",
                revision="04",
                appendix=None,
                transfer_sequence=16,
                core_stem=wir_stem,
                discipline_block="WIR-0011",
                mtime_ns=16,
            ),
            _record(
                74,
                source=SourceKind.RD,
                title="1600",
                mark="SOT",
                revision="03",
                appendix=None,
                transfer_sequence=19,
                file_kind="mto_xlsx",
                core_stem=mto_stem,
                discipline_block="MTO-0001",
                mtime_ns=19,
            ),
        ],
        {70, 72, 74},
    )[0]
    assert wir_only_inversion.rd.revision_text == "04"
    assert wir_only_inversion.rd.mto_revision_text == "03"
    assert KitFlag.TRANSFER_REVIEW not in wir_only_inversion.flags
    assert wir_only_inversion.summary is not KitSummary.TRANSFER_REVIEW
    assert not wir_only_inversion.transfer_review_notes

    wir_hides_od = build_kit_matrix(
        (),
        [
            _record(
                80,
                source=SourceKind.RD,
                title="1600",
                mark="SOT",
                revision="04",
                appendix=None,
                transfer_sequence=19,
                core_stem=od_stem,
                discipline_block="OD-0001",
                mtime_ns=19,
            ),
            _record(
                81,
                source=SourceKind.RD,
                title="1600",
                mark="SOT",
                revision="05",
                appendix=None,
                transfer_sequence=19,
                core_stem=wir_stem,
                discipline_block="WIR-0010",
                mtime_ns=20,
            ),
        ],
        {80, 81},
    )[0]
    assert wir_hides_od.rd.revision_text == "04"
    assert wir_hides_od.rd.mto_revision_text == ""
    assert KitFlag.TRANSFER_REVIEW not in wir_hides_od.flags

    mto_inverted = build_kit_matrix(
        (),
        [
            _record(
                90,
                source=SourceKind.RD,
                title="1600",
                mark="SOT",
                revision="02",
                appendix=None,
                transfer_sequence=21,
                file_kind="mto_xlsx",
                core_stem=mto_stem,
                discipline_block="MTO-0001",
                mtime_ns=21,
            ),
            _record(
                91,
                source=SourceKind.RD,
                title="1600",
                mark="SOT",
                revision="03",
                appendix=None,
                transfer_sequence=20,
                file_kind="mto_xlsx",
                core_stem=mto_stem,
                discipline_block="MTO-0001",
                mtime_ns=20,
            ),
        ],
        {90},
    )[0]
    assert mto_inverted.summary is KitSummary.TRANSFER_REVIEW
    assert mto_inverted.rd.revision_text == "02"
    assert mto_inverted.rd.mto_revision_text == "02"

    issuance_rows = [
        [
            "№ п.п.",
            "Трансмитл",
            "Марка",
            "Ревизия",
            "Титул",
            "Статус",
            "Дата отправки",
            "",
            "",
            "",
            "",
            "",
            "",
            "",
            "",
            "Дата прохождения входного контроля",
            "Примечание",
        ],
        [
            "1",
            "AGCC.287-BCC-PGS-TRM-000100",
            "POS",
            "1",
            "1513",
            "Принят вх.контр.",
            "10.01.2025",
            "",
            "",
            "",
            "",
            "",
            "",
            "",
            "",
            "11.01.2025",
            "AGCC.287-PGS-PGS-TRM-20001",
        ],
        [
            "2",
            "AGCC.287-BCC-PGS-TRM-000200",
            "POS",
            "01-AN02",
            "1513",
            "Принят вх.контр.",
            "17.12.2025",
            "",
            "",
            "",
            "",
            "",
            "",
            "",
            "",
            "18.12.2025",
            "AGCC.287-PGS-PGS-TRM-21000",
        ],
        [
            "3",
            "AGCC.287-BCC-PGS-TRM-000300",
            "3D 90",
            "3",
            "7180",
            "Принят вх.контр.",
            "01.02.2025",
            "",
            "",
            "",
            "",
            "",
            "",
            "",
            "",
            "",
            "",
        ],
        [
            "4",
            "AGCC.287-BCC-PGS-TRM-000400",
            "КСБ",
            "0",
            "2210",
            "Принят вх.контр.",
            "25.09.2024",
            "",
            "",
            "",
            "",
            "",
            "",
            "",
            "",
            "26.09.2024",
            "AGCC.287-PGS-PGS-TRM-19062",
        ],
    ]
    issuance, iss_stats = parse_issuance_matrix(issuance_rows)
    assert iss_stats.kept == 2
    by_iss = {kit.title_system: kit for kit in issuance}
    assert "1513-POS" in by_iss
    assert by_iss["1513-POS"].revision_text == "01-AN02"
    assert by_iss["1513-POS"].send_date_sortable == "2025.12.17"
    assert by_iss["1513-POS"].send_transmittal.endswith("000200")
    assert by_iss["1513-POS"].confirm_transmittal.endswith("21000")
    assert by_iss["1513-POS"].incoming_control_date_sortable == "2025.12.18"
    assert "2210-KSB" in by_iss
    assert by_iss["2210-KSB"].mark == "KSB"
    assert by_iss["2210-KSB"].revision_text == "0"
    all_sends = parse_issuance_sends(issuance_rows)
    assert len(all_sends) > len(issuance)
    pos_sends = [item for item in all_sends if item.title_system == "1513-POS"]
    assert len(pos_sends) == 2

    mixed = build_kit_matrix(
        (pos_kit,),
        [],
        set(),
        issuance_kits=issuance,
    )
    pos_row = next(row for row in mixed if row.title_system == "1513-POS")
    assert pos_row.issuance is not None
    assert pos_row.google is not None

    hour = 3_600_000_000_000
    day = 86_400_000_000_000
    base = 1_700_000_000_000_000_000

    def _mto(
        file_id: int,
        *,
        source,
        title: str,
        mark: str,
        revision: str,
        mtime_ns: int,
    ):
        return _record(
            file_id,
            source=source,
            title=title,
            mark=mark,
            revision=revision,
            appendix=None,
            file_kind="mto_xlsx",
            mtime_ns=mtime_ns,
            path=rf"\\stub\{source.value}\{title}-{mark}-{file_id}.xlsx",
        )

    close_rd = build_kit_matrix(
        (),
        [
            _mto(
                101,
                source=SourceKind.RD,
                title="8950",
                mark="POS5",
                revision="02",
                mtime_ns=base,
            ),
            _mto(
                102,
                source=SourceKind.ROBOT,
                title="8950",
                mark="POS5",
                revision="02",
                mtime_ns=base + 2 * hour,
            ),
            _mto(
                103,
                source=SourceKind.SQ,
                title="8950",
                mark="POS5",
                revision="02",
                mtime_ns=base + 10 * day,
            ),
        ],
        {101},
    )[0]
    assert close_rd.rd.mto_mtime_ns == base
    origin_rd = kit_robot_origin(close_rd)
    assert origin_rd.matched is True
    assert origin_rd.rd is True
    assert origin_rd.sq is False
    assert "РД" in origin_rd.reason

    close_sq = build_kit_matrix(
        (),
        [
            _mto(
                104,
                source=SourceKind.RD,
                title="8950",
                mark="POS5",
                revision="01",
                mtime_ns=base + 10 * day,
            ),
            _mto(
                105,
                source=SourceKind.ROBOT,
                title="8950",
                mark="POS5",
                revision="02",
                mtime_ns=base,
            ),
            _mto(
                106,
                source=SourceKind.SQ,
                title="8950",
                mark="POS5",
                revision="02",
                mtime_ns=base + hour,
            ),
        ],
        {104},
    )[0]
    origin_sq = kit_robot_origin(close_sq)
    assert origin_sq.matched is True
    assert origin_sq.sq is True
    assert origin_sq.rd is False

    both_close = build_kit_matrix(
        (),
        [
            _mto(
                107,
                source=SourceKind.RD,
                title="1111",
                mark="POS",
                revision="01",
                mtime_ns=base + 2 * day,
            ),
            _mto(
                108,
                source=SourceKind.SQ,
                title="1111",
                mark="POS",
                revision="02",
                mtime_ns=base + hour,
            ),
            _mto(
                109,
                source=SourceKind.ROBOT,
                title="1111",
                mark="POS",
                revision="02",
                mtime_ns=base,
            ),
        ],
        {107},
    )[0]
    origin_both = kit_robot_origin(both_close)
    assert origin_both.sq is True
    assert origin_both.rd is False
    origin_content = kit_robot_origin(both_close, rd_content_equal=True)
    assert origin_content.matched is True
    assert origin_content.rd is False
    assert origin_content.sq is True
    assert origin_content.content_equal is True
    assert "содержимое" in origin_content.reason

    pdf_newer = build_kit_matrix(
        (),
        [
            _record(
                110,
                source=SourceKind.RD,
                title="8950",
                mark="POS1",
                revision="02",
                appendix=None,
                file_kind="pdf",
                mtime_ns=base + 10 * day,
            ),
            _mto(
                111,
                source=SourceKind.RD,
                title="8950",
                mark="POS1",
                revision="02",
                mtime_ns=base,
            ),
            _mto(
                112,
                source=SourceKind.ROBOT,
                title="8950",
                mark="POS1",
                revision="02",
                mtime_ns=base + hour,
            ),
        ],
        {110, 111},
    )[0]
    assert pdf_newer.rd.max_mtime_ns == base + 10 * day
    assert pdf_newer.rd.mto_mtime_ns == base
    origin_pdf = kit_robot_origin(pdf_newer)
    assert origin_pdf.matched is True
    assert origin_pdf.rd is True

    far = build_kit_matrix(
        (),
        [
            _mto(
                113,
                source=SourceKind.RD,
                title="2222",
                mark="KSB",
                revision="02",
                mtime_ns=base,
            ),
            _mto(
                114,
                source=SourceKind.ROBOT,
                title="2222",
                mark="KSB",
                revision="02",
                mtime_ns=base + 10 * day,
            ),
        ],
        {113},
    )[0]
    origin_far = kit_robot_origin(far, rd_content_equal=False)
    assert origin_far.matched is False
    assert origin_far.rd is False
    assert "не близка" in origin_far.reason
    origin_far_content = kit_robot_origin(far, rd_content_equal=True)
    assert origin_far_content.matched is True
    assert origin_far_content.rd is True
    assert origin_far_content.content_equal is True

    lagged_od = build_kit_matrix(
        (),
        [
            _record(
                115,
                source=SourceKind.RD,
                title="6100",
                mark="SOS",
                revision="04",
                appendix=None,
                file_kind="pdf",
                discipline_block="OD-0001",
                mtime_ns=base,
            ),
            _mto(
                116,
                source=SourceKind.RD,
                title="6100",
                mark="SOS",
                revision="03",
                mtime_ns=base,
            ),
            _mto(
                117,
                source=SourceKind.ROBOT,
                title="6100",
                mark="SOS",
                revision="03",
                mtime_ns=base,
            ),
        ],
        {115, 116},
    )[0]
    assert lagged_od.rd.revision_text == "04"
    assert lagged_od.rd.mto_revision_text == "03"
    origin_lagged = kit_robot_origin(lagged_od)
    assert origin_lagged.matched is True
    assert origin_lagged.rd is True
    assert origin_lagged.content_equal is False
    origin_lagged_content = kit_robot_origin(lagged_od, rd_content_equal=True)
    assert origin_lagged_content.matched is True
    assert origin_lagged_content.content_equal is True

    no_robot = kit_robot_origin(filename_over_folder)
    assert no_robot.matched is None

    content_map = mto_content_equal_by_kit(
        [
            {
                "title": "8950",
                "mark": "POS5",
                "diff": {"content_status": "content_equal"},
            },
            {
                "title": "1111",
                "mark": "POS",
                "diff": {"content_status": "content_diff"},
            },
        ]
    )
    assert content_map[kit_identity_key("8950", "POS5")] is True
    assert content_map[kit_identity_key("1111", "POS")] is False

    with tempfile.TemporaryDirectory(prefix="rd_catalog_kits_") as temp:
        runtime = Path(temp, "runtime")
        runtime.mkdir()
        (runtime / "google_kits_data.json").write_text(
            json.dumps({"rows": rows}, ensure_ascii=False),
            encoding="utf-8",
        )
        cached = load_cached_google_kits(runtime)
        assert cached is not None
        assert cached.stats.kept == 2
        config = CatalogConfig(
            rd_root=Path(temp, "rd"),
            sq_root=Path(temp, "sq"),
            robot_root=Path(temp, "robot"),
            runtime_dir=runtime,
            db_path=runtime / "catalog.sqlite",
            robot_flat_structure=True,
            skip_dirs=("old",),
        )
        cache_only = fetch_google_kits(config, cache_only=True)
        assert cache_only.error is None
        assert cache_only.stats.kept == 2
        assert cache_only.issuance_kits == ()
        iss_data, _iss_meta = issuance_cache_paths(runtime)
        iss_data.write_text(
            json.dumps({"rows": issuance_rows}, ensure_ascii=False),
            encoding="utf-8",
        )
        export_ids: list[str] = []
        allow_issuance_export = False

        def fake_export(
            spreadsheet_id: str,
            sheet_name: str,
            data_path: Path,
            meta_path: Path,
        ) -> tuple[list[list[str]], dict[str, str]]:
            export_ids.append(spreadsheet_id)
            if spreadsheet_id == config.google_issuance_spreadsheet_id:
                if not allow_issuance_export:
                    raise AssertionError("issuance export must not be called")
                return issuance_rows, {}
            if spreadsheet_id != config.google_kits_spreadsheet_id:
                raise AssertionError(f"unexpected spreadsheet {spreadsheet_id}")
            return rows, {}

        with patch(
            "rd_catalog.google_kits._export_one_sheet",
            side_effect=fake_export,
        ):
            skipped = fetch_google_kits(config, include_issuance=False)
            assert skipped.error is None
            assert skipped.warning is None
            assert skipped.stats.kept == 2
            assert skipped.issuance_sends
            assert skipped.issuance_stats.kept == 2
            assert "ksb_id_export" in skipped.source
            assert "issuance_cache" in skipped.source
            assert "issuance_export" not in skipped.source
            assert config.google_kits_spreadsheet_id in export_ids
            assert config.google_issuance_spreadsheet_id not in export_ids

            export_ids.clear()
            allow_issuance_export = True
            both = fetch_google_kits(config)
            assert both.error is None
            assert config.google_kits_spreadsheet_id in export_ids
            assert config.google_issuance_spreadsheet_id in export_ids
            assert "issuance_export" in both.source
        loaded = load_config(environment={"LOCALAPPDATA": temp})
        assert loaded.google_kits_spreadsheet_id
        assert not loaded.runtime_dir.exists() or loaded.runtime_dir == Path(
            temp, "Documentation_PDF_out_NK", "rd_catalog"
        )

    mixed_titles = build_kit_matrix(
        (ksb_kit,),
        [
            _record(
                40,
                source=SourceKind.RD,
                title="2225",
                mark="KSB",
                revision="01",
                appendix=None,
                transfer_sequence=8,
                transfer_name="08_рев.01",
                path=(
                    r"\\stub\RD\2225\05_KSB\Для передачи"
                    r"\08_рев.01\PDF\AGCC.287-2225-KSB.OD-0001_01_RU.pdf"
                ),
            ),
            _record(
                41,
                source=SourceKind.RD,
                title="2225",
                mark="KSB",
                revision="01",
                appendix="01",
                transfer_sequence=9,
                transfer_name="09_рев.AN-01_AGC.287-2612-KSB_as-build",
                path=(
                    r"\\stub\RD\2612\05_KSB\Для передачи"
                    r"\09_рев.AN-01_AGC.287-2612-KSB_as-build\PDF"
                    r"\AGCC.287-2225-KSB.WIR-0010_01-AN01_RU.pdf"
                ),
            ),
            _record(42, source=SourceKind.ROBOT, title="2225", mark="KSB"),
        ],
        {40},
    )
    mixed_row = mixed_titles[0]
    assert mixed_row.summary is KitSummary.MIXED_TITLES
    assert KitFlag.MIXED_TITLES in mixed_row.flags
    mixed_notes = "\n".join(mixed_row.mixed_title_notes)
    assert "папка 2612" in mixed_notes
    mixed_folders = mixed_title_rescan_folders(
        [
            _record(
                40,
                source=SourceKind.RD,
                title="2225",
                mark="KSB",
                revision="01",
                appendix=None,
                transfer_sequence=8,
                transfer_name="08_рев.01",
                path=(
                    r"\\stub\RD\2225\05_KSB\Для передачи"
                    r"\08_рев.01\PDF\AGCC.287-2225-KSB.OD-0001_01_RU.pdf"
                ),
            ),
            _record(
                41,
                source=SourceKind.RD,
                title="2225",
                mark="KSB",
                revision="01",
                appendix="01",
                transfer_sequence=9,
                transfer_name="09_рев.AN-01_AGC.287-2612-KSB_as-build",
                path=(
                    r"\\stub\RD\2612\05_KSB\Для передачи"
                    r"\09_рев.AN-01_AGC.287-2612-KSB_as-build\PDF"
                    r"\AGCC.287-2225-KSB.WIR-0010_01-AN01_RU.pdf"
                ),
            ),
        ],
        title="2225",
        mark="KSB",
        rd_root=r"\\stub\RD",
    )
    assert mixed_folders == (r"\\stub\RD\2612\05_KSB",)
    mixed_open = mixed_title_open_folders(
        [
            _record(
                50,
                source=SourceKind.RD,
                title="2210",
                mark="KSB",
                revision="0",
                appendix=None,
                transfer_sequence=2,
                transfer_name="02_рев.02_AGCC.287-2869-SKUD",
                path=(
                    r"\\stub\RD\2869\05_SKUD\Для передачи"
                    r"\02_рев.02_AGCC.287-2869-SKUD\DWG"
                    r"\AGCC.287-2210-KSB.WIR-0007_0_RU.dwg"
                ),
            ),
            _record(
                51,
                source=SourceKind.RD,
                title="2210",
                mark="KSB",
                revision="01",
                appendix="02",
                transfer_sequence=2,
                transfer_name="02_рев.01_AGCC.287-3140-KSB2",
                path=(
                    r"\\stub\RD\3140\06_KSB2\Для передачи"
                    r"\02_рев.01_AGCC.287-3140-KSB2\PDF"
                    r"\AGCC.287-2210-KSB.OD-0001_01-AN02_RU.pdf"
                ),
            ),
        ],
        title="2210",
        mark="KSB",
        rd_root=r"\\stub\RD",
    )
    assert len(mixed_open) == 2
    assert mixed_open[0].endswith(r"\DWG")
    assert mixed_open[1].endswith(r"\PDF")
    assert r"\2869\\" in mixed_open[0].replace("/", "\\") or "\\2869\\" in mixed_open[0]
    assert "3140" in mixed_open[1]

    print("RD catalog Google kits: OK")


if __name__ == "__main__":
    main()
