"""Local checks for AN MTO filename matching and schema-8 persist."""

from __future__ import annotations

import sys
import tempfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from rd_catalog.an_index import (
    AnMtoFile,
    KitAnTargets,
    agreed_revision_target,
    an_cell_text,
    an_file_kind,
    an_is_od,
    match_an_to_kit,
    parse_an_dump_file,
    parse_an_mto_file,
    score_an_files_for_agreed,
)
from rd_catalog.monitor_views import REV_DIFF_FILL, REV_MATCH_FILL, kits_an_cell
from rd_catalog.db import SCHEMA_VERSION, CatalogDatabase
from rd_catalog.kits import kit_identity_key
from rd_catalog.models import make_path_key

_STEM = "AGCC.287-1600-POS.MTO-0001"
_POS_TARGETS = KitAnTargets(
    auto_mto="02-AN01",
    auto_mto_stem=_STEM,
    rd_mto="02",
)


def _an_file(**overrides: object) -> AnMtoFile:
    path = str(overrides.get("path") or r"C:\an\AGCC.287-1600-POS.MTO-0001_02-AN01_RU.xlsx")
    fields = {
        "path": path,
        "path_key": make_path_key(path),
        "title": "1600",
        "mark": "POS",
        "revision_text": "02-AN01",
        "core_stem": _STEM,
        "discipline_block": "MTO-0001",
        "name": Path(path).name,
        "parent_dir": str(Path(path).parent),
        "mtime_ns": 1,
        "size": 2,
    }
    fields.update(overrides)
    return AnMtoFile(**fields)  # type: ignore[arg-type]


def _parse_named(directory: Path, name: str) -> AnMtoFile | None:
    path = directory / name
    path.write_bytes(b"")
    return parse_an_mto_file(path, size=path.stat().st_size, mtime_ns=1)


def main() -> None:
    """Run matching, parse-reject, and snapshot round-trip assertions."""

    assert SCHEMA_VERSION == 14

    closing = _an_file()
    hit = match_an_to_kit((closing,), _POS_TARGETS)
    assert hit.closes_auto_mto is True
    assert hit.shown_revision == "02-AN01"
    assert hit.shown_file is closing
    assert hit.match_auto_mto is True
    assert hit.match_rd_mto is False
    assert hit.stem_matched is True
    assert hit.stem_note == ""
    assert an_cell_text(hit) == "02-AN01"
    painted_auto = kits_an_cell(hit)
    assert painted_auto.fill == REV_DIFF_FILL
    assert painted_auto.bold is False
    assert "MTO · рев." in painted_auto.tooltip
    assert "Файлы АН:" in painted_auto.tooltip
    identity_lines = [
        line
        for line in painted_auto.tooltip.splitlines()
        if "AGCC.287-1600-POS.MTO-0001_02-AN01_RU.xlsx" in line
    ]
    path_lines = [
        line
        for line in painted_auto.tooltip.splitlines()
        if line.strip().startswith("C:\\")
    ]
    assert identity_lines
    assert path_lines
    assert all("C:\\an" not in line for line in identity_lines)
    assert closing.parent_dir in path_lines[0]

    only_rd = _an_file(
        path=r"C:\an\AGCC.287-1600-POS.MTO-0001_02_RU.xlsx",
        revision_text="02",
        name="AGCC.287-1600-POS.MTO-0001_02_RU.xlsx",
    )
    hit_rd = match_an_to_kit((only_rd,), _POS_TARGETS)
    assert hit_rd.closes_auto_mto is False
    assert hit_rd.shown_revision == "02"
    assert hit_rd.shown_file is only_rd
    assert hit_rd.match_auto_mto is False
    assert hit_rd.match_rd_mto is True
    assert hit_rd.stem_matched is True
    assert hit_rd.stem_note == ""
    painted_rd = kits_an_cell(hit_rd)
    assert painted_rd.fill == REV_MATCH_FILL
    assert painted_rd.bold is True

    other_stem = _an_file(
        path=r"C:\an\AGCC.287-1600-POS.MTO-0002_02-AN01_RU.xlsx",
        core_stem="AGCC.287-1600-POS.MTO-0002",
        name="AGCC.287-1600-POS.MTO-0002_02-AN01_RU.xlsx",
    )
    hit_stem = match_an_to_kit((other_stem,), _POS_TARGETS)
    assert hit_stem.closes_auto_mto is True
    assert hit_stem.shown_revision == "02-AN01"
    assert hit_stem.stem_matched is False
    assert hit_stem.stem_note == "ствол не совпал"

    empty = match_an_to_kit((), KitAnTargets())
    assert empty.shown_revision == ""
    assert empty.shown_file is None
    assert empty.closes_auto_mto is False
    assert empty.match_auto_mto is None
    assert an_cell_text(empty) == "—"
    painted_empty = kits_an_cell(empty)
    assert painted_empty.fill is None
    assert painted_empty.bold is False

    file_02 = _an_file(
        path=r"C:\an\AGCC.287-1600-POS.MTO-0001_02_RU.xlsx",
        revision_text="02",
        name="AGCC.287-1600-POS.MTO-0001_02_RU.xlsx",
    )
    file_an = _an_file(
        path=r"C:\an\AGCC.287-1600-POS.MTO-0001_02-AN01_RU.xlsx",
        revision_text="02-AN01",
    )
    hit_pair = match_an_to_kit((file_02, file_an), _POS_TARGETS)
    assert hit_pair.shown_file is file_an
    assert hit_pair.shown_revision == "02-AN01"
    assert an_cell_text(hit_pair) == "02-AN01 · 2"

    with tempfile.TemporaryDirectory(prefix="rd_catalog_an_") as temp:
        root = Path(temp)
        accepted = _parse_named(root, "AGCC.287-1600-POS.MTO-0001_02-AN01_RU.xlsx")
        assert accepted is not None
        assert accepted.title == "1600"
        assert accepted.mark == "POS"
        assert accepted.revision_text == "02-AN01"
        assert accepted.core_stem == _STEM
        assert accepted.discipline_block.casefold().startswith("mto")

        rejects = (
            "notes.xlsx",
            "AGCC.287-1600-POS.MTO-0001_02_RU.pdf",
            "~$AGCC.287-1600-POS.MTO-0001_02_RU.xlsx",
            "SQ-PP-AGCC-KSB-00051-0_signed.xlsx",
        )
        for name in rejects:
            assert _parse_named(root, name) is None, name

        od_ok = parse_an_dump_file(
            root / "AGCC.287-1600-POS.OD-0001_02-AN01_RU.docx",
            size=1,
            mtime_ns=1,
        )
        assert od_ok is not None
        assert an_is_od(od_ok)
        assert an_file_kind(od_ok) == "OD"
        assert od_ok.revision_text == "02-AN01"
        assert parse_an_dump_file(
            root / "~$AGCC.287-1600-POS.OD-0001_02_RU.docx",
            size=1,
            mtime_ns=1,
        ) is None
        assert parse_an_dump_file(
            root / "AGCC.287-1600-POS.OD-0001_02_RU.pdf",
            size=1,
            mtime_ns=1,
        ) is None

        db_path = root / "catalog.sqlite"
        database = CatalogDatabase(db_path)
        database.initialize()
        assert database.schema_version() == SCHEMA_VERSION
        first = _an_file(path=str(root / "old.xlsx"))
        database.replace_an_snapshot((first,), scanned_at="2026-09-13T00:00:00+00:00")
        listed = database.list_an_mto_files()
        assert len(listed) == 1
        assert listed[0].path == first.path
        assert listed[0].revision_text == first.revision_text
        by_kit = database.list_an_files_by_kit()
        assert by_kit[kit_identity_key("1600", "POS")][0].path == first.path
        second = _an_file(
            path=str(root / "new.xlsx"),
            title="9110",
            mark="KSB1",
        )
        database.replace_an_snapshot((second,), scanned_at="2026-09-13T01:00:00+00:00")
        replaced = database.list_an_mto_files()
        assert len(replaced) == 1
        assert replaced[0].path == second.path
        assert first.path not in {item.path for item in replaced}
        assert kit_identity_key("1600", "POS") not in database.list_an_files_by_kit()

    assert agreed_revision_target(
        code="A",
        code_revision_text="01-AN01",
        code_date="10.12.2025",
        status="agreed",
        official_revision_text="01-AN01",
    ) == ("01-AN01", "10.12.2025")
    assert agreed_revision_target(code="B", code_revision_text="02") == ("", "")
    assert agreed_revision_target(
        status="agreed", official_revision_text="01-AN01"
    ) == ("01-AN01", "")

    old_an = _an_file(
        path=r"\\bcc\an\Комплект 3_KSB_2230_rev.AN_от 27.08.25\AGCC.287-2230-KSB.MTO-0001_AN_RU.xlsx",
        title="2230",
        mark="KSB",
        revision_text="AN",
        name="AGCC.287-2230-KSB.MTO-0001_AN_RU.xlsx",
        parent_dir=r"\\bcc\an\Комплект 3_KSB_2230_rev.AN_от 27.08.25",
        mtime_ns=1756278000_000_000_000,
    )
    agreed_pack = _an_file(
        path=r"\\bcc\an\Комплект 5_KSB_2230_rev.01-AN01_от 26.11.2025\AGCC.287-2230-KSB.MTO-0001_01-AN01_RU.xlsx",
        title="2230",
        mark="KSB",
        revision_text="01-AN01",
        name="AGCC.287-2230-KSB.MTO-0001_01-AN01_RU.xlsx",
        parent_dir=r"\\bcc\an\Комплект 5_KSB_2230_rev.01-AN01_от 26.11.2025",
        mtime_ns=1764076800_000_000_000,
    )
    scores = score_an_files_for_agreed(
        (old_an, agreed_pack),
        KitAnTargets(
            robot="01-AN02",
            issuance="01-AN01",
            google_f="01-AN01",
            agreed="01-AN01",
            agreed_date="10.12.2025",
        ),
    )
    assert scores[agreed_pack.path_key].is_best is True
    assert scores[agreed_pack.path_key].percent is not None
    assert scores[agreed_pack.path_key].percent >= 70
    assert scores[old_an.path_key].is_best is False
    assert (scores[old_an.path_key].percent or 0) < scores[agreed_pack.path_key].percent

    as_build = _an_file(
        path=r"\\bcc\an\2230\As-build\Отправку\DWG\AGCC.287-2230-KSB.MTO-0001_01-AN01_RU.xlsx",
        title="2230",
        mark="KSB",
        revision_text="01-AN01",
        name="AGCC.287-2230-KSB.MTO-0001_01-AN01_RU.xlsx",
        parent_dir=r"\\bcc\an\2230\As-build\Отправку\DWG",
        mtime_ns=1773654145_000_000_000,
    )
    sq_send = _an_file(
        path=r"\\bcc\an\2230\SQ-MFCU-AGCC-KSB-02364-0\DWG\AGCC.287-2230-KSB.MTO-0001_01-AN01_RU.xlsx",
        title="2230",
        mark="KSB",
        revision_text="01-AN01",
        name="AGCC.287-2230-KSB.MTO-0001_01-AN01_RU.xlsx",
        parent_dir=r"\\bcc\an\2230\SQ-MFCU-AGCC-KSB-02364-0\DWG",
        mtime_ns=1764076859_000_000_000,
    )
    sq_answer = _an_file(
        path=r"\\bcc\an\Ответы на SQ-запросы\SQ-MFCU-AGCC-KSB-02364-0\DWG\AGCC.287-2230-KSB.MTO-0001_01-AN01_RU.xlsx",
        title="2230",
        mark="KSB",
        revision_text="01-AN01",
        name="AGCC.287-2230-KSB.MTO-0001_01-AN01_RU.xlsx",
        parent_dir=r"\\bcc\an\Ответы на SQ-запросы\SQ-MFCU-AGCC-KSB-02364-0\DWG",
        mtime_ns=1760625000_000_000_000,
    )
    folder_scores = score_an_files_for_agreed(
        (as_build, sq_send, sq_answer),
        KitAnTargets(
            issuance="01-AN01",
            google_f="01-AN01",
            robot="01-AN02",
            agreed="01-AN01",
            agreed_date="10.12.2025",
        ),
    )
    assert folder_scores[sq_send.path_key].is_best is True
    assert folder_scores[as_build.path_key].is_best is False
    assert folder_scores[sq_answer.path_key].is_best is False
    assert folder_scores[sq_send.path_key].percent > folder_scores[as_build.path_key].percent

    june_send = _an_file(
        path=r"\\bcc\an\2235\На отправку\KSB\01_рев_AN01_AGCC.287-1-2235-KSB\DWG\AGCC.287-2235-KSB.MTO-0001_01-AN01_RU.xlsx",
        title="2235",
        mark="KSB",
        revision_text="01-AN01",
        name="AGCC.287-2235-KSB.MTO-0001_01-AN01_RU.xlsx",
        parent_dir=r"\\bcc\an\2235\На отправку\KSB\01_рев_AN01_AGCC.287-1-2235-KSB\DWG",
        mtime_ns=1750367968_000_000_000,
    )
    nov_sq = _an_file(
        path=r"\\bcc\an\2235\SQ-MFCU-AGCC-KSB-02364-0\DWG\AGCC.287-2235-KSB.MTO-0001_01-AN01_RU.xlsx",
        title="2235",
        mark="KSB",
        revision_text="01-AN01",
        name="AGCC.287-2235-KSB.MTO-0001_01-AN01_RU.xlsx",
        parent_dir=r"\\bcc\an\2235\SQ-MFCU-AGCC-KSB-02364-0\DWG",
        mtime_ns=1763732253_000_000_000,
    )
    kit_2235 = score_an_files_for_agreed(
        (june_send, nov_sq),
        KitAnTargets(
            issuance="01-AN01",
            google_f="01-AN01",
            robot="01-AN02",
            agreed="01-AN01",
            agreed_date="29.11.2025",
        ),
    )
    assert kit_2235[nov_sq.path_key].is_best is True
    assert kit_2235[june_send.path_key].is_best is False
    assert (kit_2235[nov_sq.path_key].percent or 0) > (
        kit_2235[june_send.path_key].percent or 0
    )

    empty_target = score_an_files_for_agreed((agreed_pack,), KitAnTargets())
    assert empty_target[agreed_pack.path_key].percent is None
    assert empty_target[agreed_pack.path_key].is_best is False

    od_only = _an_file(
        path=r"\\bcc\an\8260\SKUD_01_рев.AN02\АН\AGCC.287-8260-SKUD.OD-0001_01-AN02_RU.docx",
        title="8260",
        mark="SKUD",
        revision_text="01-AN02",
        core_stem="AGCC.287-8260-SKUD.OD-0001",
        discipline_block="OD-0001",
        name="AGCC.287-8260-SKUD.OD-0001_01-AN02_RU.docx",
        parent_dir=r"\\bcc\an\8260\SKUD_01_рев.AN02\АН",
        mtime_ns=1747825200_000_000_000,
    )
    lagged_mto = _an_file(
        path=r"\\bcc\an\8260\SKUD_01_рев.AN02\АН\AGCC.287-8260-SKUD.MTO-0001_01-AN01_RU.xlsx",
        title="8260",
        mark="SKUD",
        revision_text="01-AN01",
        name="AGCC.287-8260-SKUD.MTO-0001_01-AN01_RU.xlsx",
        parent_dir=r"\\bcc\an\8260\SKUD_01_рев.AN02\АН",
        mtime_ns=1747825200_000_000_000,
    )
    other_folder_mto = _an_file(
        path=r"\\bcc\an\8260\SKUD_01_рев.AN01\АН\AGCC.287-8260-SKUD.MTO-0001_01-AN01_RU.xlsx",
        title="8260",
        mark="SKUD",
        revision_text="01-AN01",
        name="AGCC.287-8260-SKUD.MTO-0001_01-AN01_RU.xlsx",
        parent_dir=r"\\bcc\an\8260\SKUD_01_рев.AN01\АН",
        mtime_ns=1758121200_000_000_000,
    )
    od_targets = KitAnTargets(
        auto_mto="01-AN01",
        issuance="01-AN02",
        google_f="01-AN02",
        agreed="01-AN02",
        agreed_date="14.08.2025",
    )
    od_scores = score_an_files_for_agreed(
        (od_only, lagged_mto, other_folder_mto),
        od_targets,
    )
    assert od_scores[od_only.path_key].is_best is True
    assert (od_scores[od_only.path_key].percent or 0) > (
        od_scores[other_folder_mto.path_key].percent or 0
    )
    assert (od_scores[lagged_mto.path_key].percent or 0) > (
        od_scores[other_folder_mto.path_key].percent or 0
    )
    assert any(
        "OD согласованной" in reason
        for reason in od_scores[lagged_mto.path_key].reasons
    )
    hit_od_kit = match_an_to_kit((od_only, lagged_mto), od_targets)
    assert hit_od_kit.shown_file is lagged_mto
    assert hit_od_kit.closes_auto_mto is True
    hit_od_only = match_an_to_kit((od_only,), od_targets)
    assert hit_od_only.shown_file is od_only
    assert hit_od_only.closes_auto_mto is False

    print("RD catalog AN index: OK")


if __name__ == "__main__":
    main()
