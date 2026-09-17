"""Local checks for D/E sync rows built from the last KSB ИД F event."""

from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from rd_catalog.kits import (
    KitFlag,
    KitMatrixRow,
    KitSummary,
    SourceKitSnapshot,
    parse_google_kit_row,
)
from rd_catalog.sheet_de_sync import (
    list_sheet_de_sync_rows,
    sheet_de_sync_jobs,
    sheet_de_sync_row,
)

_F_CODE_A_04 = "02.12.2024 код А на рев. 04 AGCC-BCC-TRM-000582"
_F_OLDER = (
    "03.05.2024 код А на рев. 02 AGCC-BCC-TRM-000339\n" + _F_CODE_A_04
)
_F_AN = "17.02.2026 код А на рев. 02-AN02 AGCC.287-BCC-PGS-TRM-000623"
_F_AN_UNICODE = "17.02.2026 код А на рев. 02\u2010AN02 AGCC.287-BCC-PGS-TRM-000623"
_F_TDO_PASSED = "19.08.2026 прошла ТДО рев. 02 AGCC.287-PGS-PGS-TRM-22049"


def _google(
    *,
    title: str = "1600",
    mark: str = "SOT",
    d_cell: str = "Рев. 01",
    e_cell: str = "старый",
    comment: str = _F_CODE_A_04,
):
    parsed = parse_google_kit_row(
        [title, mark, "", d_cell, e_cell, comment],
        row_index=4,
    )
    assert not isinstance(parsed, str), parsed
    return parsed


def _row(
    *,
    google,
    rd_revision: str = "04",
    rd_appendix: str | None = None,
    present: bool = True,
    summary: KitSummary = KitSummary.REV_MISMATCH,
    flags: tuple[KitFlag, ...] = (KitFlag.REV_MISMATCH,),
    title: str | None = None,
    mark: str | None = None,
) -> KitMatrixRow:
    title = title or google.title
    mark = mark or google.mark
    rd_text = rd_revision if not rd_appendix else f"{rd_revision}-AN{rd_appendix}"
    return KitMatrixRow(
        title=title,
        mark=mark,
        title_system=f"{title}-{mark}",
        rd=SourceKitSnapshot(
            present=present,
            revision=rd_revision if present else None,
            appendix=rd_appendix if present else None,
            revision_text=rd_text if present else "",
        ),
        robot=SourceKitSnapshot(),
        sq=SourceKitSnapshot(),
        google=google,
        issuance=None,
        flags=flags,
        summary=summary,
    )


def main() -> None:
    """Stale D/E vs last F; skip when F itself still mismatches RD."""

    stale = _row(google=_google())
    item = sheet_de_sync_row(stale)
    assert item is not None
    assert item.title == "1600"
    assert item.mark == "SOT"
    assert item.rd_revision_text == "04"
    assert item.d_now == "01"
    assert item.d_next == "04"
    assert item.e_now == "старый"
    assert item.e_next == "РД Согласовано"
    assert item.write_d is True
    assert item.write_e is True
    assert item.write_label() == "D+E"
    assert item.d_cell == "Рев. 04"
    assert item.job.f_line == _F_CODE_A_04
    assert item.job.revision == "04"
    assert item.job.stage == "code_a"
    assert item.last_f_line == _F_CODE_A_04

    already = _row(
        google=_google(d_cell="Рев. 04", e_cell="РД Согласовано"),
        summary=KitSummary.ALIGNED,
        flags=(KitFlag.ALIGNED,),
    )
    assert sheet_de_sync_row(already) is None

    prefix_only = _row(
        google=_google(d_cell="02-AN02", e_cell="РД Согласовано", comment=_F_AN),
        rd_revision="02",
        rd_appendix="02",
        summary=KitSummary.ALIGNED,
        flags=(KitFlag.ALIGNED,),
    )
    assert sheet_de_sync_row(prefix_only) is None

    f_mismatch = _row(
        google=_google(),
        rd_revision="01",
        summary=KitSummary.REV_MISMATCH,
    )
    assert sheet_de_sync_row(f_mismatch) is None

    empty_d = _row(google=_google(d_cell="", e_cell="старый"))
    empty_item = sheet_de_sync_row(empty_d)
    assert empty_item is not None
    assert empty_item.write_d is True
    assert empty_item.d_now == "—"

    e_only = _row(
        google=_google(d_cell="Рев. 04", e_cell="старый"),
        summary=KitSummary.ALIGNED,
        flags=(KitFlag.ALIGNED,),
    )
    e_item = sheet_de_sync_row(e_only)
    assert e_item is not None
    assert e_item.write_d is False
    assert e_item.write_e is True
    assert e_item.d_next == "как есть"
    assert e_item.e_next == "РД Согласовано"
    assert e_item.write_label() == "E"

    code_b_line = "04.09.2026 код B на рев. 01-AN02 PGS-BCC-TRM-000532"
    code_b_stale = _row(
        google=_google(
            d_cell="Рев. 01-AN02",
            e_cell="Прошла входной контроль",
            comment=code_b_line,
        ),
        rd_revision="01",
        rd_appendix="02",
    )
    code_b_item = sheet_de_sync_row(code_b_stale)
    assert code_b_item is not None
    assert code_b_item.write_d is False
    assert code_b_item.write_e is True
    assert code_b_item.e_next == "РД_Корректировка по зам."
    assert code_b_item.d_next == "как есть"

    code_b_ok = _row(
        google=_google(
            d_cell="Рев. 01-AN02",
            e_cell="РД_Корректировка по зам.",
            comment=code_b_line,
        ),
        rd_revision="01",
        rd_appendix="02",
    )
    assert sheet_de_sync_row(code_b_ok) is None

    no_rd = _row(google=_google(), present=False, rd_revision="04")
    assert sheet_de_sync_row(no_rd) is None

    f_has_an_rd_bare = _row(
        google=_google(d_cell="Рев. 02-AN02", e_cell="РД Согласовано", comment=_F_AN),
        rd_revision="02",
        rd_appendix=None,
    )
    assert sheet_de_sync_row(f_has_an_rd_bare) is None

    unicode_an = _row(
        google=_google(
            d_cell="Рев. 02",
            e_cell="РД Согласовано",
            comment=_F_AN_UNICODE,
        ),
        rd_revision="02",
        rd_appendix="02",
    )
    unicode_item = sheet_de_sync_row(unicode_an)
    assert unicode_item is not None
    assert unicode_item.write_d is True
    assert unicode_item.d_now == "02"
    assert unicode_item.d_next == "02-AN02"
    assert unicode_item.d_next != "как есть"

    tdo_passed = _row(
        google=_google(
            d_cell="Рев. 02-AN04",
            e_cell="Прошла входной контроль",
            comment=_F_TDO_PASSED,
        ),
        rd_revision="02",
        rd_appendix=None,
    )
    tdo_item = sheet_de_sync_row(tdo_passed)
    assert tdo_item is not None
    assert tdo_item.write_d is True
    assert tdo_item.write_e is False
    assert tdo_item.d_now == "02-AN04"
    assert tdo_item.d_next == "02"
    assert tdo_item.e_next == "как есть"

    real_d = _row(
        google=_google(
            d_cell="Рев. 0-AN01",
            e_cell="РД Согласовано",
            comment="29.06.2026 код А на рев. 0-AN02 PGS-BCC-TRM-000437",
        ),
        rd_revision="0",
        rd_appendix="02",
    )
    real_item = sheet_de_sync_row(real_d)
    assert real_item is not None
    assert real_item.d_now == "0-AN01"
    assert real_item.d_next == "0-AN02"
    assert real_item.e_next == "как есть"

    review = _row(
        google=_google(comment=_F_OLDER),
        summary=KitSummary.TRANSFER_REVIEW,
        flags=(KitFlag.TRANSFER_REVIEW, KitFlag.REV_MISMATCH),
    )
    review_item = sheet_de_sync_row(review)
    assert review_item is not None
    assert review_item.job.f_line == _F_CODE_A_04

    other = _row(
        google=_google(comment="02.12.2024 выпустить до понедельника"),
    )
    assert sheet_de_sync_row(other) is None

    listed = list_sheet_de_sync_rows(
        (
            _row(google=_google(title="9110", mark="KSB1")),
            already,
            _row(google=_google(title="1600", mark="SOT")),
        )
    )
    assert [f"{item.title}-{item.mark}" for item in listed] == [
        "1600-SOT",
        "9110-KSB1",
    ]
    jobs = sheet_de_sync_jobs(listed)
    assert len(jobs) == 2
    assert jobs[0].title == "1600"
    assert jobs[1].mark == "KSB1"

    print("RD catalog sheet D/E sync: OK")


if __name__ == "__main__":
    main()
