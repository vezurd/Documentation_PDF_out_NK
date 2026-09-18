"""Local checks for F-journal insert/replace and D/E proposals."""

from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from rd_catalog.f_journal import (
    added_history_line,
    apply_history_line,
    build_journal_patch,
    catalog_f_line,
    format_history_line,
    format_sheet_revision_cell,
    journal_diff_html,
    journal_highlight_spans,
    journal_stage_key,
    journal_write_needed,
    status_sheet_for_stage,
)
from rd_catalog.kits import parse_history_comment, parse_history_line


_SOT_F = (
    "03.05.2024 код А на рев. 02 AGCC-BCC-TRM-000339\n"
    "07.05.2024 AGCC.287-BCC-NPG-TRM-003496\n"
    "03.10.2024 отпр на ТДО AGCC.287-BCC-PGS-TRM-000029\n"
    "14.10.2024 прошла вх контр AGCC.287-PGS-PGS-TRM-19220\n"
    "21.10.2024 код B\n"
    "02.12.2024 код А на рев. 04 AGCC-BCC-TRM-000582"
)


def main() -> None:
    """Run formatter and patch assertions without Google or Outlook."""

    assert journal_stage_key("code_a") == "code_a"
    assert journal_stage_key("код А") == "code_a"
    assert journal_stage_key("отпр на ТДО") == "tdo_sent"
    assert journal_stage_key("нет такой") == ""

    code_a = format_history_line(
        date="02.12.2024",
        stage="code_a",
        revision="04",
        transmittal="AGCC-BCC-TRM-000582",
    )
    assert code_a == "02.12.2024 код А на рев. 04 AGCC-BCC-TRM-000582"
    parsed = parse_history_line(code_a)
    assert parsed.stage == "code_a"
    assert parsed.date == "02.12.2024"
    assert parsed.revision == "04"
    assert parsed.transmittals == ("AGCC-BCC-TRM-000582",)

    incoming = format_history_line(
        date="08.05.2024",
        stage="incoming_passed",
        transmittal="AGCC.287-PGS-PGS-TRM-17769",
    )
    assert incoming == "08.05.2024 прошла вх контр AGCC.287-PGS-PGS-TRM-17769"
    assert parse_history_line(incoming).stage == "incoming_passed"

    cover_tdo = format_history_line(
        date="06.03.2026",
        stage="tdo_sent",
        revision="01-AN01",
        transmittal="AGCC.287-BCC-PGS-TRM-000750",
    )
    assert cover_tdo == (
        "06.03.2026 отпр на ТДО рев. 01-AN01 AGCC.287-BCC-PGS-TRM-000750"
    )
    cover_event = parse_history_line(cover_tdo)
    assert cover_event.stage == "tdo_sent"
    assert cover_event.revision == "01"
    assert cover_event.appendix == "01"

    left, right = journal_highlight_spans(
        "06.03.2026 отпр на ТДО AGCC.287-BCC-PGS-TRM-000750",
        cover_tdo,
    )
    assert left == ()
    marked = "".join(cover_tdo[start:end] for start, end in right)
    assert "рев. 01-AN01" in marked
    assert "background-color:#FFE082" in journal_diff_html(cover_tdo, right)

    chained_before = (
        "21.01.2025 код А на рев. 0 AGCC-BCC-TRM-000629\n"
        "06.03.2026 отпр на ТДО рев. 01-AN01 AGCC.287-BCC-PGS-TRM-000750"
    )
    chained_after = (
        chained_before + "\n10.03.2026 прошла вх контр AGCC.287-PGS-PGS-TRM-21710"
    )
    left, right = journal_highlight_spans(chained_before, chained_after)
    assert left == ()
    marked = "".join(chained_after[start:end] for start, end in right)
    assert "10.03.2026" in marked
    assert "21.01.2025" not in marked

    before_trm = "10.03.2026 прошла вх контр AGCC.287-PGS-PGS-TRM-21710"
    after_trm = "10.03.2026 прошла вх контр AGCC.287-PGS-PGS-TRM-21711"
    left, right = journal_highlight_spans(before_trm, after_trm)
    marked_before = "".join(before_trm[start:end] for start, end in left)
    marked_after = "".join(after_trm[start:end] for start, end in right)
    assert "10.03.2026" not in marked_before
    assert "10.03.2026" not in marked_after
    assert marked_before != marked_after
    assert journal_highlight_spans(cover_tdo, cover_tdo) == ((), ())

    replaced_tdo, tdo_action = apply_history_line(
        "06.03.2026 отпр на ТДО AGCC.287-BCC-PGS-TRM-000750",
        cover_tdo,
    )
    assert tdo_action == "replaced"
    assert replaced_tdo == cover_tdo

    sr_upload = format_history_line(
        date="12.08.2026",
        stage="sr_upload",
        transmittal="AGCC.287-BCC-PGS-TRM-000918",
    )
    assert "загрузку" in sr_upload
    assert parse_history_line(sr_upload).stage == "sr_upload"

    inserted, action = apply_history_line(
        _SOT_F,
        format_history_line(
            date="23.05.2024",
            stage="code_a",
            revision="03",
            transmittal="AGCC-BCC-TRM-000357",
        ),
    )
    assert action == "inserted"
    lines = inserted.split("\n")
    assert lines[2].startswith("23.05.2024 код А")
    assert lines[1] == "07.05.2024 AGCC.287-BCC-NPG-TRM-003496"
    assert lines[3].startswith("03.10.2024")

    replaced, action = apply_history_line(
        _SOT_F,
        format_history_line(
            date="21.10.2024",
            stage="code_b",
            revision="04",
            transmittal="PGS-BCC-TRM-000009",
        ),
    )
    assert action == "replaced"
    assert "21.10.2024 код B на рев. 04 PGS-BCC-TRM-000009" in replaced.split("\n")
    assert replaced.count("21.10.2024") == 1
    assert "03.05.2024 код А на рев. 02 AGCC-BCC-TRM-000339" in replaced

    historical = format_history_line(
        date="09.10.2023",
        stage="code_a",
        revision="0",
        transmittal="AGCC-BCC-TRM-000136",
    )
    patch_old = build_journal_patch(_SOT_F, historical, revision="0", stage="code_a")
    assert patch_old.action == "inserted"
    assert patch_old.update_de is False
    assert patch_old.comment_after.split("\n")[0].startswith("09.10.2023")

    latest = format_history_line(
        date="09.09.2026",
        stage="code_a",
        revision="01-AN02",
        transmittal="PGS-BCC-TRM-000536",
    )
    patch_new = build_journal_patch(
        _SOT_F, latest, revision="01-AN02", stage="code_a"
    )
    assert patch_new.action == "inserted"
    assert patch_new.update_de is True
    assert patch_new.sheet_revision == "Рев. 01-AN02"
    assert patch_new.status_sheet == "РД Согласовано"
    assert patch_new.comment_after.split("\n")[-1] == latest
    events = parse_history_comment(patch_new.comment_after)
    dated = [event for event in events if event.date]
    assert dated[-1].stage == "code_a"
    assert dated[-1].date == "09.09.2026"

    same_again, same_action = apply_history_line(_SOT_F, code_a)
    assert same_action == "unchanged"
    assert same_again == _SOT_F
    already = build_journal_patch(_SOT_F, code_a, revision="04", stage="code_a")
    assert already.action == "unchanged"
    assert already.update_de is True
    assert already.comment_after == _SOT_F
    assert journal_write_needed(
        already, live_d="Рев. 04", live_e="РД Согласовано"
    ) == (False, False, False)
    assert journal_write_needed(
        already, live_d="Рев. 03", live_e="старый"
    ) == (False, True, True)
    assert journal_write_needed(already, live_d="04", live_e="РД Согласовано") == (
        False,
        False,
        False,
    )

    screenshot_f = (
        "09.10.2023 прошла вх контр AGCC.287-BCC-TPG-TRM-003144\n"
        "13.10.2023 код B на рев. 0 NPG-BCC-TRM-003097\n"
        "20.11.2023 код B на рев. 0 NPG-BCC-TRM-000362\n"
        "07.12.2023\n"
        "10.01.2024 код A на рев. 01 AGCC-BCC-TRM-000202"
    )
    filled_a = format_history_line(
        date="07.12.2023",
        stage="code_a",
        revision="0",
        transmittal="AGCC-BCC-TRM-000183",
    )
    upgraded, stub_action = apply_history_line(screenshot_f, filled_a)
    assert stub_action == "replaced"
    assert upgraded.count("07.12.2023") == 1
    assert filled_a in upgraded.split("\n")
    assert "10.01.2024 код A на рев. 01 AGCC-BCC-TRM-000202" in upgraded

    trm_stub = "07.05.2024 AGCC.287-BCC-NPG-TRM-003496"
    incoming_same_trm = format_history_line(
        date="07.05.2024",
        stage="incoming_passed",
        transmittal="AGCC.287-BCC-NPG-TRM-003496",
    )
    stub_upgraded, stub_kind = apply_history_line(trm_stub, incoming_same_trm)
    assert stub_kind == "replaced"
    assert stub_upgraded == incoming_same_trm

    other_trm = format_history_line(
        date="07.05.2024",
        stage="code_a",
        revision="02",
        transmittal="AGCC-BCC-TRM-000339",
    )
    kept_stub, other_action = apply_history_line(trm_stub, other_trm)
    assert other_action == "inserted"
    assert trm_stub in kept_stub.split("\n")
    assert other_trm in kept_stub.split("\n")

    note_line = "07.12.2023 см. письмо Иванова"
    noted, note_action = apply_history_line(note_line, filled_a)
    assert note_action == "inserted"
    assert "см. письмо Иванова" in noted
    assert filled_a in noted.split("\n")

    assert format_sheet_revision_cell("04") == "Рев. 04"
    assert format_sheet_revision_cell("0-AN02") == "Рев. 0-AN02"
    assert status_sheet_for_stage("code_b") == "РД_Корректировка по зам."
    assert status_sheet_for_stage("code_c") == "РД_Корректировка по зам."
    assert status_sheet_for_stage("incoming_passed") == "Прошла входной контроль"
    assert status_sheet_for_stage("sr_upload") == "Отпр. на входной контроль"

    with_mto = format_history_line(
        date="09.09.2026",
        stage="code_a",
        revision="04",
        transmittal="AGCC-BCC-TRM-000999",
        mto_revision="03",
    )
    assert with_mto == (
        "09.09.2026 код А на рев. 04 AGCC-BCC-TRM-000999 MTO 03"
    )
    assert with_mto == format_history_line(
        date="09.09.2026",
        stage="code_a",
        revision="04",
        transmittal="AGCC-BCC-TRM-000999",
        mto_revision="03",
        from_robot_auto=False,
    )
    mto_event = parse_history_line(with_mto)
    assert mto_event.revision == "04"
    assert mto_event.appendix is None
    assert mto_event.mto_revision == "03"
    assert mto_event.mto_appendix is None
    assert mto_event.from_robot_auto is False
    assert mto_event.transmittals == ("AGCC-BCC-TRM-000999",)
    assert "рев. 03" not in with_mto

    with_auto = format_history_line(
        date="09.09.2026",
        stage="code_a",
        revision="04",
        transmittal="AGCC-BCC-TRM-000999",
        mto_revision="03",
        from_robot_auto=True,
    )
    assert with_auto == (
        "09.09.2026 код А на рев. 04 AGCC-BCC-TRM-000999 MTO 03 auto"
    )
    auto_event = parse_history_line(with_auto)
    assert auto_event.revision == "04"
    assert auto_event.mto_revision == "03"
    assert auto_event.from_robot_auto is True
    assert auto_event.transmittals == ("AGCC-BCC-TRM-000999",)
    assert "auto" not in {token.casefold() for token in auto_event.transmittals}

    auto_only = format_history_line(
        date="09.09.2026",
        stage="code_a",
        revision="04",
        transmittal="AGCC-BCC-TRM-000999",
        from_robot_auto=True,
    )
    assert auto_only == (
        "09.09.2026 код А на рев. 04 AGCC-BCC-TRM-000999 auto"
    )
    auto_only_event = parse_history_line(auto_only)
    assert auto_only_event.mto_revision is None
    assert auto_only_event.mto_absent is False
    assert auto_only_event.from_robot_auto is True
    assert auto_only_event.transmittals == ("AGCC-BCC-TRM-000999",)
    cased = parse_history_line(
        "09.09.2026 код А на рев. 04 AGCC-BCC-TRM-000999 mto 03 AUTO"
    )
    assert cased.revision == "04"
    assert cased.mto_revision == "03"
    assert cased.from_robot_auto is True
    assert cased.transmittals == ("AGCC-BCC-TRM-000999",)

    absent = format_history_line(
        date="09.09.2026",
        stage="code_a",
        revision="04",
        transmittal="AGCC-BCC-TRM-000999",
        mto_absent=True,
        from_robot_auto=True,
    )
    assert absent == (
        "09.09.2026 код А на рев. 04 AGCC-BCC-TRM-000999 MTO Нет auto"
    )
    absent_event = parse_history_line(absent)
    assert absent_event.revision == "04"
    assert absent_event.mto_revision is None
    assert absent_event.mto_absent is True
    assert absent_event.from_robot_auto is True
    assert absent_event.transmittals == ("AGCC-BCC-TRM-000999",)
    cased_absent = parse_history_line(
        "09.09.2026 код А на рев. 04 AGCC-BCC-TRM-000999 mto нет AUTO"
    )
    assert cased_absent.mto_absent is True
    assert cased_absent.mto_revision is None
    assert cased_absent.from_robot_auto is True
    rev_wins = format_history_line(
        date="09.09.2026",
        stage="code_a",
        revision="04",
        transmittal="AGCC-BCC-TRM-000999",
        mto_revision="03",
        mto_absent=True,
        from_robot_auto=True,
    )
    assert "MTO 03" in rev_wins
    assert "MTO Нет" not in rev_wins

    an_suffix = format_history_line(
        date="09.09.2026",
        stage="code_a",
        revision="04",
        transmittal="AGCC-BCC-TRM-000999",
        mto_revision="01-AN02",
    )
    an_event = parse_history_line(an_suffix)
    assert an_event.revision == "04"
    assert an_event.mto_revision == "01"
    assert an_event.mto_appendix == "02"

    core_without_mto = format_history_line(
        date="09.09.2026",
        stage="code_a",
        revision="04",
        transmittal="AGCC-BCC-TRM-000999",
    )
    upgraded_mto, mto_action = apply_history_line(core_without_mto, with_mto)
    assert mto_action == "replaced"
    assert upgraded_mto == with_mto
    same_mto, same_mto_action = apply_history_line(with_mto, with_mto)
    assert same_mto_action == "unchanged"
    assert same_mto == with_mto

    stub_with_mto = "07.12.2023 MTO 03"
    stub_filled = format_history_line(
        date="07.12.2023",
        stage="code_a",
        revision="0",
        transmittal="AGCC-BCC-TRM-000183",
    )
    stub_upgraded_mto, stub_mto_kind = apply_history_line(stub_with_mto, stub_filled)
    assert stub_mto_kind == "replaced"
    assert stub_upgraded_mto == stub_filled

    auto_line = catalog_f_line(
        date="17.09.2026",
        stage="code_a",
        revision="01-AN02",
        transmittal="Добавлен_для_легализации_ревизии",
        mto_revision="01",
    )
    assert auto_line.endswith(" auto")
    assert "код А на рев. 01-AN02" in auto_line
    assert "Добавлен_для_легализации_ревизии" in auto_line
    assert "MTO 01" in auto_line
    auto_event = parse_history_line(auto_line)
    assert auto_event.stage == "code_a"
    assert auto_event.from_robot_auto is True
    assert auto_event.mto_revision == "01"

    before = "02.12.2024 код А на рев. 04 AGCC-BCC-TRM-000582"
    after = before + "\n" + auto_line
    assert added_history_line(before, after, fallback=auto_line) == auto_line
    edited = auto_line.replace("17.09.2026", "18.09.2026")
    edited_after = before + "\n" + edited
    assert added_history_line(before, edited_after, fallback=auto_line) == edited
    try:
        added_history_line(before, before, fallback=auto_line)
    except ValueError:
        pass
    else:
        raise AssertionError("missing new F line must raise")

    print("RD catalog F journal: OK")


if __name__ == "__main__":
    main()
