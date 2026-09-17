"""Preview chaining of F patches from parsed approval letters."""

from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from rd_catalog.approval_mail import parse_approval_mail_text
from rd_catalog.f_journal import journal_highlight_spans
from rd_catalog.approval_mail_preview import (
    build_mail_previews,
    comment_lookup_from_kits,
    kit_lookup_from_issuance,
    mail_dedupe_key,
    writable_jobs,
)
from rd_catalog.kits import IssuanceKit, parse_google_matrix

_HEADER = ["титул", "ИД/ТО/МДЗ/РД", "Наименование", "Тип/ Стадия", "Статус", "Комментарий"]
_SOT_F = (
    "03.05.2024 код А на рев. 02 AGCC-BCC-TRM-000339\n"
    "02.12.2024 код А на рев. 04 AGCC-BCC-TRM-000582"
)
_NOTIFICATION = """
Transmittal Date	09.09.2026 12:00:00
1. AGCC.287-1600-SOT.OD-0001 \\ 04 \\ IFC - x \\ x \\ \\ Offsite facilities \\ A - Замечания отсутствуют
2. AGCC.287-1600-SOT.MTO-0001 \\ 03 \\ IFC - x \\ x \\ \\ Offsite facilities \\ B - Незначительные
"""


def _issuance(
    title: str,
    mark: str,
    trm: str,
) -> IssuanceKit:
    return IssuanceKit(
        title=title,
        mark=mark,
        mark_raw=mark,
        title_system=f"{title}-{mark}",
        revision="01",
        appendix=None,
        revision_text="01",
        status="",
        send_date="",
        send_date_sortable="",
        send_transmittal=trm,
        incoming_control_date="",
        incoming_control_date_sortable="",
        confirm_transmittal="",
        note_raw="",
        row_index=2,
    )


def main() -> None:
    """Lookup, missing kit, letter counts, and chained F patches."""

    kits, _stats = parse_google_matrix(
        [
            _HEADER,
            ["1600", "SOT", "x", "Рев. 04", "РД Согласовано", _SOT_F],
        ]
    )
    lookup = comment_lookup_from_kits(kits)
    state = lookup("1600", "SOT")
    assert state is not None
    assert state.comment_raw == _SOT_F
    assert lookup("9999", "SOT") is None

    send_lookup = kit_lookup_from_issuance(
        [
            _issuance("1600", "SOT", "AGCC.287-BCC-PGS-TRM-000910"),
            _issuance("2225", "KSB", "AGCC.287-BCC-PGS-TRM-000910"),
        ]
    )
    assert send_lookup("AGCC.287-BCC-PGS-TRM-000910") == ("2225", "KSB")
    assert send_lookup("missing") is None

    mail = parse_approval_mail_text(
        subject="Notification of Transmittal Number 'AGCC-BCC-TRM-000999'",
        body=_NOTIFICATION,
    )
    assert mail.error == ""
    assert mail.kit_code == "A"
    rows = build_mail_previews([mail], comment_lookup=lookup)
    assert len(rows) == 1
    assert rows[0].writable
    assert rows[0].letter_counts_text == "A:1 · B:1"
    assert rows[0].patch is not None
    assert rows[0].patch.update_de is True
    assert rows[0].comment_before == _SOT_F
    assert "09.09.2026 код А" in rows[0].patch.comment_after
    jobs = writable_jobs(rows)
    assert len(jobs) == 1
    assert jobs[0].title == "1600"
    assert jobs[0].mark == "SOT"
    assert mail_dedupe_key(mail).startswith("f|1600|sot|")
    assert mail_dedupe_key(mail) == mail_dedupe_key(mail)

    missing_mail = parse_approval_mail_text(
        subject="Notification of Transmittal Number 'AGCC-BCC-TRM-000001'",
        body=_NOTIFICATION.replace("1600-SOT", "8888-SOT"),
    )
    missing_rows = build_mail_previews([missing_mail], comment_lookup=lookup)
    assert missing_rows[0].writable is False
    assert "Нет комплекта" in missing_rows[0].error
    assert writable_jobs(missing_rows) == ()

    second = parse_approval_mail_text(
        subject="Notification of Transmittal Number 'AGCC-BCC-TRM-001000'",
        body=_NOTIFICATION.replace("09.09.2026", "10.09.2026").replace(
            "000999", "001000"
        ),
    )
    chained = build_mail_previews([mail, second], comment_lookup=lookup)
    assert chained[1].comment_before == chained[0].patch.comment_after
    assert chained[1].patch is not None
    assert chained[1].patch.comment_after.split("\n")[-1].startswith("10.09.2026")
    left, right = journal_highlight_spans(
        chained[1].comment_before, chained[1].patch.comment_after
    )
    marked = "".join(
        chained[1].patch.comment_after[start:end] for start, end in right
    )
    assert "10.09.2026" in marked
    assert "09.09.2026" not in marked
    assert left == ()

    already_f = rows[0].patch.comment_after
    already_lookup = comment_lookup_from_kits(
        parse_google_matrix(
            [
                _HEADER,
                ["1600", "SOT", "x", "Рев. 04", "РД Согласовано", already_f],
            ]
        )[0]
    )
    already_rows = build_mail_previews([mail], comment_lookup=already_lookup)
    assert already_rows[0].already_in_f
    assert already_rows[0].already_complete
    assert already_rows[0].already_recorded
    assert already_rows[0].writable is False
    assert already_rows[0].patch is not None
    assert already_rows[0].patch.action == "unchanged"
    assert writable_jobs(already_rows) == ()

    stale_de_lookup = comment_lookup_from_kits(
        parse_google_matrix(
            [_HEADER, ["1600", "SOT", "x", "Рев. 03", "ok", already_f]]
        )[0]
    )
    stale_rows = build_mail_previews([mail], comment_lookup=stale_de_lookup)
    assert stale_rows[0].already_in_f
    assert stale_rows[0].already_complete is False
    assert stale_rows[0].already_recorded is False
    assert stale_rows[0].writable
    assert stale_rows[0].write_f is False
    assert stale_rows[0].write_d is True
    assert stale_rows[0].write_e is True
    assert len(writable_jobs(stale_rows)) == 1

    cover = parse_approval_mail_text(
        subject="Сопроводительное письмо AGCC.287-BCC-PGS-TRM-000661 КСБ",
        body=(
            "Please find attached transmittals AGCC.287-BCC-PGS-TRM-000661\n"
            "1\nAGCC.287-4130-KSB2.OD-0001\n0-AN01\nОбщие данные\n"
        ),
        sent_at="24.12.2025",
    )
    assert cover.kind == "cover_letter"
    cover_lookup = comment_lookup_from_kits(
        parse_google_matrix(
            [_HEADER, ["4130", "KSB2", "x", "Рев. 0", "ok", ""]]
        )[0]
    )
    cover_rows = build_mail_previews([cover], comment_lookup=cover_lookup)
    assert cover_rows[0].writable
    assert cover_rows[0].patch is not None
    assert "отпр на ТДО рев. 0-AN01" in cover_rows[0].patch.comment_after
    assert cover_rows[0].write_d is True
    jobs = writable_jobs(cover_rows)
    assert jobs[0].stage == "tdo_sent"
    assert jobs[0].revision == "0-AN01"

    print("approval_mail_preview ok")


if __name__ == "__main__":
    main()
