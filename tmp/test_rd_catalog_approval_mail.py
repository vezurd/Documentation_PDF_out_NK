"""Parse approval-mail samples (text fixtures + real .msg when present)."""

from __future__ import annotations

import sys
from datetime import datetime
from pathlib import Path
from zoneinfo import ZoneInfo

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from rd_catalog.approval_mail import (
    _is_mto_filename,
    decode_outlook_msg_text,
    kit_hint_from_path,
    parse_approval_mail_text,
    parse_msg_file,
)
from rd_catalog.kits import parse_history_line

_ROOT = Path(__file__).resolve().parents[1]
_SAMPLES = _ROOT / "approval_mail_samples"
_TZ = ZoneInfo("Europe/Moscow")

_NOTIFICATION_BODY = """
Уведомление о рассылке \\ Transmittal Notification
Трансмиттал\\
Transmittal	AGCC-BCC-TRM-000582
Дата трансмиттала\\
Transmittal Date	02.12.2024 16:11:16
Документы\\
Documents	1. AGCC.287-1600-SOT.MTO-0001 \\ 03 \\ IFC - Выпущено для строительства \\ Спецификация \\ \\ Offsite facilities \\ A - Замечания отсутствуют
2. AGCC.287-1600-SOT.OD-0001 \\ 04 \\ IFC - Выпущено для строительства \\ Общие данные \\ \\ Offsite facilities \\ A - Замечания отсутствуют
Создано системой Capital Projects
"""

_IFR_LETTER_BODY = """
Уведомление о рассылке \\ Transmittal Notification
Трансмиттал\\
Transmittal   AGCC-BCC-TRM-000158
Дата трансмиттала\\
Transmittal Date  10.11.2023 14:56:37
Тема\\Subject  Передача результатов рассмотрения ПО АГХК
Причина создания\\
Reason for Issue  IFR - Issued For Review\\Выпущен для рассмотрения
Документы\\
Documents  1. AGCC.287-6510-SOS.CAE-0002 \\ A \\ IFR - Выпущено для рассмотрения \\ Схема структурная \\  \\ 6000 - 6999 \\ A - Замечания отсутствуют
2. AGCC.287-6510-SOS.MTO-0001 \\ A \\ IFR - Выпущено для рассмотрения \\ Спецификация \\  \\ 6000 - 6999 \\ A - Замечания отсутствуют
7. AGCC.287-6510-SOS.OD-0001 \\ A \\ IFR - Выпущено для рассмотрения \\ Общие данные \\  \\ 6000 - 6999 \\ A - Замечания отсутствуют
Создано системой Capital Projects
"""

_MIXED_BODY = """
Transmittal Date	03.05.2024 7:43:52
1. AGCC.287-1600-SOT.LAY-0005 \\ 02 \\ IFC - Выпущено \\ x \\ \\ Offsite facilities \\ B - Незначительные замечания.
2. AGCC.287-1600-SOT.OD-0001 \\ 02 \\ IFC - Выпущено \\ x \\ \\ Offsite facilities \\ A - Замечания отсутствуют
3. AGCC.287-1600-SOT.WIR-0003 \\ 02 \\ IFC - Выпущено \\ x \\ \\ Offsite facilities \\ A - Замечания отсутствуют
"""

_TDO_LOADED_INCOMING = """
Добрый день!
Документация загружена AGCC.287-PGS-PGS-TRM-22028
С уважением,
Назаренко

From: Dolganina Elizaveta <EDolganina@bcc.ru>
Sent: Monday, August 10, 2026
Subject: RE: Сопроводительное письмо AGCC.287-BCC-PGS-TRM-000910
"""

_TDO_SR_ONLY = """
Коллеги, добрый день.
Документация загружена
С уважением,
Валеева

From: Dolganina Elizaveta <EDolganina@bcc.ru>
"""

_TDO_TWO_KITS = """
Документация загружена AGCC.287-PGS-PGS-TRM-22028
AGCC.287-3150-KSB.OD-0001_01_RU.pdf
AGCC.287-5110-KSB.OD-0001_01_RU.pdf

From: x
"""

_COVER_BODY = """
Dear All,

Please find attached transmittals AGCC.287-BCC-PGS-TRM-000661

Добрый день!

Документация по комплекту AGCC.287-BCC-PGS-TRM-000661

№ п/п
Номер основного документа
Owner Document Number
Ревизия
Revision
Наименование
Document Name
1
AGCC.287-4130-KSB2.OD-0001
0-AN01
Общие данные
2
AGCC.287-4130-KSB2.MTO-0001
0-AN01
Спецификация оборудования
3
AGCC.287-4130-KSB2.BOE-0001
0-AN01
Ведомость оборудования
4
AGCC.287-4130-KSB2.BOM-0001
0-AN01
Ведомость материалов
5
AGCC.287-4130-KSB2.BOQ-0001
0-AN01
Ведомость объёмов работ
"""


def main() -> None:
    """Run parser assertions; .msg smoke runs when samples exist."""

    mail = parse_approval_mail_text(
        subject="Notification of Transmittal Number 'AGCC-BCC-TRM-000582' / уведомление",
        body=_NOTIFICATION_BODY,
    )
    assert mail.error == ""
    assert mail.kind == "review_codes"
    assert mail.title == "1600"
    assert mail.mark == "SOT"
    assert mail.kit_code == "A"
    assert mail.od_revision == "04"
    assert mail.date == "02.12.2024"
    assert mail.send_transmittal == "AGCC-BCC-TRM-000582"
    assert mail.f_line == (
        "02.12.2024 код А на рев. 04 AGCC-BCC-TRM-000582 MTO 03"
    )
    assert mail.sheet_revision == "Рев. 04"
    assert mail.status_sheet == "РД Согласовано"
    mail_event = parse_history_line(mail.f_line)
    assert mail_event.stage == "code_a"
    assert mail_event.revision == "04"
    assert mail_event.mto_revision == "03"
    assert mail_event.from_robot_auto is False
    assert mail_event.transmittals == ("AGCC-BCC-TRM-000582",)
    assert len(mail.documents) == 2
    assert _is_mto_filename("AGCC.287-1600-SOT.MTO-0001")
    assert not _is_mto_filename("AGCC.287-1600-SOT.OD-0001")

    ifr = parse_approval_mail_text(
        subject=(
            "Notification of Transmittal Number 'AGCC-BCC-TRM-000158' / "
            "уведомление о рассылке номер 'AGCC-BCC-TRM-000158'"
        ),
        body=_IFR_LETTER_BODY,
    )
    assert ifr.error == ""
    assert ifr.kind == "review_codes"
    assert ifr.title == "6510"
    assert ifr.mark == "SOS"
    assert ifr.kit_code == "A"
    assert ifr.od_revision == "A"
    assert ifr.date == "10.11.2023"
    assert ifr.send_transmittal == "AGCC-BCC-TRM-000158"
    assert ifr.f_line == "10.11.2023 код А на рев. A AGCC-BCC-TRM-000158 MTO A"
    assert ifr.sheet_revision == "Рев. A"
    ifr_event = parse_history_line(ifr.f_line)
    assert ifr_event.revision == "A"
    assert ifr_event.mto_revision == "A"
    assert ifr_event.from_robot_auto is False
    assert len(ifr.documents) == 3
    assert all(doc.revision == "A" for doc in ifr.documents)

    wrapped = parse_approval_mail_text(
        subject="Notification of Transmittal Number 'PGS-BCC-TRM-000554'",
        body=(
            "Transmittal Date\t10.09.2026 17:20:23\n"
            "Documents\t1. AGCC.287-3140-KSB2.BOE-0001 \\ 01-AN02 \\ IFC -\n"
            "x \\ x \\ A -\n"
            "Замечания отсутствуют\n"
            "2. AGCC.287-3140-KSB2.OD-0001 \\ 01-AN02 \\ IFC - x \\ A -\n"
            "Замечания отсутствуют \\ \n"
        ),
    )
    assert wrapped.error == ""
    assert wrapped.title == "3140"
    assert wrapped.mark == "KSB2"
    assert wrapped.kit_code == "A"
    assert wrapped.letter_counts == (("A", 2),)
    assert any(doc.filename.endswith("BOE-0001") for doc in wrapped.documents)

    mojibake_body = (
        "Transmittal Notification\nTransmittal Date\t10.09.2026 17:20:23\n"
        "Documents\t1. AGCC.287-3140-KSB2.OD-0001 \\ 01-AN02 \\ IFC \\ x \\ A - "
        "Замечания отсутствуют\n"
    ).encode("cp1251").decode("cp1252")
    mojibake_html = (
        "<html><body>1. AGCC.287-3140-KSB2.OD-0001 \\ 01-AN02 \\ IFC \\ "
        "x \\ A - Замечания отсутствуют<br></body></html>"
    ).encode("utf-8")
    subject_u, body_u = decode_outlook_msg_text(
        subject="Notification of Transmittal Number 'PGS-BCC-TRM-000554' / "
        + "уведомление".encode("cp1251").decode("cp1252"),
        body=mojibake_body,
        html=mojibake_html,
    )
    assert "уведомление" in subject_u
    assert "Замечания отсутствуют" in body_u
    restored = parse_approval_mail_text(
        subject=subject_u,
        body=body_u,
    )
    assert restored.kit_code == "A"
    assert restored.mark == "KSB2"

    mixed = parse_approval_mail_text(
        subject="Notification of Transmittal Number 'AGCC-BCC-TRM-000339'",
        body=_MIXED_BODY,
    )
    assert mixed.error == ""
    assert mixed.kit_code == "A"
    assert mixed.letter_counts == (("A", 2), ("B", 1))
    assert mixed.od_revision == "02"
    assert mixed.f_line == "03.05.2024 код А на рев. 02 AGCC-BCC-TRM-000339"
    assert parse_history_line(mixed.f_line).mto_revision is None

    tdo = parse_approval_mail_text(
        subject="RE: Сопроводительное письмо AGCC.287-BCC-PGS-TRM-000910",
        body=_TDO_LOADED_INCOMING,
        sent_at=datetime(2026, 8, 10, 16, 36, tzinfo=_TZ),
        kit_from_transmittal=lambda trm: ("3150", "KSB")
        if "000910" in trm
        else None,
    )
    assert tdo.error == ""
    assert tdo.kind == "tdo_reply"
    assert tdo.stage == "incoming_passed"
    assert tdo.title == "3150"
    assert tdo.mark == "KSB"
    assert tdo.incoming_transmittal == "AGCC.287-PGS-PGS-TRM-22028"
    assert tdo.f_line == "10.08.2026 прошла вх контр AGCC.287-PGS-PGS-TRM-22028"
    assert tdo.status_sheet == "Прошла входной контроль"

    sr_only = parse_approval_mail_text(
        subject="RE: AGCC.287-BCC-PGS-TRM-000918",
        body=_TDO_SR_ONLY,
        sent_at=datetime(2026, 8, 12, 8, 28, tzinfo=_TZ),
        kit_from_transmittal=lambda trm: ("1761", "KSB") if "000918" in trm else None,
    )
    assert sr_only.stage == "sr_upload"
    assert sr_only.incoming_transmittal == ""
    assert "загрузку" in sr_only.f_line
    assert sr_only.status_sheet == "Отпр. на входной контроль"

    two = parse_approval_mail_text(
        subject="RE: Сопроводительное письмо AGCC.287-BCC-PGS-TRM-000910",
        body=_TDO_TWO_KITS,
        sent_at=datetime(2026, 8, 10, tzinfo=_TZ),
    )
    assert two.error.startswith("В письме несколько титул–марок")
    assert two.f_line == ""

    unresolved = parse_approval_mail_text(
        subject="RE: Сопроводительное письмо AGCC.287-BCC-NPG-TRM-003496",
        body="Документы загружены и направлены на рассмотрение\nAGCC.287-PGS-PGS-TRM-17769\n\nFrom: x\n",
        sent_at=datetime(2024, 5, 8, tzinfo=_TZ),
        kit_from_transmittal=lambda _trm: None,
    )
    assert unresolved.stage == "incoming_passed"
    assert "Нет комплекта в Выдаче" in unresolved.error

    quoted_only = parse_approval_mail_text(
        subject="RE: Сопроводительное письмо AGCC.287-BCC-NPG-TRM-003121",
        body=(
            "Документация принята и направлена на рассмотрение "
            "AGCC.287-NPG-NPG-TRM-20298\n\nFrom: x\n"
            "В документе AGCC.287-1600-SOT.VO-0001 некорректная дата\n"
        ),
        sent_at=datetime(2023, 9, 28, tzinfo=_TZ),
    )
    assert quoted_only.error == ""
    assert quoted_only.title == "1600"
    assert quoted_only.mark == "SOT"
    assert quoted_only.stage == "incoming_passed"
    assert quoted_only.f_line.startswith("28.09.2023 прошла вх контр")

    old_prose = parse_approval_mail_text(
        subject="RE: Сопроводительное письмо AGCC.287-BCC-NPG-TRM-003144",
        body=(
            "Документация принята и направлена на рассмотрение.\n\n"
            "From: x\n"
            "Направляю на рассмотрение откорректированную РД по титулу 6816 марка\n"
            "KSB рев.0.\n"
        ),
        sent_at=datetime(2023, 10, 9, 9, 17, tzinfo=_TZ),
    )
    assert old_prose.error == ""
    assert old_prose.kind == "tdo_reply"
    assert old_prose.stage == "incoming_passed"
    assert old_prose.title == "6816"
    assert old_prose.mark == "KSB"
    assert old_prose.od_revision == "0"
    assert old_prose.incoming_transmittal == ""
    assert old_prose.f_line == (
        "09.10.2023 прошла вх контр рев. 0 AGCC.287-BCC-NPG-TRM-003144"
    )
    assert parse_history_line(old_prose.f_line).revision == "0"

    old_stems = parse_approval_mail_text(
        subject="RE: Сопроводительное письмо AGCC.287-BCC-NPG-TRM-003144",
        body=(
            "Документация принята и направлена на рассмотрение.\n\n"
            "From: x\n"
            "AGCC.287-6816-KSB.CAE-0002_0_RU\n"
            "AGCC.287-6816-KSB.OD-0001_0_RU\n"
        ),
        sent_at=datetime(2023, 10, 9, tzinfo=_TZ),
    )
    assert old_stems.error == ""
    assert old_stems.title == "6816"
    assert old_stems.mark == "KSB"
    assert old_stems.od_revision == "0"
    assert any(doc.is_od and doc.revision == "0" for doc in old_stems.documents)
    assert old_stems.f_line == (
        "09.10.2023 прошла вх контр рев. 0 AGCC.287-BCC-NPG-TRM-003144"
    )

    hinted = parse_approval_mail_text(
        subject="RE: Сопроводительное письмо AGCC.287-BCC-NPG-TRM-003496",
        body="Документы загружены и направлены на рассмотрение\nAGCC.287-PGS-PGS-TRM-17769\n\nFrom: x\n",
        sent_at=datetime(2024, 5, 8, tzinfo=_TZ),
        kit_from_transmittal=lambda _trm: None,
        kit_hint=("1600", "SOT"),
    )
    assert hinted.error == ""
    assert hinted.title == "1600"
    assert hinted.mark == "SOT"
    assert hinted.f_line.startswith("08.05.2024 прошла вх контр")

    no_lookup = parse_approval_mail_text(
        subject="RE: Сопроводительное письмо AGCC.287-BCC-NPG-TRM-003496",
        body="Документы загружены и направлены на рассмотрение\nAGCC.287-PGS-PGS-TRM-17769\n\nFrom: x\n",
        sent_at=datetime(2024, 5, 8, tzinfo=_TZ),
    )
    assert "Загрузите комплекты Google" in no_lookup.error
    assert no_lookup.f_line == ""

    note = parse_approval_mail_text(
        subject="RE: Уточнения по AGCC.287-1600-SOT",
        body=(
            "Крайняя загруженная ревизия комплекта AGCC.287-1600-SOT — 03. "
            "Результаты проверки направлены в ваш адрес AGCC-BCC-TRM-000357.\n"
        ),
        sent_at=datetime(2024, 10, 8, tzinfo=_TZ),
    )
    assert note.title == "1600"
    assert note.mark == "SOT"
    assert note.stage == ""
    assert note.f_line == ""

    cover_plain = parse_approval_mail_text(
        subject="Сопроводительное письмо AGCC.287-BCC-PGS-TRM-000661 КСБ",
        body=_COVER_BODY,
        sent_at=datetime(2025, 12, 24, 15, 57, tzinfo=_TZ),
    )
    assert cover_plain.error == ""
    assert cover_plain.kind == "cover_letter"
    assert cover_plain.stage == "tdo_sent"
    assert cover_plain.title == "4130"
    assert cover_plain.mark == "KSB2"
    assert cover_plain.od_revision == "0-AN01"
    assert cover_plain.send_transmittal == "AGCC.287-BCC-PGS-TRM-000661"
    assert cover_plain.f_line == (
        "24.12.2025 отпр на ТДО рев. 0-AN01 AGCC.287-BCC-PGS-TRM-000661 MTO 0-AN01"
    )
    cover_event = parse_history_line(cover_plain.f_line)
    assert cover_event.revision == "0"
    assert cover_event.appendix == "01"
    assert cover_event.mto_revision == "0"
    assert cover_event.mto_appendix == "01"
    assert cover_event.from_robot_auto is False
    assert cover_plain.sheet_revision == "Рев. 0-AN01"
    assert cover_plain.status_sheet == "Отпр. на входной контроль"
    assert parse_history_line(cover_plain.f_line).stage == "tdo_sent"
    assert len(cover_plain.documents) == 5
    od_docs = [doc for doc in cover_plain.documents if doc.is_od]
    assert len(od_docs) == 1
    assert od_docs[0].filename.endswith("OD-0001")
    assert od_docs[0].revision == "0-AN01"

    cover_html_row = parse_approval_mail_text(
        subject="Сопроводительное письмо AGCC.287-BCC-PGS-TRM-000661",
        body=(
            "Please find attached transmittals AGCC.287-BCC-PGS-TRM-000661\n"
            "Owner Document Number Revision Document Name\n"
            "1 AGCC.287-4130-KSB2.OD-0001 0-AN01 Общие данные\n"
            "2 AGCC.287-4130-KSB2.MTO-0001 0-AN01 Спецификация оборудования\n"
        ),
        sent_at=datetime(2025, 12, 24, tzinfo=_TZ),
    )
    assert cover_html_row.kind == "cover_letter"
    assert cover_html_row.title == "4130"
    assert cover_html_row.mark == "KSB2"
    assert cover_html_row.od_revision == "0-AN01"
    assert len(cover_html_row.documents) == 2
    assert "MTO 0-AN01" in cover_html_row.f_line
    assert parse_history_line(cover_html_row.f_line).mto_revision == "0"

    cover_not_reply = parse_approval_mail_text(
        subject="RE: Сопроводительное письмо AGCC.287-BCC-PGS-TRM-000661 КСБ",
        body=_TDO_LOADED_INCOMING,
        sent_at=datetime(2026, 8, 10, 16, 36, tzinfo=_TZ),
        kit_from_transmittal=lambda trm: ("3150", "KSB")
        if "000910" in trm or "000661" in trm
        else None,
    )
    assert cover_not_reply.kind == "tdo_reply"
    assert cover_not_reply.stage == "incoming_passed"

    assert kit_hint_from_path(Path("x") / "17_Даты TRM" / "1600-SOT" / "a.msg") == (
        "1600",
        "SOT",
    )
    assert kit_hint_from_path(
        Path("x") / "1600-SOT" / "_вспомогательные письма" / "a.msg"
    ) == ("1600", "SOT")
    assert kit_hint_from_path(Path("x") / "a.msg") is None

    if _SAMPLES.is_dir():
        code_a = parse_msg_file(
            _SAMPLES / "2024.12.02_Код А на рев. 04 (МТО_рев.03).msg"
        )
        assert code_a.error == ""
        assert code_a.title == "1600"
        assert code_a.mark == "SOT"
        assert code_a.kit_code == "A"
        assert code_a.od_revision == "04"
        assert code_a.send_transmittal == "AGCC-BCC-TRM-000582"
        assert code_a.f_line.startswith("02.12.2024 код А на рев. 04")
        assert "MTO 03" in code_a.f_line

        code_b = parse_msg_file(
            _SAMPLES / "2024.10.21_Код Б на рев. 04 (МТО_рев.03).msg"
        )
        assert code_b.kit_code == "B"
        assert code_b.od_revision == "04"
        assert code_b.send_transmittal == "PGS-BCC-TRM-000009"

        mixed_msg = parse_msg_file(
            _SAMPLES / "2024.05.03_Кода А рев. 02 (МТО_рев.02).msg"
        )
        assert mixed_msg.kit_code == "A"
        counts = dict(mixed_msg.letter_counts)
        assert counts.get("B", 0) >= 1
        assert counts.get("A", 0) >= 1

        tdo_folder = _SAMPLES / "это на отправку и подтверждение прохождения ТДО"
        tdo_msg = parse_msg_file(
            tdo_folder / "RE  Сопроводительное письмо AGCC 287-BCC-PGS-TRM-000910.msg",
            kit_from_transmittal=lambda trm: ("3150", "KSB") if "000910" in trm else None,
        )
        assert tdo_msg.kind == "tdo_reply"
        assert tdo_msg.stage == "incoming_passed"
        assert tdo_msg.incoming_transmittal.endswith("TRM-22028")
        assert tdo_msg.title == "3150"

    dropped = (
        Path.home()
        / "AppData/Local/Documentation_PDF_out_NK/rd_catalog/approval_mail_drop"
    )
    if dropped.is_dir():
        sample = next(dropped.glob("*PGS-BCC-TRM-000554*.msg"), None)
        if sample is not None:
            live = parse_msg_file(sample)
            assert live.error == ""
            assert live.title == "3140"
            assert live.mark == "KSB2"
            assert live.kit_code == "A"
            assert live.letter_counts[0][0] == "A"
            assert live.f_line.startswith("10.09.2026 код А")

    cover_msg = next(
        Path(r"C:\Users\ydruzev\AppData\Local\Temp").glob("*TRM-000661*.msg"),
        None,
    )
    if cover_msg is not None:
        live_cover = parse_msg_file(cover_msg)
        assert live_cover.error == ""
        assert live_cover.kind == "cover_letter"
        assert live_cover.title == "4130"
        assert live_cover.mark == "KSB2"
        assert live_cover.od_revision == "0-AN01"
        assert live_cover.stage == "tdo_sent"
        assert live_cover.send_transmittal.endswith("TRM-000661")
        assert live_cover.f_line.startswith("24.12.2025 отпр на ТДО рев.")
        assert any(doc.is_od and doc.revision == "0-AN01" for doc in live_cover.documents)

    ifr_msg = next(
        Path(r"C:\Users\ydruzev\AppData\Local\Temp").glob("*TRM-000158*.msg"),
        None,
    )
    if ifr_msg is not None:
        live_ifr = parse_msg_file(ifr_msg)
        assert live_ifr.error == ""
        assert live_ifr.kind == "review_codes"
        assert live_ifr.title == "6510"
        assert live_ifr.mark == "SOS"
        assert live_ifr.kit_code == "A"
        assert live_ifr.od_revision == "A"
        assert live_ifr.f_line.startswith("10.11.2023 код А на рев. A")
        assert len(live_ifr.documents) == 8

    old_tdo_paths = list(
        Path(r"C:\Users\ydruzev\AppData\Local\Temp").glob("*NPG-TRM-003144*.msg")
    )
    drop_dir = (
        Path.home()
        / "AppData/Local/Documentation_PDF_out_NK/rd_catalog/approval_mail_drop"
    )
    if drop_dir.is_dir():
        old_tdo_paths.extend(drop_dir.glob("*NPG-TRM-003144*.msg"))
    if old_tdo_paths:
        live_old_tdo = parse_msg_file(old_tdo_paths[0])
        assert live_old_tdo.error == ""
        assert live_old_tdo.kind == "tdo_reply"
        assert live_old_tdo.stage == "incoming_passed"
        assert live_old_tdo.title == "6816"
        assert live_old_tdo.mark == "KSB"
        assert live_old_tdo.od_revision == "0"
        assert live_old_tdo.send_transmittal.endswith("TRM-003144")
        assert live_old_tdo.f_line == (
            "09.10.2023 прошла вх контр рев. 0 AGCC.287-BCC-NPG-TRM-003144 MTO 0"
        )
        assert parse_history_line(live_old_tdo.f_line).mto_revision == "0"
        assert parse_history_line(live_old_tdo.f_line).from_robot_auto is False
        assert len(live_old_tdo.documents) == 6

    print("RD catalog approval mail: OK")


if __name__ == "__main__":
    main()
