"""Text dump of an Outlook/Explorer drop for approval-mail debugging.

Qt-free. Does not write Google. Re-reads a saved ``.msg`` with ``extract_msg``
the same way the parser does, so the dump matches what the tab actually saw.
"""

from __future__ import annotations

from pathlib import Path

from rd_catalog.approval_mail import ApprovalMail, decode_outlook_msg_text
from rd_catalog.outlook_drop_qt import DroppedMsg

DUMP_HEADER = "RD Catalog · Дамп письма Outlook"
TEXT_PREVIEW_HEADER = "RD Catalog · Текст письма"
_BODY_LIMIT = 12_000
_HTML_LIMIT = 4_000


def build_outlook_dump(
    dropped: DroppedMsg,
    mail: ApprovalMail | None = None,
) -> str:
    """Build a pasteable dump of one dropped letter.

    Args:
        dropped: Item stored at ingest (MIME names, path, Outlook IDs).
        mail: Parsed preview mail for the same row, if any.

    Returns:
        Multiline UTF-8 text ending with a newline.
    """

    lines = [DUMP_HEADER, ""]
    lines.extend(_drop_lines(dropped))
    source = (mail.source_path if mail is not None else "") or dropped.path
    if source:
        lines.append("")
        lines.extend(_msg_file_lines(source))
    if mail is not None:
        lines.append("")
        lines.extend(_parsed_lines(mail))
    return "\n".join(lines).rstrip() + "\n"


def build_mail_text_preview(
    dropped: DroppedMsg,
    mail: ApprovalMail | None = None,
) -> str:
    """Build a simple subject/body view of one letter.

    Uses the same ``extract_msg`` + CP1251 repair as the parser.

    Args:
        dropped: Item stored at ingest.
        mail: Parsed preview mail for the same row, if any.

    Returns:
        Multiline UTF-8 text ending with a newline.
    """

    lines = [TEXT_PREVIEW_HEADER, ""]
    if mail is not None and mail.error:
        lines.append(f"Ошибка: {mail.error}")
        lines.append("")
    if mail is not None:
        kit = f"{mail.title}-{mail.mark}" if mail.title else "—"
        lines.extend(
            [
                f"Тема: {mail.subject or '—'}",
                f"Дата: {mail.date or '—'}",
                f"Вид: {mail.kind or '—'}",
                f"Титул–марка: {kit}",
                f"Стадия: {mail.stage or '—'}",
                f"Строка F: {mail.f_line or '—'}",
            ]
        )
    else:
        lines.append(f"Имя: {dropped.name or '—'}")
        if dropped.error:
            lines.append(f"Ошибка дропа: {dropped.error}")
    source = (mail.source_path if mail is not None else "") or dropped.path
    body = _decoded_letter_body(source)
    lines.append("")
    lines.append("— текст —")
    lines.append(body or "Текст письма недоступен.")
    return "\n".join(lines).rstrip() + "\n"


def _decoded_letter_body(path: str) -> str:
    file_path = Path(path) if path else None
    if file_path is None or not file_path.is_file():
        return ""
    try:
        import extract_msg
    except ImportError:
        return ""
    try:
        message = extract_msg.openMsg(str(file_path))
    except Exception:
        return ""
    try:
        _subject, body = decode_outlook_msg_text(
            subject=str(message.subject or ""),
            body=str(message.body or ""),
            html=message.htmlBody,
        )
    finally:
        message.close()
    return _clip(body, _BODY_LIMIT)


def _drop_lines(dropped: DroppedMsg) -> list[str]:
    source = "Explorer .msg"
    if dropped.payload:
        source = "Outlook FileContents"
    elif dropped.outlook_entry_id:
        source = "Outlook RenPrivateMessages → SaveAs"
    elif dropped.error:
        source = "ошибка дропа"
    lines = [
        f"Источник дропа: {source}",
        f"Имя: {dropped.name or '—'}",
        f"Путь дропа: {dropped.path or '—'}",
        f"FileContents байт: {len(dropped.payload)}",
        f"Outlook EntryID: {dropped.outlook_entry_id or '—'}",
        f"Outlook StoreID: {dropped.outlook_store_id or '—'}",
        f"Ошибка дропа: {dropped.error or '—'}",
    ]
    if dropped.payload:
        lines.append(f"FileContents hex: {_hex_head(dropped.payload)}")
    if dropped.debug_note:
        lines.append(f"OLE/разбор: {dropped.debug_note}")
    if dropped.mime_formats:
        size_by_fmt = {name: size for name, size in dropped.mime_sizes}
        lines.append("MIME-форматы:")
        for fmt in dropped.mime_formats:
            if fmt in size_by_fmt:
                lines.append(f"  {fmt} · {size_by_fmt[fmt]} байт")
            else:
                lines.append(f"  {fmt}")
    else:
        lines.append("MIME-форматы: —")
    return lines


def _msg_file_lines(path: str) -> list[str]:
    file_path = Path(path)
    lines = [f"Сохранённый .msg: {file_path}"]
    if not file_path.is_file():
        lines.append("Файл .msg отсутствует.")
        return lines
    lines.append(f"Размер .msg: {file_path.stat().st_size} байт")
    try:
        import extract_msg
    except ImportError:
        lines.append("extract_msg не установлен.")
        return lines
    try:
        message = extract_msg.openMsg(str(file_path))
    except Exception as exc:
        lines.append(f"extract_msg: {type(exc).__name__}: {exc}")
        return lines
    try:
        subject = str(message.subject or "")
        body = str(message.body or "")
        sent_at = message.date
        html = message.htmlBody
    finally:
        message.close()
    html_text, html_info = _decode_html(html)
    recoded = _recode_cp1251(body)
    lines.append(f"extract_msg date: {sent_at!r}")
    lines.append(f"extract_msg subject: {subject}")
    lines.append(
        "Кириллица в subject: "
        f"{_yes_no(_has_cyrillic(subject))}"
    )
    lines.append(f"extract_msg body байт/символов: {len(body)}")
    lines.append(f"Кириллица в body: {_yes_no(_has_cyrillic(body))}")
    if recoded and recoded != body:
        lines.append(
            "Кириллица после cp1252→cp1251: "
            f"{_yes_no(_has_cyrillic(recoded))}"
        )
    lines.append(html_info)
    lines.append("— body —")
    lines.append(_clip(body, _BODY_LIMIT))
    if recoded and recoded != body and _has_cyrillic(recoded):
        lines.append("— body cp1252→cp1251 —")
        lines.append(_clip(recoded, _BODY_LIMIT))
    if html_text:
        lines.append("— html utf-8 (фрагмент) —")
        lines.append(_clip(html_text, _HTML_LIMIT))
    return lines


def _parsed_lines(mail: ApprovalMail) -> list[str]:
    lines = [
        "— разбор парсера —",
        f"kind: {mail.kind or '—'}",
        f"title: {mail.title or '—'}",
        f"mark: {mail.mark or '—'}",
        f"date: {mail.date or '—'}",
        f"stage: {mail.stage or '—'}",
        f"kit_code: {mail.kit_code or '—'}",
        f"od_revision: {mail.od_revision or '—'}",
        f"send_trm: {mail.send_transmittal or '—'}",
        f"incoming_trm: {mail.incoming_transmittal or '—'}",
        f"letter_counts: {mail.letter_counts or '—'}",
        f"f_line: {mail.f_line or '—'}",
        f"error: {mail.error or '—'}",
        f"documents: {len(mail.documents)}",
    ]
    for item in mail.documents:
        lines.append(
            f"  {item.filename} rev={item.revision} code={item.code} "
            f"od={item.is_od} {item.title}-{item.mark}"
        )
    return lines


def _decode_html(raw: object) -> tuple[str, str]:
    if raw is None:
        return "", "HTML: нет"
    if isinstance(raw, (bytes, bytearray)):
        data = bytes(raw)
        try:
            text = data.decode("utf-8")
            enc = "utf-8"
        except UnicodeDecodeError:
            text = data.decode("cp1251", errors="replace")
            enc = "cp1251/replace"
        return text, (
            f"HTML: {len(data)} байт, {enc}, "
            f"кириллица={_yes_no(_has_cyrillic(text))}"
        )
    text = str(raw)
    return text, (
        f"HTML: строка {len(text)} символов, "
        f"кириллица={_yes_no(_has_cyrillic(text))}"
    )


def _recode_cp1251(text: str) -> str:
    for encoding in ("cp1252", "latin-1"):
        try:
            return text.encode(encoding).decode("cp1251")
        except UnicodeError:
            continue
    return ""


def _has_cyrillic(text: str) -> bool:
    return any("\u0400" <= char <= "\u04ff" for char in text)


def _yes_no(value: bool) -> str:
    return "да" if value else "нет"


def _clip(text: str, limit: int) -> str:
    if len(text) <= limit:
        return text
    return text[:limit] + f"\n… обрезано, всего {len(text)} символов"


def _hex_head(payload: bytes, size: int = 32) -> str:
    return payload[:size].hex()
