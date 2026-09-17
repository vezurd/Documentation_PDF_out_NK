"""Parse AGCC review-code, cover, and TDO confirmation letters.

Qt-free. Understands Capital Projects «Notification of Transmittal» bodies
(IFC numeric / AN revisions and older IFR letter revisions such as ``\\ A \\``),
outgoing BCC «Сопроводительное письмо» document tables, and the top of
``RE: Сопроводительное письмо …`` replies (including older threads whose
kit/revision sit in prose such as ``по титулу 6816 марка KSB рев.0`` or
in AGCC stems with a ``_0_RU`` tail and no extension). File IO uses
``extract_msg`` when opening ``.msg``; tests may call
:func:`parse_approval_mail_text` with already extracted subject/body.
"""

from __future__ import annotations

import re
from collections import Counter
from collections.abc import Callable
from dataclasses import dataclass, replace
from datetime import datetime
from html import unescape
from pathlib import Path

from rd_catalog.f_journal import format_history_line, format_sheet_revision_cell, status_sheet_for_stage
from rd_catalog.kits import is_rd_kit_mark, kit_identity_key
from rd_catalog.parse import normalize_unicode_dashes
from utils.file_name_converts import AgccFilenamePatterns

KitLookup = Callable[[str], tuple[str, str] | None]

_TRM_RE = re.compile(r"[A-Z0-9]+(?:[.-][A-Z0-9]+)*-TRM-\d+", re.IGNORECASE)
_NOTIFICATION_TRM_RE = re.compile(
    r"Notification of Transmittal Number\s+'([^']+)'",
    re.IGNORECASE,
)
_TRANSMITTAL_DATE_RE = re.compile(
    r"Transmittal Date\s+(\d{1,2}\.\d{1,2}\.\d{4})",
    re.IGNORECASE,
)
# Second field after the filename: IFC ``04`` / ``01-AN02`` / ``V`` / ``S``,
# or the older IFR letter ``A`` / ``B`` / ``C``.
_DOC_START_RE = re.compile(
    r"(?:^|[\n\t ])\s*\d+\.\s+"
    r"(AGCC\.287-\d{4}-[A-Za-z0-9.]+(?:\.[A-Za-z0-9.-]+)*)"
    r"\s*\\\s*(\d+(?:-AN\d+)?|[A-Z])\s*\\",
    re.IGNORECASE,
)
_TAIL_CODE_RE = re.compile(
    r"\\\s*([ABC])\s*-\s*"
    r"(Замечания отсутствуют|Незначительные|Существенные|Критические)",
    re.IGNORECASE,
)
_FROM_SPLIT_RE = re.compile(r"\nFrom:\s", re.IGNORECASE)
_STEM_REV_RE = re.compile(
    rf"(?P<stem>{AgccFilenamePatterns.pattern_core()})"
    rf"(?:{AgccFilenamePatterns.SEP_STEM_REVISION_RX}"
    rf"(?P<rev>{AgccFilenamePatterns.REVISION_BODY})"
    rf"(?:{AgccFilenamePatterns.SEP_REVISION_LANGUAGE_RX}"
    rf"{AgccFilenamePatterns.LANGUAGE})?)?",
    re.IGNORECASE,
)
_AGCC_KIT_RE = re.compile(r"AGCC\.287-\d{4}-[A-Za-z0-9]+", re.IGNORECASE)
_KIT_PROSE_RE = re.compile(
    r"титул[ауе]?\s+"
    r"(?:AGCC\.287-)?(\d{4})"
    r"(?:\s*-\s*([A-Za-z0-9.]+))?"
    r"[\s,;]*марка\s+([A-Za-z0-9.]+)"
    r"[\s,;]*рев\.?\s*"
    r"(\d{1,2}(?:-AN\d{1,2})?|[A-Z])(?![A-Za-z0-9А-Яа-я])",
    re.IGNORECASE,
)
_KIT_AGCC_PROSE_RE = re.compile(
    r"титул[ауе]?\s+AGCC\.287-(\d{4})-([A-Za-z0-9.]+)\s*_"
    r"(\d{1,2}(?:-AN\d{1,2})?|[A-Z])(?![A-Za-z0-9А-Яа-я])",
    re.IGNORECASE,
)
_COVER_DOC_RE = re.compile(
    rf"(?:^|[\n\t ])\s*\d+\s+"
    rf"({AgccFilenamePatterns.pattern_core()})"
    r"\s+"
    rf"({AgccFilenamePatterns.REVISION_BODY})"
    r"(?![A-Za-z0-9-])",
    re.IGNORECASE | re.MULTILINE,
)
_REPLY_SUBJECT_RE = re.compile(r"^(?:re|отв(?:ет)?)\s*[:：]", re.IGNORECASE)
_COVER_SUBJECT_RE = re.compile(r"сопроводительн\w*\s+письм", re.IGNORECASE)
_COVER_BODY_HINT_RE = re.compile(
    r"please\s+find\s+attached\s+transmittals|"
    r"owner\s+document\s+number|"
    r"документац\w*.{0,80}комплект|"
    r"направляем\w*.{0,80}документ",
    re.IGNORECASE | re.DOTALL,
)
_CODE_STAGE = {"A": "code_a", "B": "code_b", "C": "code_c"}
_KIT_FOLDER_RE = re.compile(r"^(\d{4})-([A-Za-z0-9]+)$")
_DOC_LOADED_RE = re.compile(
    r"(?:документац\w*|документы)\s+загружен",
    re.IGNORECASE,
)
_DOC_ACCEPTED_RE = re.compile(
    r"(?:документац\w*|документы)\s+принят",
    re.IGNORECASE,
)


@dataclass(frozen=True, slots=True)
class MailDocument:
    """One document row from a transmittal notification."""

    filename: str
    title: str
    mark: str
    revision: str
    code: str | None
    is_od: bool


@dataclass(frozen=True, slots=True)
class ApprovalMail:
    """Structured result of parsing one letter."""

    kind: str
    subject: str
    date: str
    stage: str
    send_transmittal: str
    incoming_transmittal: str
    title: str
    mark: str
    od_revision: str
    kit_code: str
    letter_counts: tuple[tuple[str, int], ...]
    f_line: str
    sheet_revision: str
    status_sheet: str
    error: str
    documents: tuple[MailDocument, ...]
    source_path: str = ""


def parse_msg_file(
    path: str | Path,
    *,
    kit_from_transmittal: KitLookup | None = None,
) -> ApprovalMail:
    """Open a ``.msg`` file and parse subject/body.

    Args:
        path: Saved Outlook message.
        kit_from_transmittal: Optional ``send TRM → (title, mark)`` lookup
            used when the reply has no AGCC filenames.

    Returns:
        Parsed mail, with ``error`` set instead of raising on domain problems.

    Raises:
        FileNotFoundError: If ``path`` does not exist.
        RuntimeError: If ``extract_msg`` is not installed.
    """

    file_path = Path(path)
    if not file_path.is_file():
        raise FileNotFoundError(str(file_path))
    try:
        import extract_msg
    except ImportError as exc:
        raise RuntimeError(
            "Package extract-msg is required to read .msg files."
        ) from exc
    message = extract_msg.openMsg(str(file_path))
    try:
        subject, body = decode_outlook_msg_text(
            subject=str(message.subject or ""),
            body=str(message.body or ""),
            html=message.htmlBody,
        )
        sent_at = message.date
    finally:
        message.close()
    parsed = parse_approval_mail_text(
        subject=subject,
        body=body,
        sent_at=sent_at,
        kit_from_transmittal=kit_from_transmittal,
        kit_hint=kit_hint_from_path(file_path),
    )
    return replace(parsed, source_path=str(file_path))


def kit_hint_from_path(path: str | Path) -> tuple[str, str] | None:
    """Return title+mark from a parent folder named ``1600-SOT``.

    Archive drops under ``17_Даты TRM/<title>-<mark>/`` (and one extra
    folder such as ``_вспомогательные письма``) have no AGCC filenames
    in the TDO reply top. Looks at most three parents.

    Args:
        path: Saved ``.msg`` path.

    Returns:
        ``(title, mark)`` when a parent looks like a kit folder.
    """

    current = Path(path)
    for folder in list(current.parents)[:3]:
        match = _KIT_FOLDER_RE.fullmatch(folder.name)
        if match is None:
            continue
        title, mark = match.group(1), match.group(2)
        if is_rd_kit_mark(mark):
            return title, mark
    return None


def decode_outlook_msg_text(
    *,
    subject: str,
    body: str,
    html: bytes | str | None = None,
) -> tuple[str, str]:
    """Repair Outlook 8-bit ``.msg`` text and prefer Unicode HTML when needed.

    Outlook ``SaveAs`` olMSG often stores CP1251 bytes that ``extract_msg``
    reads as Latin-1/CP1252. The HTML body is usually UTF-8.

    Args:
        subject: ``extract_msg`` subject string.
        body: ``extract_msg`` plain body.
        html: Raw HTML bytes or already-decoded HTML.

    Returns:
        ``(subject, body)`` with Cyrillic restored when possible.
    """

    subject_fixed = _restore_cyrillic(subject)
    body_fixed = _restore_cyrillic(body)
    html_text = _html_to_text(_decode_html(html))
    if html_text and _should_prefer_html(body_fixed, html_text):
        return subject_fixed, html_text
    return subject_fixed, body_fixed


def parse_approval_mail_text(
    *,
    subject: str,
    body: str,
    sent_at: datetime | str | None = None,
    kit_from_transmittal: KitLookup | None = None,
    kit_hint: tuple[str, str] | None = None,
) -> ApprovalMail:
    """Parse already extracted mail text.

    Args:
        subject: Message subject.
        body: Plain-text body (Outlook ``body``, not HTML).
        sent_at: Message datetime; used when Transmittal Date is absent.
        kit_from_transmittal: Optional issuance lookup by send TRM.
        kit_hint: Optional ``(title, mark)`` from the parent folder name.

    Returns:
        Parsed mail. Domain failures set ``error`` and leave ``f_line`` empty.
    """

    subject_norm = normalize_unicode_dashes(subject or "").strip()
    body_norm = normalize_unicode_dashes(body or "").replace("\r\n", "\n")
    notification_trm = _notification_trm(subject_norm)
    if notification_trm or "transmittal notification" in body_norm.casefold():
        return _parse_notification(
            subject=subject_norm,
            body=body_norm,
            sent_at=sent_at,
            fallback_trm=notification_trm,
        )
    if _is_cover_letter(subject_norm, body_norm):
        return _parse_cover_letter(
            subject=subject_norm,
            body=body_norm,
            sent_at=sent_at,
            kit_from_transmittal=kit_from_transmittal,
            kit_hint=kit_hint,
        )
    return _parse_tdo_reply(
        subject=subject_norm,
        body=body_norm,
        sent_at=sent_at,
        kit_from_transmittal=kit_from_transmittal,
        kit_hint=kit_hint,
    )


def _parse_notification(
    *,
    subject: str,
    body: str,
    sent_at: datetime | str | None,
    fallback_trm: str,
) -> ApprovalMail:
    documents = _parse_notification_documents(body)
    date = _transmittal_date(body) or _format_sent_date(sent_at)
    trm = fallback_trm or _first_trm(subject) or _first_trm(body)
    kits = _unique_kits(documents)
    error = ""
    title = ""
    mark = ""
    if not documents:
        error = "В уведомлении нет строк документов."
    elif len(kits) > 1:
        error = (
            "В письме несколько титул–марок: "
            + ", ".join(f"{item[0]}-{item[1]}" for item in kits)
        )
    elif len(kits) == 1:
        title, mark = kits[0]
        if not is_rd_kit_mark(mark):
            error = f"Марка {mark!r} не является комплектом РД."
    od_docs = [item for item in documents if item.is_od]
    od_revision = od_docs[0].revision if od_docs else _highest_revision(documents)
    kit_code = ""
    if od_docs and od_docs[0].code:
        kit_code = od_docs[0].code
    elif documents:
        counts = Counter(item.code for item in documents if item.code)
        if counts:
            kit_code = counts.most_common(1)[0][0]
    stage = _CODE_STAGE.get(kit_code, "")
    letter_counts = tuple(
        sorted(Counter(item.code or "?" for item in documents).items())
    )
    f_line = ""
    if not error and date and stage:
        f_line = _catalog_f_line(
            date=date,
            stage=stage,
            revision=od_revision,
            transmittal=trm,
            documents=documents,
            record_mto_absent=True,
        )
    return ApprovalMail(
        kind="review_codes",
        subject=subject,
        date=date,
        stage=stage,
        send_transmittal=trm,
        incoming_transmittal="",
        title=title,
        mark=mark,
        od_revision=od_revision,
        kit_code=kit_code,
        letter_counts=letter_counts,
        f_line=f_line,
        sheet_revision=format_sheet_revision_cell(od_revision) if not error else "",
        status_sheet=status_sheet_for_stage(stage) if not error else "",
        error=error,
        documents=tuple(documents),
    )


def _parse_cover_letter(
    *,
    subject: str,
    body: str,
    sent_at: datetime | str | None,
    kit_from_transmittal: KitLookup | None,
    kit_hint: tuple[str, str] | None = None,
) -> ApprovalMail:
    """Parse an outgoing BCC cover letter into a ``tdo_sent`` journal event."""

    documents = _parse_cover_documents(body)
    if not documents:
        documents = _documents_from_stems(body)
    date = _format_sent_date(sent_at)
    trm = _first_trm(subject) or _first_trm(body)
    kits = _unique_kits(documents)
    error = ""
    title = ""
    mark = ""
    if len(kits) > 1:
        error = (
            "В письме несколько титул–марок: "
            + ", ".join(f"{item[0]}-{item[1]}" for item in kits)
        )
    elif len(kits) == 1:
        title, mark = kits[0]
        if not is_rd_kit_mark(mark):
            error = f"Марка {mark!r} не является комплектом РД."
    else:
        title, mark, error = _resolve_tdo_kit(
            send_trm=trm,
            kit_from_transmittal=kit_from_transmittal,
            kit_hint=kit_hint,
        )
    od_docs = [item for item in documents if item.is_od]
    od_revision = od_docs[0].revision if od_docs else _highest_revision(documents)
    stage = "tdo_sent"
    f_line = ""
    if not error and date and stage:
        f_line = _catalog_f_line(
            date=date,
            stage=stage,
            revision=od_revision or None,
            transmittal=trm,
            documents=documents,
            record_mto_absent=True,
        )
    return ApprovalMail(
        kind="cover_letter",
        subject=subject,
        date=date,
        stage=stage,
        send_transmittal=trm,
        incoming_transmittal="",
        title=title,
        mark=mark,
        od_revision=od_revision,
        kit_code="",
        letter_counts=(),
        f_line=f_line,
        sheet_revision=(
            format_sheet_revision_cell(od_revision)
            if od_revision and not error
            else ""
        ),
        status_sheet=status_sheet_for_stage(stage) if not error else "",
        error=error,
        documents=tuple(documents),
    )


def _parse_tdo_reply(
    *,
    subject: str,
    body: str,
    sent_at: datetime | str | None,
    kit_from_transmittal: KitLookup | None,
    kit_hint: tuple[str, str] | None = None,
) -> ApprovalMail:
    top = _top_reply(body)
    send_trm = _first_trm(subject)
    incoming = ""
    for token in _TRM_RE.findall(top):
        if token.casefold() != send_trm.casefold():
            incoming = token
            break
    stage = _classify_tdo_top(top, has_incoming=bool(incoming))
    date = _format_sent_date(sent_at)
    documents = _documents_from_stems(top)
    kits = _unique_kits(documents)
    if not kits:
        thread_docs = _documents_from_stems(body)
        thread_kits = _unique_kits(thread_docs)
        if len(thread_kits) == 1:
            documents = thread_docs
            kits = thread_kits
    prose = _pick_kit_prose(body)
    error = ""
    title = ""
    mark = ""
    if len(kits) > 1:
        error = (
            "В письме несколько титул–марок: "
            + ", ".join(f"{item[0]}-{item[1]}" for item in kits)
        )
    elif len(kits) == 1:
        title, mark = kits[0]
        if not is_rd_kit_mark(mark):
            error = f"Марка {mark!r} не является комплектом РД."
    elif prose is not None:
        title, mark = prose[0], prose[1]
        if not is_rd_kit_mark(mark):
            error = f"Марка {mark!r} не является комплектом РД."
    else:
        title, mark, error = _resolve_tdo_kit(
            send_trm=send_trm,
            kit_from_transmittal=kit_from_transmittal,
            kit_hint=kit_hint,
        )
    od_docs = [item for item in documents if item.is_od]
    od_revision = od_docs[0].revision if od_docs else _highest_revision(documents)
    if (
        not od_revision
        and prose is not None
        and title
        and kit_identity_key(title, mark) == kit_identity_key(prose[0], prose[1])
    ):
        od_revision = prose[2]
    f_line = ""
    if not error and date and stage:
        trm_for_line = incoming if stage == "incoming_passed" and incoming else send_trm
        f_line = _catalog_f_line(
            date=date,
            stage=stage,
            revision=od_revision or None,
            transmittal=trm_for_line,
            documents=documents,
            record_mto_absent=False,
        )
    return ApprovalMail(
        kind="tdo_reply",
        subject=subject,
        date=date,
        stage=stage,
        send_transmittal=send_trm,
        incoming_transmittal=incoming,
        title=title,
        mark=mark,
        od_revision=od_revision,
        kit_code="",
        letter_counts=(),
        f_line=f_line,
        sheet_revision=format_sheet_revision_cell(od_revision) if od_revision else "",
        status_sheet=status_sheet_for_stage(stage) if not error else "",
        error=error,
        documents=tuple(documents),
    )


def _parse_notification_documents(body: str) -> list[MailDocument]:
    matches = list(_DOC_START_RE.finditer(body))
    documents: list[MailDocument] = []
    for index, match in enumerate(matches):
        end = matches[index + 1].start() if index + 1 < len(matches) else len(body)
        block = body[match.start() : end]
        filename = match.group(1)
        revision = match.group(2)
        codes = _TAIL_CODE_RE.findall(block)
        letter = codes[-1][0].upper() if codes else None
        doc = _mail_document_from_filename(filename, revision=revision, code=letter)
        if doc is not None:
            documents.append(doc)
    return documents


def _parse_cover_documents(body: str) -> list[MailDocument]:
    documents: list[MailDocument] = []
    seen: set[str] = set()
    for match in _COVER_DOC_RE.finditer(body or ""):
        filename = match.group(1)
        revision = match.group(2)
        key = filename.casefold()
        if key in seen:
            continue
        seen.add(key)
        doc = _mail_document_from_filename(filename, revision=revision)
        if doc is not None:
            documents.append(doc)
    return documents


def _documents_from_stems(text: str) -> list[MailDocument]:
    by_stem: dict[str, MailDocument] = {}
    order: list[str] = []
    for match in _STEM_REV_RE.finditer(text or ""):
        stem = match.group("stem")
        revision = match.group("rev") or ""
        key = stem.casefold()
        doc = _mail_document_from_filename(stem, revision=revision)
        if doc is None:
            continue
        prev = by_stem.get(key)
        if prev is None:
            by_stem[key] = doc
            order.append(key)
        elif doc.revision and not prev.revision:
            by_stem[key] = doc
    if order:
        return [by_stem[key] for key in order]
    documents: list[MailDocument] = []
    seen: set[str] = set()
    for raw in _AGCC_KIT_RE.findall(text or ""):
        key = raw.casefold()
        if key in seen:
            continue
        seen.add(key)
        doc = _mail_document_from_filename(raw)
        if doc is not None:
            documents.append(doc)
    return documents


def _pick_kit_prose(text: str) -> tuple[str, str, str] | None:
    """Return a unique ``(title, mark, revision)`` from old TDO/cover wording.

    Matches ``по титулу 6816 марка KSB рев.0`` (``марка`` may wrap onto the
    next line) and ``титулу AGCC.287-6816-KSB _0``. Returns ``None`` when
    several title–marks appear. A single kit with conflicting revisions
    keeps the kit and an empty revision.

    Args:
        text: Plain body (quoted thread included).

    Returns:
        Unique kit triple, or ``None``.
    """

    found: list[tuple[str, str, str]] = []
    seen: set[tuple[str, str, str]] = set()
    for match in _KIT_PROSE_RE.finditer(text or ""):
        title = match.group(1)
        mark = match.group(3) or match.group(2) or ""
        revision = match.group(4) or ""
        item = (title, mark, revision)
        if not mark or item in seen:
            continue
        seen.add(item)
        found.append(item)
    for match in _KIT_AGCC_PROSE_RE.finditer(text or ""):
        item = (match.group(1), match.group(2), match.group(3) or "")
        if not item[1] or item in seen:
            continue
        seen.add(item)
        found.append(item)
    if not found:
        return None
    kits = {(title, mark.casefold()) for title, mark, _revision in found}
    if len(kits) != 1:
        return None
    title, mark = found[0][0], found[0][1]
    revisions = {revision for _title, _mark, revision in found if revision}
    if len(revisions) != 1:
        return title, mark, ""
    return title, mark, next(iter(revisions))


def _mail_document_from_filename(
    filename: str,
    *,
    revision: str = "",
    code: str | None = None,
) -> MailDocument | None:
    title, mark = _title_mark_from_filename(filename)
    if not title:
        return None
    return MailDocument(
        filename=filename,
        title=title,
        mark=mark,
        revision=revision or _revision_from_filename(filename),
        code=code,
        is_od=_is_od_filename(filename),
    )


def _revision_from_filename(filename: str) -> str:
    parts = AgccFilenamePatterns.parse_strict(filename) or AgccFilenamePatterns.parse_loose(
        filename
    )
    if parts is None:
        return ""
    tail = AgccFilenamePatterns.split_revision_tail(
        getattr(parts, "revision_tail", None)
    )
    if tail is None:
        return ""
    revision = tail.rev_sheet
    if tail.an:
        revision = f"{revision}-AN{tail.an}"
    return revision


def _is_od_filename(filename: str) -> bool:
    parts = AgccFilenamePatterns.parse_strict(filename) or AgccFilenamePatterns.parse_loose(
        filename
    )
    if parts is not None:
        code, _serial = AgccFilenamePatterns._split_discipline(parts.discipline_block)
        return (code or "").casefold() == "od"
    upper = filename.upper()
    return ".OD-" in upper or upper.endswith(".OD")


def _is_mto_filename(filename: str) -> bool:
    """Return whether *filename* is an AGCC MTO document.

    Discipline block must start with ``mto`` (same ``AgccFilenamePatterns``
    split as :func:`_is_od_filename`).
    """

    parts = AgccFilenamePatterns.parse_strict(filename) or AgccFilenamePatterns.parse_loose(
        filename
    )
    if parts is not None:
        code, _serial = AgccFilenamePatterns._split_discipline(parts.discipline_block)
        return (code or "").casefold().startswith("mto")
    upper = filename.upper()
    return ".MTO-" in upper or upper.endswith(".MTO")


def _mto_revision_from_documents(documents: list[MailDocument]) -> str:
    """Return MTO filename revision from letter documents, or empty."""

    return _highest_revision(
        [item for item in documents if _is_mto_filename(item.filename)]
    )


def _catalog_f_line(
    *,
    date: str,
    stage: str,
    revision: str | None,
    transmittal: str,
    documents: list[MailDocument],
    record_mto_absent: bool,
) -> str:
    """Format an F line written by the catalog (always ``auto``).

    Transfer letters (notification / cover) pass ``record_mto_absent=True``
    so a package without an MTO file is stored as ``MTO Нет``. TDO replies
    only attach ``MTO <rev>`` when a document stem is present.

    Args:
        date: ``DD.MM.YYYY``.
        stage: Journal stage key.
        revision: OD revision, or empty.
        transmittal: TRM token, or empty.
        documents: Parsed letter documents.
        record_mto_absent: Write ``MTO Нет`` when no MTO document exists.

    Returns:
        Canonical F line from :func:`format_history_line`.
    """

    mto_revision = _mto_revision_from_documents(documents) or None
    return format_history_line(
        date=date,
        stage=stage,
        revision=revision,
        transmittal=transmittal,
        mto_revision=mto_revision,
        mto_absent=bool(record_mto_absent and mto_revision is None),
        from_robot_auto=True,
    )


def _title_mark_from_filename(filename: str) -> tuple[str, str]:
    parts = AgccFilenamePatterns.parse_strict(filename) or AgccFilenamePatterns.parse_loose(
        filename
    )
    if parts is None:
        match = re.match(
            r"AGCC\.287-(\d{4})-([A-Za-z0-9.]+)",
            filename,
            re.IGNORECASE,
        )
        if not match:
            return "", ""
        return match.group(1), match.group(2)
    title, mark = AgccFilenamePatterns._split_title_system(parts.title_system)
    return title or "", mark or ""


def _unique_kits(documents: list[MailDocument]) -> list[tuple[str, str]]:
    ordered: list[tuple[str, str]] = []
    seen: set[tuple[str, str]] = set()
    for item in documents:
        if not item.title or not item.mark:
            continue
        key = kit_identity_key(item.title, item.mark)
        if key in seen:
            continue
        seen.add(key)
        ordered.append((item.title, item.mark))
    return ordered


def _highest_revision(documents: list[MailDocument]) -> str:
    for item in documents:
        if item.revision:
            return item.revision
    return ""


def _is_cover_letter(subject: str, body: str) -> bool:
    if _REPLY_SUBJECT_RE.match((subject or "").lstrip()):
        return False
    if _COVER_SUBJECT_RE.search(subject or ""):
        return True
    return bool(
        _COVER_BODY_HINT_RE.search(body or "") and _parse_cover_documents(body)
    )


def _should_prefer_html(body: str, html_text: str) -> bool:
    if _has_cyrillic(html_text) and not _has_cyrillic(body):
        return True
    html_cover = len(_parse_cover_documents(html_text))
    body_cover = len(_parse_cover_documents(body))
    if html_cover > body_cover:
        return True
    if html_text and _has_cyrillic(html_text) and not _has_review_tokens(body):
        return True
    html_codes = len(_TAIL_CODE_RE.findall(html_text))
    body_codes = len(_TAIL_CODE_RE.findall(body))
    return html_codes > body_codes


def _resolve_tdo_kit(
    *,
    send_trm: str,
    kit_from_transmittal: KitLookup | None,
    kit_hint: tuple[str, str] | None,
) -> tuple[str, str, str]:
    """Return title, mark, error for a TDO reply without filenames."""

    if send_trm and kit_from_transmittal is not None:
        looked = kit_from_transmittal(send_trm)
        if looked:
            return looked[0], looked[1], ""
    if kit_hint:
        title, mark = kit_hint
        if is_rd_kit_mark(mark):
            return title, mark, ""
        return "", "", f"Марка {mark!r} не является комплектом РД."
    if send_trm and kit_from_transmittal is not None:
        return "", "", f"Нет комплекта в Выдаче для TRM {send_trm}."
    if send_trm:
        return (
            "",
            "",
            f"Не удалось определить титул и марку по TRM {send_trm}. "
            "Загрузите комплекты Google (Выдача РД ПД) "
            "или откройте письмо из папки вида 1600-SOT.",
        )
    return "", "", "Не удалось определить титул и марку (нет имён файлов и TRM)."


def _classify_tdo_top(top: str, *, has_incoming: bool) -> str:
    low = top.casefold()
    if "загруз" in low and "ср" in low:
        return "incoming_passed" if has_incoming else "sr_upload"
    if "направлен" in low and "на рассмотрение" in low:
        return "incoming_passed"
    if "на рассмотрении" in low:
        return "incoming_passed"
    if _DOC_ACCEPTED_RE.search(top or ""):
        return "incoming_passed"
    if _DOC_LOADED_RE.search(top or ""):
        return "incoming_passed" if has_incoming else "sr_upload"
    return ""


def _top_reply(body: str) -> str:
    match = _FROM_SPLIT_RE.search(body or "")
    if match:
        return (body or "")[: match.start()].strip()
    return (body or "").strip()


def _notification_trm(subject: str) -> str:
    match = _NOTIFICATION_TRM_RE.search(subject or "")
    return match.group(1).strip() if match else ""


def _first_trm(text: str) -> str:
    match = _TRM_RE.search(text or "")
    return match.group(0) if match else ""


def _transmittal_date(body: str) -> str:
    flat = re.sub(r"\s+", " ", body or "")
    match = _TRANSMITTAL_DATE_RE.search(flat)
    return match.group(1) if match else ""


def _format_sent_date(sent_at: datetime | str | None) -> str:
    if sent_at is None:
        return ""
    if isinstance(sent_at, datetime):
        return sent_at.strftime("%d.%m.%Y")
    raw = str(sent_at).strip()
    for fmt in ("%d.%m.%Y", "%Y-%m-%d"):
        try:
            return datetime.strptime(raw[:10], fmt).strftime("%d.%m.%Y")
        except ValueError:
            continue
    match = re.match(r"(\d{4})-(\d{2})-(\d{2})", raw)
    if match:
        return f"{match.group(3)}.{match.group(2)}.{match.group(1)}"
    return ""


def _restore_cyrillic(text: str) -> str:
    if _has_cyrillic(text):
        return text
    recoded = _recode_cp1251(text)
    return recoded if recoded and _has_cyrillic(recoded) else text


def _recode_cp1251(text: str) -> str:
    for encoding in ("cp1252", "latin-1"):
        try:
            return text.encode(encoding).decode("cp1251")
        except UnicodeError:
            continue
    return ""


def _has_cyrillic(text: str) -> bool:
    return any("\u0400" <= char <= "\u04ff" for char in text or "")


def _has_review_tokens(text: str) -> bool:
    return bool(
        re.search(
            r"замечан|незначительн|существенн|критическ|загруз|"
            r"направлен|рассмотр|transmittal date",
            text or "",
            re.IGNORECASE,
        )
    )


def _decode_html(raw: bytes | str | None) -> str:
    if raw is None:
        return ""
    if isinstance(raw, str):
        return _restore_cyrillic(raw)
    data = bytes(raw)
    try:
        text = data.decode("utf-8")
    except UnicodeDecodeError:
        return data.decode("cp1251", errors="replace")
    restored = _restore_cyrillic(text)
    if _has_cyrillic(restored):
        return restored
    try:
        cp_text = data.decode("cp1251")
    except UnicodeError:
        return restored
    return cp_text if _has_cyrillic(cp_text) else restored


def _html_to_text(html: str) -> str:
    if not html:
        return ""
    text = re.sub(r"(?is)<(script|style)[^>]*>.*?</\1>", " ", html)
    text = re.sub(r"(?i)<br\s*/?>", "\n", text)
    text = re.sub(r"(?i)</(p|div|tr|h1|h2|li)>", "\n", text)
    text = re.sub(r"(?i)</t[dh]>", " ", text)
    text = re.sub(r"<[^>]+>", " ", text)
    text = unescape(text)
    text = text.replace("\xa0", " ")
    text = re.sub(r"[ \t]+\n", "\n", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()
