"""Extract an RFP materials table from PDF into an xlsx workbook.

Uses PyMuPDF ``page.find_tables()`` only (no pdfplumber / pdf_parsing_v2).
Output is a subfolder next to the source PDF:
``результат распознавания PDF_<YYYY.MM.DD_HH.MM>``.
The xlsx and ``pdf_rfp_report.txt`` both go there. Nothing is written into
``RFP_Зиновьев``.
"""

from __future__ import annotations

import re
import time
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any, Iterable

import fitz
from openpyxl import Workbook, load_workbook

from RFQ.rfp_parts.analyze_rfp_parts import (
    EXPECTED_SHEET_NAME,
    REQUIRED_FIELDS,
    _is_total_row,
    _match_header,
    _parse_decimal,
    _tags_text_contains_rfq,
)
from RFQ.rfp_parts.file_status import format_field_list

REPORT_NAME = "pdf_rfp_report.txt"
OUTSIDE_LIST_CAP = 50
RESULT_DIR_PREFIX = "результат распознавания PDF_"
MATERIALS_HEADER_MARKERS = ("Наименование МТР", "Закупка по")
IDENTITY_FIELDS = (
    "DS_TITLE",
    "DS_SPECIFICATION",
    "TAGS",
    "DS_CODE_1C",
    "CODE",
    "UNITS",
    "VALUES",
)
# Comma between digits is a decimal/thousands mark (``1,000``), not a tag list.
_SPLIT_KEEP_RE = re.compile(r"(;|\n\s*\n|(?<!\d),(?!\d))")
# PyMuPDF + common Windows fonts remap ASCII '-' / ';' in extracted text.
_PDF_CHAR_MAP = str.maketrans(
    {
        "\u00ad": "-",  # soft hyphen
        "\u2010": "-",
        "\u2011": "-",
        "\u2012": "-",
        "\u2013": "-",
        "\u2014": "-",
        "\u2212": "-",
        "\u037e": ";",  # Greek question mark (Arial semicolon)
        "\xa0": " ",
    }
)
_DS_SLASH_RE = re.compile(
    r"№\s*(\d+)\s*/\s*(\d+)([А-Яа-яЁё])?",
    re.IGNORECASE,
)
_DS_AGREEMENT_RE = re.compile(
    r"дополнительн\w*\s+соглашен\w*\s+№\s*(\d+)",
    re.IGNORECASE,
)
_DS_BARE_RE = re.compile(r"№\s*(\d+)")
_APPENDIX_PREFIX_RE = re.compile(r"приложен", re.IGNORECASE)
_FURNITURE_WORDS = frozenset({"страница", "из"})
_STAMP_NEAR_RE = re.compile(r"Диадок|Передан|GMT\+|Страница|[0-9a-fA-F]{8}-")
_STAMP_WORD_RE = re.compile(
    r"^(?:Передан|через|Диадок|Страница|из|"
    r"GMT\+\d{2}:\d{2}|\d{2}:\d{2}|\d{2}\.\d{2}\.\d{4}|\d{1,4}|"
    r"[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12})$",
    re.IGNORECASE,
)
_STAMP_CELL_RE = re.compile(
    r"Передан\s*через\s*Диадок|Диадок|GMT\+\d{2}:\d{2}|"
    r"[0-9a-fA-F]{8}(?:-[0-9a-fA-F]{4}){3}-[0-9a-fA-F]{12}|"
    r"Страница\s*\d+\s*из\s*\d+",
    re.IGNORECASE,
)
_HEADER_NPP = "№ п/п"
_GROUP_RD = "Закупка по РД"
_GROUP_LOT = 'Закупка по "Лоту"'


def _log_phase(phase: str, started: float, *, detail: str = "") -> None:
    """Print one timed line. ``print`` is mirrored into the job monitor."""
    stamp = datetime.now().strftime("%H:%M:%S")
    extra = f" — {detail}" if detail else ""
    elapsed = time.perf_counter() - started
    print(f"[pdf rfp] {stamp} {phase}{extra} — {elapsed:.2f} с", flush=True)


@dataclass(frozen=True)
class _MaterialsPage:
    """Rows and cell boxes copied while the page table is still valid.

    PyMuPDF table objects are overwritten by the next ``find_tables`` call,
    so text and rectangles must be snapshotted immediately.
    """

    page_index: int
    rows: tuple[tuple[str, ...], ...]
    col_count: int
    cell_rects: tuple[tuple[float, float, float, float], ...]


@dataclass(frozen=True)
class PdfRfpIssue:
    """One extract or post-check remark."""

    level: str  # "ERROR" or "WARN"
    message: str
    page: int | None = None
    row_number: str = ""


@dataclass(frozen=True)
class PdfRfpResult:
    """Paths and counters after a successful extract (xlsx is always written)."""

    xlsx_path: Path
    report_path: Path
    out_dir: Path
    ds_label: str
    row_count: int
    contract_errors: int
    outside_words: int
    issues: tuple[PdfRfpIssue, ...]
    cleaned_pdf_path: Path | None = None


@dataclass(frozen=True)
class _OutsideWord:
    page: int
    word: str
    x: float
    y: float


def parse_ds_label_from_title_text(text: str) -> str:
    """From title-page text return ``ДС98_35Б`` or ``""``.

    Match a phrase like::

        Приложение №2 к Дополнительному соглашению №98/35Б от 21.07.2026

    Number before the slash is actual (98). Number after the slash is
    historical (35). Optional Cyrillic letter after the historical number
    is the revision (Б). Return ``ДС{actual}_{historical}{letter}`` with an
    underscore, never a slash. If there is no slash, ``№98`` alone →
    ``ДС98``. No match → empty string.

    Args:
        text: Raw text of the title page (or any haystack).

    Returns:
        Sanitized DS label, or an empty string.
    """
    if not text:
        return ""
    slash = _DS_SLASH_RE.search(text)
    if slash:
        actual, historical, letter = slash.group(1), slash.group(2), slash.group(3) or ""
        return f"ДС{actual}_{historical}{letter.upper()}"
    agreement = _DS_AGREEMENT_RE.search(text)
    if agreement:
        return f"ДС{agreement.group(1)}"
    for match in _DS_BARE_RE.finditer(text):
        prefix = text[max(0, match.start() - 24) : match.start()]
        if _APPENDIX_PREFIX_RE.search(prefix):
            continue
        return f"ДС{match.group(1)}"
    return ""


def preview_ds_label(pdf_path: Path) -> str:
    """Read page 0 text only and return :func:`parse_ds_label_from_title_text`.

    Args:
        pdf_path: RFP PDF path.

    Returns:
        Detected DS label, or an empty string.
    """
    doc = fitz.open(Path(pdf_path))
    try:
        if doc.page_count < 1:
            return ""
        return parse_ds_label_from_title_text(doc[0].get_text() or "")
    finally:
        doc.close()


def pdf_rfp_stamp_dir(pdf_path: Path, *, when: datetime | None = None) -> Path:
    """Return a result folder next to the source PDF.

    Name: ``результат распознавания PDF_<YYYY.MM.DD_HH.MM>``. If that folder
    already exists, append ``_2``, ``_3``, … The directory is not created here.

    Args:
        pdf_path: Source PDF. The folder is created in its parent directory.
        when: Timestamp for the folder name. Default is now.

    Returns:
        Unique stamp path beside the PDF.
    """
    parent = Path(pdf_path).resolve().parent
    stamp = (when or datetime.now()).strftime("%Y.%m.%d_%H.%M")
    candidate = parent / f"{RESULT_DIR_PREFIX}{stamp}"
    if not candidate.exists():
        return candidate
    suffix = 2
    while True:
        candidate = parent / f"{RESULT_DIR_PREFIX}{stamp}_{suffix}"
        if not candidate.exists():
            return candidate
        suffix += 1


def _rect_covers_table_word(rect: fitz.Rect, words: list[tuple]) -> bool:
    """True when a stamp box also covers a table word (qty, unit, date)."""
    for word in words:
        token = str(word[4])
        if _STAMP_WORD_RE.match(token):
            continue
        if rect.intersects(fitz.Rect(word[:4])):
            return True
    return False


def strip_diadoc_from_cell(value: Any) -> str:
    """Drop Diadoc fragments that share a cell with table text.

    The UUID line sits on the same baseline as the lot quantity, so a
    redaction rectangle would erase the number. Those glyphs stay in the
    PDF and are removed from the cell string instead.

    Args:
        value: Raw cell text.

    Returns:
        Cell text without the transfer line, UUID, or stamp page counter.
    """
    if value is None:
        return ""
    text = str(value)
    if not _STAMP_CELL_RE.search(text) and not re.search(
        r"[0-9a-fA-F]{4,}-", text
    ):
        return text
    cleaned = _STAMP_CELL_RE.sub(" ", text)
    cleaned = re.sub(r"Передан|через|Диадок", " ", cleaned, flags=re.IGNORECASE)
    cleaned = re.sub(r"[0-9a-fA-F]{4,}-[0-9a-fA-F,-]{2,}", " ", cleaned)
    cleaned = re.sub(r"\d{1,2}\.\d{2}\.\d{4}", " ", cleaned)
    cleaned = re.sub(r"\b\d{1,2}:\d{2}\b", " ", cleaned)
    return cleaned


def clean_code_cell(value: Any) -> str:
    """Glue wrapped identity text without joining semicolon/comma/blank-line lists.

    Replaces NBSP with space and drops ``\\r``. Splits on ``;``, a comma that
    is not between digits, and a blank line (``\\n\\s*\\n``), keeping the
    separator. ``1,000`` stays one token. Inside each non-separator piece,
    deletes all whitespace (so ``8950-\\nSOT4`` becomes ``8950-SOT4`` and
    ``1,\\n000`` becomes ``1,000``). Rejoins: semicolon and blank-line
    separators become ``; ``, comma separators become ``, ``.

    Args:
        value: Raw cell value from ``find_tables`` extract.

    Returns:
        Cleaned identity string.
    """
    if value is None:
        return ""
    text = strip_diadoc_from_cell(value).replace("\xa0", " ").replace("\r", "")
    if not text.strip():
        return ""
    pieces: list[str] = []
    for part in _SPLIT_KEEP_RE.split(text):
        if part == ";" or (part.startswith("\n") and part.endswith("\n")):
            pieces.append("; ")
        elif part == ",":
            pieces.append(", ")
        else:
            pieces.append("".join(ch for ch in part if not ch.isspace()))
    return "".join(pieces).strip()


def clean_text_cell(value: Any) -> str:
    """Collapse prose-cell whitespace; do not invent missing tails.

    Args:
        value: Raw cell value from ``find_tables`` extract.

    Returns:
        NBSP-normalized text with all whitespace folded to a single space.
    """
    if value is None:
        return ""
    text = strip_diadoc_from_cell(value).replace("\xa0", " ").replace("\r", " ")
    return re.sub(r"\s+", " ", text).strip()


def _stamp_font_keys(page: fitz.Page) -> list[str]:
    """Resource names of the Diadoc subset font, not body Segoe text."""
    keys: list[str] = []
    for font in page.get_fonts() or []:
        name = str(font[3])
        key = str(font[4])
        if name.startswith("EPDLTT+") and "SegoeUI" in name:
            keys.append(key)
    return keys


def _cut_stamp_stream(raw: bytes, font_key: str) -> bytes | None:
    """Drop the trailing blue stamp block that paints ``font_key``.

    The block is a ``q`` … ``Q`` at the end of the page stream: blue fill,
    the transfer lines, the icon and the circle. Table text is earlier and
    uses another font, so it stays even where the stamp bbox overlaps it.
    """
    token = b"/" + font_key.encode("ascii") + b" "
    pos = raw.rfind(token)
    if pos < 0:
        return None
    color = b"0 0.33333 0.7451 rg"
    start = raw.rfind(color, 0, pos)
    if start < 0:
        return None
    qpos = raw.rfind(b"q", 0, start)
    if qpos < 0 or start - qpos > 24:
        qpos = start
    return raw[:qpos]


def _apply_stamp_stream_cut(page: fitz.Page) -> bool:
    keys = _stamp_font_keys(page)
    if not keys:
        return False
    xrefs = page.get_contents()
    if not xrefs:
        return False
    raw = page.read_contents()
    updated = raw
    for key in keys:
        cut = _cut_stamp_stream(updated, key)
        if cut is not None and cut != updated:
            updated = cut
    if updated == raw:
        return False
    page.clean_contents()
    xrefs = page.get_contents()
    if len(xrefs) != 1:
        return False
    page.parent.update_stream(xrefs[0], updated)
    return True


def strip_diadoc_stamps(doc: fitz.Document) -> int:
    """Remove the Diadoc transfer line and its icon from every page.

    The stamp is real text plus a small image in the bottom band, including
    a page counter whose number changes per sheet. The document's own
    «Страница N из M» in the header is a different block and is left in place.
    Vector grid lines are kept.

    Args:
        doc: Open PDF, mutated in place.

    Returns:
        Number of pages where a stamp was redacted.
    """
    pages_hit = 0
    for page in doc:
        if _apply_stamp_stream_cut(page):
            pages_hit += 1
            continue
        anchor: fitz.Rect | None = None
        text_rects: list[fitz.Rect] = []
        data = page.get_text("dict") or {}
        for block in data.get("blocks", []):
            if block.get("type") != 0:
                continue
            text = "".join(
                span.get("text", "")
                for line in block.get("lines", [])
                for span in line.get("spans", [])
            )
            if "Диадок" not in text:
                continue
            anchor = fitz.Rect(block["bbox"])
            text_rects.append(anchor)
        if anchor is None:
            continue
        # Only the stamp cluster at the right. A full-width table row can
        # share the same y and must not be painted over: a white box becomes
        # an extra column for find_tables.
        band = fitz.Rect(
            anchor.x0 - 30,
            anchor.y0 - 8,
            page.rect.x1,
            anchor.y1 + 36,
        )
        rects = list(text_rects)
        for block in data.get("blocks", []):
            box = fitz.Rect(block["bbox"])
            if box.x0 < anchor.x0 - 30 or not box.intersects(band):
                continue
            if block.get("type") == 1:
                rects.append(box)
                continue
            if block.get("type") != 0:
                continue
            text = "".join(
                span.get("text", "")
                for line in block.get("lines", [])
                for span in line.get("spans", [])
            )
            if _STAMP_NEAR_RE.search(text) and box not in rects:
                rects.append(box)
        words = page.get_text("words") or []
        added = False
        for rect in rects:
            if _rect_covers_table_word(rect, words):
                continue
            page.add_redact_annot(rect)
            added = True
        if not added:
            pages_hit += 1
            continue
        page.apply_redactions(
            images=fitz.PDF_REDACT_IMAGE_REMOVE,
            graphics=fitz.PDF_REDACT_LINE_ART_NONE,
        )
        pages_hit += 1
    return pages_hit


def extract_rfp_pdf(
    pdf_path: Path,
    *,
    out_dir: Path | None = None,
    ds_label: str | None = None,
    strip_stamp: bool = False,
) -> PdfRfpResult:
    """Extract the materials table to xlsx plus ``pdf_rfp_report.txt``.

    ``ds_label=None`` detects the label from page 0. A passed string
    (including ``""``) overrides detection. An empty label after sanitize
    warns and names the xlsx from the PDF stem only.

    Args:
        pdf_path: Source RFP PDF.
        out_dir: Stamp folder. Default is a dated folder next to the PDF.
        ds_label: Override for the DS filename prefix.
        strip_stamp: When True, redact the Diadoc overlay before table extract
            and write ``<stem>_без_печати.pdf`` into the result folder.

    Returns:
        Paths, row count, contract/outside counters, and issues.

    Raises:
        RuntimeError: No materials table, or the header cannot be mapped.
            No half xlsx is written in that case.
    """
    pdf_path = Path(pdf_path)
    issues: list[PdfRfpIssue] = []
    started = time.perf_counter()
    _log_phase("старт", started, detail=pdf_path.name)
    doc = fitz.open(pdf_path)
    cleaned_pdf: Path | None = None
    try:
        _log_phase("pdf открыт", started, detail=f"страниц {doc.page_count}")
        stripped_pages = strip_diadoc_stamps(doc) if strip_stamp else 0
        if strip_stamp:
            _log_phase(
                "печать Диадок",
                started,
                detail=f"снята на листах {stripped_pages}",
            )
        page0_text = doc[0].get_text() if doc.page_count else ""
        if ds_label is None:
            raw_label = parse_ds_label_from_title_text(page0_text)
        else:
            raw_label = ds_label
        chosen_label = _sanitize_ds_label(raw_label)
        if not chosen_label:
            issues.append(
                PdfRfpIssue(
                    "WARN",
                    "не удалось определить метку ДС по титульному листу",
                )
            )

        page_scan = _collect_materials_page_tables(doc, started=started)
        page_tables = page_scan.pages
        if not page_tables:
            raise RuntimeError(
                "в PDF не найдена таблица материалов "
                "(нет ячейки «Наименование МТР» или «Закупка по»)"
            )

        col_count = page_tables[0].col_count
        all_rows: list[list[str]] = []
        for materials_page in page_tables:
            all_rows.extend(list(row) for row in materials_page.rows)
        _log_phase(
            "строки сняты",
            started,
            detail=f"страниц {len(page_tables)}, строк {len(all_rows)}",
        )
        coverage = materials_coverage_issue(
            pdf_page_count=doc.page_count,
            first_page=page_tables[0].page_index + 1,
            last_page=page_tables[-1].page_index + 1,
            expected_cols=page_tables[0].col_count,
            stop_page=page_scan.stop_page,
            stop_cols=page_scan.stop_cols,
        )
        if coverage is not None:
            issues.append(coverage)
            _log_phase("охват листов", started, detail=coverage.message)

        header_row, left_cols, right_cols = _find_materials_header(all_rows)
        if header_row is None or right_cols is None or left_cols is None:
            _log_header_miss(all_rows)
            raise RuntimeError(
                "не удалось разобрать шапку таблицы материалов "
                f"(нужны столбцы: {format_field_list(REQUIRED_FIELDS)})"
            )
        _log_phase(
            "шапка",
            started,
            detail="лот=" + ", ".join(
                f"{field}:{right_cols[field]}" for field in REQUIRED_FIELDS
            ),
        )

        identity_idx = _identity_column_indexes(left_cols, right_cols)
        data_rows = [
            row
            for row in all_rows
            if not _is_dropped_header_or_total(row)
        ]
        cleaned_rows = [
            _clean_data_row(row, identity_idx) for row in data_rows
        ]
        code_idx = right_cols.get("CODE", left_cols.get("CODE"))
        if code_idx is not None:
            for excel_offset, row in enumerate(cleaned_rows, start=3):
                code = row[code_idx] if code_idx < len(row) else ""
                if _code_looks_truncated_agcc(code):
                    issues.append(
                        PdfRfpIssue(
                            "WARN",
                            (
                                f"код «{code}» похож на номер документа AGCC "
                                "и обрезан (не заканчивается цифрой)"
                            ),
                            row_number=str(excel_offset),
                        )
                    )

        outside = _scan_outside_words(doc, page_tables)
        _log_phase("слова вне ячеек", started, detail=str(len(outside)))
        stamp = Path(out_dir) if out_dir is not None else pdf_rfp_stamp_dir(pdf_path)
        if stripped_pages:
            stamp.mkdir(parents=True, exist_ok=True)
            cleaned_pdf = stamp / f"{pdf_path.stem}_без_печати.pdf"
            doc.save(cleaned_pdf, garbage=3, deflate=True)
            _log_phase("pdf без печати", started, detail=cleaned_pdf.name)
    finally:
        doc.close()

    stamp.mkdir(parents=True, exist_ok=True)
    _log_phase("папка результата", started, detail=str(stamp))
    xlsx_name = _xlsx_filename(chosen_label, pdf_path.stem)
    xlsx_path = stamp / xlsx_name
    _write_materials_xlsx(
        xlsx_path,
        header_row=header_row,
        data_rows=cleaned_rows,
        left_cols=left_cols,
        right_cols=right_cols,
        col_count=max(col_count, len(header_row)),
    )

    issues.extend(_contract_check(xlsx_path))
    if outside:
        issues.append(
            PdfRfpIssue(
                "WARN",
                f"на страницах материалов есть слова вне ячеек таблицы: {len(outside)}",
            )
        )

    contract_errors = sum(1 for item in issues if item.level == "ERROR")
    report_path = stamp / REPORT_NAME
    _write_report(
        report_path,
        ds_label=chosen_label,
        row_count=len(cleaned_rows),
        contract_errors=contract_errors,
        outside_words=len(outside),
        issues=issues,
        outside=outside,
        cleaned_pdf_path=cleaned_pdf,
    )
    _log_phase(
        "готово",
        started,
        detail=(
            f"строк {len(cleaned_rows)}, ошибок контракта {contract_errors}, "
            f"слов вне ячеек {len(outside)}"
        ),
    )
    return PdfRfpResult(
        xlsx_path=xlsx_path,
        report_path=report_path,
        out_dir=stamp,
        ds_label=chosen_label,
        row_count=len(cleaned_rows),
        contract_errors=contract_errors,
        outside_words=len(outside),
        issues=tuple(issues),
        cleaned_pdf_path=cleaned_pdf,
    )


def _sanitize_ds_label(raw: str) -> str:
    """Keep ДС / digits / underscore / Cyrillic; replace slash with underscore."""
    text = (raw or "").replace("/", "_")
    kept = "".join(
        ch
        for ch in text
        if ch.isdigit() or ch == "_" or _is_cyrillic_letter(ch)
    )
    return kept.upper()


def _is_cyrillic_letter(ch: str) -> bool:
    return "\u0400" <= ch <= "\u04FF"


def _xlsx_filename(ds_label: str, pdf_stem: str) -> str:
    if ds_label:
        return f"{ds_label}. {pdf_stem}.xlsx"
    return f"{pdf_stem}.xlsx"


def _table_extract(table: Any) -> list[list[Any]]:
    extracted = table.extract() or []
    return [list(row) for row in extracted]


def _table_col_count(table: Any, extracted: list[list[Any]]) -> int:
    n = getattr(table, "col_count", None)
    if isinstance(n, int) and n > 0:
        return n
    return max((len(row) for row in extracted), default=0)


def _table_area(table: Any) -> float:
    bbox = getattr(table, "bbox", None)
    if bbox is None:
        return 0.0
    x0, y0, x1, y1 = bbox
    return max(0.0, float(x1 - x0) * float(y1 - y0))


def _normalize_pdf_text(value: Any) -> str:
    if value is None:
        return ""
    return str(value).translate(_PDF_CHAR_MAP)


def _pad_row(row: Iterable[Any], col_count: int) -> list[str]:
    values = [_normalize_pdf_text(cell) for cell in row]
    if len(values) < col_count:
        values.extend([""] * (col_count - len(values)))
    return values[:col_count] if col_count else values


def _row_has_materials_marker(row: Iterable[Any]) -> bool:
    for cell in row:
        text = _normalize_pdf_text(cell)
        if any(marker in text for marker in MATERIALS_HEADER_MARKERS):
            return True
    return False


def _is_materials_table(table: Any) -> bool:
    for row in _table_extract(table):
        if _row_has_materials_marker(row):
            return True
    return False


def _page_tables(page: fitz.Page) -> list[Any]:
    finder = page.find_tables()
    tables = getattr(finder, "tables", None)
    if not tables:
        return []
    return list(tables)


@dataclass(frozen=True)
class _MaterialsScan:
    """Materials pages plus why the scan stopped."""

    pages: list[_MaterialsPage]
    stop_page: int | None
    stop_cols: int | None


def _collect_materials_page_tables(
    doc: fitz.Document,
    *,
    started: float,
) -> _MaterialsScan:
    """Return snapshotted materials pages.

    Text and cell boxes are copied before the next ``find_tables`` call.
    A later extract of the same table object returns another page's text.
    """
    found: list[_MaterialsPage] = []
    started_run = False
    expected_cols: int | None = None
    stop_page: int | None = None
    stop_cols: int | None = None
    for page_index in range(doc.page_count):
        page_started = time.perf_counter()
        page = doc[page_index]
        tables = _page_tables(page)
        find_s = time.perf_counter() - page_started
        if not tables:
            _log_phase(
                "страница",
                started,
                detail=f"стр. {page_index + 1}: таблиц нет, find {find_s:.2f} с",
            )
            if started_run:
                stop_page = page_index + 1
                break
            continue
        largest = max(tables, key=_table_area)
        extracted = _table_extract(largest)
        ncols = _table_col_count(largest, extracted)
        rows = tuple(
            tuple(_pad_row(raw, ncols if ncols else len(raw))) for raw in extracted
        )
        if not started_run:
            if not _is_materials_table(largest):
                _log_phase(
                    "страница",
                    started,
                    detail=(
                        f"стр. {page_index + 1}: не материалы, "
                        f"таблиц {len(tables)}, find {find_s:.2f} с"
                    ),
                )
                continue
            started_run = True
            expected_cols = ncols
            found.append(
                _MaterialsPage(
                    page_index=page_index,
                    rows=rows,
                    col_count=ncols,
                    cell_rects=tuple(_table_cell_rects(largest)),
                )
            )
            _log_phase(
                "страница",
                started,
                detail=(
                    f"стр. {page_index + 1}: старт материалов, "
                    f"колонок {ncols}, строк {len(rows)}, find {find_s:.2f} с"
                ),
            )
            continue
        if ncols == expected_cols and _rows_continue_materials(rows):
            found.append(
                _MaterialsPage(
                    page_index=page_index,
                    rows=rows,
                    col_count=ncols,
                    cell_rects=tuple(_table_cell_rects(largest)),
                )
            )
            _log_phase(
                "страница",
                started,
                detail=(
                    f"стр. {page_index + 1}: продолжение, "
                    f"строк {len(rows)}, find {find_s:.2f} с"
                ),
            )
        else:
            stop_page = page_index + 1
            stop_cols = ncols
            _log_phase(
                "страница",
                started,
                detail=(
                    f"стр. {page_index + 1}: стоп, колонок {ncols}, find {find_s:.2f} с"
                ),
            )
            break
    return _MaterialsScan(pages=found, stop_page=stop_page, stop_cols=stop_cols)


def materials_coverage_issue(
    *,
    pdf_page_count: int,
    first_page: int,
    last_page: int,
    expected_cols: int,
    stop_page: int | None,
    stop_cols: int | None,
) -> PdfRfpIssue | None:
    """Warn when the materials table stops while many pages are still left.

    A sharp drop in column count is treated as an appendix and is not a
    truncated extract. A similar or larger column count with a long tail is
    the case where a stamp split the grid and the scan stopped early.

    Args:
        pdf_page_count: Pages in the PDF after the optional stamp strip.
        first_page: First materials page, 1-based.
        last_page: Last page written into the xlsx, 1-based.
        expected_cols: Column count of the materials table.
        stop_page: Page that ended the scan, 1-based, if the scan broke.
        stop_cols: Column count on ``stop_page``, if that page had a table.

    Returns:
        A WARN issue, or None when the tail is short or looks like an appendix.
    """
    if pdf_page_count <= 0 or last_page <= 0:
        return None
    tail = pdf_page_count - last_page
    if tail <= 0:
        return None
    if (
        stop_cols is not None
        and expected_cols > 0
        and stop_cols <= expected_cols * 0.6
    ):
        return None
    if tail <= max(3, int(pdf_page_count * 0.15)):
        return None
    stop = ""
    if stop_page is not None:
        cols = f", колонок {stop_cols}" if stop_cols is not None else ""
        stop = f", стоп на стр. {stop_page}{cols}"
    return PdfRfpIssue(
        "WARN",
        (
            f"в xlsx попали листы {first_page}–{last_page} из {pdf_page_count}: "
            f"после листа {last_page} осталось {tail}{stop}"
        ),
    )


def _rows_continue_materials(rows: tuple[tuple[str, ...], ...]) -> bool:
    for row in rows:
        padded = list(row)
        if _is_dropped_header_or_total(padded):
            continue
        first = padded[0].strip() if padded else ""
        return bool(first) and first[0].isdigit()
    return False


def _log_header_miss(rows: list[list[str]]) -> None:
    """Print the closest header candidate so a failed match is visible."""
    best_missing: list[str] | None = None
    best_preview = ""
    for index, row in enumerate(rows[:8]):
        columns = _match_header(row, prefer_lot_qty=True)
        missing = [field for field in REQUIRED_FIELDS if field not in columns]
        preview = " | ".join(
            (cell or "").replace("\n", "/")[:30] for cell in row[:8]
        )
        print(
            f"[pdf rfp] шапка строка {index + 1}: нет {', '.join(missing) or '—'} "
            f"— {preview}",
            flush=True,
        )
        if best_missing is None or len(missing) < len(best_missing):
            best_missing = missing
            best_preview = preview
    if best_missing is not None:
        print(
            f"[pdf rfp] ближайшая шапка, не хватает: {', '.join(best_missing)} "
            f"— {best_preview}",
            flush=True,
        )


def _is_dropped_header_or_total(row: list[str]) -> bool:
    if _is_total_row(row):
        return True
    for cell in row:
        folded = re.sub(r"\s+", " ", cell or "").strip()
        if not folded:
            continue
        if folded == _HEADER_NPP or _HEADER_NPP in folded:
            return True
        if _GROUP_RD in folded or "Закупка по" in folded:
            return True
    columns = _match_header(row, prefer_lot_qty=True)
    return all(field in columns for field in REQUIRED_FIELDS)


def _find_materials_header(
    rows: list[list[str]],
) -> tuple[list[str] | None, dict[str, int] | None, dict[str, int] | None]:
    for row in rows:
        right_cols = _match_header(row, prefer_lot_qty=True)
        if not all(field in right_cols for field in REQUIRED_FIELDS):
            continue
        left_cols = _match_header(row, prefer_lot_qty=False)
        return row, left_cols, right_cols
    return None, None, None


def _identity_column_indexes(
    left_cols: dict[str, int],
    right_cols: dict[str, int],
) -> set[int]:
    indexes: set[int] = set()
    for field in IDENTITY_FIELDS:
        if field in left_cols:
            indexes.add(left_cols[field])
        if field in right_cols:
            indexes.add(right_cols[field])
    return indexes


def _clean_data_row(row: list[str], identity_idx: set[int]) -> list[str]:
    cleaned: list[str] = []
    for idx, cell in enumerate(row):
        if idx in identity_idx:
            cleaned.append(clean_code_cell(cell))
        else:
            cleaned.append(clean_text_cell(cell))
    return cleaned


def _code_looks_truncated_agcc(code: str) -> bool:
    if "AGCC." not in code.upper():
        return False
    stripped = code.rstrip()
    return not (stripped and stripped[-1].isdigit())


def _as_rect(cell: Any) -> tuple[float, float, float, float] | None:
    if cell is None:
        return None
    if hasattr(cell, "x0"):
        return (float(cell.x0), float(cell.y0), float(cell.x1), float(cell.y1))
    if isinstance(cell, (tuple, list)) and len(cell) >= 4:
        try:
            x0, y0, x1, y1 = (float(cell[0]), float(cell[1]), float(cell[2]), float(cell[3]))
        except (TypeError, ValueError):
            return None
        if x1 > x0 and y1 > y0:
            return (x0, y0, x1, y1)
    return None


def _table_cell_rects(table: Any) -> list[tuple[float, float, float, float]]:
    rects: list[tuple[float, float, float, float]] = []
    cells = getattr(table, "cells", None) or []
    for cell in cells:
        rect = _as_rect(cell)
        if rect is not None:
            rects.append(rect)
    if rects:
        return rects
    rows = getattr(table, "rows", None) or []
    for row in rows:
        for cell in getattr(row, "cells", None) or []:
            rect = _as_rect(cell)
            if rect is not None:
                rects.append(rect)
    return rects


def _point_in_rect(
    x: float,
    y: float,
    rect: tuple[float, float, float, float],
) -> bool:
    x0, y0, x1, y1 = rect
    return x0 <= x <= x1 and y0 <= y <= y1


def _is_furniture_word(word: str) -> bool:
    token = (word or "").strip()
    if not token:
        return True
    if token.casefold() in _FURNITURE_WORDS:
        return True
    return bool(re.fullmatch(r"\d+", token))


def _scan_outside_words(
    doc: fitz.Document,
    page_tables: list[_MaterialsPage],
) -> list[_OutsideWord]:
    found: list[_OutsideWord] = []
    for materials_page in page_tables:
        page = doc[materials_page.page_index]
        rects = materials_page.cell_rects
        words = page.get_text("words") or []
        for item in words:
            if len(item) < 5:
                continue
            x0, y0, x1, y1, token = item[0], item[1], item[2], item[3], str(item[4])
            if _is_furniture_word(token):
                continue
            cx = (float(x0) + float(x1)) / 2.0
            cy = (float(y0) + float(y1)) / 2.0
            if any(_point_in_rect(cx, cy, rect) for rect in rects):
                continue
            found.append(
                _OutsideWord(
                    page=materials_page.page_index + 1,
                    word=token,
                    x=cx,
                    y=cy,
                )
            )
    return found


def _write_materials_xlsx(
    path: Path,
    *,
    header_row: list[str],
    data_rows: list[list[str]],
    left_cols: dict[str, int],
    right_cols: dict[str, int],
    col_count: int,
) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = EXPECTED_SHEET_NAME
    width = max(col_count, len(header_row))
    labels = _pad_row(header_row, width)
    group = [""] * width
    left_values = left_cols.get("VALUES")
    right_values = right_cols.get("VALUES")
    if left_values is not None and right_values is not None and left_values != right_values:
        group[left_values] = _GROUP_RD
        group[right_values] = _GROUP_LOT
    elif right_values is not None:
        group[right_values] = _GROUP_LOT
    elif left_values is not None:
        group[left_values] = _GROUP_LOT
    ws.append(group)
    ws.append([clean_text_cell(cell) for cell in labels])
    for row in data_rows:
        ws.append(_pad_row(row, width))
    wb.save(path)
    wb.close()


def _contract_check(xlsx_path: Path) -> list[PdfRfpIssue]:
    issues: list[PdfRfpIssue] = []
    wb = load_workbook(xlsx_path, read_only=True, data_only=True)
    try:
        if EXPECTED_SHEET_NAME not in wb.sheetnames:
            issues.append(
                PdfRfpIssue(
                    "ERROR",
                    f"в книге нет листа «{EXPECTED_SHEET_NAME}»",
                )
            )
            return issues
        ws = wb[EXPECTED_SHEET_NAME]
        rows = [list(row) for row in ws.iter_rows(values_only=True)]
    finally:
        wb.close()

    header_row: list[Any] | None = None
    header_idx = -1
    columns: dict[str, int] = {}
    for idx, row in enumerate(rows[:100]):
        columns = _match_header(row, prefer_lot_qty=True)
        if all(field in columns for field in REQUIRED_FIELDS):
            header_row = row
            header_idx = idx
            break
    if header_row is None:
        missing = list(REQUIRED_FIELDS)
        issues.append(
            PdfRfpIssue(
                "ERROR",
                f"не найдены обязательные столбцы: {format_field_list(missing)}",
            )
        )
        return issues
    missing_fields = [field for field in REQUIRED_FIELDS if field not in columns]
    if missing_fields:
        issues.append(
            PdfRfpIssue(
                "ERROR",
                f"не найдены обязательные столбцы: {format_field_list(missing_fields)}",
            )
        )
        return issues

    title_i = columns["DS_TITLE"]
    code_i = columns["CODE"]
    name_i = columns["NAME"]
    qty_i = columns["VALUES"]
    unit_i = columns["UNITS"]
    tag_i = columns.get("TAGS")

    def _cell(row: list[Any], index: int) -> str:
        if index >= len(row) or row[index] is None:
            return ""
        return str(row[index])

    for offset, row in enumerate(rows[header_idx + 1 :], start=header_idx + 2):
        if not any(str(cell).strip() for cell in row if cell is not None):
            continue
        if _is_total_row(["" if cell is None else str(cell) for cell in row]):
            continue
        excel_row = str(offset)
        title = _cell(row, title_i)
        code = _cell(row, code_i)
        name = _cell(row, name_i)
        qty = _cell(row, qty_i)
        unit = _cell(row, unit_i)
        tags = _cell(row, tag_i) if tag_i is not None else ""
        if not title.strip():
            issues.append(PdfRfpIssue("ERROR", "пустой титул", row_number=excel_row))
        if not code.strip():
            issues.append(PdfRfpIssue("ERROR", "пустой код", row_number=excel_row))
        if not name.strip():
            issues.append(
                PdfRfpIssue("ERROR", "пустое наименование", row_number=excel_row)
            )
        if not qty.strip():
            issues.append(
                PdfRfpIssue("ERROR", "пустое количество лота", row_number=excel_row)
            )
        elif _parse_decimal(qty) is None:
            issues.append(
                PdfRfpIssue(
                    "ERROR",
                    f"количество лота не является числом: {qty}",
                    row_number=excel_row,
                )
            )
        if not unit.strip():
            issues.append(
                PdfRfpIssue("ERROR", "пустая единица лота", row_number=excel_row)
            )
        if _has_space_or_newline(title):
            issues.append(
                PdfRfpIssue(
                    "ERROR",
                    "в титуле остались пробел или перевод строки",
                    row_number=excel_row,
                )
            )
        if _has_space_or_newline(code):
            issues.append(
                PdfRfpIssue(
                    "ERROR",
                    "в коде остались пробел или перевод строки",
                    row_number=excel_row,
                )
            )
        if _tag_tokens_have_space_or_newline(tags):
            issues.append(
                PdfRfpIssue(
                    "ERROR",
                    "в теге остались пробел или перевод строки",
                    row_number=excel_row,
                )
            )
        if _tags_text_contains_rfq(tags):
            issues.append(
                PdfRfpIssue(
                    "ERROR",
                    "в поле Tag недопустима подстрока RFQ",
                    row_number=excel_row,
                )
            )
    return issues


def _has_space_or_newline(text: str) -> bool:
    return any(ch.isspace() for ch in text)


def _tag_tokens_have_space_or_newline(tags: str) -> bool:
    """True when a single tag token still contains space/newline.

    ``; `` / ``, `` joiners from :func:`clean_code_cell` are not a leftover
    wrap; they split tokens first.
    """
    if not tags:
        return False
    for token in re.split(r"[;,]", tags):
        piece = token.strip()
        if piece and _has_space_or_newline(piece):
            return True
    return False


def _write_report(
    path: Path,
    *,
    ds_label: str,
    row_count: int,
    contract_errors: int,
    outside_words: int,
    issues: list[PdfRfpIssue],
    outside: list[_OutsideWord],
    cleaned_pdf_path: Path | None = None,
) -> None:
    lines = [
        f"Метка ДС: {ds_label or '(нет)'}",
        f"Строк данных: {row_count}",
        f"Ошибок контракта: {contract_errors}",
        f"Слов вне ячеек: {outside_words}",
        "",
        "Замечания:",
    ]
    if cleaned_pdf_path is not None:
        lines.insert(4, f"PDF без печати: {cleaned_pdf_path.name}")
    if issues:
        for item in issues:
            loc = []
            if item.page is not None:
                loc.append(f"стр.{item.page}")
            if item.row_number:
                loc.append(f"строка {item.row_number}")
            prefix = f" ({', '.join(loc)})" if loc else ""
            lines.append(f"[{item.level}] {item.message}{prefix}")
    else:
        lines.append("(нет)")
    lines.append("")
    shown = outside[:OUTSIDE_LIST_CAP]
    lines.append(
        f"Слова вне ячеек (показано {len(shown)} из {len(outside)}):"
    )
    for item in shown:
        lines.append(
            f"стр.{item.page} {item.word} x={item.x:.1f} y={item.y:.1f}"
        )
    omitted = len(outside) - len(shown)
    if omitted:
        lines.append(f"ещё {omitted} строк опущено")
    path.write_text("\n".join(lines) + "\n", encoding="utf-8")


_last_pdf_rfp_result: PdfRfpResult | None = None


def get_last_pdf_rfp_result() -> PdfRfpResult | None:
    """Return the last result from ``run_pdf_rfp_job``, if any."""
    return _last_pdf_rfp_result


@dataclass(frozen=True, slots=True)
class PdfRfpJobResult:
    """Duck-typed ``FunctionJobRunner`` result (no GUI import)."""

    success: bool
    message: str
    result_path: str | None = None


def run_pdf_rfp_job(
    pdf_path: str | Path,
    ds_label: str = "",
    strip_stamp: bool = False,
) -> PdfRfpJobResult:
    """Extract the PDF, remember the result for the GUI tab.

    ``ds_label`` is always forwarded to :func:`extract_rfp_pdf`, including an
    empty string (explicit override of title-page detection). Stamp folder is
    Stamp folder is a dated directory next to the PDF (``out_dir=None``).

    Args:
        pdf_path: Source RFP PDF.
        ds_label: Explicit DS prefix override.
        strip_stamp: Redact the Diadoc overlay before recognition.

    Returns:
        Job result. ``result_path`` is the written xlsx.

    Raises:
        RuntimeError: Propagated from :func:`extract_rfp_pdf` (no half xlsx).
    """
    global _last_pdf_rfp_result
    result = extract_rfp_pdf(
        Path(pdf_path), ds_label=ds_label, strip_stamp=strip_stamp
    )
    _last_pdf_rfp_result = result
    print(
        f"ДС: {result.ds_label or '(нет)'}\n"
        f"Строк: {result.row_count}\n"
        f"Ошибок контракта: {result.contract_errors}\n"
        f"Слов вне ячеек: {result.outside_words}\n"
        f"xlsx: {result.xlsx_path}\n"
        f"отчёт: {result.report_path}\n"
        f"pdf без печати: {result.cleaned_pdf_path or '(нет)'}",
        flush=True,
    )
    summary = (
        f"Строк {result.row_count}, ошибок контракта {result.contract_errors}, "
        f"слов вне ячеек {result.outside_words}"
    )
    return PdfRfpJobResult(
        success=True,
        message=summary,
        result_path=str(result.xlsx_path),
    )


__all__ = (
    "PdfRfpIssue",
    "PdfRfpJobResult",
    "PdfRfpResult",
    "clean_code_cell",
    "clean_text_cell",
        "extract_rfp_pdf",
        "materials_coverage_issue",
    "strip_diadoc_stamps",
    "get_last_pdf_rfp_result",
    "parse_ds_label_from_title_text",
    "pdf_rfp_stamp_dir",
    "preview_ds_label",
    "run_pdf_rfp_job",
)
