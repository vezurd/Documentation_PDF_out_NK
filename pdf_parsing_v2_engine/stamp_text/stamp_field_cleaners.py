"""Domain stamp field cleaners: titles, BBB sheet lines, revisions (non-primitives).

Contract and strict/relaxed patterns: see ``stamp_text/README.md``.
"""

from __future__ import annotations

import re

from pdf_parsing_v2_engine.stamp_text.context import StampTextContext
from pdf_parsing_v2_engine.stamp_text.junk import JunkMode, junk_clean_lines
from pdf_parsing_v2_engine.stamp_text.outcome import CleanOutcome, ParseWarning, strict_clean_outcome
from pdf_parsing_v2_engine.stamp_text.primitives import *
from utils.file_name_converts import ProjectFileName

# --- Titles / document lines -------------------------------------------------

# Heuristic: hard slice (AGCC → first ``\\n``) may include the next stamp line;
# if the cleaned title looks too long or contains sheet vocabulary, use relaxed
# boundaries and emit ``DOC_TITLE_RELAXED``.
_DOC_TITLE_HARD_MAX_CHARS = 100
_DOC_TITLE_HARD_MAX_WORDS = 18

_DOC_TITLE_RELAXED_MARKERS_LOWER = (" лист", " листов", " из ", " л.")


def _doc_title_join(blob: str) -> str:
    res = junk_clean_lines(blob, JunkMode.DEFAULT)
    return " ".join(res).strip() if res else ""


def _doc_title_strict_raw(text: str) -> str:
    """Legacy slice: AGCC line only when both AGCC and a following newline exist."""
    output_text = text
    start_index = text.find("AGCC")
    end_index = text.find("\n", start_index + 1)
    if start_index != -1 and end_index != -1:
        output_text = text[start_index:end_index]
    return output_text


def _doc_title_hard_suspicious(cleaned: str, raw_slice: str) -> bool:
    if len(cleaned) > _DOC_TITLE_HARD_MAX_CHARS:
        return True
    if len(cleaned.split()) > _DOC_TITLE_HARD_MAX_WORDS:
        return True
    low = cleaned.lower()
    if "лист" in low and len(cleaned) > 30:
        return True
    return False


def _doc_title_relaxed_raw(text: str) -> str:
    """AGCC to the earliest of newline, tab, or common sheet-line markers."""
    start = text.find("AGCC")
    if start == -1:
        return text
    candidates: list[int] = []
    ne = text.find("\n", start + 1)
    if ne != -1:
        candidates.append(ne)
    te = text.find("\t", start + 1)
    if te != -1:
        candidates.append(te)
    lower = text.lower()
    for needle in _DOC_TITLE_RELAXED_MARKERS_LOWER:
        p = lower.find(needle, start + 1)
        if p != -1:
            candidates.append(p)
    if not candidates:
        return text[start:]
    return text[start : min(candidates)]


def _doc_title_fallback_first_junk_line(text: str) -> str:
    """When strict and relaxed share the same raw slice, keep only first junk-surviving line."""
    start = text.find("AGCC")
    if start == -1:
        return _doc_title_join(text)
    lines = junk_clean_lines(text[start:], JunkMode.DEFAULT)
    return lines[0].strip() if lines else ""


def pipeline_doc_title(text: str, ctx: StampTextContext) -> CleanOutcome:
    """Field clean ``doc_title``: strict AGCC→newline slice; relaxed fallback + warning if suspicious."""
    del ctx
    strict_raw = _doc_title_strict_raw(text)
    cleaned_strict = _doc_title_join(strict_raw)
    if not _doc_title_hard_suspicious(cleaned_strict, strict_raw):
        return CleanOutcome(
            value=cleaned_strict,
            raw_input=text,
            warnings=[],
            via="doc_title",
            clean_tier="strict",
        )

    relaxed_raw = _doc_title_relaxed_raw(text)
    if relaxed_raw == strict_raw:
        cleaned = _doc_title_fallback_first_junk_line(text)
    else:
        cleaned = _doc_title_join(relaxed_raw)

    warn = ParseWarning(
        code="DOC_TITLE_RELAXED",
        message=(
            "Strict AGCC→newline slice looked over-inclusive; used relaxed boundaries "
            "or first junk-cleaned line."
        ),
        tier="relaxed",
    )
    return CleanOutcome(
        value=cleaned,
        raw_input=text,
        warnings=[warn],
        via="doc_title",
        clean_tier="relaxed",
    )


def _facility_name(text: str, ctx: StampTextContext) -> str:
    """Field clean ``facility_name`` (body)."""
    del ctx
    output_text = text
    if "AGCC" in text:
        start_index = text.find("AGCC")
        end_index = text.find("\n", start_index + 1)
        output_text = text[end_index + 1 : len(text)] if end_index != -1 else text

    s_list = junk_clean_lines(output_text, JunkMode.WITH_SHEET_LABELS)
    return " ".join(s_list).strip()


def pipeline_facility_name(text: str, ctx: StampTextContext) -> CleanOutcome:
    """Field clean ``facility_name``."""
    return strict_clean_outcome(text, _facility_name(text, ctx), via="facility_name")


def _document_name(text: str, ctx: StampTextContext) -> str:
    """Field clean ``document_name`` (body; dates removed, no junk pass)."""
    del ctx
    input_text = remove_dates(text)
    return remove_newlines_with_space(input_text)


def pipeline_document_name(text: str, ctx: StampTextContext) -> CleanOutcome:
    """Field clean ``document_name``."""
    return strict_clean_outcome(text, _document_name(text, ctx), via="document_name")


def _documentation_type(text: str, ctx: StampTextContext) -> str:
    """Field clean ``documentation_type`` (body)."""
    del ctx
    out = remove_newlines(text)
    cleaned = junk_clean_lines(out, JunkMode.DEFAULT)
    return " ".join(cleaned).strip() if cleaned else ""


def pipeline_documentation_type(text: str, ctx: StampTextContext) -> CleanOutcome:
    """Field clean ``documentation_type``."""
    return strict_clean_outcome(text, _documentation_type(text, ctx), via="documentation_type")


def _total_number_of_sheets(text: str, ctx: StampTextContext) -> str:
    """Field clean ``total_number_of_sheets`` (body)."""
    del ctx
    out = remove_newlines(text)
    cleaned = junk_clean_lines(out, JunkMode.WITH_SHEET_LABELS)
    return " ".join(cleaned).strip() if cleaned else ""


def pipeline_total_number_of_sheets(text: str, ctx: StampTextContext) -> CleanOutcome:
    """Field clean ``total_number_of_sheets``."""
    return strict_clean_outcome(text, _total_number_of_sheets(text, ctx), via="total_number_of_sheets")


def _unit_title(text: str, ctx: StampTextContext) -> str:
    """Field clean ``unit_title`` (body)."""
    del ctx
    return text.replace("\n", " ")


def pipeline_unit_title(text: str, ctx: StampTextContext) -> CleanOutcome:
    """Field clean ``unit_title``."""
    return strict_clean_outcome(text, _unit_title(text, ctx), via="unit_title")


# --- BBB sheet numbering (6.x) -----------------------------------------------

# Document types that use BBB-style sheet blocks (legacy ``list_of_BBB``).
LIST_OF_BBB: frozenset[str] = frozenset({"BOM", "BOE", "BOQ", "MTO"})

# BBB page 2+: ``… л. X из Y листов …`` (same line as code / title is OK).
_BBB_LY_IZ_LISTOV_RE = re.compile(
    r"л\.\s*(\d+)\s*из\s*(\d+)\s*листов",
    re.IGNORECASE,
)


def _parse_ly_iz_listov_numbers(blob: str) -> tuple[str, str] | None:
    """Parse ``л. X из Y листов``; returns ``(X, Y)`` digit groups or ``None``."""
    m = _BBB_LY_IZ_LISTOV_RE.search(blob)
    if not m:
        return None
    return (m.group(1), m.group(2))


def _extract_sheet_6_1_regex_matches(text: str) -> list[str]:
    """Legacy ``get_6_1_att_reg_exp``: prefer ``d+.d+`` tokens, else short digits."""
    reg_long = r"\d{1,2}.\d{1,2}"
    reg_short = r"\d{1,2}"
    match_long = re.findall(reg_long, text)
    if match_long:
        return match_long
    match_short = re.findall(reg_short, text)
    if match_short:
        return match_short
    return []


def _sheet_number_6_1(text: str, ctx: StampTextContext) -> str:
    """Current sheet index (field clean ``sheet_number_6_1``, body)."""
    doc_type = ctx.doc_type
    page_num = ctx.page_num
    if doc_type not in LIST_OF_BBB or page_num < 2:
        res = _extract_sheet_6_1_regex_matches(text)
        return " ".join(str(x) for x in res) if res else ""

    blob = " ".join(junk_clean_lines(text, JunkMode.DEFAULT))
    ly_iz = _parse_ly_iz_listov_numbers(blob)
    if ly_iz:
        return ly_iz[0]

    
    return ""


def pipeline_sheet_number_6_1(text: str, ctx: StampTextContext) -> CleanOutcome:
    """Field clean ``sheet_number_6_1``."""
    return strict_clean_outcome(text, _sheet_number_6_1(text, ctx), via="sheet_number_6_1")


def _sheet_number_6_2(text: str, ctx: StampTextContext) -> str:
    """Total sheets (field clean ``sheet_number_6_2``, body)."""
    doc_type = ctx.doc_type
    page_num = ctx.page_num
    if doc_type not in LIST_OF_BBB or page_num < 2:
        lines = junk_clean_lines(text, JunkMode.DEFAULT)
        names = ("на", "На")
        indexes: list[int] = []
        for i, line in enumerate(lines):
            for n in names:
                if n in line:
                    indexes.append(i)
                    break
        res_lines = [lines[i] for i in range(len(lines)) if i in indexes]
        sting_f = str(res_lines).lower()
        start_index = sting_f.find("на")
        end_index = sting_f.find("л.", start_index + 1)
        if start_index < 0 or end_index < 0:
            return ""
        res = sting_f[start_index + 2 : end_index].strip()
    else:
        blob = " ".join(junk_clean_lines(text, JunkMode.DEFAULT))
        ly_iz = _parse_ly_iz_listov_numbers(blob)
        if ly_iz:
            return ly_iz[1]

        lines = junk_clean_lines(text, JunkMode.DEFAULT)
        names = ("AGCC",)
        indexes: list[int] = []
        for i, line in enumerate(lines):
            for n in names:
                if n in line:
                    indexes.append(i)
                    break
        res_lines: list[str] = []
        for i, line in enumerate(lines):
            if i not in indexes:
                res_lines.append(line)
        sting_f = str(res_lines)
        start_index = sting_f.find("из")
        end_index = sting_f.find("листов", start_index + 1)
        if start_index < 0 or end_index < 0:
            return ""
        res = sting_f[start_index + 2 : end_index].strip()

    return res.replace("'", "")


def pipeline_sheet_number_6_2(text: str, ctx: StampTextContext) -> CleanOutcome:
    """Field clean ``sheet_number_6_2``."""
    return strict_clean_outcome(text, _sheet_number_6_2(text, ctx), via="sheet_number_6_2")


# --- Revisions and stamp metadata ------------------------------------------


def _revision(text: str, ctx: StampTextContext) -> str:
    """Field clean ``revision`` (body)."""
    del ctx
    out_lines = junk_clean_lines(text)
    if not out_lines:
        return ""
    first = out_lines[0]
    return first.split(" ")[0]


def pipeline_revision(text: str, ctx: StampTextContext) -> CleanOutcome:
    """Field clean ``revision``."""
    return strict_clean_outcome(text, _revision(text, ctx), via="revision")


################################################################################

def pipeline_file_name_stamp(text: str, ctx: StampTextContext, dbg: bool = False) -> CleanOutcome:
    """Field clean ``file_name_stamp``."""
    #Headers for CleanOutcome
    warnings: list[ParseWarning] = []
    raw_input: str = text
    via: str = "file_name_stamp"
    clean_tier: str = "strict"
    
    text = v2_normalize_line_text(text)
    if dbg:
        print(f"pipeline_file_name_stamp: text = {text}")
    # Пытаемся извелечь правильное имя файла из текста штампа с расширением файла    
    value_strict = ProjectFileName.get_file_name_with_file_extension(text)
    if dbg:
        print(f"pipeline_file_name_stamp: value_strict = {value_strict}")
    if value_strict and text == value_strict[0]:
        return strict_clean_outcome(text, value_strict[0], via=via)
    else:
        # Проверяем на задвоение надписи
        double_name = ProjectFileName.get_file_name_with_file_extension(text)
        if dbg:
            print(f"pipeline_file_name_stamp: double_name = {double_name}")
        if double_name and len(double_name) > 1:
            warn = ParseWarning(code="FILE_NAME_STAMP_RELAXED", message="Задвоение надписи", tier="relaxed")
            return CleanOutcome(
                value=double_name[0], 
                raw_input=text, 
                warnings=[warn], 
                via=via, 
                clean_tier="relaxed")
        #Проверяем есть ли вхождение полного имени файла с расширением но с мусором в text
        full_name = ProjectFileName.get_file_name_with_file_extension(text)
        if full_name:
            warn = ParseWarning(
                code="FILE_NAME_STAMP_RELAXED", 
                message="Имя фала найдено но есть лишние символы", 
                tier="relaxed")
            return CleanOutcome(
                value=full_name[0], 
                raw_input=text, 
                warnings=[warn], 
                via="file_name_stamp", 
                clean_tier="relaxed")

        # Проверяем на наличие имени документа но без расширения, т.к. с раширинем уже откинули
        document_name = ProjectFileName.get_file_name_without_file_extension(text)
        if document_name and text == document_name[0]:
            warn = ParseWarning(code="FILE_NAME_STAMP_RELAXED", message="Отсутствие расширения файла", tier="relaxed")
            return CleanOutcome(
                value=text, 
                raw_input=text, 
                warnings=[warn], 
                via="file_name_stamp", 
                clean_tier="relaxed")
        # Если ничего не найдено, то возвращаем пустую строку
        warn = ParseWarning(code="FILE_NAME_STAMP_RELAXED", message="Не удалось извлечь имя файла", tier="relaxed")
        return CleanOutcome(
                value="", 
                raw_input=text, 
                warnings=[warn], 
                via="file_name_stamp", 
                clean_tier="relaxed")
        
###################################################



def _page_format(text: str, ctx: StampTextContext) -> str:
    """Field clean ``page_format`` (body)."""
    del ctx
    out_lines = junk_clean_lines(text)
    if not out_lines:
        return ""
    for el in out_lines:
        if "Формат" in el:
            return el
    return " ".join(out_lines)


def pipeline_page_format(text: str, ctx: StampTextContext) -> CleanOutcome:
    """Field clean ``page_format``."""
    return strict_clean_outcome(text, _page_format(text, ctx), via="page_format")


def _current_revision_18_1(text: str, ctx: StampTextContext) -> str:
    """Field clean ``current_revision_18_1`` (body)."""
    del ctx
    if "Кабели одного типа" in text:
        return ""
    lines = junk_clean_lines(text)
    res: list[list[str]] = []
    for line in lines:
        res.append(line.split())
    flat = list_flatter(res)
    names = ("IFC", "IFR", "IFU")
    indexes: list[int] = []
    for i, item in enumerate(flat):
        for n in names:
            if n in item or len(item.split(".")) == 3:
                indexes.append(i)
                break
    kept: list[str] = []
    for i, item in enumerate(flat):
        if i not in indexes:
            kept.append(item)
    if not kept:
        return ""
    return kept[0]


def pipeline_current_revision_18_1(text: str, ctx: StampTextContext) -> CleanOutcome:
    """Field clean ``current_revision_18_1``."""
    return strict_clean_outcome(text, _current_revision_18_1(text, ctx), via="current_revision_18_1")


def _revision_date_18_2(text: str, ctx: StampTextContext) -> str:
    """Field clean ``revision_date_18_2`` (18-column date, body)."""
    del ctx
    return _list_18_2_inner(text, flag=18)


def pipeline_revision_date_18_2(text: str, ctx: StampTextContext) -> CleanOutcome:
    """Field clean ``revision_date_18_2``."""
    return strict_clean_outcome(text, _revision_date_18_2(text, ctx), via="revision_date_18_2")


def _purpose_of_issue_18_3(text: str, ctx: StampTextContext) -> str:
    """Field clean ``purpose_of_issue_18_3`` (body)."""
    del ctx
    lines = junk_clean_lines(text)
    a = str(lines)
    a = a.strip()
    a = re.sub(r"[\][\']", "", a)

    data_junk = re.sub(r"[/\\.]", ".", a)
    m = re.search(r"(\d+.*?\d+.*?\d+)", data_junk)
    data_junk = m.group(1) if m else None

    if data_junk is not None and data_junk in a:
        stat_index = a.find(data_junk)
        result = a[stat_index + len(data_junk) : len(a)].strip()
    else:
        result = a

    items = result.split(",")
    return items[0].strip()


def pipeline_purpose_of_issue_18_3(text: str, ctx: StampTextContext) -> CleanOutcome:
    """Field clean ``purpose_of_issue_18_3``."""
    return strict_clean_outcome(text, _purpose_of_issue_18_3(text, ctx), via="purpose_of_issue_18_3")


def _signatures_date_10(text: str, ctx: StampTextContext) -> str:
    """Field clean ``signatures_date_10`` (column 10 in 18_2 table, body)."""
    del ctx
    return _list_18_2_inner(text, flag=10)


def pipeline_signatures_date_10(text: str, ctx: StampTextContext) -> CleanOutcome:
    """Field clean ``signatures_date_10``."""
    return strict_clean_outcome(text, _signatures_date_10(text, ctx), via="signatures_date_10")


def _list_18_2_inner(input_text: str, flag: int) -> str:
    if "Кабели одного типа" in input_text:
        return ""
    try:
        lines = junk_clean_lines(input_text)
        a = str(lines)
        a = a.replace(" ", "")
        if flag == 10:
            m = re.search(r"(\d\d.\d\d.\d\d)", a)
            if not m:
                return ""
            return re.sub(r"[/\\.]", ".", m.group(1))
        if flag == 18:
            m = re.search(r"(\d\d.\d\d.\d\d\d\d)", a)
            if not m:
                return ""
            return re.sub(r"[/\\.]", ".", m.group(1))
    except Exception:
        return ""
    return ""
