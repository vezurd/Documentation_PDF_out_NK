"""OD table parsing — v2 port of pdf_parsing.OD_tab_parsing.

Extracts document metadata from the "General Data" (ОД) PDF table.
"""

from __future__ import annotations

import re
from os import listdir
from os.path import isfile, join
from typing import Any

from utils import string_parsing
from utils.file_name_converts import ProjectFileName
from utils.formats_A import get_a4_count_in_format

from pdf_parsing_v2_engine.document import V2Document
from pdf_parsing_v2_engine.models import StampTemplate
from pdf_parsing_v2_od.od_debug_trace import append_od_debug_trace
from pdf_parsing_v2_od.od_manifest_geometry import (
    OdManifestGeometryRow,
    extract_od_manifest_geometry_rows,
)
from pdf_parsing_v2_od.od_parse_warnings import (
    PAGE_FORMAT_UNKNOWN_KEY_CODE,
    PAGE_FORMAT_WHITESPACE_WARNING_CODE,
    append_od_parse_warning,
)
from pdf_parsing_v2_od.od_raw_reader import od_read_raw

_TITLE_OSNOVNOI = "ВЕДОМОСТЬ ДОКУМЕНТОВ ОСНОВНОГО КОМПЛЕКТА РАБОЧИХ ЧЕРТЕЖЕЙ"
_TITLE_PRILAGAEMIE = "ВЕДОМОСТЬ ПРИЛАГАЕМЫХ ДОКУМЕНТОВ"
_TAB_SEPARATOR = "None"

DOC_OSNOVNOI = "Основной комплект"
DOC_PRILAGAEMIE = "Прилагаемые документы"

_MSG_OD_PATH_EMPTY = "Не указан путь к PDF ведомости (ОД)."


def error_message_for_invalid_od_parse_result(pdf_full_path: str) -> str:
    """Human-readable reason when ``od_table_parsing_f`` did not return a list."""
    p = str(pdf_full_path).strip()
    if not p:
        return _MSG_OD_PATH_EMPTY
    if not isfile(p):
        return f"Файл ОД не найден на диске: {p}"
    return "Разбор ОД не выполнен: неожиданный результат парсера (ожидался список строк)."


def find_od_pdf_path(pdf_path: str) -> str:
    """Return full path to the first PDF in *pdf_path* whose doc type is OD, else empty string.

    Args:
        pdf_path: Folder containing PDF files.

    Returns:
        Path string, or empty if no OD PDF found.
    """
    onlyfiles = [f for f in listdir(pdf_path) if isfile(join(pdf_path, f))]
    for next_file in sorted(onlyfiles):
        if not next_file.lower().endswith(".pdf"):
            continue
        try:
            if string_parsing.getDocTypeFromFile(next_file) != "OD":
                continue
        except Exception:
            continue
        return join(pdf_path, next_file)
    return ""


def build_od_table_payload(
    proj_od_list: list[DocOdAttributes],
    od_pdf_path: str,
    *,
    status: str = "ok",
    error: str | None = None,
    warnings: list[str] | None = None,
) -> dict[str, Any]:
    """Build JSON-safe table payload for GUI / pipeline summary (matches stdout table).

    Args:
        proj_od_list: Parsed OD rows.
        od_pdf_path: Path to the OD PDF that was read (may be empty if none).
        status: ``ok`` | ``no_od_file`` | ``no_rows`` | ``error``.
        error: Set when ``status`` is ``error``.
        warnings: Optional diagnostic lines (e.g. pdfplumber / header heuristics).
            May include ``3001:`` / ``3002:`` lines from ``od_parse_warnings`` (см. НК).

    Returns:
        Dict with ``status``, ``error``, ``warnings``, ``od_pdf_path``, ``columns``, ``rows``.
    """
    total_page_count = sum(
        el.doc_page_count
        for el in proj_od_list
        if el.doc_range == DOC_OSNOVNOI
        and isinstance(el.doc_page_count, int)
        and "V" not in str(el.Document_Revision)
    )
    columns = [
        "Doc_Title",
        "Page_Format",
        f"Page_Count:{total_page_count}",
        "Document_name",
        "Document_Revision",
        "Комплект",
    ]
    rows: list[list[str]] = []
    for el in proj_od_list:
        rows.append(
            [
                str(el.Doc_Title),
                str(el.Page_Format),
                str(el.doc_page_count),
                str(el.Document_name),
                str(el.Document_Revision),
                str(el.doc_range),
            ]
        )
    warn_list = [str(w).strip() for w in (warnings or []) if str(w).strip()]
    return {
        "status": status,
        "error": error,
        "warnings": warn_list,
        "od_pdf_path": od_pdf_path,
        "columns": columns,
        "rows": rows,
    }


class DocOdAttributes:
    """One row from the OD table (document entry)."""

    def __init__(
        self,
        doc_title: str,
        page_format: Any,
        document_name: str,
        document_revision: str,
        doc_range: str,
    ) -> None:
        self.Doc_Title = doc_title
        self.Page_Format = page_format
        self.Document_name = document_name
        self.Document_Revision = document_revision
        self.doc_page_count: int = -1
        self.doc_page_count_a4: int = -1
        self.doc_range = doc_range
        self.doc_OD_style_file_name = self.Doc_Title


def _string_to_att_list(input_string: str) -> list[str]:
    input_string = re.sub(r"[|]", "\n", input_string)
    input_string = input_string.replace(_TAB_SEPARATOR, "\n")
    return list(filter(None, input_string.split("\n")))


def _page_format_token_parts(page_format: Any) -> tuple[list[str], bool]:
    """Split ``Page_Format`` into tokens for ``get_a4_count_in_format``.

    Commas separate groups; within a group, one or more spaces / newlines / tabs
    separate formats (same as an implicit comma). Returns ``(tokens, inferred_ws)`` where
    ``inferred_ws`` is True when whitespace was used to split inside a comma segment.
    """
    if page_format is None or page_format == -1:
        return [], False
    s = str(page_format)
    s = s.replace("\r\n", " ").replace("\n", " ").replace("\r", " ").replace("\t", " ")
    s = s.replace("А", "A")
    segments = [seg.strip() for seg in s.split(",") if seg.strip()]
    parts: list[str] = []
    inferred = False
    for seg in segments:
        sub = [t for t in re.split(r"\s+", seg.strip()) if t]
        if len(sub) > 1:
            inferred = True
        parts.extend(sub)
    return parts, inferred


def _get_page_count(
    page_format: Any,
    *,
    debug_trace_out: list[dict[str, str]] | None = None,
    warnings_out: list[str] | None = None,
    row_context: str = "",
) -> tuple[int | str, int | str]:
    if page_format is None:
        return "", ""
    if page_format == -1:
        return -1, -1
    parts, inferred_ws = _page_format_token_parts(page_format)
    fmt_joined = ",".join(parts)
    ctx = f" ({row_context})" if row_context else ""
    append_od_debug_trace(
        debug_trace_out,
        function="od_parsing._get_page_count",
        stage=f"input{ctx}",
        body=(
            f"page_format repr: {page_format!r}\n"
            f"token parts (comma + whitespace between formats): {parts!r}\n"
            f"inferred_whitespace_separator: {inferred_ws!r}"
        ),
    )
    if inferred_ws:
        append_od_parse_warning(
            warnings_out,
            code=PAGE_FORMAT_WHITESPACE_WARNING_CODE,
            message=(
                "Между обозначениями формата встречены пробелы или переносы строк без запятой; "
                f"выполнен разбор как перечисление через запятую: {fmt_joined!r}{ctx}."
            ),
        )
    count = 0
    count_a4 = 0
    for part in parts:
        try:
            if part[0] == "A":
                curr_count = 1
                curr_count_a4 = get_a4_count_in_format(part)
            else:
                idx = part.find("A")
                if idx > 0:
                    curr_count = int(part[:idx])
                    curr_format = part[idx:]
                else:
                    curr_count = 1
                    curr_format = part
                curr_count_a4 = curr_count * get_a4_count_in_format(curr_format)
            count += curr_count
            count_a4 += curr_count_a4
        except KeyError as e:
            append_od_debug_trace(
                debug_trace_out,
                function="od_parsing._get_page_count",
                stage=f"key_error{ctx}",
                body=(
                    f"get_a4_count_in_format: unknown format key {e!r}.\n"
                    f"part being looked up: {part!r}\n"
                    f"(Often caused by merged cells: e.g. 'A1x4A2x4' without a comma.)"
                ),
            )
            append_od_parse_warning(
                warnings_out,
                code=PAGE_FORMAT_UNKNOWN_KEY_CODE,
                message=(
                    f"Неизвестный ключ формата листа {e!s} для части {part!r} "
                    f"(исходное поле Формат: {page_format!r}){ctx}."
                ),
            )
            return -1, -1
        except (IndexError, ValueError) as e:
            print(f"ERROR page format <{page_format}>: {e}")
            append_od_debug_trace(
                debug_trace_out,
                function="od_parsing._get_page_count",
                stage=f"parse_error{ctx}",
                body=f"{type(e).__name__}: {e!r}; part={part!r}",
            )
    append_od_debug_trace(
        debug_trace_out,
        function="od_parsing._get_page_count",
        stage=f"result{ctx}",
        body=f"doc_page_count={count!r}, doc_page_count_a4={count_a4!r}",
    )
    return count, count_a4


def od_table_parsing_f(
    *,
    pdf_full_path: str,
    save_to_file: int = 1,
    debug_print_out_list: int = 1,
    warnings_out: list[str] | None = None,
    cfg: dict[str, Any] | None = None,
    templates: list[StampTemplate] | None = None,
    curr_proj: list[V2Document] | None = None,
    manifest_geometry_out: list[OdManifestGeometryRow] | None = None,
    debug_trace_out: list[dict[str, str]] | None = None,
) -> list[DocOdAttributes] | int:
    """Parse OD PDF and return list of document entries.

    Args:
        pdf_full_path: Absolute or relative path to the OD PDF file.
        save_to_file: unused (kept for signature compat).
        debug_print_out_list: print summary table to stdout.
        warnings_out: When set, ``od_read_raw`` and row heuristics append diagnostic strings.
        cfg: v2 config — with non-empty ``templates``, enables manifest clip for pdfplumber.
        templates: Loaded stamp templates; when provided with ``cfg``, pdfplumber uses
            engine-aligned crop (frame minus stamp) per page where geometry succeeds.
        curr_proj: Optional extraction cache — reuse ``V2PageResult`` bbox when paths match OD file.
        manifest_geometry_out: When passed (typically empty list), extended with per-page
            geometry rows so postprocess can run ``find_tables`` debug without a second extract.
        debug_trace_out: When passed (e.g. empty list), append structured debug blocks with
            ``function`` / ``stage`` / ``body`` (GUI standalone OD read).

    Returns:
        List of ``DocOdAttributes`` or ``0`` if the path is missing or the file
        is not found on disk.
    """
    resolved = str(pdf_full_path).strip()
    if not resolved:
        print(_MSG_OD_PATH_EMPTY)
        return 0
    if not isfile(resolved):
        print(f"Файл ОД не найден: {resolved}")
        return 0
    pdf_full_path = resolved

    append_od_debug_trace(
        debug_trace_out,
        function="od_parsing.od_table_parsing_f",
        stage="entry",
        body=(
            f"pdf_full_path={pdf_full_path!r}\n"
            f"cfg={'set' if cfg is not None else 'None'}, "
            f"templates_count={len(templates or [])}\n"
            f"curr_proj={'set' if curr_proj else 'None'}\n"
            f"manifest_geometry_out={'list' if manifest_geometry_out is not None else 'None'}"
        ),
    )

    geom_rows: list[OdManifestGeometryRow] | None = None
    clip_rects: list[tuple[float, float, float, float] | None] | None = None
    use_engine_clip = bool(cfg is not None and templates)
    if use_engine_clip:
        geom_rows = extract_od_manifest_geometry_rows(
            pdf_full_path,
            templates or [],
            cfg,
            curr_proj=curr_proj,
            warnings_out=warnings_out,
        )
        clip_rects = [r.clip for r in geom_rows]
        if manifest_geometry_out is not None:
            manifest_geometry_out.clear()
            manifest_geometry_out.extend(geom_rows)
        geom_lines = []
        for r in geom_rows:
            geom_lines.append(
                f"  page {r.page_num}: source={r.geometry_source!r} "
                f"template={r.template_name!r} clip={r.clip!r} "
                f"skip_reason={r.skip_reason!r}"
            )
        append_od_debug_trace(
            debug_trace_out,
            function="od_parsing.od_table_parsing_f",
            stage="after_extract_od_manifest_geometry_rows",
            body="engine clip per page:\n" + "\n".join(geom_lines),
        )
    else:
        append_od_debug_trace(
            debug_trace_out,
            function="od_parsing.od_table_parsing_f",
            stage="geometry_skipped",
            body="No cfg or empty templates — pdfplumber uses full page (legacy).",
        )

    input_list = od_read_raw(
        pdf_full_path,
        warnings_out=warnings_out,
        page_clip_rects=clip_rects,
        debug_trace_out=debug_trace_out,
    )

    proj_od_list: list[DocOdAttributes] = []
    page_idx = 0
    for page_content in input_list:
        page_idx += 1
        page_text = "".join(page_content)
        append_od_debug_trace(
            debug_trace_out,
            function="od_parsing.od_table_parsing_f",
            stage=f"merged_text_page_{page_idx}",
            body=(
                f"len(page_text)={len(page_text)}\n"
                f"--- snippet (first 4000 chars) ---\n"
                f"{page_text[:4000]}"
            ),
        )

        if _TITLE_OSNOVNOI in page_text or (
            "Обозначение" in page_text
            and "Формат" in page_text
            and "Наименование" in page_text
            and "Примечание" in page_text
        ):
            for line in page_text.split("\n"):
                string_arr = _string_to_att_list(line)
                if not string_arr:
                    break
                file_name = string_parsing.string_remove_spaces(string_arr[0])
                if ProjectFileName.find_stem_matches(file_name):
                    append_od_debug_trace(
                        debug_trace_out,
                        function="od_parsing.od_table_parsing_f",
                        stage=f"row_osnovnoi_page_{page_idx}",
                        body=(
                            "Table row after pdfplumber → pipe/newline split (_string_to_att_list).\n"
                            f"raw_line repr: {line!r}\n"
                            f"column count: {len(string_arr)}\n"
                            + "\n".join(
                                f"  [{i}] {string_arr[i]!r}"
                                for i in range(len(string_arr))
                            )
                        ),
                    )
                    try:
                        proj_od_list.append(
                            DocOdAttributes(
                                string_parsing.string_remove_spaces(string_arr[0]),
                                string_arr[1],
                                string_arr[2],
                                string_arr[3],
                                DOC_OSNOVNOI,
                            )
                        )
                    except IndexError:
                        continue

        if _TITLE_PRILAGAEMIE in page_text:
            for line in page_text.split("\n"):
                string_arr = _string_to_att_list(line)
                if not string_arr:
                    break
                file_name = string_parsing.string_remove_spaces(string_arr[0])
                if ProjectFileName.find_stem_matches(file_name):
                    sar = [repr(string_arr[i]) for i in range(len(string_arr))]
                    append_od_debug_trace(
                        debug_trace_out,
                        function="od_parsing.od_table_parsing_f",
                        stage=f"row_prilagaemie_page_{page_idx}",
                        body=(
                            f"raw_line repr: {line!r}\n"
                            f"string_arr reprs: {sar!r}"
                        ),
                    )
                    name = string_arr[1] if len(string_arr) > 1 else "Наименование не найдено"
                    revision = string_arr[2] if len(string_arr) > 2 else "ревизия не найдена"
                    proj_od_list.append(
                        DocOdAttributes(
                            string_parsing.string_remove_spaces(string_arr[0]),
                            -1,
                            name,
                            revision,
                            DOC_PRILAGAEMIE,
                        )
                    )

    for el in proj_od_list:
        label = f"{el.Doc_Title!r} / {el.doc_range!r}"
        el.doc_page_count, el.doc_page_count_a4 = _get_page_count(
            el.Page_Format,
            debug_trace_out=debug_trace_out,
            warnings_out=warnings_out,
            row_context=label,
        )

    if not proj_od_list and warnings_out is not None:
        warnings_out.append(
            "Не распознаны строки ведомости (нет совпадений с шаблоном имён или "
            "не найдены ожидаемые заголовки/колонки)."
        )

    if debug_print_out_list:
        _print_od_summary(proj_od_list, pdf_full_path)

    return proj_od_list


def _print_od_summary(
    proj_od_list: list[DocOdAttributes], pdf_path: str
) -> None:
    total_page_count = sum(
        el.doc_page_count
        for el in proj_od_list
        if el.doc_range == DOC_OSNOVNOI
        and isinstance(el.doc_page_count, int)
        and "V" not in str(el.Document_Revision)
    )
    print(pdf_path)
    try:
        from prettytable import PrettyTable

        table = PrettyTable()
        table.field_names = [
            "Doc_Title",
            "Page_Format",
            f"Page_Count:{total_page_count}",
            "Document_name",
            "Document_Revision",
            "Комплект",
        ]
        table.align["Doc_Title"] = "l"
        table.align["Document_name"] = "l"
        for el in proj_od_list:
            table.add_row([
                el.Doc_Title,
                el.Page_Format,
                el.doc_page_count,
                el.Document_name,
                el.Document_Revision,
                el.doc_range,
            ])
        print(str(table).encode("cp1251", errors="replace").decode("cp1251"))
    except ImportError:
        print(f"  ({len(proj_od_list)} entries, prettytable not available)")
