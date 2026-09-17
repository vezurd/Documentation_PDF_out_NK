"""Structured tag analysis payload for v2 pipeline and GUI (no Qt)."""

from __future__ import annotations

import os
from collections import Counter
from itertools import groupby
from typing import Any

from prettytable import PrettyTable

from tags.tag_classes import (
    TagClass,
    check_system_code_94s_4_1,
    check_tag_94s_4_2,
    check_tag_94s_4_4_additional_code,
)
from tags.tag_parser import add_to_dict, add_to_dict_doubles, diff_lists_MTO_B
from utils import string_parsing


def _sorted_source_items(source_dict: dict) -> list[tuple[Any, list[str]]]:
    return sorted(source_dict.items(), key=lambda kv: (kv[0][0], kv[0][1], kv[0][2]))


def _sheet_to_pdf_index(source_dict: dict) -> dict[str, list[dict[str, Any]]]:
    acc: dict[str, set[tuple[str, int]]] = {}
    for (doc_name, page_num, short_name), _ in _sorted_source_items(source_dict):
        acc.setdefault(short_name, set()).add((doc_name, int(page_num)))
    out: dict[str, list[dict[str, Any]]] = {}
    for sn in sorted(acc.keys()):
        pairs = sorted(acc[sn], key=lambda t: (t[0], t[1]))
        out[sn] = [{"pdf_path": p, "page_num": n} for p, n in pairs]
    return out


def _resolve_tag_location(
    sheet: str, tag: str, source_dict: dict
) -> tuple[str, int]:
    for (doc_name, page_num, sn), lst in _sorted_source_items(source_dict):
        if sn != sheet:
            continue
        if tag in lst:
            return str(doc_name), int(page_num)
    return "", -1


def _invalid_key(sheet: str, tag: str) -> tuple[str, str]:
    return (sheet, tag)


def build_tag_analysis_payload(
    source_dict: dict,
    *,
    main_doc_title: str = "МТО",
    main_doc_signature: str = "MTO-0001",
    high_frequency_threshold: int = 3,
) -> dict[str, Any]:
    """Build ``tag_analysis`` dict for GUI / summary (mirrors legacy ``analyze`` logic)."""
    out_dict: dict[str, list[str]] = {}
    mto_doubles_check: dict[str, list[str]] = {}
    count_check: dict[str, list[str]] = {}
    for i, v in _sorted_source_items(source_dict):
        short_name = i[2]
        content = v
        for x in range(len(content)):
            add_to_dict(out_dict, short_name, content[x])
            add_to_dict_doubles(mto_doubles_check, short_name, content[x])
            add_to_dict_doubles(count_check, f"{short_name};{i[1]}", content[x])

    for key in mto_doubles_check:
        out_dict.setdefault(key, [])
        out_dict[key].sort()
        mto_doubles_check[key].sort()

    sheet_to_pdf = _sheet_to_pdf_index(source_dict)
    warnings: list[str] = []
    invalid_tags: list[dict[str, Any]] = []
    invalid_seen: set[tuple[str, str]] = set()

    def _note_invalid(sheet: str, tag: str, err: str) -> None:
        k = _invalid_key(sheet, tag)
        if k in invalid_seen:
            return
        invalid_seen.add(k)
        pdf_path, page_num = _resolve_tag_location(sheet, tag, source_dict)
        invalid_tags.append(
            {
                "sheet": sheet,
                "tag": tag,
                "error": err,
                "pdf_path": pdf_path,
                "page_num": page_num,
            }
        )

    # --- Cyrillic ---
    cyrillic_rows: list[dict[str, Any]] = []
    for sheet in out_dict:
        content_list = out_dict[sheet]
        if not content_list:
            continue
        for tag in content_list:
            tc = TagClass(tag, strict=False)
            if not tc.is_valid:
                _note_invalid(sheet, tag, tc.parse_error)
                continue
            if not string_parsing.check_1_cirillic_letters(tag):
                letters = string_parsing.check_2_cirillic_letters(tag)
                pdf_path, page_num = _resolve_tag_location(sheet, tag, source_dict)
                cyrillic_rows.append(
                    {
                        "sheet": sheet,
                        "tag": tag,
                        "letters": letters,
                        "pdf_path": pdf_path,
                        "page_num": page_num,
                    }
                )

    # --- 94S ---
    amur_rows: list[dict[str, Any]] = []
    for sheet in out_dict:
        content_list = out_dict[sheet]
        if not content_list:
            continue
        for tag in content_list:
            tc = TagClass(tag, strict=False)
            if not tc.is_valid:
                _note_invalid(sheet, tag, tc.parse_error)
                continue
            if not check_system_code_94s_4_1(tc):
                continue
            pdf_path, page_num = _resolve_tag_location(sheet, tag, source_dict)
            if not check_tag_94s_4_2(tc):
                amur_rows.append(
                    {
                        "sheet": sheet,
                        "tag": tag,
                        "kind": "4.2",
                        "detail": f"<{tc.equipment}>",
                        "pdf_path": pdf_path,
                        "page_num": page_num,
                    }
                )
            if not check_tag_94s_4_4_additional_code(tc):
                amur_rows.append(
                    {
                        "sheet": sheet,
                        "tag": tag,
                        "kind": "4.4",
                        "detail": f"<{tc.additional_code}>",
                        "pdf_path": pdf_path,
                        "page_num": page_num,
                    }
                )

    mto_list: list[str] = []
    mto_flag = False
    mto_sheet_name = ""
    file_out_name = "NO_TAGS_IN_PROJECT_MTO"
    for i in out_dict:
        short_name = i
        if not mto_flag:
            file_out_name = short_name
        content = out_dict[i]
        if main_doc_signature in short_name:
            mto_list = list(content)
            file_out_name = short_name
            mto_sheet_name = short_name
            mto_flag = True

    mto_list_doubles: list[str] = []
    for i in mto_doubles_check:
        if main_doc_signature in i:
            mto_list_doubles = list(mto_doubles_check[i])

    wbs_counts: Counter[str] = Counter()
    for t in mto_list:
        tc = TagClass(t, strict=False)
        if not tc.is_valid:
            continue
        w = (tc.wbs or "").strip()
        if w:
            wbs_counts[w] += 1
    mto_wbs = ""
    if wbs_counts:
        max_c = max(wbs_counts.values())
        winners = sorted([w for w, c in wbs_counts.items() if c == max_c])
        if len(winners) == 1:
            mto_wbs = winners[0]
        else:
            warnings.append(
                "Невозможно однозначно определить титул МТО: " + ", ".join(winners)
            )

    mto_set = set(mto_list)
    doubles_counter = Counter(mto_list_doubles)
    mto_dup_rows = [
        {"tag": k, "count": v}
        for k, v in sorted(doubles_counter.items())
        if v > 1
    ]

    extra_rows: list[dict[str, Any]] = []
    for sheet_name in out_dict:
        if sheet_name == main_doc_signature:
            continue
        for tag in out_dict[sheet_name]:
            tc = TagClass(tag, strict=False)
            if not tc.is_valid:
                _note_invalid(sheet_name, tag, tc.parse_error)
                continue
            if tag in mto_set:
                continue
            if tc.is_hidden():
                reason = "hidden_by_rule"
            else:
                reason = "not_found"
            detail = ""
            if mto_wbs and (tc.wbs or "").strip() != mto_wbs.strip():
                detail = "Титул отличается от титула МТО"
            pdf_path, page_num = _resolve_tag_location(sheet_name, tag, source_dict)
            extra_rows.append(
                {
                    "sheet": sheet_name,
                    "tag": tag,
                    "wbs": tc.wbs,
                    "category_name": tc.category_name,
                    "out_row": tc.out_row,
                    "reason": reason,
                    "detail": detail,
                    "pdf_path": pdf_path,
                    "page_num": page_num,
                }
            )

    lists_list: list[str] = []
    for i in out_dict:
        if main_doc_signature not in i:
            for j in out_dict[i]:
                lists_list.append(j)
    missing_tags = diff_lists_MTO_B(lists_list, mto_list)
    missing_rows: list[dict[str, Any]] = []
    for el in missing_tags:
        tc = TagClass(el, strict=False)
        if not tc.is_valid:
            _note_invalid(mto_sheet_name or main_doc_signature, el, tc.parse_error)
            continue
        pdf_path, page_num = ("", -1)
        if mto_sheet_name:
            pdf_path, page_num = _resolve_tag_location(mto_sheet_name, el, source_dict)
        line4 = [main_doc_signature, "", "", ""]
        orow = tc.out_row
        if isinstance(orow, int) and 0 <= orow <= 3:
            line4[orow] = el
        else:
            line4[3] = el
        missing_rows.append(
            {
                "sheet": line4[0],
                "col_not_cab": line4[1],
                "col_in_cab": line4[2],
                "col_foreign": line4[3],
                "tag": el,
                "pdf_path": pdf_path,
                "page_num": page_num,
            }
        )

    out_list: list[list[Any]] = []
    high_frequency_warning = False
    if count_check:
        for k, v in count_check.items():
            temp_list = str(k).split(";")
            sheet_name = temp_list[0]
            if "CJ" in sheet_name:
                continue
            sheet_number = temp_list[1] if len(temp_list) > 1 else ""
            ctr = Counter(v)
            if any(cnt > 2 for cnt in ctr.values()):
                high_frequency_warning = True
            hi = {
                element: count
                for element, count in ctr.items()
                if count >= high_frequency_threshold
            }
            for tag, cnt in sorted(hi.items(), key=lambda x: x[1], reverse=True):
                pdf_path, page_num = _resolve_tag_location(sheet_name, tag, source_dict)
                out_list.append([sheet_name, sheet_number, tag, cnt, pdf_path, page_num])

    status = "ok" if mto_flag else "no_mto"

    checks: list[dict[str, Any]] = [
        {
            "id": "cyrillic",
            "title": "Проверка тегов на кириллицу",
            "severity": "warning" if cyrillic_rows else "info",
            "columns": ["Лист", "Тег", "Кириллица (заглавные)"],
            "column_keys": ["sheet", "tag", "letters"],
            "rows": cyrillic_rows,
        },
        {
            "id": "amur_94s",
            "title": "Проверка по процедуре AMUR-9000-94S-0010",
            "severity": "warning" if amur_rows else "info",
            "columns": ["Лист", "Тег", "Вид", "Деталь"],
            "column_keys": ["sheet", "tag", "kind", "detail"],
            "rows": amur_rows,
        },
        {
            "id": "mto_duplicates",
            "title": f"Проверка тегов на дублирование (по {main_doc_title})",
            "severity": "warning" if mto_dup_rows else "info",
            "columns": ["Тег", "Кол-во"],
            "column_keys": ["tag", "count"],
            "rows": mto_dup_rows,
        },
        {
            "id": "extra_in_sheets",
            "title": f"Теги которых НЕТ в {main_doc_title}, но ЕСТЬ на Листе",
            "severity": "warning" if extra_rows else "info",
            "columns": [
                "Лист",
                "Тег",
                "Титул",
                "Категория",
                "Причина",
                "Детали",
            ],
            "column_keys": [
                "sheet",
                "tag",
                "wbs",
                "category_name",
                "reason",
                "detail",
            ],
            "rows": extra_rows,
        },
        {
            "id": "missing_in_mto",
            "title": f"Теги которые ЕСТЬ в {main_doc_title}, но НЕТ на Листах",
            "severity": "warning" if missing_rows else "info",
            "columns": [
                "Имя листа",
                "Теги КСБ (НЕ в шкафах)",
                "Теги КСБ (обор. в шкафах)",
                "Теги НЕ КСБ",
            ],
            "column_keys": ["sheet", "col_not_cab", "col_in_cab", "col_foreign"],
            "rows": [],
        },
        {
            "id": "high_frequency",
            "title": "Количество тегов на листе",
            "severity": "warning" if high_frequency_warning else "info",
            "columns": ["Имя листа", "N листа", "Тег", "Кол-во"],
            "column_keys": ["sheet", "sheet_num", "tag", "count"],
            "rows": [
                {
                    "sheet": r[0],
                    "sheet_num": r[1],
                    "tag": r[2],
                    "count": r[3],
                    "pdf_path": r[4],
                    "page_num": r[5],
                }
                for r in out_list
            ],
            "meta": {"threshold": high_frequency_threshold},
        },
    ]

    checks[4]["rows"] = missing_rows

    return {
        "status": status,
        "mto_sheet_name": mto_sheet_name,
        "mto_wbs": mto_wbs,
        "text_report_path": "",
        "checks": checks,
        "invalid_tags": invalid_tags,
        "warnings": warnings,
        "sheet_to_pdf": sheet_to_pdf,
        "high_frequency_threshold": high_frequency_threshold,
        "main_doc_title": main_doc_title,
        "main_doc_signature": main_doc_signature,
        "file_out_name": file_out_name,
    }


def render_tag_analysis_text_from_payload(
    payload: dict[str, Any],
    *,
    main_doc_title: str,
    main_doc_signature: str,
) -> str:
    """Rebuild legacy ``*_tag_analyze.txt`` body (all extra tags, including hidden_by_rule)."""
    lines: list[str] = []

    def p(s: str = "") -> None:
        lines.append(s)

    cyrillic_check = next(c for c in payload["checks"] if c["id"] == "cyrillic")
    p("###############################")
    p("# Проверка тегов на кириллицу #")
    p("###############################")
    if not cyrillic_check["rows"]:
        p("\tКириллица в тегах не обнаружена")
    else:
        for r in cyrillic_check["rows"]:
            p(f"\tКириллица в теге (заглавные буквы): {r['letters']}, лист {r['sheet']}")
    p("\t--------------------------------")
    p()

    amur = next(c for c in payload["checks"] if c["id"] == "amur_94s")
    p("######################################################################")
    p("# Проверка тегов по процедуре AMUR-9000-94S-0010 приложение 4.2; 4.4 #")
    p("######################################################################")
    if not amur["rows"]:
        p("\tОшибок в тегировании по приложениям не обнаружено")
    else:
        for r in amur["rows"]:
            if r["kind"] == "4.2":
                p(f"{r['tag']}: {r['detail']}- не соответствует приложению 4.2 ({r['sheet']})")
            else:
                p(f"{r['tag']}: {r['detail']}- не соответствует приложению 4.4 ({r['sheet']})")
    p("\t--------------------------------")
    p()

    if payload.get("status") == "no_mto":
        p(f"\tОШИБКА: МТО не содержит тегов!!!")

    mto_dup = next(c for c in payload["checks"] if c["id"] == "mto_duplicates")
    p("###########################################")
    p(f" Проверка тегов на дублирование (по {main_doc_title}) ")
    p("###########################################")
    if not mto_dup["rows"]:
        p(f"\tДублирование тегов в {main_doc_title} не обнаружено")
    else:
        for r in mto_dup["rows"]:
            p(f"{r['tag']}: {r['count']} шт.")
    p("\t--------------------------------")
    p()

    extra = next(c for c in payload["checks"] if c["id"] == "extra_in_sheets")
    extra_rows_txt = [r for r in extra["rows"] if isinstance(r, dict)]

    p("############################################################")
    p(f" Теги которых НЕТ в {main_doc_title}, но ЕСТЬ на Листе     ")
    p("############################################################")
    table = PrettyTable()
    table.field_names = ["Имя листа", "Теги КСБ (НЕ в шкафах)", "Теги КСБ (обор. в шкафах)", "Теги НЕ КСБ"]
    table.align = "l"
    for sheet, group in groupby(extra_rows_txt, key=lambda r: r["sheet"]):
        for r in group:
            line = [sheet, "", "", ""]
            orow = r.get("out_row")
            if isinstance(orow, int) and 0 <= orow <= 3:
                line[orow] = r["tag"]
            else:
                line[3] = r["tag"]
            table.add_row(line)
        table.add_row(["", "", "", ""])
    lines.append(table.get_string())
    p("")
    p("############################################################")
    p(f"  Теги которые ЕСТЬ в {main_doc_title}, но НЕТ на Листах   ")
    p("############################################################")
    missing = next(c for c in payload["checks"] if c["id"] == "missing_in_mto")
    table2 = PrettyTable()
    table2.field_names = ["Имя листа", "Теги КСБ (НЕ в шкафах)", "Теги КСБ (обор. в шкафах)", "Теги НЕ КСБ"]
    table2.align = "l"
    if missing["rows"]:
        for r in missing["rows"]:
            table2.add_row(
                [
                    str(r.get("sheet", "")),
                    str(r.get("col_not_cab", "")),
                    str(r.get("col_in_cab", "")),
                    str(r.get("col_foreign", "")),
                ]
            )
        table2.add_row(["", "", "", ""])
    lines.append(table2.get_string())
    p("")
    p("############################################################")
    p(f"  Количество тегов на листе (выводим если больше 2)  ")
    p("############################################################")
    hf = next(c for c in payload["checks"] if c["id"] == "high_frequency")
    if hf["rows"]:
        thr = int(hf.get("meta", {}).get("threshold", 3))
        rows_thr = [r for r in hf["rows"] if int(r["count"]) >= thr]
        table3 = PrettyTable()
        table3.field_names = ["Имя листа", "N листа", "Тег", "Кол-во"]
        table3.align = "l"
        for r in sorted(rows_thr, key=lambda x: int(x["count"]), reverse=True):
            table3.add_row([r["sheet"], r["sheet_num"], r["tag"], r["count"]])
        lines.append(table3.get_string())
    for w in payload.get("warnings") or []:
        p(w)
    return "\n".join(lines)


def apply_text_report_path(payload: dict[str, Any], file_name: str) -> None:
    payload["text_report_path"] = os.path.abspath(file_name)
