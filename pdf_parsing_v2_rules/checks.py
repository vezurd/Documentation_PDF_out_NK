"""Normcontrol checks (v2): ported from ``pdf_parsing.rules_check`` with ``RulesProfile``."""

from __future__ import annotations

import difflib
import re
from collections import Counter
from typing import Any

from pdf_parsing_v2_engine.stamp_fields import *
from pdf_parsing_v2_engine.doc_types import LIST_OF_BBB
from pdf_parsing_v2_od import DOC_OSNOVNOI, DOC_PRILAGAEMIE
from pdf_parsing_v2_od.od_parsing import (
    DocOdAttributes,
    _get_page_count,
    _page_format_token_parts,
)
from utils.file_name_converts import ProjectFileName
from utils.string_parsing import getOdStyleFileName

import utils.string_parsing as string_parsing

from pdf_parsing_v2_rules._an_revision import (
    _STAMP_AN_TOKEN,
    _an_revision_long_from_18_1_rows,
    _first_cell_text,
    c18_2_field_id_for_max_literal_revision,
    c18_2_field_id_for_max_numeric_revision,
)
from pdf_parsing_v2_rules.output import CheckRow, formats_convolution, make_check_row
from pdf_parsing_v2_rules.profile import RulesProfile
from pdf_parsing_v2_engine.document import V2Document


def _is_enabled(c_code_int: int, enabled: frozenset[int] | None) -> bool:
    if enabled is None:
        return True
    return c_code_int in enabled


def _bbb_types(profile: RulesProfile) -> frozenset[str]:
    """BBB document types for this profile (fallback: engine ``LIST_OF_BBB``)."""
    if profile.list_of_bbb:
        return frozenset(profile.list_of_bbb)
    return frozenset(LIST_OF_BBB)


def run_all_checks(
    curr_proj: list[V2Document],
    proj_od_list: list | int,
    profile: RulesProfile,
    *,
    od_pdf_path: str = "",
) -> list[CheckRow]:
    """Run all enabled checks; return flat list of results."""
    enabled = profile.enabled_checks
    results: list[CheckRow] = []
    brake = False

    if _is_enabled(1000, enabled):
        rows = check_file_names(curr_proj, profile)
        results.extend(rows)
        if any(not r.result for r in rows):
            brake = True

    if not brake and _is_enabled(100, enabled):
        results.extend(prepare_revision_selection(curr_proj, profile))
    if not brake and _is_enabled(1001, enabled):
        results.extend(check_cyrillic(curr_proj, profile))
    if not brake and _is_enabled(1002, enabled):
        results.extend(check_layers(curr_proj, profile))
    if not brake and _is_enabled(1003, enabled):
        results.extend(check_annotations(curr_proj, profile))
    if not brake and _is_enabled(1004, enabled):
        results.extend(check_format_vs_real(curr_proj, profile))
    if not brake and _is_enabled(1005, enabled):
        results.extend(check_revisions_internal(curr_proj, profile))
    if not brake and _is_enabled(1006, enabled):
        results.extend(check_revisions_vs_od(curr_proj, proj_od_list, profile))
    if not brake and _is_enabled(1007, enabled):
        results.extend(check_formats_vs_od(curr_proj, proj_od_list, profile))
    if not brake and _is_enabled(1008, enabled):
        results.extend(check_page_count_stamp(curr_proj, profile))
    if not brake and _is_enabled(1009, enabled):
        results.extend(check_doc_numbers(curr_proj, profile))
    if not brake and _is_enabled(1010, enabled):
        results.extend(check_page_numbers(curr_proj, profile))
    if not brake and _is_enabled(1011, enabled):
        results.extend(check_doc_name_vs_od(curr_proj, proj_od_list, profile))
    if not brake and _is_enabled(1012, enabled):
        results.extend(check_facility_title(curr_proj, proj_od_list, profile))
    if not brake and _is_enabled(1013, enabled):
        results.extend(check_dates(curr_proj, profile))
    if not brake and _is_enabled(1014, enabled):
        results.extend(check_doc_code(curr_proj, profile))
    if not brake and _is_enabled(1015, enabled):
        results.extend(
            check_od_page_count(
                curr_proj, proj_od_list, profile, od_pdf_path=str(od_pdf_path or "").strip()
            )
        )
    if not brake and _is_enabled(1016, enabled):
        results.extend(check_pdf_extension(curr_proj, profile))
    if not brake and _is_enabled(1017, enabled):
        results.extend(check_ifc_ifr(curr_proj, profile))
    if not brake and _is_enabled(1018, enabled):
        results.extend(check_mark(curr_proj, profile))
    if not brake and _is_enabled(1019, enabled):
        results.extend(check_duplicate_ext(curr_proj, profile))

    return results


def check_file_names(curr_proj: list[V2Document], profile: RulesProfile) -> list[CheckRow]:
    rows: list[CheckRow] = []
    c_code = [1000, "Проверка имен файлов"]
    for doc in curr_proj:
        file_name = doc.file_name
        check_file_name = ProjectFileName.find_full_matches(
            file_name, project_name=profile.project_name
        )
        if len(check_file_name) != 1:
            result = False
            text = f"Имя файла некорректное: {file_name}, проверка комплекта остановлена"
        else:
            result = True
            text = f"Имя файла корректное: {file_name}"
        rows.append(make_check_row(result, c_code, doc.doc_OD_style_file_name, "нет", text))
    return rows


_REVISION_ROW_FIELDS = (
    (c_18_1_1, c_18_2_1, c_18_3_1),
    (c_18_1_2, c_18_2_2, c_18_3_2),
    (c_18_1_3, c_18_2_3, c_18_3_3),
)


def _row_fields_for_c18_2(c18_2_key: str | None) -> tuple[str, str, str] | None:
    """Return parallel ``18_1``/``18_2``/``18_3`` field ids for a selected row."""
    if c18_2_key is None:
        return None
    for row_fields in _REVISION_ROW_FIELDS:
        if row_fields[1] == c18_2_key:
            return row_fields
    return None


def _selected_revision_row_fields(temp_p: dict, profile: RulesProfile) -> tuple[str, str, str] | None:
    """Choose the current revision row using the shared 18.1 revision parser."""
    rev_full, rev_number, rev_an, stamp_error = _an_revision_long_from_18_1_rows(
        temp_p,
        [c_18_1_1, c_18_1_2, c_18_1_3],
        rev_literal=profile.rev_literal,
    )
    if stamp_error is not None or not rev_full:
        return None
    if rev_number and rev_number in profile.rev_literal and rev_an is None:
        c18_2_key = c18_2_field_id_for_max_literal_revision(
            temp_p,
            rev_literal=profile.rev_literal,
        )
    else:
        # For AN revisions, the full revision is RD+AN, while the working row for
        # date/purpose remains the last numeric RD row by project rule 5.3.4.
        c18_2_key = c18_2_field_id_for_max_numeric_revision(temp_p)
    return _row_fields_for_c18_2(c18_2_key)


def prepare_revision_selection(curr_proj: list[V2Document], profile: RulesProfile) -> list[CheckRow]:
    rows: list[CheckRow] = []
    c_code = [
        100,
        "Подготовка выбора ревизии:\n"
        "    18_1_1, 18_1_2, 18_1_3;\n",
    ]
    for doc in curr_proj:
        for page in doc.pages:
            if page.page_num == 1:
                temp_p = page.dict_attributes
                rev_full, _, _, stamp_parse_error = _an_revision_long_from_18_1_rows(
                    temp_p,
                    [c_18_1_1, c_18_1_2, c_18_1_3],
                    rev_literal=profile.rev_literal,
                )
                selected_fields = _selected_revision_row_fields(temp_p, profile)
                if stamp_parse_error is not None:
                    result = False
                    text = stamp_parse_error
                elif selected_fields is not None and rev_full:
                    row_18_1, row_18_2, row_18_3 = selected_fields
                    page.dict_attributes[c_18_1] = [rev_full]
                    page.dict_attributes[c_18_2] = page.dict_attributes[row_18_2]
                    page.dict_attributes[c_18_3] = page.dict_attributes[row_18_3]
                    result = True
                    text = f"Рабочая ревизия выбрана: {rev_full} ({row_18_1})"
                else:
                    result = False
                    text = "Рабочая ревизия НЕ найдена в поле 18"
                rows.append(make_check_row(result, c_code, doc.doc_OD_style_file_name, page, text))
    return rows


def check_cyrillic(curr_proj: list[V2Document], profile: RulesProfile) -> list[CheckRow]:
    rows: list[CheckRow] = []
    c_code = [
        1001,
        "Проверка на кириллицу:\n"
        "    в имени файла; \n"
        "    1_DOC_TITLE;\n"
        "    в полях ревизий 18_1_1, 18_1_2, 18_1_3;\n"
        "    26_Document_Revision; \n"
        "    c_50_File_Name_Stamp.",
    ]
    attr_list_to_check = [
        c_75_File_Full_Name,
        c_1_DOC_TITLE,
        c_18_1_1,
        c_18_1_2,
        c_18_1_3,
        c_26_Document_Revision,
        c_50_File_Name_Stamp,
    ]
    for doc in curr_proj:
        for page in doc.pages:
            value_list: dict[str | int, str | Any] = {}
            for attr in attr_list_to_check:
                if not page.dict_attributes[attr]:
                    result = True
                    text = "Кириллица в атрибуте " + attr + " не найдена"
                else:
                    value_list[attr] = page.dict_attributes[attr][0]
            for i in value_list:
                result = string_parsing.check_1_cirillic_letters(value_list[i])
                if not result:
                    text = "Кириллица в атрибуте " + i + ": " + value_list[i]
                else:
                    text = "Кириллица в атрибуте " + i + " не найдена"
                rows.append(make_check_row(result, c_code, doc.doc_OD_style_file_name, page, text))
    return rows


def check_layers(curr_proj: list[V2Document], profile: RulesProfile) -> list[CheckRow]:
    rows: list[CheckRow] = []
    c_code = [1002, "Проверка на многослойность"]
    attr_list_to_check = [c_61_Page_Layers]
    for doc in curr_proj:
        for page in doc.pages:
            for attr in attr_list_to_check:
                value_list = page.dict_attributes[attr]
                if value_list == [-1]:
                    result = True
                    text = "Нет многослойности " + attr
                else:
                    result = False
                    text = "Есть многослойность " + attr
                rows.append(make_check_row(result, c_code, doc.doc_OD_style_file_name, page, text))
    return rows


def check_annotations(curr_proj: list[V2Document], profile: RulesProfile) -> list[CheckRow]:
    rows: list[CheckRow] = []
    c_code = [1003, "Проверка на заметки, аннотации, комментарии"]
    attr_list_to_check = [c_62_Page_Bookmarks]
    for doc in curr_proj:
        for page in doc.pages:
            for attr in attr_list_to_check:
                value_list = page.dict_attributes[attr]
                if value_list == [-1]:
                    result = True
                    text = "Нет Аннотаций " + attr
                else:
                    result = False
                    text = "Есть Аннотации " + attr
                rows.append(make_check_row(result, c_code, doc.doc_OD_style_file_name, page, text))
    return rows


def check_format_vs_real(curr_proj: list[V2Document], profile: RulesProfile) -> list[CheckRow]:
    rows: list[CheckRow] = []
    c_code = [1004, "Проверка на форматы листов в штампе и реального (51)"]
    attr_list_to_check = [c_51_Page_Format]
    bbb = _bbb_types(profile)
    for doc in curr_proj:
        for page in doc.pages:
            if not (doc.doc_Type in bbb and page.page_num > 1):
                for attr in attr_list_to_check:
                    try:
                        value = page.dict_attributes[attr][0]
                    except IndexError:
                        value = "Формат не найден"

                    page_format = string_parsing.get_Format_from_51(value)
                    real_format = page.dict_attributes[c_65_Page_Real_Format][0]
                    if page_format == real_format:
                        result = True
                        text = "Формат правильный"
                    else:
                        result = False
                        text = (
                            f"Формат не совпадет {attr}:\n"
                            f"    {page_format} | Формат под основной надписью\n"
                            f"    {real_format} | Реальный формат страницы"
                        )

                    rows.append(make_check_row(result, c_code, doc.doc_OD_style_file_name, page, text))
    return rows


def check_revisions_internal(curr_proj: list[V2Document], profile: RulesProfile) -> list[CheckRow]:
    rows: list[CheckRow] = []
    c_code = [1005, "Проверка на Ревизии, внутри Документа (File_name; 18_1; 26; 50)"]
    attr_list_to_check = [
        c_18_1,
        c_26_Document_Revision,
        c_50_File_Name_Stamp,
    ]
    bbb = _bbb_types(profile)
    rev_literal = profile.rev_literal
    for doc in curr_proj:
        if doc.doc_Revision == "V":
            continue
        rev_full = ""
        an_flag = False
        stamp_parse_error: str | None = None

        for page in doc.pages:
            if page.page_num == 1:
                list_att_of_18_1 = [
                    c_18_1_1,
                    c_18_1_2,
                    c_18_1_3,
                ]
                temp_p = doc.pages[0].dict_attributes
                rev_full, _, rev_an, stamp_parse_error = _an_revision_long_from_18_1_rows(
                    temp_p, list_att_of_18_1, rev_literal=rev_literal
                )
                an_flag = rev_an is not None
                if stamp_parse_error is None and doc.doc_Revision and _STAMP_AN_TOKEN.search(doc.doc_Revision):
                    an_flag = True

            value_list_1005: dict[str, Any] = {}

            value_list_1005["Из имени файла"] = doc.doc_Revision

            for attr in attr_list_to_check:
                if attr == c_18_1 and page.page_num > 1:
                    continue
                elif (
                    attr == c_26_Document_Revision
                    and doc.doc_Type in bbb
                    and page.page_num > 1
                ):
                    continue
                else:
                    try:
                        if attr == c_50_File_Name_Stamp:
                            t = string_parsing.getRevisionFromFileName(page.dict_attributes[attr][0])
                        else:
                            t = page.dict_attributes[attr][0]

                        value_list_1005[str(attr)] = t
                    except IndexError:
                        t = "Нет ревизии"
                        value_list_1005[str(attr)] = t

            result = True

            if stamp_parse_error is not None:
                if page.page_num == 1:
                    rows.append(
                        make_check_row(
                            False,
                            c_code,
                            doc.doc_OD_style_file_name,
                            page,
                            stamp_parse_error,
                        )
                    )
                continue

            if an_flag:
                if (
                    value_list_1005[c_50_File_Name_Stamp] != rev_full
                    or value_list_1005["Из имени файла"] != rev_full
                ):
                    text = (
                        f"Ревизии НЕ совпадают:\n"
                        f"    {value_list_1005['Из имени файла']} | Из имени файла\n"
                        f"    {value_list_1005[c_50_File_Name_Stamp]} | Из имени файла под штампом\n"
                        f"    {rev_full} | Из штампа ревизий (рев. РД + рев. АН)"
                    )
                    rows.append(make_check_row(False, c_code, doc.doc_OD_style_file_name, page, text))
            if not an_flag:
                x = list(value_list_1005.values())
                result = False
                freq = Counter(x)
                if len(freq) == 1:
                    result = True

                if result:
                    text = "Ревизии совпадают "
                else:
                    text = "Ревизии НЕ совпадают:"
                    for key, value in value_list_1005.items():
                        text += f"\n    {value} | {key}"
                rows.append(make_check_row(result, c_code, doc.doc_OD_style_file_name, page, text))
    return rows


def check_revisions_vs_od(
    curr_proj: list[V2Document],
    proj_od_list: list | int,
    profile: RulesProfile,
) -> list[CheckRow]:
    rows: list[CheckRow] = []
    c_code = [1006, "Проверка на Ревизии, в Документе (26) и ОД "]
    if proj_od_list == 0:
        return rows
    dict_od_doc_name: dict[Any, Any] = {}
    for doc in proj_od_list:
        dict_od_doc_name[doc.Doc_Title] = doc.Document_Revision

    for doc in curr_proj:
        doc_od_style_format_file_name = doc.doc_OD_style_file_name
        doc_document_name_4 = doc.doc_Revision
        try:
            od_rev = dict_od_doc_name[doc_od_style_format_file_name]
            od_rev = od_rev.replace("Рев.", "").strip()
            od_rev = od_rev.replace("Новый", "").strip()
            od_rev = od_rev.replace("Выпущен взамен", "").strip()
        except:
            od_rev = "в ОД нет документа " + doc_od_style_format_file_name
        result = False
        if doc_document_name_4 == od_rev:
            result = True

        if not result:
            text = (
                f"Ревизии НЕ совпадают:\n"
                f"    {str(doc_document_name_4)} | в {doc.doc_Short_File_Name}\n"
                f"    {str(od_rev)} | в Общих Данных"
            )
        else:
            text = "Ревизии совпадают"

        rows.append(make_check_row(result, c_code, doc.doc_OD_style_file_name, "нет", text))
    return rows


def check_formats_vs_od(
    curr_proj: list[V2Document],
    proj_od_list: list | int,
    profile: RulesProfile,
) -> list[CheckRow]:
    rows: list[CheckRow] = []
    c_code = [1007, "Проверка на форматы листов ОД и документов (51==ОД_форматы)"]
    if proj_od_list == 0:
        return rows
    dict_doc_formats: dict[str, list[Any]] = {}
    for doc in curr_proj:
        curr_doc_name = doc.doc_OD_style_file_name
        dict_doc_formats[curr_doc_name] = []
        for page in doc.pages:
            try:
                value_list = page.dict_attributes[c_51_Page_Format][0]
            except IndexError:
                value_list = "Формат не найден"
            page_format = string_parsing.get_Format_from_51(value_list)
            dict_doc_formats[curr_doc_name].append(page_format)
    for doc in proj_od_list:
        if doc.doc_range == DOC_PRILAGAEMIE:
            try:
                dict_doc_formats.pop(doc.Doc_Title)
            except KeyError:
                continue
    dict_od_formats: dict[str, list[Any]] = {}
    for doc in proj_od_list:
        if doc.doc_range == DOC_OSNOVNOI:
            title = doc.Doc_Title
            dict_od_formats[title] = []
            format_list = doc.Page_Format.split(",")
            for x in format_list:
                count = string_parsing.get_Format_Count(x)

                single_format = string_parsing.get_Format_Single(x)
                for i in range(count):
                    dict_od_formats[title].append(single_format)

    for x in dict_doc_formats:
        dict_doc_formats[x].sort()
    for x in dict_od_formats:
        dict_od_formats[x].sort()
    for x in dict_doc_formats:
        try:
            if dict_doc_formats[x] != dict_od_formats[x]:
                result = False
                text = (
                    f"Форматы НЕ совпадают:\n"
                    f"    {formats_convolution(dict_doc_formats[x])} | в {x}\n"
                    f"    {formats_convolution(dict_od_formats[x])} | в Общих данных"
                )
            else:
                result = True
                text = "Форматы совпадают"
        except IndexError:
            result = False
            text = "Ошибка, <" + x + "> нет в Общих Данных "
        except KeyError:
            result = False
            text = "Ошибка, <" + x + "> нет в Общих Данных "

        rows.append(make_check_row(result, c_code, x, "нет", text))
    return rows


def check_page_count_stamp(curr_proj: list[V2Document], profile: RulesProfile) -> list[CheckRow]:
    rows: list[CheckRow] = []
    c_code = [1008, "Проверка на кол-во листов в штампе (6.2) и реальное кол-во"]
    for doc in curr_proj:
        for page in doc.pages:
            if page.page_num < 2:
                try:
                    count_6_2 = page.dict_attributes[c_6_2_Quantity_of_sheets][0]
                except IndexError:
                    count_6_2 = "Атрибут 6_2 не найден"
                try:
                    if count_6_2 == "Атрибут 6_2 не найден":
                        result = False
                        text = count_6_2
                    elif int(count_6_2) == len(doc.pages):
                        result = True
                        text = "Кол-во листов в штампе совпадают"
                    else:
                        result = False
                        text = (
                            f"Кол-во листов в штампе ({count_6_2})  и реальное кол-во "
                            f"({len(doc.pages)}) - НЕ совпадают"
                        )
                except ValueError as err:
                    text = (
                        f"ERROR: {doc.doc_OD_style_file_name} page:{page.page_num} "
                        f'6_2:"{count_6_2}" err:{err}'
                    )
                    result = False
                    print(text)

                rows.append(make_check_row(result, c_code, doc.doc_OD_style_file_name, page, text))
    return rows


def check_doc_numbers(curr_proj: list[V2Document], profile: RulesProfile) -> list[CheckRow]:
    rows: list[CheckRow] = []
    c_code = [1009, "Проверка номеров документа (File_name==1==50==6.1_prefix)"]
    attr_list_to_check = [
        c_1_DOC_TITLE,
        c_50_File_Name_Stamp,
        c_6_1_Sheet_number,
    ]
    bbb = _bbb_types(profile)
    for doc in curr_proj:
        for page in doc.pages:
            page_num_list_6_1: list[Any] = []
            try:
                page_num_list_6_1.append((int(doc.doc_Number), "В имени файла"))
            except:
                page_num_list_6_1.append(("номер не найден", "В имени файла"))
            if not (doc.doc_Type in bbb and page.page_num > 1):
                for attr in attr_list_to_check:
                    try:
                        page_num_list_6_1.append(
                            (string_parsing.get_page_num_prefix(page.dict_attributes[attr][0]), attr)
                        )
                    except IndexError:
                        page_num_list_6_1.append(("не найден", attr))
                    except ValueError:
                        page_num_list_6_1.append(("не корректный", attr))

                result = True
                x = page_num_list_6_1[0][0]
                for i in range(len(page_num_list_6_1)):
                    if x != page_num_list_6_1[i][0]:
                        result = False

                if result:
                    text = f"Номера документов совпадают: {page_num_list_6_1}"
                else:
                    text = "Ревизии НЕ совпадают:"
                    for i in range(len(page_num_list_6_1)):
                        text = (
                            f"{text}\n"
                            f"    {page_num_list_6_1[i][0]} ({page_num_list_6_1[i][1]}) "
                        )
                rows.append(make_check_row(result, c_code, doc.doc_OD_style_file_name, page, text))
    return rows


def check_page_numbers(curr_proj: list[V2Document], profile: RulesProfile) -> list[CheckRow]:
    rows: list[CheckRow] = []
    c_code = [1010, "Проверка номеров страниц (№ стр == 6.1_suffix)"]
    for doc in curr_proj:
        for page in doc.pages:
            if page.page_num < 2:
                try:
                    count = page.dict_attributes[c_6_2_Quantity_of_sheets][0]
                except IndexError:
                    print(f"ERROR: {c_code[0]} - {doc.doc_Short_File_Name} - page:{page.page_num} ")
                    count = ""
            else:
                count = ""
            value_list: Any = None
            try:
                if count != "":
                    first_page_check = int(count)
            except ValueError:
                first_page_check = 0
                value_list = (
                    f"{c_6_2_Quantity_of_sheets} не корректный first_page_check (ValueError) ({count})"
                )

            if "не корректный" in str(value_list):
                pass
            elif count != "" and first_page_check == 1:
                value_list = 1
            else:
                try:
                    value_list = string_parsing.get_page_num_suffix(
                        page.dict_attributes[c_6_1_Sheet_number][0]
                    )
                except IndexError:
                    value_list = f"{c_6_2_Quantity_of_sheets} не корректный (IndexError)"
                except ValueError:
                    value_list = f"{c_6_2_Quantity_of_sheets} не корректный (ValueError)"

            if "не корректный" in str(value_list):
                result = False
                text = value_list
            elif value_list == page.page_num:
                result = True
                text = f"Номера страниц совпадают: {value_list} и {page.page_num}"
            else:
                result = False
                text = f"Номера страниц НЕ совпадают: {value_list} вместо {page.page_num}"
            rows.append(make_check_row(result, c_code, doc.doc_OD_style_file_name, page, text))
    return rows


def check_doc_name_vs_od(
    curr_proj: list[V2Document],
    proj_od_list: list | int,
    profile: RulesProfile,
) -> list[CheckRow]:
    rows: list[CheckRow] = []
    c_code = [1011, "Проверка на Наименования листов в Документе и ОД (п.4)"]
    if proj_od_list == 0:
        return rows
    dict_od_doc_name: dict[Any, Any] = {}
    for doc in proj_od_list:
        dict_od_doc_name[doc.Doc_Title] = doc.Document_name

    for doc in curr_proj:
        doc_od_style_format_file_name = ""
        doc_document_name_4 = ""
        for page in doc.pages:
            if page.page_num == 1:
                doc_od_style_format_file_name = doc.doc_OD_style_file_name
                try:
                    doc_document_name_4 = page.dict_attributes[c_4_Document_name][0]
                except IndexError:
                    print(
                        f"IndexError: list index out of range: {doc.doc_OD_style_file_name}, "
                        f"page:{page.page_num}, {c_4_Document_name}"
                    )
                    print("Программа завершена с ошибкой!")
                    exit(0)
                try:
                    od_doc_name_4 = dict_od_doc_name[doc_od_style_format_file_name]
                except:
                    od_doc_name_4 = "в ОД нет документа " + doc_od_style_format_file_name

                doc_document_name_4_r = doc_document_name_4.strip()
                od_doc_name_4_r = od_doc_name_4.strip()

                doc_document_name_4_cmp = doc_document_name_4.replace(" ", "").strip().lower()
                od_doc_name_4_cmp = od_doc_name_4.replace(" ", "").strip().lower()

                diff = difflib.unified_diff(doc_document_name_4_cmp, od_doc_name_4_cmp)
                diff = "".join(diff)
                result = False

                if doc_document_name_4_cmp == od_doc_name_4_cmp:
                    result = True

                if result:
                    text = "Наименования совпадают"
                else:
                    text = (
                        f"Наименования НЕ совпадают.\n"
                        f"в {doc.doc_Short_File_Name}:\n"
                        f"    {str(doc_document_name_4_r)}\n"
                        f"в Общих Данных:\n"
                        f"    {str(od_doc_name_4_r)}\n "
                        f"Отличающиеся фрагменты:\n"
                        f"    {diff.split('@')[-1].strip()}"
                    )

                rows.append(make_check_row(result, c_code, doc.doc_OD_style_file_name, "нет", text))
    return rows


def check_facility_title(
    curr_proj: list[V2Document],
    proj_od_list: list | int,
    profile: RulesProfile,
) -> list[CheckRow]:
    rows: list[CheckRow] = []
    c_code = [1012, "Проверка наименования титула и объекта (относительно ОД) (п.2, п.3)"]
    if proj_od_list == 0:
        return rows
    attr_list_to_check = [
        c_2_Facility_name,
        c_3_Unit_title_name,
    ]
    for attr in attr_list_to_check:
        value_list: list[str] = []
        val_for_show_list: list[str] = []
        for doc in curr_proj:
            for page in doc.pages:
                if page.page_num == 1:
                    try:
                        value = page.dict_attributes[attr][0]
                    except IndexError:
                        value = f"Ошибка чтения атрибута:{attr}"

                    val_for_show = value
                    val_for_show_list.append(val_for_show)

                    value = value.replace(" ", "").strip().lower()
                    value = value.replace(".", "")
                    value_list.append(value)

                    if value == value_list[0]:
                        result = True
                        text = f"{attr} совпадают."
                    else:
                        result = False
                        text = (
                            f"{attr} НЕ совпадают:\n"
                            f"    {val_for_show} |текущ.|\n"
                            f"    {val_for_show_list[0]} |ОД(или первый документ)|"
                        )

                    rows.append(make_check_row(result, c_code, doc.doc_OD_style_file_name, page, text))
    return rows


_C10_SIGNATURE_FIELDS = (
    c_10_1,
    c_10_2,
    c_10_3,
    c_10_4,
    c_10_5,
)


def _parse_stamp_date_loose(value: str):
    """Return ``date`` or ``None`` if string is not ``dd.mm.yyyy`` / ``dd.mm.yy``."""
    from datetime import datetime

    s = value.strip()
    for fmt in ("%d.%m.%Y", "%d.%m.%y"):
        try:
            return datetime.strptime(s, fmt).date()
        except Exception:
            continue
    return None


def _stamp_10_dates_consistent(page) -> tuple[bool, str]:
    """Non-empty ``10_x`` dates must be the same calendar day; not all empty."""
    parts: list[tuple[str, str]] = []
    for att in _C10_SIGNATURE_FIELDS:
        try:
            raw = page.dict_attributes[att][0]
        except (IndexError, KeyError):
            raw = ""
        v = str(raw).strip() if raw else ""
        parts.append((att, v))
    nonempty = [(a, v) for a, v in parts if v]
    if not nonempty:
        listing = " ".join(f"{a}=<пусто>" for a, _ in parts)
        return False, f"Даты подписей в основном штампе некорректны: {listing}"
    d0 = _parse_stamp_date_loose(nonempty[0][1])
    if d0 is None:
        listing = " ".join(f"{a}=<{v}>" for a, v in parts)
        return False, f"Даты подписей в основном штампе некорректны: {listing}"
    for _att, val in nonempty[1:]:
        d = _parse_stamp_date_loose(val)
        if d is None or d != d0:
            listing = " ".join(f"{a}=<{v}>" for a, v in parts)
            return False, f"Даты подписей в основном штампе некорректны: {listing}"
    return True, ""


def _first_nonempty_c10_date(page) -> str:
    for att in _C10_SIGNATURE_FIELDS:
        try:
            raw = page.dict_attributes[att][0]
        except (IndexError, KeyError):
            continue
        v = str(raw).strip() if raw else ""
        if v:
            return v
    return ""


_C18_2_ROW_KEYS_TOP_FIRST = (
    c_18_2_3,
    c_18_2_2,
    c_18_2_1,
)


def _first_nonempty_c18_2_row_key_top_first(temp_p: dict) -> str | None:
    """First non-empty date among ``18_2_3`` … ``18_2_1`` (верх → низ штампа)."""
    for att in _C18_2_ROW_KEYS_TOP_FIRST:
        raw = temp_p.get(att, [])
        if not raw:
            continue
        v = str(raw[0]).strip() if raw[0] else ""
        if v:
            return att
    return None


_ROW_ATTS_18_1 = (c_18_1_1, c_18_1_2, c_18_1_3)


def _rev_number_matches_rev_numeral(rev_number: str, rev_numeral: tuple[str, ...]) -> bool:
    """Часть РД (без AN) — только цифры и входит в ``rev_numeral`` профиля."""
    if not rev_number or _STAMP_AN_TOKEN.search(rev_number):
        return False
    if rev_number in rev_numeral:
        return True
    if rev_number.isdigit():
        return any(r.isdigit() and int(rev_number) == int(r) for r in rev_numeral)
    return False


def check_dates(curr_proj: list[V2Document], profile: RulesProfile) -> list[CheckRow]:
    rows: list[CheckRow] = []
    c_code = [
        1013,
        "Проверка дат: цифровые ревизии (10==18.2); AN-ревизии — только формат 18.2\n"
        "    (5.3.4: дата подписания в поле 10 остаётся от последней цифровой ревизии)",
    ]
    rev_literal = profile.rev_literal
    rev_numeral = profile.rev_numeral
    for doc in curr_proj:
        for page in doc.pages:
            if page.page_num != 1:
                continue
            ok10, err10 = _stamp_10_dates_consistent(page)
            if not ok10:
                rows.append(make_check_row(False, c_code, doc.doc_OD_style_file_name, page, err10))
                continue

            temp_p = page.dict_attributes
            row_atts = list(_ROW_ATTS_18_1)
            rev_full, rev_number, rev_an, stamp_err = _an_revision_long_from_18_1_rows(
                temp_p, row_atts, rev_literal=rev_literal
            )
            if stamp_err:
                rows.append(make_check_row(False, c_code, doc.doc_OD_style_file_name, page, stamp_err))
                continue

            stamp_rev_label = rev_full if rev_full else "—"

            format_only = (
                rev_an is not None
                or (rev_full and _STAMP_AN_TOKEN.search(rev_full))
                or (rev_number and rev_number in rev_literal)
            )
            if format_only:
                if rev_an is not None or (rev_full and _STAMP_AN_TOKEN.search(rev_full)):
                    c18_key = c18_2_field_id_for_max_numeric_revision(temp_p)
                else:
                    c18_key = c18_2_field_id_for_max_literal_revision(temp_p, rev_literal=rev_literal)
                if c18_key is None:
                    c18_key = _first_nonempty_c18_2_row_key_top_first(temp_p)
                if c18_key is None:
                    result = False
                    text = (
                        "1013: нет даты в полях 18_2_1…18_2_3 "
                        f"(литерал/АН по штампу, канон {stamp_rev_label})"
                    )
                else:
                    try:
                        date_18_2_raw = page.dict_attributes[c18_key][0]
                    except (IndexError, KeyError):
                        date_18_2_raw = ""
                    date_18_2_fmt = string_parsing.check_data_format_long(date_18_2_raw)
                    if date_18_2_fmt != string_parsing.check_data_format_long(11):
                        result = True
                        text = (
                            f"Корректный формат даты (штамп {stamp_rev_label}, {c18_key}): "
                            f"{date_18_2_fmt}"
                        )
                    else:
                        result = False
                        text = (
                            f"НЕ корректный формат даты (штамп {stamp_rev_label}, {c18_key}): "
                            f"{date_18_2_fmt}"
                        )
            elif rev_number and _rev_number_matches_rev_numeral(rev_number, rev_numeral):
                c18_key = c18_2_field_id_for_max_numeric_revision(temp_p)
                if c18_key is None:
                    result = False
                    text = (
                        "1013: в 18_1_1…18_1_3 нет числовой ревизии РД — нельзя выбрать дату "
                        f"в 18_2_1…18_2_3 (канон штампа {stamp_rev_label})"
                    )
                else:
                    try:
                        date_18_2_raw = page.dict_attributes[c18_key][0]
                    except (IndexError, KeyError):
                        date_18_2_raw = ""
                    date_18_2_r = string_parsing.check_data_format_long(date_18_2_raw)
                    date_10 = _first_nonempty_c10_date(page)
                    date_10_r = string_parsing.check_data_format_short(date_10)
                    if date_18_2_r == string_parsing.check_data_format_long(11):
                        result = False
                        text = (
                            f"НЕ корректный формат даты (штамп {stamp_rev_label}, {c18_key}): "
                            f"{date_18_2_raw}"
                        )
                    elif date_10_r == string_parsing.check_data_format_long(11):
                        result = False
                        text = f"НЕ корректный формат даты (штамп {stamp_rev_label}, поле 10): {date_10}"
                    else:
                        if string_parsing.compare_data(date_10, date_18_2_raw):
                            result = True
                            text = (
                                f"Даты 18_2_x и 10 совпадают (канон {stamp_rev_label}, дата из {c18_key}): "
                                f"<{date_18_2_raw}> <{date_10}>"
                            )
                        else:
                            result = False
                            text = (
                                f"Даты 18_2_x и 10 НЕ совпадают (канон {stamp_rev_label}, дата из {c18_key}): "
                                f"<{date_18_2_raw}> <{date_10}>"
                            )
            else:
                result = True
                text = (
                    f"1013: канон штампа не в rev_literal / не числовой из профиля ({stamp_rev_label!r}), "
                    "сравнение 18_2_x с 10 не выполнялось"
                )

            rows.append(make_check_row(result, c_code, doc.doc_OD_style_file_name, page, text))
    return rows


def check_doc_code(curr_proj: list[V2Document], profile: RulesProfile) -> list[CheckRow]:
    rows: list[CheckRow] = []
    c_code = [1014, "Проверка шифра внутри документа (1_DOC_TITLE)"]
    for doc in curr_proj:
        for page in doc.pages:
            value_list: list[Any] = [doc.doc_OD_style_file_name]
            try:
                value_list.append(page.dict_attributes[c_1_DOC_TITLE][0].strip())
                result = all(x == value_list[0] for x in value_list)
            except IndexError:
                result = False
                text = f"ERROR 1014: {doc.doc_OD_style_file_name} p{page.page_num}"
                value_list.append("не найден")
                print(text)

            if result:
                text = "Шифры совпадают"
            else:
                try:
                    text = (
                        f"Шифры НЕ совпадают:\n"
                        f"    {value_list[0]} | В имени фала\n"
                        f'    {value_list[1]} | В поле "1_DOC_TITLE"'
                    )
                except IndexError:
                    text = (
                        "ОШИБКА"
                        ", не найден или не корректный 1_DOC_TITLE (value_list ([file_name, 1_DOC_TITLE])"
                        f":{value_list})"
                    )
                    print(text)
            rows.append(make_check_row(result, c_code, doc.doc_OD_style_file_name, page, text))
    return rows


def _od_legacy_row_sheet_count(page_format: Any) -> int:
    """Sheets for one OD row using legacy ``split()`` + ``get_Format_*`` (spaces only)."""
    if page_format is None or page_format == -1:
        return 0
    total = 0
    for raw in str(page_format).split():
        token = string_parsing.get_Format_from_51(raw)
        try:
            cnt = int(string_parsing.get_Format_Count(token))
        except (ValueError, IndexError, TypeError):
            return -1
        total += cnt
    return total


def _od_v2_row_sheet_count(page_format: Any) -> int:
    """Sheets for one OD row using the same rules as ``od_parsing._get_page_count``."""
    pc, _ = _get_page_count(page_format)
    if isinstance(pc, int):
        return pc
    if pc == "":
        return 0
    return -1


def _od_page_format_token_debug(page_format: Any) -> tuple[str, str]:
    """Human-readable token lists: (ОД-парсер / запятые, legacy split по пробелам)."""
    v2_parts, _ = _page_format_token_parts(page_format)
    if page_format is None or page_format == -1:
        legacy_raw: list[str] = []
    else:
        legacy_raw = [t for t in str(page_format).split() if t]
    return " | ".join(v2_parts) if v2_parts else "(пусто)", " | ".join(legacy_raw) if legacy_raw else "(пусто)"


def _od_single_whitespace_token_sheet_counts(raw: str) -> tuple[int, int]:
    """Legacy vs ОД-парсер sheet counts for one whitespace-delimited fragment."""
    leg = _od_legacy_row_sheet_count(raw)
    v2 = _od_v2_row_sheet_count(raw)
    return leg, v2


def _od_suggest_nk_token_after_commas(raw: str) -> str:
    """Insert space after commas when the next character is non-space (NK-friendly)."""
    return re.sub(r",(?=\S)", ", ", raw).strip()


def _od_format_token_fix_lines(page_format: Any) -> list[str]:
    """Per whitespace token where legacy count differs from ОД — concrete «было» / «должно быть»."""
    if page_format is None or page_format == -1:
        return []
    single_line = (
        str(page_format)
        .replace("\r\n", " ")
        .replace("\n", " ")
        .replace("\t", " ")
    )
    lines: list[str] = []
    for raw in single_line.split():
        if not raw:
            continue
        leg, v2 = _od_single_whitespace_token_sheet_counts(raw)
        if leg < 0 or v2 < 0:
            lines.append(f"«{raw}»: проверьте формат")
            continue
        if leg == v2:
            continue
        sug = _od_suggest_nk_token_after_commas(raw)
        if sug != raw:
            lines.append(f"«{raw}»→«{sug}»")
        else:
            lines.append(f"«{raw}»: НК{leg}≠ОД{v2}")
    return lines


def _od_nc_hints_for_page_format(page_format: Any) -> str:
    """Short NC-oriented hints where commas/spaces likely break legacy ``split`` + ``get_Format_*``."""
    if page_format is None or page_format == -1:
        return "Проверьте непустое значение поля «Формат»."
    single_line = (
        str(page_format)
        .replace("\r\n", " ")
        .replace("\n", " ")
        .replace("\t", " ")
    )
    hints: list[str] = []
    comma_tokens = [t for t in single_line.split() if "," in t]
    if comma_tokens:
        joined = ", ".join(f"«{t}»" for t in comma_tokens)
        hints.append(
            f"Токены с запятой внутри (одно «слово» у split по пробелам): {joined}. "
            f"`get_Format_from_51` удаляет запятые — несколько форматов склеиваются. "
            f"Разбейте на отдельные слова пробелами или поставьте пробел сразу после каждой запятой."
        )
    else:
        if re.search(r",[АA]\d", single_line):
            hints.append(
                "Встречается «,A…» / «,А…» без пробела после запятой между обозначениями — "
                "добавьте пробел (например «A1, A4»)."
            )
        if re.search(r",\d+[AАa]", single_line):
            hints.append(
                "Встречается «,число+формат» без пробела после запятой — добавьте пробел "
                "(например «2A0, 2A4x3»)."
            )
    if not hints:
        hints.append(
            "Сверьте токены «как в ОД» и «split по пробелам» выше; "
            "между обозначениями форматов должны быть явные пробелы или запятые с пробелами."
        )
    return " ".join(hints)


def _od_page_count_detail_text(
    d: DocOdAttributes,
    *,
    legacy_r: int,
    v2_r: int,
) -> str:
    """Compact body for one OD row (fits NK table cell; full detail in xlsx)."""
    fmt_repr = str(d.Page_Format).replace("\n", " ").strip()
    if len(fmt_repr) > 64:
        fmt_repr = fmt_repr[:61] + "…"
    fix_lines = _od_format_token_fix_lines(d.Page_Format)
    parts: list[str] = [f"Строка: {d.Doc_Title}"]
    if fix_lines:
        parts.append("Правка: " + " ".join(fix_lines))
    else:
        od_tok, leg_tok = _od_page_format_token_debug(d.Page_Format)
        parts.append(_od_nc_hints_for_page_format(d.Page_Format)[:140])
        parts.append(f"ОД:{od_tok} | split:{leg_tok}")
    parts.append(f"«Формат»: {fmt_repr}  (НК {legacy_r} ≠ ОД {v2_r} л.)")
    return "\n".join(parts)


def check_od_page_count(
    curr_proj: list[V2Document],
    proj_od_list: list | int,
    profile: RulesProfile,
    *,
    od_pdf_path: str = "",
) -> list[CheckRow]:
    rows: list[CheckRow] = []
    c_code = [1015, "Проверка ОД: кол-во листов (7==len(ОД_форматы)"]
    if proj_od_list == 0:
        return rows
    dict_od_formats: dict[str, list[Any]] = {}
    for doc in proj_od_list:
        if doc.doc_range == DOC_OSNOVNOI and "V" not in doc.Document_Revision:
            title = doc.Doc_Title
            dict_od_formats[title] = []
            format_list = doc.Page_Format.split()
            for x in format_list:
                x = string_parsing.get_Format_from_51(x)
                count = string_parsing.get_Format_Count(x)
                single_format = string_parsing.get_Format_Single(x)
                for i in range(count):
                    dict_od_formats[title].append(single_format)

    od_by_formats_count = 0

    for x in dict_od_formats:
        od_by_formats_count = od_by_formats_count + len(dict_od_formats[x])

    count_7 = 0
    doc_for_result = "No_Found"

    for doc in curr_proj:
        for page in doc.pages:
            if doc.doc_Type == "OD" and page.page_num == 1:
                doc_for_result = doc.doc_OD_style_file_name
                try:
                    count_7 = int(page.dict_attributes[c_7_Total_number_of_sheets][0])
                except Exception:
                    count_7 = -1

    if count_7 == -1:
        result = False
        text = "Ошибка, не найдено кол-во листов в поле 7"
    elif count_7 == od_by_formats_count:
        result = True
        text = "Кол-во листов совпадает"
    else:
        od_rows: list[DocOdAttributes] = [
            d
            for d in proj_od_list
            if d.doc_range == DOC_OSNOVNOI and "V" not in d.Document_Revision
        ]
        mismatch_docs: list[tuple[DocOdAttributes, int, int]] = []
        v2_total = 0
        v2_any_invalid = False
        legacy_row_any_invalid = False
        for d in od_rows:
            legacy_r = _od_legacy_row_sheet_count(d.Page_Format)
            v2_r = _od_v2_row_sheet_count(d.Page_Format)
            if legacy_r < 0:
                legacy_row_any_invalid = True
            if v2_r < 0:
                v2_any_invalid = True
            else:
                v2_total += v2_r
            if legacy_r >= 0 and v2_r >= 0 and legacy_r != v2_r:
                mismatch_docs.append((d, legacy_r, v2_r))

        row_pdf = str(od_pdf_path or "").strip()

        if not v2_any_invalid and count_7 == v2_total:
            prefix = (
                f"1015 предупреждение: штамп/ОД {count_7} л., старый НК {od_by_formats_count} л. "
                f"(исправьте «Формат» в строке ниже)."
            )
            if mismatch_docs:
                for d, lr, vr in mismatch_docs:
                    detail = _od_page_count_detail_text(d, legacy_r=lr, v2_r=vr)
                    text = f"{prefix}\n{detail}"
                    rows.append(
                        make_check_row(False, c_code, doc_for_result, "нет", text, pdf_path=row_pdf)
                    )
            elif od_by_formats_count != v2_total:
                text = (
                    f"{prefix}\n"
                    "Старый НК по словарю Doc_Title ≠ сумма по строкам — проверьте дубликаты шифра в ОД."
                )
                rows.append(
                    make_check_row(False, c_code, doc_for_result, "нет", text, pdf_path=row_pdf)
                )
        else:
            text_head = (
                f"1015: штамп {count_7} л. ≠ ст.НК {od_by_formats_count} л.; ОД {v2_total} л."
            )
            if v2_any_invalid:
                text_head += " Неразобранный формат в «Формат»."
            if legacy_row_any_invalid:
                text_head += " Ошибка legacy-разбора."

            if mismatch_docs:
                for d, lr, vr in mismatch_docs:
                    detail = _od_page_count_detail_text(d, legacy_r=lr, v2_r=vr)
                    full_text = f"{text_head}\n{detail}"
                    rows.append(
                        make_check_row(False, c_code, doc_for_result, "нет", full_text, pdf_path=row_pdf)
                    )
            else:
                rows.append(
                    make_check_row(False, c_code, doc_for_result, "нет", text_head, pdf_path=row_pdf)
                )

        return rows

    rows.append(make_check_row(result, c_code, doc_for_result, "нет", text))
    return rows


def check_pdf_extension(curr_proj: list[V2Document], profile: RulesProfile) -> list[CheckRow]:
    rows: list[CheckRow] = []
    c_code = [1016, 'Проверка расширения ".pdf" (50)']
    for doc in curr_proj:
        for page in doc.pages:
            try:
                att_50 = page.dict_attributes[c_50_File_Name_Stamp][0]
                result_raw = string_parsing.check_pdf_suffix(att_50)
            except IndexError:
                result_raw = "Ошибка IndexError"

            if result_raw == "True":
                text = f"Расширение файла корректное: {result_raw}"
                result = True
            else:
                text = f"Расширение файла в поле 50 НЕ корректное: {result_raw}"
                result = False
            rows.append(make_check_row(result, c_code, doc.doc_OD_style_file_name, page, text))
    return rows


def check_ifc_ifr(curr_proj: list[V2Document], profile: RulesProfile) -> list[CheckRow]:
    rows: list[CheckRow] = []
    c_code = [1017, "Проверка написания IFC и IFR (18_3)"]

    text_IFC = profile.text_ifc
    text_IFC_cirillic = "IFС - Выпущен для строительства"
    text_IFR = profile.text_ifr
    text_SUP = profile.text_sup
    text_CAN = profile.text_can

    rev_literal = profile.rev_literal
    rev_numeral = profile.rev_numeral

    for doc in curr_proj:
        for page in doc.pages:
            if page.page_num == 1:
                try:
                    att_18_3 = page.dict_attributes[c_18_3][0]
                except IndexError:
                    att_18_3 = "IndexError"
                att_18_3 = att_18_3.replace("\u2013", "-")
                try:
                    curr_revision = page.dict_attributes[c_18_1][0]
                except IndexError:
                    curr_revision = "не найдена"
                revision_for_issue = curr_revision
                if _STAMP_AN_TOKEN.search(curr_revision) and "-" in curr_revision:
                    revision_for_issue = curr_revision.split("-", 1)[0]
                if revision_for_issue in rev_literal:
                    if att_18_3 == text_IFR:
                        result = True
                        text = "Описание причины выпуска совпадает с номером ревизии"
                    else:
                        result = False
                        text = (
                            f"Некорректное название ревизии:\n"
                            f"    {att_18_3} | текущее название\n"
                            f"    {text_IFR} | должно быть"
                        )
                elif revision_for_issue in rev_numeral:
                    if att_18_3 == text_IFC:
                        result = True
                        text = "Описание причины выпуска совпадает с номером ревизии"
                    elif att_18_3 == text_IFC_cirillic:
                        result = False
                        text = (
                            f"Некорректное название ревизии: {att_18_3}. "
                            f"В <IFC> использована кириллическая <С>"
                        )
                    else:
                        result = False
                        text = (
                            f"Некорректное название ревизии:\n"
                            f"    {att_18_3} | текущее название\n"
                            f"    {text_IFC} | должно быть"
                        )
                elif revision_for_issue == "S":
                    if att_18_3 == text_SUP:
                        result = True
                        text = "Описание причины выпуска совпадает с номером ревизии"
                    else:
                        result = False
                        text = (
                            f"Некорректное название ревизии:\n"
                            f"    {att_18_3} | текущее название\n"
                            f"    {text_SUP} | должно быть"
                        )
                elif revision_for_issue == "V":
                    if att_18_3 == text_CAN:
                        result = True
                        text = "Описание причины выпуска совпадает с номером ревизии"
                    else:
                        result = False
                        text = (
                            f"Некорректное название ревизии:\n"
                            f"    {att_18_3} | текущее название\n"
                            f"    {text_CAN} | должно быть"
                        )
                else:
                    result = False
                    text = (
                        f"Ошибка: нет правила для сочетания ревизии <{curr_revision}> "
                        f"и описания {att_18_3}"
                    )

                rows.append(make_check_row(result, c_code, doc.doc_OD_style_file_name, page, text))
    return rows


def check_mark(curr_proj: list[V2Document], profile: RulesProfile) -> list[CheckRow]:
    rows: list[CheckRow] = []
    c_code = [1018, "Проверка марки (File_name==1==50)"]
    attr_list_to_check = [
        c_1_DOC_TITLE,
        c_50_File_Name_Stamp,
    ]
    for doc in curr_proj:
        for page in doc.pages:
            page_doc_mark: list[Any] = []
            try:
                page_doc_mark.append((doc.doc_OD_style_file_name, "В имени файла"))
            except:
                page_doc_mark.append(("марка не найдена", "В имени файла"))

            for attr in attr_list_to_check:
                try:
                    if attr == c_50_File_Name_Stamp:
                        doc_mark = getOdStyleFileName(page.dict_attributes[attr][0], mode="Clear")
                    else:
                        doc_mark = page.dict_attributes[attr][0]
                    page_doc_mark.append((doc_mark, attr))
                except IndexError:
                    page_doc_mark.append(("не найден", attr))
                except ValueError:
                    page_doc_mark.append(("не корректный", attr))

            result = True
            x = page_doc_mark[0][0]
            for i in range(len(page_doc_mark)):
                if x != page_doc_mark[i][0]:
                    result = False

            if result:
                text = f"Номера документов совпадают: {page_doc_mark}"
            else:
                text = "Титулы НЕ совпадают:"
                for i in range(len(page_doc_mark)):
                    text = (
                        f"{text}\n"
                        f"    {page_doc_mark[i][0]} ({page_doc_mark[i][1]}) "
                    )
            rows.append(make_check_row(result, c_code, doc.doc_OD_style_file_name, page, text))
    return rows


def check_duplicate_ext(curr_proj: list[V2Document], profile: RulesProfile) -> list[CheckRow]:
    rows: list[CheckRow] = []
    c_code = [1019, "Проверка задвоения расширений файлов (File_name==1==50)"]
    attr_list_to_check = [
        c_50_File_Name_Stamp,
    ]
    suff_list = (".doc", ".docx", ".xls", ".xlsx", ".dwg")
    for doc in curr_proj:
        for page in doc.pages:
            page_doc_mark: list[Any] = []
            try:
                page_doc_mark.append((doc.file_name, "Имя файла"))
            except:
                page_doc_mark.append(("имя файла не найдено", "Имя файла"))

            for attr in attr_list_to_check:
                try:
                    page_doc_mark.append((page.dict_attributes[attr][0], attr))
                except IndexError:
                    page_doc_mark.append(("не найден", attr))
                except ValueError:
                    page_doc_mark.append(("не корректный", attr))

            for i in range(len(page_doc_mark)):
                check_text = page_doc_mark[i][0]
                for suffix in suff_list:
                    find_suff = re.findall(suffix, check_text)
                    if len(find_suff) > 1:
                        result = False
                        text = (
                            f"Дублирование расширения файла:\n"
                            f"    {page_doc_mark[i][0]} ({find_suff})"
                        )
                        rows.append(make_check_row(result, c_code, doc.doc_OD_style_file_name, page, text))
    return rows
