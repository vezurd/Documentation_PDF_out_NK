from __future__ import annotations

from pdf_parsing_v2_engine.document import V2Document, V2FieldDiagnostic

from pdf_parsing_v2_rules.output import CheckRow

_CLEANING_CODES: dict[int, str] = {
    2100: "Ошибки очистки полей штампа",
    2101: "Неизвестный cleaner поля штампа",
    2102: "Исключение внутри cleaner поля штампа",
    2103: "Relaxed parse / fallback cleaning поля штампа",
}


def _check_code_for_diagnostic(diag: V2FieldDiagnostic) -> int:
    """Map one field diagnostic to a normcontrol check code."""
    if diag.warning_code == "UNKNOWN_CLEANER":
        return 2101
    if diag.warning_code == "CLEANER_EXCEPTION":
        return 2102
    if diag.warning_tier == "relaxed" or diag.clean_tier == "relaxed":
        return 2103
    return 2100


def _format_diagnostic_text(diag: V2FieldDiagnostic) -> str:
    """Build a human-readable message for the final normcontrol checklist."""
    return (
        f"{diag.field_id} ({diag.warning_code})\n"
        f"{diag.warning_message}\n"
        f"Прочитанная строка: {diag.raw_value!r}\n"
        f"После очистки: {diag.cleaned_value!r}"
    )


def build_extraction_check_rows(curr_proj: list[V2Document]) -> list[CheckRow]:
    """Convert field cleaning diagnostics into normcontrol checklist rows."""
    rows: list[CheckRow] = []
    for doc in curr_proj:
        for diag in doc.iter_field_diagnostics():
            c_code = _check_code_for_diagnostic(diag)
            rows.append(
                CheckRow(
                    result=False,
                    c_code=c_code,
                    c_description=_CLEANING_CODES[c_code],
                    doc_name=diag.doc_name,
                    page_num=diag.page_num,
                    text=_format_diagnostic_text(diag),
                )
            )
    return rows
