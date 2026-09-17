from __future__ import annotations

from pathlib import Path

from prettytable import PrettyTable

import utils.path
from base.base_classes import RowStd, RowType
from base.tables_columns import *
from RFQ.tags_rfp_compare.mto_file_filter import filter_mto_files

debug_print_files_list = True

# UTF-8 report next to the DS workbook (same folder as analyze_ds_specification output).
GET_MTO_SPEC_MISMATCH_TXT = "GET_MTO_LIST_FROM_DS_spec_mismatch.txt"
# UTF-8: title-system keys with no MTO file under mto_path (for next run / adding files).
GET_MTO_MTO_NOT_FOUND_TXT = "GET_MTO_LIST_FROM_DS_mto_not_found.txt"
_NOT_FOUND_MARKER = "Файл МТО не найден"

# DS columns aligned with DS_VS_MTO_OUTPUT_COLUMNS_CONFIG (ds_vs_mto_excel_columns): no RFQ / 1С / supplier name.
_SPEC_MISMATCH_COLUMNS: tuple[tuple[str, str], ...] = (
    (DS_NAME, "Имя ДС"),
    (DS_NUMBER, "№ п/п"),
    (DS_TITLE, "Титул"),
    (DS_SYSTEM, "Раздел"),
    (DS_SPECIFICATION, "Линия/ TAG-Номер/ Спецификация"),
    (NAME, "Наименование Позиций Товара по РД"),
    (CODE, "Код РД"),
    (TYPE_MARK, "Технические требования (ГОСТ/ ТУ и др.)"),
    (UNITS, "Ед. изм."),
    (VALUES, "Кол-во"),
    (VENDOR, "Прим."),
)

_SPEC_MISMATCH_CELL_MAX_LEN = 48


def _format_spec_mismatch_cell(value, max_len: int = _SPEC_MISMATCH_CELL_MAX_LEN) -> str:
    if value is None:
        s = ""
    else:
        s = str(value).replace("\r\n", " ").replace("\n", " ").replace("\r", " ")
    if len(s) <= max_len:
        return s
    if max_len <= 1:
        return s[:max_len]
    return s[: max_len - 1] + "…"


def _resolve_error_report_dir(
    ds_std: list[RowStd],
    error_report_out_dir: str | None,
) -> str | None:
    if error_report_out_dir:
        return error_report_out_dir
    if not ds_std:
        return None
    fp = getattr(ds_std[0].t_com, "file_full_path", None) or ""
    if not fp:
        return None
    return str(Path(fp).parent)


def _write_spec_mismatch_report(
    bad_rows: list[tuple[int, RowStd, str, str, str]],
    out_dir: str,
    ds_source_path: str | None,
) -> str:
    """PrettyTable UTF-8 txt; each bad row: list index, row, specification, title_system, source xlsx."""
    table = PrettyTable()
    field_names = [h for _, h in _SPEC_MISMATCH_COLUMNS] + ["Диагностика"]
    table.field_names = field_names
    for fn in field_names:
        table.align[fn] = "l"
    table.max_width["Линия/ TAG-Номер/ Спецификация"] = 40
    table.max_width["Наименование Позиций Товара по РД"] = 36
    table.max_width["Диагностика"] = 56

    for row_index, row, specification, title_system, src in bad_rows:
        cells = [_format_spec_mismatch_cell(row.get_value(key)) for key, _ in _SPEC_MISMATCH_COLUMNS]
        diag = _format_spec_mismatch_cell(
            f"list_index_0based={row_index}; title_system={title_system!r}; source_xlsx={src!r}",
            max_len=120,
        )
        cells.append(diag)
        table.add_row(cells)

    header_lines = [
        "GET_MTO_LIST_FROM_DS: position_row where (DS_TITLE + '-' + DS_SYSTEM) is not a substring of DS_SPECIFICATION",
        f"DS workbook: {ds_source_path or ''}",
        f"bad_rows: {len(bad_rows)}",
        "",
    ]
    body = table.get_string()
    text_out = "\n".join(header_lines) + body + "\n"

    utils.path.make_dir(out_dir)
    out_path = str(Path(out_dir) / GET_MTO_SPEC_MISMATCH_TXT)
    Path(out_path).write_text(text_out, encoding="utf-8")
    print(text_out, end="")
    print(f"GET_MTO_LIST_FROM_DS: отчёт по {len(bad_rows)} строкам -> {out_path}")
    return out_path


def _write_mto_not_found_list(
    spec_dict: dict,
    out_dir: str,
    mto_path: str,
    ds_source_path: str | None,
    flat_structure: bool,
) -> str | None:
    """Write UTF-8 list of title-system keys with no MTO file match; one key per line."""
    missing = sorted(k for k, v in spec_dict.items() if v == _NOT_FOUND_MARKER)
    if not missing:
        return None

    lines = [
        "GET_MTO_LIST_FROM_DS: спецификации (ключ Титул-Раздел), для которых не найден файл МТО в каталоге закупки.",
        f"DS workbook: {ds_source_path or ''}",
        f"mto_path: {mto_path}",
        f"flat_structure: {flat_structure}",
        f"count: {len(missing)}",
        "",
        "Список ключей (по одному на строку) — добавьте соответствующие MTO в дерево или поправьте имена файлов:",
        "",
    ]
    lines.extend(missing)
    lines.append("")

    text_out = "\n".join(lines)
    utils.path.make_dir(out_dir)
    out_path = str(Path(out_dir) / GET_MTO_MTO_NOT_FOUND_TXT)
    Path(out_path).write_text(text_out, encoding="utf-8")
    print(
        f"GET_MTO_LIST_FROM_DS: не найдено МТО для {len(missing)} ключ(ей) — список: {out_path}"
    )
    return out_path


def get_mto_list_from_ds(
    ds_std: list[RowStd],
    mto_path,
    flat_structure: bool = False,
    *,
    error_report_out_dir: str | None = None,
    ds_source_path: str | None = None,
) -> dict:
    spec_list = set()
    bad_rows: list[tuple[int, RowStd, str, str, str]] = []

    if ds_std:
        for row_index, row in enumerate(ds_std):
            if row.row_type == RowType.position_row:
                specification = str(row.get_value(DS_SPECIFICATION)).strip()
                title = str(row.get_value(DS_TITLE)).strip()
                system = str(row.get_value(DS_SYSTEM)).strip()
                title_system = title + "-" + system
                if title_system in specification:
                    if specification not in spec_list:
                        spec_list.add(title_system)
                else:
                    src = getattr(row.t_com, "file_name", "") or ""
                    bad_rows.append((row_index, row, specification, title_system, src))

        if bad_rows:
            out_dir = _resolve_error_report_dir(ds_std, error_report_out_dir)
            report_path: str | None = None
            if out_dir:
                report_path = _write_spec_mismatch_report(bad_rows, out_dir, ds_source_path)
            first = bad_rows[0]
            msg = (
                "GET_MTO_LIST_FROM_DS: для одной или нескольких position_row "
                "(DS_TITLE + '-' + DS_SYSTEM) не входит в DS_SPECIFICATION как подстрока. "
                f"Всего таких строк: {len(bad_rows)}. "
                f"Первая: specification={first[2]!r}, title_system={first[3]!r}, "
                f"source_xlsx={first[4]!r}, list_index_0based={first[0]}."
            )
            if report_path:
                msg += f" Отчёт (PrettyTable): {report_path}"
            print(f"Error: GET_MTO_LIST_FROM_DS:\n\t{len(bad_rows)} строк(и) — см. отчёт или сообщение выше.\n\tПрограмма завершена")
            raise RuntimeError(msg) from None
    else:
        print("Error: GET_MTO_LIST_FROM_DS - список ДС пуст. Программа завершена")
        raise RuntimeError("GET_MTO_LIST_FROM_DS: список ДС пуст") from None

    # Массив для хранения всех Документов (файлов) текущего проекта
    curr_proj = filter_mto_files(mto_path, debug=False, flat_structure=flat_structure)
    if not curr_proj:
        print(f"По пути <{mto_path}> MTO файлы не найдены.")
        exit(0)

    spec_dict = {}
    for spec in spec_list:
        for document in curr_proj:
            spec_bom = str(spec).replace("BOM", "MTO")
            spec_bom = str(spec_bom).replace("DS", "MTO")
            if spec in document.file_name or spec_bom in document.file_name:
                if spec in spec_dict.keys():
                    continue
                spec_dict[spec] = document.file_full_path
        if spec not in spec_dict:
            spec_dict[spec] = _NOT_FOUND_MARKER

    out_dir = _resolve_error_report_dir(ds_std, error_report_out_dir)
    if out_dir:
        _write_mto_not_found_list(
            spec_dict, out_dir, str(mto_path), ds_source_path, flat_structure
        )

    ############################
    if debug_print_files_list:
        table = PrettyTable()
        table.field_names = ["Спецификация из ДС", "Проверяемая МТО из МТО для Закупки"]
        table.border = 0
        for fn in table.field_names:
            table.align[fn] = "l"
        for k, v in spec_dict.items():
            table.add_row([k, v])
        print(table)

    return spec_dict
