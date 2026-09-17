"""Extra Step4 Excel sheets: match quality and UL allocation diagnosis."""

from __future__ import annotations

import re
from collections import Counter, defaultdict
from typing import Any

from base.base_classes import RowStd, RowType
from base.tables_columns import (
    CODE,
    CODE_MTO,
    DS_ACTUAL,
    DS_MANAGER,
    DS_NAME,
    MATCH_STATUS,
    MATCH_STATUS_VO,
    POSITION_STATUS,
    UL_COMPARE_STATUS,
    UL_DATA_STATUS,
    UL_TAG_MATCH_STATUS,
    UL_VALUES,
    VALUES,
    VALUES_MTO,
)
from utils.colors import Color

from RFQ.tags_rfp_compare.step4.step4_packing_compare import (
    FALLBACK_COMMENT_MARKER,
    STATUS_COMPLETE,
    STATUS_DATA_PROBLEM,
    STATUS_GEM_SUPPLY,
    STATUS_MTO_DELIVERED,
    STATUS_MTO_DELIVERED_SHORT,
    STATUS_MTO_ONLY,
    STATUS_OPEN_UPD,
    STATUS_OVERDELIVERY,
    STATUS_PACKING_ONLY,
    STATUS_EXCLUDED_FROM_SUPPLY,
    STATUS_RFP_ONLY_EXTRA,
    STATUS_SHORTFALL,
    STATUS_UNAVAILABLE,
    STATUS_VO_DELIVERED,
    STATUS_VO_DELIVERED_SHORT,
    STATUS_VO_ONLY,
    STATUS_ZIP_SMR_PNR,
    RfpPackingAudit,
)

_ORIGIN_RFP = "RFP"
_ORIGIN_MTO = "Добавлен из МТО"
_ORIGIN_VO = "Добавлен из VO"
_ORIGIN_UL = "Только в УЛ"
_ORIGIN_OTHER = "Прочее"
_ORIGIN_ORDER = (_ORIGIN_RFP, _ORIGIN_MTO, _ORIGIN_VO, _ORIGIN_UL, _ORIGIN_OTHER)
_EMPTY_UL_STATUS = "(пусто)"
_KNOWN_UL_STATUS_ORDER = (
    STATUS_SHORTFALL,
    STATUS_OPEN_UPD,
    STATUS_EXCLUDED_FROM_SUPPLY,
    STATUS_PACKING_ONLY,
    STATUS_COMPLETE,
    STATUS_OVERDELIVERY,
    STATUS_GEM_SUPPLY,
    STATUS_ZIP_SMR_PNR,
    STATUS_RFP_ONLY_EXTRA,
    STATUS_MTO_ONLY,
    STATUS_MTO_DELIVERED,
    STATUS_MTO_DELIVERED_SHORT,
    STATUS_VO_ONLY,
    STATUS_VO_DELIVERED,
    STATUS_VO_DELIVERED_SHORT,
    STATUS_DATA_PROBLEM,
    STATUS_UNAVAILABLE,
)
_DS_LABEL_NUMBER_RE = re.compile(r"(?:^|[ДD][СC])(\d+)", re.IGNORECASE)
_DS_STATS_HEADER_TOTAL_RE = re.compile(r"\s*\((\d+)\)\s*$")
_DS_STATS_OK_FILL = f"#{Color.match_matched}"
_DS_STATS_PROBLEM_FILL = f"#{Color.yellow}"


def _text(row: RowStd, column: str) -> str:
    if column not in row.el:
        return ""
    return str(row.get_value(column) or "").strip()


def _number(row: RowStd, column: str) -> float | None:
    if column not in row.el:
        return None
    value = row.get_value(column)
    if value is None or value == "":
        return None
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def _comment(row: RowStd, column: str) -> str:
    if column not in row.el:
        return ""
    return str(row.el[column].comment or "")


def _allocated_from_comment(comment: str) -> float | None:
    for line in comment.splitlines():
        if line.startswith("Распределено:"):
            raw = line.split(":", 1)[1].strip()
            token = raw.split()[0] if raw else ""
            try:
                return float(token.replace(",", "."))
            except ValueError:
                return None
    return None


def _row_origin(row: RowStd) -> str:
    status = _text(row, UL_COMPARE_STATUS)
    if status == STATUS_PACKING_ONLY:
        return _ORIGIN_UL
    if _text(row, MATCH_STATUS) == "Добавлен из МТО":
        return _ORIGIN_MTO
    if _text(row, MATCH_STATUS_VO) == "Добавлен из VO":
        return _ORIGIN_VO
    if _text(row, CODE):
        return _ORIGIN_RFP
    if status == STATUS_ZIP_SMR_PNR:
        return _ORIGIN_UL
    return _ORIGIN_OTHER


def _position_rows(result_rows: list[RowStd]) -> list[RowStd]:
    return [
        row
        for row in result_rows
        if row.row_type not in (RowType.empty_row, RowType.other_row)
    ]


def _ds_label_sort_key(label: str) -> tuple[int, int, int, str]:
    """Sort key for DS labels: empty last, then numeric ``ДС{n}``, then text.

    The tuple is homogeneous so ``ДС24`` never compares an ``int`` to a
    leftover junk label as ``str``.
    """
    text = label.strip()
    if not text:
        return (1, 0, 0, "")
    match = _DS_LABEL_NUMBER_RE.search(text)
    if match:
        return (0, 0, int(match.group(1)), text.casefold())
    digits = re.search(r"(\d+)", text)
    if digits:
        return (0, 0, int(digits.group(1)), text.casefold())
    return (0, 1, 0, text.casefold())


def _ds_group_sort_key(group: tuple[str, str, str]) -> tuple:
    actual, sequential, manager = group
    return (
        _ds_label_sort_key(actual),
        _ds_label_sort_key(sequential),
        (1, "") if not manager.strip() else (0, manager.casefold()),
    )


def build_ds_ul_status_stats(
    rows: list[RowStd],
) -> tuple[list[list[Any]], list[str], dict[str, int]]:
    """Build per-DS UL status count table rows, headers, and column totals.

    Args:
        rows: Step4 result rows (position rows are counted; empty/other skipped).

    Returns:
        Tuple of (data rows, header labels, status-name -> global total).
    """
    position_rows = _position_rows(rows)
    status_totals: Counter[str] = Counter()
    group_counts: dict[tuple[str, str, str], Counter[str]] = defaultdict(Counter)

    for row in position_rows:
        actual = _text(row, DS_ACTUAL)
        sequential = _text(row, DS_NAME)
        manager = _text(row, DS_MANAGER)
        raw_status = _text(row, UL_COMPARE_STATUS)
        status = raw_status or _EMPTY_UL_STATUS
        group_key = (actual, sequential, manager)
        group_counts[group_key][status] += 1
        status_totals[status] += 1

    ordered_statuses: list[str] = list(_KNOWN_UL_STATUS_ORDER)
    seen_statuses = set(status_totals)
    unexpected = sorted(
        seen_statuses - set(_KNOWN_UL_STATUS_ORDER) - {_EMPTY_UL_STATUS}
    )
    ordered_statuses.extend(unexpected)
    if _EMPTY_UL_STATUS in seen_statuses:
        ordered_statuses.append(_EMPTY_UL_STATUS)

    identity_headers = ["Фактический ДС", "Порядковый ДС", "Фамилия менеджера"]
    status_headers = [
        f"{status} ({status_totals.get(status, 0)})" for status in ordered_statuses
    ]
    headers = identity_headers + status_headers

    records: list[list[Any]] = []
    for group_key in sorted(group_counts, key=_ds_group_sort_key):
        actual, sequential, manager = group_key
        counts = group_counts[group_key]
        record: list[Any] = [actual, sequential, manager]
        record.extend(counts.get(status, 0) for status in ordered_statuses)
        records.append(record)

    totals = {status: status_totals.get(status, 0) for status in ordered_statuses}
    return records, headers, totals


def ds_stats_status_from_header(header: str) -> str:
    """Strip the book-total suffix from a DS stats status header.

    Args:
        header: Header like ``Поставка комплектна (30850)``.

    Returns:
        Status name without the `` (N)`` suffix.
    """
    return _DS_STATS_HEADER_TOTAL_RE.sub("", str(header or "")).strip()


def ds_stats_cell_tone(status: str, count: Any) -> str:
    """Return fill tone for a DS stats count cell.

    Zero means no issues of this type (green). A positive count is a
    problem (yellow), except ``Поставка комплектна`` which stays green.

    Args:
        status: UL status name without the header total suffix.
        count: Integer count for that status in the DS group.

    Returns:
        ``ok`` (green) or ``problem`` (yellow).
    """
    try:
        n = int(count)
    except (TypeError, ValueError):
        n = 0
    if n <= 0:
        return "ok"
    if status == STATUS_COMPLETE:
        return "ok"
    return "problem"


def ds_stats_row_has_problems(record: list[Any], headers: list[str]) -> bool:
    """Return True if any status count cell on the DS stats row is yellow.

    Args:
        record: One data row from ``build_ds_ul_status_stats``.
        headers: Sheet headers; status columns start at index 3.

    Returns:
        True if the row has at least one problem (yellow) count.
    """
    limit = min(len(record), len(headers))
    for col in range(3, limit):
        status = ds_stats_status_from_header(headers[col])
        if ds_stats_cell_tone(status, record[col]) == "problem":
            return True
    return False


def write_quality_sheets(
    workbook,
    result_rows: list[RowStd],
    packing_audit: RfpPackingAudit | None = None,
) -> None:
    """Write diagnostic sheets after the main Step4 worksheet.

    Args:
        workbook: Open xlsxwriter workbook.
        result_rows: Rows about to be (or already) written to the main sheet.
        packing_audit: Optional UL matcher audit (issues and loader counters).
    """
    rows = _position_rows(result_rows)
    header_fmt = workbook.add_format(
        {
            "bold": True,
            "bg_color": "#c9daf8",
            "border": 1,
            "text_wrap": True,
            "valign": "vcenter",
        }
    )
    cell_fmt = workbook.add_format({"border": 1, "valign": "vcenter"})
    note_fmt = workbook.add_format(
        {"border": 1, "valign": "vcenter", "text_wrap": True}
    )
    _write_summary_sheet(
        workbook, rows, packing_audit, header_fmt, cell_fmt, note_fmt
    )
    _write_ds_ul_status_stats_sheet(workbook, rows, header_fmt, cell_fmt)
    _write_landing_sheet(workbook, rows, header_fmt, cell_fmt)
    _write_status_counts_sheet(workbook, rows, header_fmt, cell_fmt)
    if packing_audit is not None and packing_audit.issues:
        _write_issues_sheet(workbook, packing_audit, header_fmt, cell_fmt, note_fmt)


def _write_table(
    worksheet,
    headers: list[str],
    records: list[list[Any]],
    header_fmt,
    cell_fmt,
    *,
    widths: list[int] | list[float] | None = None,
    note_fmt=None,
    note_col: int | None = None,
    int_cols: set[int] | None = None,
    int_fmt=None,
    cell_fmt_fn=None,
) -> None:
    worksheet.freeze_panes(1, 0)
    for col, header in enumerate(headers):
        worksheet.write(0, col, header, header_fmt)
        width = 18
        if widths and col < len(widths):
            width = widths[col]
        worksheet.set_column(col, col, width)
    for row_idx, record in enumerate(records, start=1):
        for col, value in enumerate(record):
            fmt = cell_fmt
            if cell_fmt_fn is not None:
                override = cell_fmt_fn(col, value, record)
                if override is not None:
                    fmt = override
            elif int_fmt is not None and int_cols is not None and col in int_cols:
                fmt = int_fmt
            elif note_fmt is not None and note_col is not None and col == note_col:
                fmt = note_fmt
            worksheet.write(row_idx, col, value, fmt)
    last_col = max(len(headers) - 1, 0)
    worksheet.autofilter(0, 0, max(len(records), 1), last_col)


def _write_summary_sheet(
    workbook,
    rows: list[RowStd],
    packing_audit: RfpPackingAudit | None,
    header_fmt,
    cell_fmt,
    note_fmt,
) -> None:
    ul_status = Counter(_text(row, UL_COMPARE_STATUS) or "(пусто)" for row in rows)
    complete_zero = 0
    complete_zero_fallback = 0
    complete_zero_empty_code = 0
    complete_zero_added_mto = 0
    fallback_any = 0
    for row in rows:
        comment = _comment(row, UL_COMPARE_STATUS)
        allocated = _allocated_from_comment(comment)
        has_fallback = FALLBACK_COMMENT_MARKER in comment
        if has_fallback:
            fallback_any += 1
        if _text(row, UL_COMPARE_STATUS) != STATUS_COMPLETE:
            continue
        delivered = _number(row, UL_VALUES)
        if allocated == 0 or delivered == 0:
            complete_zero += 1
            if has_fallback:
                complete_zero_fallback += 1
            if not _text(row, CODE):
                complete_zero_empty_code += 1
            if _text(row, MATCH_STATUS) == "Добавлен из МТО":
                complete_zero_added_mto += 1

    records = [
        ["Строк в Excel (без пустых)", len(rows), "Все position_row после УЛ"],
        [
            "Поставка комплектна",
            ul_status.get(STATUS_COMPLETE, 0),
            "остаток заказа 0, включая зелёные 0/0",
        ],
        [
            "комплектна и allocated=0",
            complete_zero,
            "очередь УЛ могла найтись, но списано 0 штук; комментарий на статусе",
        ],
        [
            "из них fallback CODE_MTO/VO и allocated=0",
            complete_zero_fallback,
            "пример: Fallback CODE_MTO, Распределено 0",
        ],
        [
            "из них пустой CODE RFP и allocated=0",
            complete_zero_empty_code,
            "типично «Добавлен из МТО»: qty в VALUES_MTO, снимка RFP нет",
        ],
        [
            "из них статус «Добавлен из МТО» и allocated=0",
            complete_zero_added_mto,
            "H1: не списывать очередь УЛ с пустым RFP CODE / VALUES",
        ],
        ["Любой fallback кода УЛ", fallback_any, "ключ очереди не по CODE RFP"],
        ["Недопоставка", ul_status.get(STATUS_SHORTFALL, 0), ""],
        ["Перепоставка", ul_status.get(STATUS_OVERDELIVERY, 0), ""],
        [
            STATUS_OPEN_UPD,
            ul_status.get(STATUS_OPEN_UPD, 0),
            "есть RFP и МТО, поставки УЛ нет",
        ],
        [
            "Поставка ГЭМ",
            ul_status.get(STATUS_GEM_SUPPLY, 0),
            "коды ГЭМ 8950, статус был «Неотгружено по УЛ»",
        ],
        [
            STATUS_ZIP_SMR_PNR,
            ul_status.get(STATUS_ZIP_SMR_PNR, 0),
            "комплект ЗИП для СМР/ПНР в Наименовании УЛ; статус был «Только в УЛ» или комплектна",
        ],
        [
            "Только в RFP (лишняя)",
            ul_status.get(STATUS_RFP_ONLY_EXTRA, 0),
            "есть RFP, нет МТО и УЛ",
        ],
        [
            STATUS_MTO_ONLY,
            ul_status.get(STATUS_MTO_ONLY, 0),
            "есть МТО, нет RFP и УЛ",
        ],
        [
            STATUS_MTO_DELIVERED,
            ul_status.get(STATUS_MTO_DELIVERED, 0),
            "leftover УЛ сел на строку без RFP, qty совпало",
        ],
        [
            STATUS_MTO_DELIVERED_SHORT,
            ul_status.get(STATUS_MTO_DELIVERED_SHORT, 0),
            "leftover УЛ сел частично",
        ],
        [
            "Только в РКД",
            ul_status.get(STATUS_VO_ONLY, 0),
            "есть VO, нет RFP/МТО и УЛ",
        ],
        [
            STATUS_VO_DELIVERED,
            ul_status.get(STATUS_VO_DELIVERED, 0),
            "leftover УЛ сел на строку без RFP, qty совпало",
        ],
        [
            STATUS_VO_DELIVERED_SHORT,
            ul_status.get(STATUS_VO_DELIVERED_SHORT, 0),
            "leftover УЛ сел частично",
        ],
        ["Только в УЛ", ul_status.get(STATUS_PACKING_ONLY, 0), "остаток очереди"],
        [
            "Проблема данных УЛ",
            ul_status.get(STATUS_DATA_PROBLEM, 0),
            "дробное qty УЛ, кривой титул, ед. изм.",
        ],
    ]
    if packing_audit is not None:
        stats = packing_audit.stats
        records.extend(
            [
                ["УЛ allocated_units (аудит)", stats.allocated_units, "сумма списанных штук"],
                [
                    "УЛ complete_zero_allocated (аудит)",
                    stats.complete_zero_allocated,
                    "счётчик matcher, без разбора комментария",
                ],
                ["УЛ fallback_code_matches (аудит)", stats.fallback_code_matches, ""],
                ["УЛ проблем в отчёте", len(packing_audit.issues), "см. лист Проблемы УЛ"],
            ]
        )
    ws = workbook.add_worksheet("Сводка")
    _write_table(
        ws,
        ["Показатель", "Значение", "Пояснение"],
        records,
        header_fmt,
        cell_fmt,
        widths=[42, 14, 70],
        note_fmt=note_fmt,
        note_col=2,
    )
    ws.set_row(0, 28)


def _ds_stats_column_widths(
    headers: list[str],
    records: list[list[Any]],
) -> list[float]:
    """Return Excel column widths for the DS UL-status stats sheet.

    Identity columns follow data/header length. Status headers wrap at
    row height 36 (~two lines), so width tracks half the header length
    with a floor for short names like «Недопоставка».
    """
    widths: list[float] = []
    identity_floors = (12.0, 18.0, 16.0)
    identity_caps = (16.0, 26.0, 22.0)
    for col, header in enumerate(headers):
        data_len = 0
        for record in records:
            if col < len(record):
                data_len = max(data_len, len(str(record[col] or "")))
        header_len = len(header)
        if col < 3:
            floor, cap = identity_floors[col], identity_caps[col]
            widths.append(min(cap, max(floor, header_len + 1, data_len + 2)))
            continue
        wrapped = max(14.0, min(26.0, header_len * 0.55 + 2.5))
        widths.append(max(wrapped, min(14.0, data_len + 2.0)))
    return widths


def _write_ds_ul_status_stats_sheet(
    workbook,
    rows: list[RowStd],
    header_fmt,
    cell_fmt,
) -> None:
    records, headers, _totals = build_ds_ul_status_stats(rows)
    ok_int = workbook.add_format(
        {
            "border": 1,
            "valign": "vcenter",
            "num_format": "0",
            "bg_color": _DS_STATS_OK_FILL,
        }
    )
    problem_int = workbook.add_format(
        {
            "border": 1,
            "valign": "vcenter",
            "num_format": "0",
            "bg_color": _DS_STATS_PROBLEM_FILL,
        }
    )
    ok_text = workbook.add_format(
        {
            "border": 1,
            "valign": "vcenter",
            "bg_color": _DS_STATS_OK_FILL,
        }
    )
    status_by_col = {
        col: ds_stats_status_from_header(headers[col])
        for col in range(3, len(headers))
    }

    def cell_fmt_fn(col: int, value: Any, record: list[Any]):
        if col < 3:
            if ds_stats_row_has_problems(record, headers):
                return None
            return ok_text
        status = status_by_col.get(col)
        if status is None:
            return None
        if ds_stats_cell_tone(status, value) == "problem":
            return problem_int
        return ok_int

    ws = workbook.add_worksheet("Статистика ДС")
    _write_table(
        ws,
        headers,
        records,
        header_fmt,
        cell_fmt,
        widths=_ds_stats_column_widths(headers, records),
        cell_fmt_fn=cell_fmt_fn,
    )
    ws.set_row(0, 40)


def _write_landing_sheet(workbook, rows: list[RowStd], header_fmt, cell_fmt) -> None:
    buckets: dict[str, dict[str, int]] = {
        origin: defaultdict(int) for origin in _ORIGIN_ORDER
    }
    for row in rows:
        origin = _row_origin(row)
        bucket = buckets.setdefault(origin, defaultdict(int))
        bucket["всего"] += 1
        ul_status = _text(row, UL_COMPARE_STATUS) or "(пусто)"
        bucket[ul_status] += 1
        comment = _comment(row, UL_COMPARE_STATUS)
        allocated = _allocated_from_comment(comment)
        has_fallback = FALLBACK_COMMENT_MARKER in comment
        if has_fallback:
            bucket["fallback"] += 1
        if ul_status == STATUS_COMPLETE and (allocated == 0 or _number(row, UL_VALUES) == 0):
            bucket["комплектна allocated=0"] += 1
            if has_fallback:
                bucket["комплектна fallback allocated=0"] += 1
        if not _text(row, CODE):
            bucket["пустой CODE"] += 1
        if _text(row, CODE_MTO):
            bucket["есть CODE_MTO"] += 1
        values = _number(row, VALUES)
        values_mto = _number(row, VALUES_MTO)
        if (values is None or values == 0) and values_mto:
            bucket["qty только VALUES_MTO"] += 1

    headers = [
        "Тип строки",
        "Всего",
        "Поставка комплектна",
        "комплектна allocated=0",
        "комплектна fallback allocated=0",
        "Недопоставка",
        "Перепоставка",
        STATUS_OPEN_UPD,
        STATUS_GEM_SUPPLY,
        STATUS_ZIP_SMR_PNR,
        "Только в RFP (лишняя)",
        STATUS_MTO_ONLY,
        STATUS_MTO_DELIVERED,
        STATUS_MTO_DELIVERED_SHORT,
        "Только в РКД",
        STATUS_VO_DELIVERED,
        STATUS_VO_DELIVERED_SHORT,
        "Только в УЛ",
        "Проблема данных УЛ",
        "Fallback",
        "Пустой CODE",
        "Есть CODE_MTO",
        "qty только VALUES_MTO",
    ]
    records = []
    for origin in _ORIGIN_ORDER:
        bucket = buckets.get(origin) or {}
        if not bucket.get("всего"):
            continue
        records.append(
            [
                origin,
                bucket.get("всего", 0),
                bucket.get(STATUS_COMPLETE, 0),
                bucket.get("комплектна allocated=0", 0),
                bucket.get("комплектна fallback allocated=0", 0),
                bucket.get(STATUS_SHORTFALL, 0),
                bucket.get(STATUS_OVERDELIVERY, 0),
                bucket.get(STATUS_OPEN_UPD, 0),
                bucket.get(STATUS_GEM_SUPPLY, 0),
                bucket.get(STATUS_ZIP_SMR_PNR, 0),
                bucket.get(STATUS_RFP_ONLY_EXTRA, 0),
                bucket.get(STATUS_MTO_ONLY, 0),
                bucket.get(STATUS_MTO_DELIVERED, 0),
                bucket.get(STATUS_MTO_DELIVERED_SHORT, 0),
                bucket.get(STATUS_VO_ONLY, 0),
                bucket.get(STATUS_VO_DELIVERED, 0),
                bucket.get(STATUS_VO_DELIVERED_SHORT, 0),
                bucket.get(STATUS_PACKING_ONLY, 0),
                bucket.get(STATUS_DATA_PROBLEM, 0),
                bucket.get("fallback", 0),
                bucket.get("пустой CODE", 0),
                bucket.get("есть CODE_MTO", 0),
                bucket.get("qty только VALUES_MTO", 0),
            ]
        )
    ws = workbook.add_worksheet("Посадка")
    _write_table(
        ws,
        headers,
        records,
        header_fmt,
        cell_fmt,
        widths=[22] + [16] * (len(headers) - 1),
    )
    ws.set_row(0, 36)


def _write_status_counts_sheet(workbook, rows: list[RowStd], header_fmt, cell_fmt) -> None:
    groups = (
        ("Статус УЛ", UL_COMPARE_STATUS),
        ("Статус тегов УЛ", UL_TAG_MATCH_STATUS),
        ("Качество данных УЛ", UL_DATA_STATUS),
        ("MTO, сравнение", MATCH_STATUS),
        ("VO, сравнение", MATCH_STATUS_VO),
        ("MTO, статус позиции", POSITION_STATUS),
    )
    records: list[list[Any]] = []
    for group_name, column in groups:
        counts: Counter[str] = Counter()
        for row in rows:
            counts[_text(row, column) or "(пусто)"] += 1
        for value, count in counts.most_common():
            records.append([group_name, value, count])
    ws = workbook.add_worksheet("Статусы")
    _write_table(
        ws,
        ["Группа", "Значение", "Строк"],
        records,
        header_fmt,
        cell_fmt,
        widths=[24, 40, 12],
    )


def _write_issues_sheet(
    workbook,
    packing_audit: RfpPackingAudit,
    header_fmt,
    cell_fmt,
    note_fmt,
) -> None:
    by_code: Counter[str] = Counter(issue.code for issue in packing_audit.issues)
    samples: dict[str, str] = {}
    for issue in packing_audit.issues:
        if issue.code not in samples:
            samples[issue.code] = issue.format_line()
    records = [
        [code, count, samples.get(code, "")]
        for code, count in by_code.most_common()
    ]
    ws = workbook.add_worksheet("Проблемы УЛ")
    _write_table(
        ws,
        ["Код проблемы", "Число", "Пример"],
        records,
        header_fmt,
        cell_fmt,
        widths=[28, 12, 90],
        note_fmt=note_fmt,
        note_col=2,
    )
    ws.set_row(0, 28)
