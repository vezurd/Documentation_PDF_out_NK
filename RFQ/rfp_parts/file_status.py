"""Per-file OK / error status for RFP parts collection.

Written next to ``rfp_parts_net.xlsx`` as ``rfp_parts_file_status.json``.
The GUI table shows only critical (ERROR) remarks; tag duplicates and
tag-count mismatches live in ``Отчет по тегам - сбор частей.xlsx``.
Empty-CODE skipped rows live in ``Отчет по позициям без кода - сбор частей.xlsx``
(not in net; Launch copies the file into result_dir when present).
Older stamp folders (diagnostics.xlsx + sources.json) share the same payload.
"""

from __future__ import annotations

import json
import re
import shutil
from collections import defaultdict
from pathlib import Path
from typing import Any, Iterable

from RFQ.rfp_parts.parts_net_preflight import SOURCES_JSON_NAME

FILE_STATUS_JSON_NAME = "rfp_parts_file_status.json"
DIAGNOSTICS_XLSX_NAME = "rfp_parts_diagnostics.xlsx"
DUPLICATE_TAGS_XLSX_NAME = "Отчет по тегам - сбор частей.xlsx"
EMPTY_CODE_XLSX_NAME = "Отчет по позициям без кода - сбор частей.xlsx"

STATUS_OK = "OK"
STATUS_WARN = "WARN"
STATUS_ERROR = "ERROR"
STATUS_RANK = {STATUS_ERROR: 0, STATUS_WARN: 1, STATUS_OK: 2}
STATUS_LABEL_RU = {
    STATUS_OK: "ОК",
    STATUS_WARN: "Предупреждение",
    STATUS_ERROR: "Ошибка",
}

FIELD_LABELS_RU: dict[str, str] = {
    "DS_NUMBER": "№ п/п",
    "DS_TITLE": "Титул",
    "DS_SPECIFICATION": "Спецификация",
    "TAGS": "Tag",
    "DS_CODE_1C": "Код 1С",
    "CODE": "Код РД",
    "NAME_2": "Наименование (доп.)",
    "NAME": "Наименование МТР",
    "TYPE_MARK": "Технические характеристики",
    "VALUES_2": "Кол-во (доп.)",
    "VALUES": "Кол-во (лот)",
    "UNITS": "Ед. изм. (лот)",
    "VENDOR": "Поставщик",
    "RFP_SUPPLY_STATUS": "Статус поставки",
}

_FIELD_KEYS_DESC = tuple(sorted(FIELD_LABELS_RU, key=len, reverse=True))
_LEVEL_FROM_RU = {
    "ошибка": STATUS_ERROR,
    "предупреждение": STATUS_WARN,
    "информация": "INFO",
    "ok": STATUS_OK,
    "ошиб": STATUS_ERROR,
}
_ROW_PREFIX_RE = re.compile(r"^(?:row|строка)\s+(\d+)\s*:\s*(.*)$", re.IGNORECASE)
_SHEET_RE = re.compile(r"sheet\s+['\"]([^'\"]+)['\"]", re.IGNORECASE)
_TAG_REMARK_RE = re.compile(
    r"количество тегов|теги\s*=|дублир\w*\s+тег|тег\s+.+\s+дублир|"
    r"duplicate\s+tag|tags?\s*(?:≠|!=)|tag.?count",
    re.IGNORECASE,
)


def empty_file_status_payload() -> dict[str, Any]:
    """Empty GUI payload when no stamp folder is available."""
    return {"ok": 0, "warn": 0, "error": 0, "tag_remarks": 0, "files": []}


def is_tag_remark_message(message: str) -> bool:
    """Return True for duplicate-tag / tag-count mismatch texts (not GUI-critical)."""
    return bool(message) and bool(_TAG_REMARK_RE.search(message))


def status_label_ru(status: str) -> str:
    """Return the Russian status word for a payload status code."""
    return STATUS_LABEL_RU.get(status, status)


def format_field_list(fields: Iterable[str]) -> str:
    """Join internal field keys as Russian column titles."""
    labels = [FIELD_LABELS_RU.get(field.strip(), field.strip()) for field in fields]
    return ", ".join(label for label in labels if label)


def replace_field_tokens(text: str) -> str:
    """Replace ``NAME`` / ``CODE`` tokens with Russian column titles."""
    for key in _FIELD_KEYS_DESC:
        text = re.sub(rf"\b{re.escape(key)}\b", FIELD_LABELS_RU[key], text)
    return text


def _quote_sheet(name: str) -> str:
    return f"«{name}»"


def humanize_parts_message(message: str) -> str:
    """Turn an extract/diagnostic string into a Russian UI sentence.

    Accepts current Russian emits and older English diagnostics so a stamp
    folder without ``rfp_parts_file_status.json`` still fills the table.

    Args:
        message: Raw warning text (optionally with ``row N:`` / ``строка N:``).

    Returns:
        Russian text; already-Russian sentences are left intact aside from
        field-token substitution.
    """
    raw = (message or "").strip()
    if not raw:
        return ""
    row_match = _ROW_PREFIX_RE.match(raw)
    prefix = ""
    body = raw
    if row_match:
        prefix = f"строка {row_match.group(1)}: "
        body = row_match.group(2).strip()

    translated = _translate_message_body(body)
    return prefix + translated


def _translate_message_body(body: str) -> str:
    text = body.strip()

    match = re.fullmatch(r"required columns were not found:\s*(.+)", text, re.I)
    if match:
        fields = [part.strip() for part in match.group(1).split(",") if part.strip()]
        return f"не найдены обязательные столбцы: {format_field_list(fields)}"

    match = re.fullmatch(
        r"required sheet ['\"](.+)['\"] was not found in workbook", text, re.I
    )
    if match:
        return f"в книге нет листа {_quote_sheet(match.group(1))}"

    match = re.match(
        r"RFP header was not found in first 100 rows of sheet ['\"](.+)['\"](.*)",
        text,
        re.I,
    )
    if match:
        extra = match.group(2).strip()
        result = (
            "шапка RFP не найдена в первых 100 строках листа "
            f"{_quote_sheet(match.group(1))}"
        )
        return f"{result} {extra}".strip() if extra else result

    match = re.match(r"cannot open workbook:\s*(.*)", text, re.I)
    if match:
        return f"не удалось открыть книгу: {match.group(1)}"

    match = re.match(
        r"primary RFP columns differ from the expected template:\s*(.*)", text, re.I
    )
    if match:
        rest = replace_field_tokens(match.group(1))
        rest = rest.replace("expected", "ожидалось").replace("actual", "факт")
        return f"порядок основных столбцов RFP не совпадает с шаблоном: {rest}"

    match = re.match(
        r"skipped non-empty row with missing fields:\s*(.+)", text, re.I
    )
    if match:
        fields = [part.strip() for part in match.group(1).split(",") if part.strip()]
        return f"пропущена, нет полей: {format_field_list(fields)}"

    match = re.match(
        r"cannot parse Lot VALUES \(Excel №17\)=(.*)", text, re.I
    )
    if match:
        return f"не разобрать количество лота (Excel №17)={match.group(1)}"

    match = re.fullmatch(r"empty Lot UNITS \(Excel №18\)", text, re.I)
    if match:
        return "пустая единица измерения лота (Excel №18)"

    match = re.match(r"negative VALUES=(.*)", text, re.I)
    if match:
        return f"отрицательное количество={match.group(1)}"

    match = re.match(
        r"(.+):\s*openpyxl stopped while parsing sheet XML:\s*(.*)", text, re.I
    )
    if match:
        return (
            f"{match.group(1)}: openpyxl остановился при разборе XML листа: "
            f"{match.group(2)}"
        )

    sheet_match = _SHEET_RE.search(text)
    if sheet_match:
        text = text.replace(sheet_match.group(0), f"листа {_quote_sheet(sheet_match.group(1))}")
    return replace_field_tokens(text)


def _collapse_detail_texts(texts: list[str]) -> str:
    """Join unique texts; cluster ``строка N:`` variants of the same remainder."""
    row_groups: dict[str, list[int]] = defaultdict(list)
    other: list[str] = []
    seen_other: set[str] = set()
    for text in texts:
        match = _ROW_PREFIX_RE.match(text)
        if match:
            row_groups[match.group(2).strip()].append(int(match.group(1)))
            continue
        if text and text not in seen_other:
            seen_other.add(text)
            other.append(text)
    parts = list(other)
    for rest, rows in row_groups.items():
        unique_rows = sorted(dict.fromkeys(rows))
        shown = ", ".join(str(n) for n in unique_rows[:8])
        extra = f" и ещё {len(unique_rows) - 8}" if len(unique_rows) > 8 else ""
        if len(unique_rows) == 1:
            parts.append(f"строка {unique_rows[0]}: {rest}")
        else:
            parts.append(f"строки {shown}{extra}: {rest}")
    return "; ".join(parts)


def _gui_status_from_error_texts(error_texts: list[str]) -> str:
    return STATUS_ERROR if error_texts else STATUS_OK


def _ok_detail(rows: int) -> str:
    return f"{rows} позиций" if rows else "ОК"


def _strip_tag_remarks_from_detail(detail: str) -> str:
    if not detail:
        return ""
    parts = [
        part.strip()
        for part in re.split(r";\s*", detail)
        if part.strip() and not is_tag_remark_message(part)
    ]
    return "; ".join(parts)


def build_file_status_payload(
    file_stats: Iterable[Any],
    warnings: list[tuple[str, str, str]],
    *,
    tag_remarks: int | None = None,
) -> dict[str, Any]:
    """Build the GUI/JSON payload from extract stats and diagnostic tuples.

    The table lists every parts workbook. Status is **ERROR** or **OK** only:
    tag duplicates / tag-count mismatches are counted in ``tag_remarks`` and
    written to ``Отчет по тегам - сбор частей.xlsx``, not shown as GUI warnings.

    Args:
        file_stats: ``FileStats``-like objects (``kind``, ``file_name``,
            ``ds_name``, ``sheet``, ``rows``). Summary workbooks are skipped.
        warnings: ``(level, file_name, message)`` from the collection run.
        tag_remarks: Rows in the duplicate-tags workbook. ``None`` counts
            matching diagnostic messages instead.

    Returns:
        Dict with ``ok`` / ``error`` / ``tag_remarks`` counts and a ``files``
        list sorted errors first (A–Z), then OK (A–Z).
    """
    by_file: dict[str, list[tuple[str, str]]] = defaultdict(list)
    counted_tag_remarks = 0
    for level, file_name, message in warnings:
        if not file_name:
            continue
        human = humanize_parts_message(message)
        if is_tag_remark_message(message) or is_tag_remark_message(human):
            counted_tag_remarks += 1
        by_file[file_name].append((level, human))

    files: list[dict[str, Any]] = []
    for stats in file_stats:
        kind = str(getattr(stats, "kind", "") or "")
        if kind == "summary":
            continue
        file_name = str(getattr(stats, "file_name", "") or "")
        if not file_name:
            continue
        msgs = by_file.get(file_name, [])
        error_texts: list[str] = []
        seen: set[str] = set()
        for level, text in msgs:
            if level != STATUS_ERROR or not text or text in seen:
                continue
            if is_tag_remark_message(text):
                continue
            seen.add(text)
            error_texts.append(text)
        status = _gui_status_from_error_texts(error_texts)
        rows = int(getattr(stats, "rows", 0) or 0)
        detail = (
            _collapse_detail_texts(error_texts)
            if status == STATUS_ERROR
            else _ok_detail(rows)
        )
        files.append(
            {
                "file_name": file_name,
                "ds_name": str(getattr(stats, "ds_name", "") or ""),
                "sheet": str(getattr(stats, "sheet", "") or ""),
                "rows": rows,
                "status": status,
                "status_label": status_label_ru(status),
                "detail": detail,
            }
        )

    files.sort(
        key=lambda item: (
            STATUS_RANK.get(str(item["status"]), 9),
            str(item["file_name"]).casefold(),
        )
    )
    resolved_tags = (
        counted_tag_remarks if tag_remarks is None else int(tag_remarks)
    )
    return _counts_payload(files, tag_remarks=resolved_tags)


def _counts_payload(
    files: list[dict[str, Any]],
    *,
    tag_remarks: int = 0,
) -> dict[str, Any]:
    ok = sum(1 for item in files if item.get("status") == STATUS_OK)
    warn = sum(1 for item in files if item.get("status") == STATUS_WARN)
    error = sum(1 for item in files if item.get("status") == STATUS_ERROR)
    return {
        "ok": ok,
        "warn": warn,
        "error": error,
        "tag_remarks": int(tag_remarks or 0),
        "files": files,
    }


def write_file_status_json(out_dir: Path, payload: dict[str, Any]) -> Path:
    """Write ``rfp_parts_file_status.json`` next to the net workbook.

    Args:
        out_dir: Stamp folder of this collection run.
        payload: Output of ``build_file_status_payload``.

    Returns:
        Path to the written JSON file.
    """
    out_path = Path(out_dir) / FILE_STATUS_JSON_NAME
    out_path.write_text(
        json.dumps(payload, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    return out_path


def load_file_status_payload(run_dir: Path | None) -> dict[str, Any]:
    """Load status JSON, or reconstruct from sources + diagnostics.

    Args:
        run_dir: Stamp folder; ``None`` yields an empty payload.

    Returns:
        Payload dict. Missing/unreadable folders yield empty counts.
    """
    if run_dir is None:
        return empty_file_status_payload()
    folder = Path(run_dir)
    try:
        json_path = folder / FILE_STATUS_JSON_NAME
        if json_path.is_file():
            data = json.loads(json_path.read_text(encoding="utf-8"))
            if isinstance(data, dict) and isinstance(data.get("files"), list):
                return _normalize_loaded_payload(data)
        return reconstruct_file_status_payload(folder)
    except (OSError, json.JSONDecodeError, TypeError, ValueError):
        return empty_file_status_payload()


def resolve_duplicate_tags_xlsx(run_dir: Path | None) -> Path | None:
    """Return ``Отчет по тегам - сбор частей.xlsx`` in the stamp folder, if present."""
    if run_dir is None:
        return None
    path = Path(run_dir) / DUPLICATE_TAGS_XLSX_NAME
    try:
        return path if path.is_file() else None
    except OSError:
        return None


def resolve_empty_code_xlsx(run_dir: Path | None) -> Path | None:
    """Return ``Отчет по позициям без кода - сбор частей.xlsx`` if present.

    Args:
        run_dir: Stamp folder next to ``rfp_parts_net.xlsx``.

    Returns:
        Path to the sidecar workbook, or ``None`` when missing or unreadable.
    """
    if run_dir is None:
        return None
    path = Path(run_dir) / EMPTY_CODE_XLSX_NAME
    try:
        return path if path.is_file() else None
    except OSError:
        return None


def copy_empty_code_report_to_result_dir(
    stamp_dir: Path | str | None,
    result_dir: Path | str | None,
) -> Path | None:
    """Copy the empty-CODE sidecar from a parts stamp into Launch result_dir.

    Missing source, a non-file path, or ``OSError`` is a silent no-op.

    Args:
        stamp_dir: Stamp folder next to the current ``rfp_parts_net.xlsx``.
        result_dir: Launch ``_результат_проверки_*`` folder.

    Returns:
        Destination path if the file was copied; ``None`` otherwise.
    """
    if stamp_dir is None or result_dir is None:
        return None
    try:
        src = Path(stamp_dir) / EMPTY_CODE_XLSX_NAME
        if not src.is_file():
            return None
        dest = Path(result_dir) / EMPTY_CODE_XLSX_NAME
        dest.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(src, dest)
        return dest
    except OSError:
        return None


def _normalize_loaded_payload(data: dict[str, Any]) -> dict[str, Any]:
    files: list[dict[str, Any]] = []
    for raw in data.get("files") or []:
        if not isinstance(raw, dict):
            continue
        status = str(raw.get("status") or STATUS_OK)
        rows = int(raw.get("rows") or 0)
        detail = str(raw.get("detail") or "")
        if status == STATUS_ERROR:
            detail = _strip_tag_remarks_from_detail(detail)
            if not detail:
                detail = "ошибка"
        else:
            # GUI is ERROR-or-OK only; leftover WARN rows become OK.
            status = STATUS_OK
            detail = _ok_detail(rows)
        files.append(
            {
                "file_name": str(raw.get("file_name") or ""),
                "ds_name": str(raw.get("ds_name") or ""),
                "sheet": str(raw.get("sheet") or ""),
                "rows": rows,
                "status": status,
                "status_label": status_label_ru(status),
                "detail": detail,
            }
        )
    files = [item for item in files if item["file_name"]]
    files.sort(
        key=lambda item: (
            STATUS_RANK.get(str(item["status"]), 9),
            str(item["file_name"]).casefold(),
        )
    )
    stored_tags = data.get("tag_remarks")
    try:
        tag_remarks = int(stored_tags or 0)
    except (TypeError, ValueError):
        tag_remarks = 0
    return _counts_payload(files, tag_remarks=tag_remarks)


def reconstruct_file_status_payload(run_dir: Path) -> dict[str, Any]:
    """Build a payload from ``rfp_parts_sources.json`` + diagnostics.xlsx.

    Used for stamp folders created before ``rfp_parts_file_status.json``.
    OK files have empty detail (row counts are not in those artifacts).
    """
    names = _source_file_names(run_dir)
    warnings = _diagnostics_as_warnings(run_dir)
    if not names and not warnings:
        return empty_file_status_payload()
    if not names:
        names = sorted({fname for _, fname, _ in warnings if fname})

    class _Stub:
        def __init__(self, file_name: str) -> None:
            self.kind = "parts"
            self.file_name = file_name
            self.ds_name = ""
            self.sheet = ""
            self.rows = 0

    stats = [_Stub(name) for name in names]
    return build_file_status_payload(stats, warnings)


def _source_file_names(run_dir: Path) -> list[str]:
    path = Path(run_dir) / SOURCES_JSON_NAME
    try:
        if not path.is_file():
            return []
        data = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError, TypeError, ValueError):
        return []
    files = data.get("files") if isinstance(data, dict) else None
    if not isinstance(files, list):
        return []
    names: list[str] = []
    for item in files:
        if isinstance(item, dict) and item.get("name"):
            names.append(str(item["name"]))
        elif isinstance(item, str) and item:
            names.append(item)
    return names


def _parse_diagnostic_level(raw: Any) -> str:
    text = str(raw or "").strip()
    upper = text.upper()
    if upper in {STATUS_ERROR, STATUS_WARN, "INFO"}:
        return upper
    folded = text.casefold()
    if folded in _LEVEL_FROM_RU:
        return _LEVEL_FROM_RU[folded]
    if "ошиб" in folded:
        return STATUS_ERROR
    if "предупр" in folded:
        return STATUS_WARN
    return STATUS_WARN


def _diagnostics_as_warnings(run_dir: Path) -> list[tuple[str, str, str]]:
    path = Path(run_dir) / DIAGNOSTICS_XLSX_NAME
    try:
        if not path.is_file():
            return []
        from openpyxl import load_workbook

        wb = load_workbook(path, read_only=True, data_only=True)
    except OSError:
        return []
    try:
        sheet_name = "diagnostics" if "diagnostics" in wb.sheetnames else wb.sheetnames[-1]
        ws = wb[sheet_name]
        rows = ws.iter_rows(min_row=2, values_only=True)
        warnings: list[tuple[str, str, str]] = []
        for row in rows:
            if not row:
                continue
            level = _parse_diagnostic_level(row[0] if len(row) > 0 else "")
            file_name = str(row[2] if len(row) > 2 else "").strip()
            row_number = str(row[3] if len(row) > 3 else "").strip()
            message = str(row[4] if len(row) > 4 else "").strip()
            if not file_name and not message:
                continue
            if row_number.isdigit() and message and not _ROW_PREFIX_RE.match(message):
                message = f"строка {row_number}: {message}"
            warnings.append((level, file_name, message))
        return warnings
    finally:
        wb.close()
