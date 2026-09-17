"""Deep links from catalog TRM / Google cells into Google Sheets.

Qt-free. Builds ``https://docs.google.com/spreadsheets/d/...`` URLs that
open a worksheet and highlight one A1 cell. Cached ``row_index`` values
come from the last Google snapshot, not a live A+B locate.
"""

from __future__ import annotations

import json
import os
import subprocess
import sys
import webbrowser
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Mapping
from urllib.parse import quote

from rd_catalog.config import CatalogConfig
from rd_catalog.kits import GoogleKit, IssuanceKit

PINS_FILENAME = "google_sheet_pins.json"
GOOGLE_HREF_TIP = (
    "Клик открывает эту ячейку в Google Sheets. "
    "Двойной клик — копирование текста."
)
KITS_SHEET_TITLE_FALLBACK = "Контроль выдачи"
_SHEETS_PREFIX = "https://docs.google.com/spreadsheets/d/"

# КСБ ИД: A=title, B=mark, D=rev, E=status, F=journal (TRM lives in F).
KITS_GOOGLE_COLUMNS: dict[str, str] = {
    "Google · рев.": "D",
    "Google · статус": "E",
    "Google · дата F": "F",
    "Google · этап F": "F",
    "Google · рев. F": "F",
    "Google · TRM F": "F",
}

# Выдача РД ПД: B=send TRM, D=rev, E=title, F=status, G=date, P=incoming, Q=note.
KITS_ISSUANCE_COLUMNS: dict[str, str] = {
    "Выдача · рев.": "D",
    "Выдача · статус": "F",
    "Выдача · дата отпр.": "G",
    "Выдача · TRM отпр.": "B",
    "Выдача · дата вх.контр.": "P",
    "Выдача · TRM подтв.": "Q",
}

JOURNAL_ISSUANCE_COLUMNS: dict[str, str] = {
    "Рев.": "D",
    "Дата отпр.": "G",
    "TRM": "B",
    "Статус листа": "F",
    "Вх.контр.": "P",
    "TRM подтв.": "Q",
    "Примечание": "Q",
}


@dataclass(frozen=True, slots=True)
class SheetLinkContext:
    """Spreadsheet ids / titles / gids used to build cell URLs."""

    kits_spreadsheet_id: str
    kits_sheet_title: str
    kits_sheet_id: int | None
    issuance_spreadsheet_id: str
    issuance_sheet_title: str
    issuance_sheet_id: int | None = None


def pins_path(runtime_dir: str | Path) -> Path:
    """Return the runtime JSON path for cached worksheet pins."""

    return Path(runtime_dir) / PINS_FILENAME


def a1_cell(column_letter: str, row_index: int) -> str:
    """Return an A1 token such as ``F12``, or empty when the row is invalid.

    Args:
        column_letter: Sheet column (``B``, ``F``, ``Q``, …).
        row_index: 1-based sheet row from the last snapshot.

    Returns:
        ``{column}{row}`` or ``""``.
    """

    if int(row_index or 0) < 1:
        return ""
    column = (column_letter or "").strip().upper()
    if not column.isalpha():
        return ""
    return f"{column}{int(row_index)}"


def google_sheet_cell_url(
    spreadsheet_id: str,
    a1: str,
    *,
    sheet_id: int | None = None,
    sheet_title: str = "",
) -> str:
    """Return a browser URL that opens one Google Sheets cell.

    When ``sheet_id`` (gid) is known, the range is the bare A1 on that
    tab. Otherwise the range includes ``'Title'!A1`` so Google can switch
    sheets without a gid. ``range`` is always in the query string (and
    repeated in the hash) so Windows can drop the fragment and still
    land on the cell.

    Args:
        spreadsheet_id: Spreadsheet id from catalog config.
        a1: Cell such as ``F12``.
        sheet_id: Worksheet gid, or None.
        sheet_title: Live worksheet title (trailing space is significant).

    Returns:
        HTTPS edit URL, or ``""`` when the target is incomplete.
    """

    sid = (spreadsheet_id or "").strip()
    cell = (a1 or "").strip()
    if not sid or not cell:
        return ""
    # ``range`` must live in the query string. Windows often drops the
    # ``#fragment``, and ``webbrowser``/``start`` may split on ``&`` in the
    # hash, so ``#gid=…&range=F12`` opens the tab but never selects the cell.
    if sheet_id is not None:
        gid = int(sheet_id)
        query = f"gid={gid}&range={cell}"
        fragment = f"gid={gid}&range={cell}"
        return f"{_SHEETS_PREFIX}{sid}/edit?{query}#{fragment}"
    title = sheet_title or ""
    if title:
        escaped = title.replace("'", "''")
        range_body = f"'{escaped}'!{cell}"
    else:
        range_body = cell
    encoded = quote(range_body, safe="!:'")
    return f"{_SHEETS_PREFIX}{sid}/edit?range={encoded}#range={encoded}"


def is_google_sheets_url(url: str) -> bool:
    """Return whether ``url`` is a Google Sheets edit link we built."""

    text = (url or "").strip()
    return text.startswith(_SHEETS_PREFIX)


def open_google_sheet_url(url: str) -> bool:
    """Open a Sheets URL in the default browser.

    Args:
        url: Value from :func:`google_sheet_cell_url`.

    Returns:
        True when the URL looked valid and a browser launch was attempted.
    """

    if not is_google_sheets_url(url):
        return False
    if sys.platform == "win32":
        try:
            os.startfile(url)
            return True
        except OSError:
            subprocess.Popen(
                ["cmd", "/c", "start", "", url],
                close_fds=True,
            )
            return True
    webbrowser.open(url)
    return True


def save_kits_sheet_pin(
    runtime_dir: str | Path,
    *,
    spreadsheet_id: str,
    sheet_id: int,
    title: str,
) -> bool:
    """Persist the КСБ ИД worksheet gid and live title.

    Args:
        runtime_dir: Catalog runtime directory.
        spreadsheet_id: Spreadsheet id of the pin.
        sheet_id: Worksheet gid.
        title: Live tab title, including a trailing space when present.

    Returns:
        True when the JSON was written.
    """

    sid = (spreadsheet_id or "").strip()
    name = title or ""
    if not sid or not name.strip():
        return False
    path = pins_path(runtime_dir)
    payload = _read_pins_payload(path)
    payload["kits"] = {
        "spreadsheet_id": sid,
        "sheet_id": int(sheet_id),
        "title": name,
    }
    try:
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(
            json.dumps(payload, ensure_ascii=False, indent=2) + "\n",
            encoding="utf-8",
        )
    except OSError:
        return False
    return True


def sheet_link_context_from_config(config: CatalogConfig) -> SheetLinkContext:
    """Build link context from config plus an optional pin cache.

    Args:
        config: Resolved catalog configuration.

    Returns:
        Ids and titles for КСБ ИД and «Выдача РД ПД».
    """

    pins = _read_pins_payload(pins_path(config.runtime_dir))
    kits_pin = pins.get("kits") if isinstance(pins.get("kits"), dict) else {}
    issuance_pin = (
        pins.get("issuance") if isinstance(pins.get("issuance"), dict) else {}
    )
    kits_id = (config.google_kits_spreadsheet_id or "").strip()
    kits_title, kits_gid = _pin_or_config(
        kits_pin,
        spreadsheet_id=kits_id,
        config_title=config.google_kits_sheet_name,
        fallback_title=KITS_SHEET_TITLE_FALLBACK,
    )
    issuance_id = (config.google_issuance_spreadsheet_id or "").strip()
    issuance_title, issuance_gid = _pin_or_config(
        issuance_pin,
        spreadsheet_id=issuance_id,
        config_title=config.google_issuance_sheet_name,
        fallback_title="Выдача РД ПД",
    )
    return SheetLinkContext(
        kits_spreadsheet_id=kits_id,
        kits_sheet_title=kits_title,
        kits_sheet_id=kits_gid,
        issuance_spreadsheet_id=issuance_id,
        issuance_sheet_title=issuance_title,
        issuance_sheet_id=issuance_gid,
    )


def kits_google_hrefs(
    *,
    google: GoogleKit | None,
    issuance: IssuanceKit | None,
    links: SheetLinkContext,
) -> dict[str, str]:
    """Return Комплекты header → Sheets URL for Google-backed cells.

    Args:
        google: КСБ ИД kit, or None.
        issuance: Latest effective «Выдача РД ПД» send, or None.
        links: Spreadsheet ids and worksheet pin.

    Returns:
        Headers that have a jump URL. Empty TRM text still links when the
        sheet row is known.
    """

    result: dict[str, str] = {}
    if google is not None:
        for header, column in KITS_GOOGLE_COLUMNS.items():
            href = google_sheet_cell_url(
                links.kits_spreadsheet_id,
                a1_cell(column, google.row_index),
                sheet_id=links.kits_sheet_id,
                sheet_title=links.kits_sheet_title,
            )
            if href:
                result[header] = href
    if issuance is not None:
        for header, column in KITS_ISSUANCE_COLUMNS.items():
            href = google_sheet_cell_url(
                links.issuance_spreadsheet_id,
                a1_cell(column, issuance.row_index),
                sheet_id=links.issuance_sheet_id,
                sheet_title=links.issuance_sheet_title,
            )
            if href:
                result[header] = href
    return result


def journal_cell_href(
    header: str,
    sheet_row_index: int,
    links: SheetLinkContext,
) -> str:
    """Return a Sheets URL for one issuance-journal column.

    Args:
        header: Journal column title.
        sheet_row_index: 1-based «Выдача РД ПД» row, or 0 for orphans.
        links: Spreadsheet ids and worksheet pin.

    Returns:
        Jump URL, or ``""``.
    """

    column = JOURNAL_ISSUANCE_COLUMNS.get(header, "")
    if not column:
        return ""
    return google_sheet_cell_url(
        links.issuance_spreadsheet_id,
        a1_cell(column, sheet_row_index),
        sheet_id=links.issuance_sheet_id,
        sheet_title=links.issuance_sheet_title,
    )


def _read_pins_payload(path: Path) -> dict[str, Any]:
    try:
        raw = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, ValueError, TypeError):
        return {}
    return raw if isinstance(raw, dict) else {}


def _pin_or_config(
    pin: Mapping[str, Any],
    *,
    spreadsheet_id: str,
    config_title: str,
    fallback_title: str,
) -> tuple[str, int | None]:
    pin_sid = str(pin.get("spreadsheet_id") or "").strip()
    pin_title = str(pin.get("title") or "")
    pin_gid: int | None = None
    raw_gid = pin.get("sheet_id")
    try:
        if raw_gid is not None:
            pin_gid = int(raw_gid)
    except (TypeError, ValueError):
        pin_gid = None
    if pin_sid and spreadsheet_id and pin_sid != spreadsheet_id:
        pin_title = ""
        pin_gid = None
    title = pin_title or (config_title or "").strip() or fallback_title
    return title, pin_gid
