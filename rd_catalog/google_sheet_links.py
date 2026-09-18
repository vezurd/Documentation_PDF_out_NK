"""Deep links from catalog TRM / Google cells into Google Sheets.

Qt-free. Builds ``https://docs.google.com/spreadsheets/d/...`` URLs that
open a worksheet and highlight one A1 cell. Cached ``row_index`` values
come from the last Google snapshot, not a live A+B locate.
"""

from __future__ import annotations

import json
import os
import shlex
import subprocess
import sys
import webbrowser
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Mapping
from urllib.parse import parse_qs, urlparse

from rd_catalog.config import CatalogConfig
from rd_catalog.kits import GoogleKit, IssuanceKit

PINS_FILENAME = "google_sheet_pins.json"
GOOGLE_HREF_TIP = (
    "Shift+клик открывает эту ячейку в Google Sheets. "
    "Обычный клик и двойной клик — как раньше (выбор / копирование)."
)
KITS_SHEET_TITLE_FALLBACK = "Контроль выдачи"
_SHEETS_PREFIX = "https://docs.google.com/spreadsheets/d/"
_ISSUANCE_EXPORT_META = "google_issuance_meta.json"

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

    Google's «Get link to this cell» shape puts ``range`` in the
    **fragment** (``#gid=…&range=F12``). ``range`` in the query string is
    ignored. Never emit ``'Title'!A1``: quotes and ``!`` break Windows
    ``start`` (the URL becomes ``http://"https//…%22``). Without a gid
    the fragment is only ``#range=F12`` — the spreadsheet id already
    selects the workbook.

    Args:
        spreadsheet_id: Spreadsheet id from catalog config.
        a1: Cell such as ``F12``.
        sheet_id: Worksheet gid, or None.
        sheet_title: Unused in the URL; kept for callers that pass a pin.

    Returns:
        HTTPS edit URL, or ``""`` when the target is incomplete.
    """

    del sheet_title  # URL must not include 'Title'!A1 — Windows breaks on quotes/`!`.
    sid = (spreadsheet_id or "").strip()
    cell = (a1 or "").strip()
    if not sid or not cell:
        return ""
    if sheet_id is not None:
        gid = int(sheet_id)
        return (
            f"{_SHEETS_PREFIX}{sid}/edit?gid={gid}#gid={gid}&range={cell}"
        )
    return f"{_SHEETS_PREFIX}{sid}/edit#range={cell}"


def gid_from_google_url(url: str) -> int | None:
    """Return a worksheet gid from a Sheets URL query or fragment.

    Args:
        url: Export, edit, or redirect URL that may contain ``gid=``.

    Returns:
        Integer gid (including ``0``), or ``None``.
    """

    text = (url or "").strip()
    if not text:
        return None
    parsed = urlparse(text)
    for part in (parsed.query, parsed.fragment):
        if not part:
            continue
        for raw in parse_qs(part, keep_blank_values=True).get("gid") or []:
            try:
                return int(raw)
            except (TypeError, ValueError):
                continue
    return None


def is_google_sheets_url(url: str) -> bool:
    """Return whether ``url`` is a Google Sheets edit link we built."""

    text = (url or "").strip()
    return text.startswith(_SHEETS_PREFIX)


def open_google_sheet_url(url: str) -> bool:
    """Open a Sheets URL in the default browser.

    Windows ``os.startfile`` / unquoted ``start`` drop ``#range=`` or
    split on ``&``, so the spreadsheet opens without selecting the cell.
    Launch the HTTPS handler with the URL as one argv, else a quoted
    ``start``.

    Args:
        url: Value from :func:`google_sheet_cell_url`.

    Returns:
        True when the URL looked valid and a browser launch was attempted.
    """

    if not is_google_sheets_url(url):
        return False
    if sys.platform == "win32":
        _open_url_windows(url)
        return True
    webbrowser.open(url, new=2)
    return True


def windows_start_command(url: str) -> str:
    """Return a quoted ``start`` command that keeps ``#`` and ``&``.

    Args:
        url: HTTPS Sheets URL.

    Returns:
        ``cmd`` ``start`` line with the URL in double quotes.
    """

    cleaned = (url or "").replace('"', "").replace("'", "")
    return f'start "" "{cleaned}"'


def _open_url_windows(url: str) -> None:
    argv = _windows_https_argv(url)
    if argv:
        subprocess.Popen(argv, close_fds=True)
        return
    subprocess.Popen(windows_start_command(url), shell=True)


def _windows_https_argv(url: str) -> list[str] | None:
    command = _windows_https_open_command()
    if not command:
        return _windows_known_browser_argv(url)
    try:
        parts = shlex.split(command, posix=False)
    except ValueError:
        return _windows_known_browser_argv(url)
    if not parts:
        return _windows_known_browser_argv(url)
    exe = parts[0].strip('"')
    if not Path(exe).is_file():
        return _windows_known_browser_argv(url)
    lowered = command.casefold()
    if "rundll32" in exe.casefold() or "url.dll" in lowered:
        return _windows_known_browser_argv(url)
    args = [exe]
    replaced = False
    for part in parts[1:]:
        token = part.strip('"')
        if token in {"%1", "%~1"}:
            args.append(url)
            replaced = True
        elif "%1" in token:
            args.append(token.replace("%1", url))
            replaced = True
        else:
            args.append(token)
    if not replaced:
        args.append(url)
    return args


def _windows_https_open_command() -> str:
    try:
        import winreg
    except ImportError:
        return ""
    progid = ""
    try:
        with winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            r"Software\Microsoft\Windows\Shell\Associations"
            r"\UrlAssociations\https\UserChoice",
        ) as key:
            progid = str(winreg.QueryValueEx(key, "ProgId")[0] or "")
    except OSError:
        progid = ""
    if not progid:
        return ""
    try:
        with winreg.OpenKey(
            winreg.HKEY_CLASSES_ROOT,
            rf"{progid}\shell\open\command",
        ) as key:
            return str(winreg.QueryValueEx(key, None)[0] or "")
    except OSError:
        return ""


def _windows_known_browser_argv(url: str) -> list[str] | None:
    candidates = (
        Path(os.environ.get("PROGRAMFILES", r"C:\Program Files"))
        / "Google"
        / "Chrome"
        / "Application"
        / "chrome.exe",
        Path(os.environ.get("LOCALAPPDATA", ""))
        / "Google"
        / "Chrome"
        / "Application"
        / "chrome.exe",
        Path(os.environ.get("PROGRAMFILES", r"C:\Program Files"))
        / "Microsoft"
        / "Edge"
        / "Application"
        / "msedge.exe",
        Path(os.environ.get("PROGRAMFILES(X86)", r"C:\Program Files (x86)"))
        / "Microsoft"
        / "Edge"
        / "Application"
        / "msedge.exe",
    )
    for exe in candidates:
        if exe and exe.is_file():
            return [str(exe), url]
    return None


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

    return save_sheet_pin(
        runtime_dir,
        kind="kits",
        spreadsheet_id=spreadsheet_id,
        sheet_id=sheet_id,
        title=title,
    )


def save_issuance_sheet_pin(
    runtime_dir: str | Path,
    *,
    spreadsheet_id: str,
    sheet_id: int,
    title: str,
) -> bool:
    """Persist the «Выдача РД ПД» worksheet gid.

    Args:
        runtime_dir: Catalog runtime directory.
        spreadsheet_id: Spreadsheet id of the pin.
        sheet_id: Worksheet gid (``0`` is valid).
        title: Tab title from config or the live sheet.

    Returns:
        True when the JSON was written.
    """

    return save_sheet_pin(
        runtime_dir,
        kind="issuance",
        spreadsheet_id=spreadsheet_id,
        sheet_id=sheet_id,
        title=title,
    )


def save_sheet_pin(
    runtime_dir: str | Path,
    *,
    kind: str,
    spreadsheet_id: str,
    sheet_id: int,
    title: str,
) -> bool:
    """Persist one worksheet pin into ``google_sheet_pins.json``.

    Args:
        runtime_dir: Catalog runtime directory.
        kind: ``kits`` or ``issuance``.
        spreadsheet_id: Spreadsheet id of the pin.
        sheet_id: Worksheet gid (``0`` is valid).
        title: Live or configured tab title.

    Returns:
        True when the JSON was written.
    """

    if kind not in {"kits", "issuance"}:
        return False
    sid = (spreadsheet_id or "").strip()
    name = title or ""
    if not sid:
        return False
    if kind == "kits" and not name.strip():
        return False
    try:
        gid = int(sheet_id)
    except (TypeError, ValueError):
        return False
    path = pins_path(runtime_dir)
    payload = _read_pins_payload(path)
    payload[kind] = {
        "spreadsheet_id": sid,
        "sheet_id": gid,
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
    if issuance_gid is None:
        issuance_gid = _sheet_id_from_meta(
            Path(config.runtime_dir) / _ISSUANCE_EXPORT_META,
            spreadsheet_id=issuance_id,
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


def _sheet_id_from_meta(
    path: Path, *, spreadsheet_id: str = ""
) -> int | None:
    try:
        raw = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, ValueError, TypeError):
        return None
    if not isinstance(raw, dict):
        return None
    sid = str(raw.get("spreadsheet_id") or "").strip()
    expected = (spreadsheet_id or "").strip()
    if expected and sid and sid != expected:
        return None
    value = raw.get("sheet_id")
    try:
        if value is None:
            return None
        return int(value)
    except (TypeError, ValueError):
        return None


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
