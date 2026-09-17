"""Open Outlook Instant Search for an RD kit OD token.

Uses the desktop Outlook object model (same integrity level as the catalog).
``olSearchScopeAllFolders`` is the Instant Search scope «Все почтовые ящики»:
mail folders of the current type across stores selected for search. Do not use
``olSearchScopeAllOutlookItems`` (that is «Все элементы Outlook»).
"""

from __future__ import annotations

from typing import Any

OL_FOLDER_INBOX = 6
OL_MAIL_ITEM = 0
OL_SEARCH_SCOPE_ALL_FOLDERS = 1
OL_FOLDER_DISPLAY_NORMAL = 0
_SW_RESTORE = 9


def open_outlook_instant_search(query: str) -> None:
    """Connect to Outlook and run Instant Search across all mailboxes.

    Args:
        query: Instant Search text, including quotes, e.g. ``"1600-SOT.OD"``.

    Raises:
        RuntimeError: Empty query, missing pywin32, Instant Search disabled,
            or Outlook automation failure.
    """

    query_text = (query or "").strip()
    if not query_text:
        raise RuntimeError("Пустая строка поиска Outlook.")
    try:
        import pythoncom
        import win32com.client
    except ImportError as exc:
        raise RuntimeError("Для поиска в Outlook требуется pywin32.") from exc

    pythoncom.CoInitialize()
    try:
        try:
            try:
                outlook = win32com.client.GetActiveObject("Outlook.Application")
            except Exception:
                outlook = win32com.client.Dispatch("Outlook.Application")
            run_outlook_instant_search(outlook, query_text)
        except RuntimeError:
            raise
        except Exception as exc:
            raise RuntimeError(
                "Не удалось открыть поиск Outlook. "
                "Каталог и Outlook должны быть с одинаковыми правами Windows. "
                f"{exc}"
            ) from exc
    finally:
        pythoncom.CoUninitialize()


def run_outlook_instant_search(outlook: Any, query: str) -> None:
    """Run Instant Search on an existing Outlook.Application COM object.

    Args:
        outlook: ``Outlook.Application`` (or a test double).
        query: Instant Search text, including quotes.

    Raises:
        RuntimeError: Instant Search is off, or Outlook has no explorer window.
    """

    namespace = outlook.GetNamespace("MAPI")
    inbox = namespace.GetDefaultFolder(OL_FOLDER_INBOX)
    _require_instant_search(namespace)
    explorer = _mail_explorer(outlook, inbox)
    explorer.Search(query, OL_SEARCH_SCOPE_ALL_FOLDERS)
    _activate_explorer(explorer)


def _require_instant_search(namespace: Any) -> None:
    try:
        store = namespace.DefaultStore
        enabled = bool(store.IsInstantSearchEnabled)
    except RuntimeError:
        raise
    except Exception:
        return
    if not enabled:
        raise RuntimeError(
            "В Outlook выключен мгновенный поиск (Instant Search)."
        )


def _mail_explorer(outlook: Any, inbox: Any) -> Any:
    explorer = outlook.ActiveExplorer()
    created = False
    if explorer is None:
        explorers = outlook.Explorers
        count = int(getattr(explorers, "Count", 0) or 0)
        if count > 0:
            explorer = explorers.Item(1)
        else:
            explorer = explorers.Add(inbox, OL_FOLDER_DISPLAY_NORMAL)
            created = True
    if explorer is None:
        raise RuntimeError("Outlook не открыл окно поиска.")
    current = getattr(explorer, "CurrentFolder", None)
    item_type = (
        getattr(current, "DefaultItemType", None)
        if current is not None
        else None
    )
    if item_type != OL_MAIL_ITEM:
        explorer.CurrentFolder = inbox
    if created:
        display = getattr(explorer, "Display", None)
        if callable(display):
            display()
    return explorer


def _activate_explorer(explorer: Any) -> None:
    for name in ("Activate", "Display"):
        method = getattr(explorer, name, None)
        if not callable(method):
            continue
        try:
            method()
            break
        except Exception:
            continue
    hwnd = getattr(explorer, "HWND", None)
    if not hwnd:
        return
    try:
        import ctypes

        handle = int(hwnd)
        ctypes.windll.user32.ShowWindow(handle, _SW_RESTORE)
        ctypes.windll.user32.SetForegroundWindow(handle)
    except Exception:
        return
