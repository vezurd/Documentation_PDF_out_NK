"""Local checks for Outlook Instant Search (no live Outlook)."""

from __future__ import annotations

import sys
from pathlib import Path
from unittest.mock import MagicMock, patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from rd_catalog.outlook_search import (
    OL_FOLDER_DISPLAY_NORMAL,
    OL_FOLDER_INBOX,
    OL_SEARCH_SCOPE_ALL_FOLDERS,
    open_outlook_instant_search,
    run_outlook_instant_search,
)


def _outlook(*, explorer: MagicMock | None, instant: bool = True) -> MagicMock:
    outlook = MagicMock()
    namespace = MagicMock()
    inbox = MagicMock(name="inbox")
    inbox.DefaultItemType = 0
    store = MagicMock()
    store.IsInstantSearchEnabled = instant
    namespace.DefaultStore = store
    namespace.GetDefaultFolder.return_value = inbox
    outlook.GetNamespace.return_value = namespace
    outlook.ActiveExplorer.return_value = explorer
    if explorer is not None:
        explorer.HWND = None
    outlook._inbox = inbox
    return outlook


def main() -> None:
    """Drive Instant Search mocks: All Mailboxes scope, mail folder, errors."""

    try:
        open_outlook_instant_search("  ")
        raise AssertionError("blank query must fail before COM")
    except RuntimeError as exc:
        assert "Пустая" in str(exc)

    query = '"6550-SOS.OD"'
    mail_folder = MagicMock(name="inbox_view")
    mail_folder.DefaultItemType = 0
    explorer = MagicMock()
    explorer.CurrentFolder = mail_folder
    outlook = _outlook(explorer=explorer)
    run_outlook_instant_search(outlook, query)
    outlook.GetNamespace.assert_called_once_with("MAPI")
    outlook.GetNamespace.return_value.GetDefaultFolder.assert_called_once_with(
        OL_FOLDER_INBOX
    )
    explorer.Search.assert_called_once_with(query, OL_SEARCH_SCOPE_ALL_FOLDERS)
    assert explorer.CurrentFolder is mail_folder
    explorer.Activate.assert_called()

    calendar = MagicMock(name="calendar")
    calendar.DefaultItemType = 1
    cal_explorer = MagicMock()
    cal_explorer.CurrentFolder = calendar
    cal_outlook = _outlook(explorer=cal_explorer)
    run_outlook_instant_search(cal_outlook, query)
    assert cal_explorer.CurrentFolder is cal_outlook._inbox

    explorers = MagicMock()
    explorers.Count = 0
    new_explorer = MagicMock()
    new_explorer.CurrentFolder = None
    new_explorer.HWND = None
    explorers.Add.return_value = new_explorer
    fresh = _outlook(explorer=None)
    fresh.Explorers = explorers
    run_outlook_instant_search(fresh, query)
    explorers.Add.assert_called_once_with(
        fresh._inbox, OL_FOLDER_DISPLAY_NORMAL
    )
    new_explorer.Display.assert_called()
    new_explorer.Search.assert_called_once_with(
        query, OL_SEARCH_SCOPE_ALL_FOLDERS
    )

    disabled = _outlook(explorer=explorer, instant=False)
    try:
        run_outlook_instant_search(disabled, query)
        raise AssertionError("disabled Instant Search must fail")
    except RuntimeError as exc:
        assert "Instant Search" in str(exc)

    pythoncom = MagicMock()
    client = MagicMock()
    live = _outlook(explorer=explorer)
    client.GetActiveObject.return_value = live
    with patch.dict(
        sys.modules,
        {
            "pythoncom": pythoncom,
            "win32com": MagicMock(client=client),
            "win32com.client": client,
        },
    ):
        open_outlook_instant_search(query)
    client.GetActiveObject.assert_called_once_with("Outlook.Application")
    pythoncom.CoInitialize.assert_called_once()
    pythoncom.CoUninitialize.assert_called_once()
    explorer.Search.assert_called_with(query, OL_SEARCH_SCOPE_ALL_FOLDERS)

    print("test_rd_catalog_outlook_search: ok")


if __name__ == "__main__":
    main()
