"""Save an exact Outlook item identified by drag-and-drop MAPI IDs."""

from __future__ import annotations

from pathlib import Path


def save_outlook_item(
    *,
    entry_id: str,
    store_id: str,
    destination: str | Path,
) -> Path:
    """Save an Outlook item as ``.msg`` without reading current selection.

    Args:
        entry_id: MAPI EntryID from ``RenPrivateMessages``.
        store_id: MAPI StoreID from ``RenPrivateMessages``.
        destination: Local output ``.msg`` path.

    Returns:
        The saved local path.

    Raises:
        RuntimeError: If Outlook automation is unavailable or saving fails.
    """

    try:
        import pythoncom
        import win32com.client
    except ImportError as exc:
        raise RuntimeError(
            "Для резервного чтения письма Outlook требуется pywin32."
        ) from exc

    path = Path(destination)
    path.parent.mkdir(parents=True, exist_ok=True)
    pythoncom.CoInitialize()
    try:
        try:
            try:
                outlook = win32com.client.GetActiveObject(
                    "Outlook.Application"
                )
            except Exception:
                outlook = win32com.client.Dispatch("Outlook.Application")
            namespace = outlook.GetNamespace("MAPI")
            item = namespace.GetItemFromID(entry_id, store_id)
            if item is None:
                raise RuntimeError("Outlook не нашёл перетаскиваемое письмо.")
            item.SaveAs(str(path), 9)  # 9 = Outlook olMSGUnicode
        except Exception as exc:
            raise RuntimeError(
                f"Не удалось сохранить перетаскиваемое письмо: {exc}"
            ) from exc
    finally:
        pythoncom.CoUninitialize()
    if not path.is_file() or path.stat().st_size == 0:
        raise RuntimeError("Outlook сохранил пустой файл письма.")
    return path
