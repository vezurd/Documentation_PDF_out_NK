import tkinter as tk
from tkinter import filedialog as fd

# Ссылка на главное окно приложения (устанавливается из main.py).
# Диалог как дочернее окно выводится на передний план при возврате фокуса.
_main_window = None


def set_main_window(window):
    """Устанавливает главное окно для диалогов выбора (вызывается из main.py)."""
    global _main_window
    _main_window = window


def _get_parent():
    """Возвращает родительское окно для диалога: главное окно или временное."""
    global _main_window
    if _main_window is not None and _main_window.winfo_exists():
        return _main_window
    root = tk.Tk()
    root.withdraw()
    return root


def getDirectory():
    try:
        parent = _get_parent()
        use_temp = parent is not _main_window
        if not use_temp:
            parent.lift()  # главное окно на передний план — диалог откроется поверх
        result = fd.askdirectory(
            master=parent,
            mustexist=True
        )
        if use_temp:
            parent.destroy()
    except FileNotFoundError as e:
        print(e)
        exit(0)
    return result


def getFile():
    result = None
    try:
        parent = _get_parent()
        use_temp = parent is not _main_window
        if not use_temp:
            parent.lift()
        result = fd.askopenfile(master=parent)
        if use_temp:
            parent.destroy()
        if result is not None:
            result = result.name
    except FileNotFoundError as e:
        print(e)
    except AttributeError as e:
        print(e)
    return result


if __name__ == "__main__":
    pdf_path = getDirectory()
    if pdf_path == "":
        print("No Directory")
    else:
        print(pdf_path)
