"""
Полный пересчёт книги в Microsoft Excel (COM) для записи актуального кэша формул в xlsx.

Требования: Windows, установленный Excel, пакет xlwings (pip install xlwings).

Проверка вручную (BOE с формулами): после пайплайна «Для 1C» с включённым флагом
пересчёта открыть файл через openpyxl с data_only=True и сравнить долю None
в ячейках, где раньше были формулы, с прогоном без Excel-проходов.
"""

from __future__ import annotations

import os
import time


def recalculate_workbook_save(path: str, *, label: str = "") -> None:
    """
    Открывает xlsx в Excel, выполняет CalculateFullRebuild, сохраняет и закрывает книгу.

    :param path: путь к файлу
    :param label: подпись для лога (например «до замены формул»)
    :raises FileNotFoundError: нет файла
    :raises RuntimeError: нет xlwings / Excel недоступен
    """
    path = os.path.abspath(os.path.normpath(path))
    if not os.path.isfile(path):
        raise FileNotFoundError(path)

    try:
        import xlwings as xw
    except ImportError as e:
        raise RuntimeError(
            "Установите пакет xlwings (pip install xlwings) и Microsoft Excel для пересчёта книги."
        ) from e

    app = None
    wb = None
    t0 = time.perf_counter()
    try:
        try:
            app = xw.App(visible=False, add_book=False)
        except Exception as e:
            raise RuntimeError(
                "Не удалось запустить Microsoft Excel (нужен установленный Office, Windows)."
            ) from e
        app.display_alerts = False
        # False — не спрашивать об обновлении внешних связей (снижает риск зависаний)
        wb = app.books.open(path, update_links=False)
        wb.activate()
        app.api.CalculateFullRebuild()
        wb.save()
    finally:
        if wb is not None:
            try:
                wb.close()
            except Exception:
                pass
        if app is not None:
            try:
                app.quit()
            except Exception:
                pass

    elapsed = time.perf_counter() - t0
    suffix = f" ({label})" if label else ""
    print(f"[Excel]{suffix} Пересчёт и сохранение: {path} — {elapsed:.1f} с")
