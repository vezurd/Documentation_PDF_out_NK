"""Qt bind helpers for ``monitor_views.MonitorCell``.

Clients copy DTO text/fill/tooltip onto widgets. They must not recompute
Auto MTO targets, `` AB`` suffixes, robot-origin fill, or Google labels.
"""

from __future__ import annotations

import html

from PySide6.QtCore import Qt
from PySide6.QtGui import QBrush, QColor
from PySide6.QtWidgets import QTableWidgetItem

from rd_catalog.monitor_views import MonitorCell

ROLE_SORT = Qt.ItemDataRole.UserRole + 1
_DEFAULT_FOREGROUND = "#202124"
_LIGHT_FOREGROUND = "#ffffff"


def _qt_tooltip(text: str) -> str:
    """Keep tabular summary tooltips aligned in a monospace tooltip."""

    if "\t" not in text and "Пути от папки передачи:" not in text:
        return text
    escaped = html.escape(text, quote=False)
    return (
        "<qt><pre style=\"font-family:'Consolas','Courier New',monospace;"
        "margin:0;tab-size:8\">"
        f"{escaped}</pre></qt>"
    )


def apply_monitor_cell(
    item: QTableWidgetItem,
    cell: MonitorCell,
    *,
    sort_role: int = ROLE_SORT,
) -> None:
    """Copy painted DTO fields onto a Qt table item.

    Args:
        item: Target table cell.
        cell: Join-layer cell (text/fill/tooltip).
        sort_role: Qt role for ``cell.sort_key``. Defaults to UserRole+1.
    """

    item.setText(cell.text)
    if cell.tooltip:
        item.setToolTip(_qt_tooltip(cell.tooltip))
    if cell.fill:
        color = QColor(cell.fill)
        if color.isValid():
            item.setBackground(QBrush(color))
            fg = cell.foreground or (
                _LIGHT_FOREGROUND if color.lightness() < 140 else _DEFAULT_FOREGROUND
            )
            item.setForeground(QBrush(QColor(fg)))
    elif cell.foreground:
        item.setForeground(QBrush(QColor(cell.foreground)))
    if cell.bold or cell.underline:
        font = item.font()
        font.setBold(cell.bold)
        font.setUnderline(cell.underline)
        item.setFont(font)
    if cell.sort_key is not None:
        item.setData(sort_role, cell.sort_key)
