"""Apply cached context-menu usage styles and record a chosen action.

I/O stays in :mod:`rd_catalog.context_menu_usage`. This module only paints
the already-computed bold set and increments an in-memory counter.
"""

from __future__ import annotations

from PySide6.QtCore import QPoint
from PySide6.QtGui import QAction, QFont
from PySide6.QtWidgets import QMenu

from rd_catalog.context_menu_usage import (
    action_key,
    get_context_menu_usage,
    normalize_action_label,
)


def apply_usage_styles(menu: QMenu, menu_id: str) -> None:
    """Bold frequent actions from the last flush. No disk access."""

    store = get_context_menu_usage()
    if store is None:
        return
    hot = store.hot_keys()
    if not hot:
        return
    for action in menu.actions():
        if action.isSeparator():
            continue
        label = normalize_action_label(action.text())
        if not label:
            continue
        if action_key(menu_id, label) not in hot:
            continue
        font = QFont(action.font())
        font.setBold(True)
        action.setFont(font)


def exec_tracked_menu(
    menu: QMenu,
    menu_id: str,
    global_pos: QPoint,
) -> QAction | None:
    """Show *menu*, bold hot items, and count the chosen command.

    Args:
        menu: Already populated context menu.
        menu_id: Stable menu identity (``MENU_KITS``, …).
        global_pos: ``viewport.mapToGlobal(position)``.

    Returns:
        The action passed to ``QMenu.exec``, or ``None`` if dismissed.
    """

    apply_usage_styles(menu, menu_id)
    chosen = menu.exec(global_pos)
    if chosen is None or chosen.isSeparator():
        return chosen
    label = normalize_action_label(chosen.text())
    if not label:
        return chosen
    store = get_context_menu_usage()
    if store is not None:
        store.note(menu_id, label)
    return chosen
