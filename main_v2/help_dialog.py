"""Simple read-only help viewer backed by ``help_descriptions``."""

from __future__ import annotations

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QComboBox,
    QDialog,
    QDialogButtonBox,
    QPlainTextEdit,
    QVBoxLayout,
    QWidget,
)

from help_descriptions import get_all_keys, get_help_description


def show_help_dialog(parent: QWidget | None, *, initial_key: str = "open_pdf_folder") -> None:
    """Opens a modal dialog listing help keys and showing the selected text.

    Args:
        parent: Parent widget for modality, if any.
        initial_key: Help dictionary key selected on open.
    """
    dlg = QDialog(parent)
    dlg.setWindowTitle("Справка")
    dlg.resize(560, 420)
    layout = QVBoxLayout(dlg)
    combo = QComboBox()
    keys = sorted(get_all_keys())
    for k in keys:
        combo.addItem(k, k)
    idx = combo.findData(initial_key)
    combo.setCurrentIndex(max(0, idx))
    view = QPlainTextEdit()
    view.setReadOnly(True)

    def _load(key: str) -> None:
        info = get_help_description(key)
        title = info.get("title", key)
        desc = info.get("description", "")
        icon = info.get("icon", "")
        prefix = f"{icon} {title}\n\n" if icon else f"{title}\n\n"
        view.setPlainText(f"{prefix}{desc}".strip())

    _load(str(combo.currentData() or initial_key))
    combo.currentIndexChanged.connect(
        lambda _i: _load(str(combo.currentData() or initial_key))
    )
    layout.addWidget(combo)
    layout.addWidget(view)
    buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Close)
    buttons.rejected.connect(dlg.reject)
    buttons.accepted.connect(dlg.accept)
    layout.addWidget(buttons)
    dlg.setAttribute(Qt.WidgetAttribute.WA_DeleteOnClose, True)
    dlg.exec()
