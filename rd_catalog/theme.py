"""Light high-contrast theme adapter for the RD catalog."""

from __future__ import annotations

from PySide6.QtCore import Qt
from PySide6.QtWidgets import QApplication

from main_v2.appearance import apply_theme


def apply_catalog_theme(app: QApplication | None = None) -> str:
    """Apply the shared light high-contrast application theme."""

    qapp = app or QApplication.instance()
    if qapp is None:
        return "light_high_contrast"
    applied = apply_theme(qapp, "light_high_contrast")
    qapp.styleHints().setColorScheme(Qt.ColorScheme.Light)
    return applied
