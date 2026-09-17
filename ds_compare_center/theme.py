"""UI theme for ``ds_compare_center`` (reuses ``main_v2.appearance`` palettes)."""

from __future__ import annotations

from PySide6.QtCore import Qt
from PySide6.QtWidgets import QApplication

from main_v2.appearance import apply_theme

_DS_COMPARE_UI_THEME = "light_high_contrast"


def apply_ds_compare_ui_theme(app: QApplication | None = None) -> str:
    """Apply light high-contrast Fusion theme to the control center."""
    qapp = app or QApplication.instance()
    if qapp is None:
        return _DS_COMPARE_UI_THEME
    applied = apply_theme(qapp, _DS_COMPARE_UI_THEME)
    qapp.styleHints().setColorScheme(Qt.ColorScheme.Light)
    return applied
