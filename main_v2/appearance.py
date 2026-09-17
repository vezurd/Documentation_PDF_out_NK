"""Qt Fusion palettes, light QSS, and QSettings persistence for main_v2 themes."""

from __future__ import annotations

from dataclasses import dataclass
from typing import TYPE_CHECKING

from PySide6.QtCore import QSettings
from PySide6.QtGui import QColor, QPalette
from PySide6.QtWidgets import QApplication, QStyleFactory

if TYPE_CHECKING:
    from PySide6.QtWidgets import QWidget

SETTINGS_ORGANIZATION = "Documentation_PDF_out_NK"
SETTINGS_APPLICATION = "MainV2"
THEME_SETTINGS_KEY = "ui/theme"


@dataclass(frozen=True)
class ThemeEntry:
    """One selectable UI theme."""

    theme_id: str
    label: str


THEME_ENTRIES: tuple[ThemeEntry, ...] = (
    ThemeEntry("system", "Системная"),
    ThemeEntry("light_high_contrast", "Светлая контрастная"),
    ThemeEntry("dark_high_contrast", "Тёмная контрастная"),
    ThemeEntry("dark_blue", "Тёмно-синяя"),
)


def available_themes() -> tuple[ThemeEntry, ...]:
    """Returns all built-in theme entries (id + localized label)."""
    return THEME_ENTRIES


def theme_qsettings() -> QSettings:
    """Returns a scoped :class:`QSettings` for main_v2 UI preferences."""
    return QSettings(SETTINGS_ORGANIZATION, SETTINGS_APPLICATION)


def read_saved_theme(default: str = "system") -> str:
    """Reads the persisted theme id, or ``default`` if missing or invalid.

    Args:
        default: Theme id used when nothing is stored or the value is unknown.

    Returns:
        A theme id string (possibly ``default``).
    """
    raw = theme_qsettings().value(THEME_SETTINGS_KEY, default)
    tid = default if raw is None else str(raw)
    valid = {e.theme_id for e in THEME_ENTRIES}
    return tid if tid in valid else default


def write_saved_theme(theme_id: str) -> None:
    """Persists ``theme_id`` under ``ui/theme``.

    Args:
        theme_id: Theme id to store.
    """
    theme_qsettings().setValue(THEME_SETTINGS_KEY, theme_id)


def theme_label(theme_id: str) -> str:
    """Returns the human-readable label for ``theme_id``, or the id if unknown.

    Args:
        theme_id: Internal theme id.

    Returns:
        Localized label suitable for menus or status text.
    """
    for entry in THEME_ENTRIES:
        if entry.theme_id == theme_id:
            return entry.label
    return theme_id


def _resolve_application(app_or_widget: QApplication | QWidget) -> QApplication:
    if isinstance(app_or_widget, QApplication):
        return app_or_widget
    app = QApplication.instance()
    if app is None:
        msg = "QApplication instance must exist before applying a theme."
        raise RuntimeError(msg)
    return app


def _apply_system_theme(app: QApplication) -> None:
    """Restores native style, clears app QSS, and resets palette from the style."""
    app.setStyleSheet("")
    app.setStyle(None)
    style = app.style()
    if style is not None:
        app.setPalette(style.standardPalette())
    else:
        app.setPalette(QPalette())


def _palette_light_high_contrast() -> QPalette:
    pal = QPalette()
    pal.setColor(QPalette.ColorRole.Window, QColor("#ffffff"))
    pal.setColor(QPalette.ColorRole.WindowText, QColor("#000000"))
    pal.setColor(QPalette.ColorRole.Base, QColor("#f5f5f5"))
    pal.setColor(QPalette.ColorRole.AlternateBase, QColor("#e8e8e8"))
    pal.setColor(QPalette.ColorRole.Text, QColor("#000000"))
    pal.setColor(QPalette.ColorRole.Button, QColor("#e0e0e0"))
    pal.setColor(QPalette.ColorRole.ButtonText, QColor("#000000"))
    pal.setColor(QPalette.ColorRole.Highlight, QColor("#0050aa"))
    pal.setColor(QPalette.ColorRole.HighlightedText, QColor("#ffffff"))
    pal.setColor(QPalette.ColorRole.Mid, QColor("#888888"))
    pal.setColor(QPalette.ColorRole.Light, QColor("#f8f8f8"))
    pal.setColor(QPalette.ColorRole.Dark, QColor("#404040"))
    return pal


def _palette_dark_high_contrast() -> QPalette:
    pal = QPalette()
    pal.setColor(QPalette.ColorRole.Window, QColor("#121212"))
    pal.setColor(QPalette.ColorRole.WindowText, QColor("#f5f5f5"))
    pal.setColor(QPalette.ColorRole.Base, QColor("#1e1e1e"))
    pal.setColor(QPalette.ColorRole.AlternateBase, QColor("#262626"))
    pal.setColor(QPalette.ColorRole.Text, QColor("#ffffff"))
    pal.setColor(QPalette.ColorRole.Button, QColor("#2d2d2d"))
    pal.setColor(QPalette.ColorRole.ButtonText, QColor("#ffffff"))
    pal.setColor(QPalette.ColorRole.Highlight, QColor("#ffcc00"))
    pal.setColor(QPalette.ColorRole.HighlightedText, QColor("#000000"))
    pal.setColor(QPalette.ColorRole.Mid, QColor("#666666"))
    pal.setColor(QPalette.ColorRole.Light, QColor("#3a3a3a"))
    pal.setColor(QPalette.ColorRole.Dark, QColor("#0a0a0a"))
    return pal


def _palette_dark_blue() -> QPalette:
    pal = QPalette()
    pal.setColor(QPalette.ColorRole.Window, QColor("#1a2332"))
    pal.setColor(QPalette.ColorRole.WindowText, QColor("#e8eef7"))
    pal.setColor(QPalette.ColorRole.Base, QColor("#152028"))
    pal.setColor(QPalette.ColorRole.AlternateBase, QColor("#1c2a3a"))
    pal.setColor(QPalette.ColorRole.Text, QColor("#dce5f5"))
    pal.setColor(QPalette.ColorRole.Button, QColor("#243044"))
    pal.setColor(QPalette.ColorRole.ButtonText, QColor("#eef4ff"))
    pal.setColor(QPalette.ColorRole.Highlight, QColor("#4a9eff"))
    pal.setColor(QPalette.ColorRole.HighlightedText, QColor("#ffffff"))
    pal.setColor(QPalette.ColorRole.Mid, QColor("#4a5a70"))
    pal.setColor(QPalette.ColorRole.Light, QColor("#3d4f66"))
    pal.setColor(QPalette.ColorRole.Dark, QColor("#0f141c"))
    return pal


def _qss_light_high_contrast() -> str:
    border = "#6a6a6a"
    btn = "#dedede"
    btn_hover = "#d0d0d0"
    btn_pressed = "#c0c0c0"
    tab_sel = "#cce4ff"
    menubar = "#f0f0f0"
    return _QSS_TEMPLATE % {
        "border": border,
        "btn": btn,
        "btn_hover": btn_hover,
        "btn_pressed": btn_pressed,
        "tab_selected": tab_sel,
        "menubar": menubar,
    }


def _qss_dark_high_contrast() -> str:
    border = "#666666"
    btn = "#333333"
    btn_hover = "#3d3d3d"
    btn_pressed = "#2a2a2a"
    tab_sel = "#3a3a3a"
    menubar = "#1a1a1a"
    return _QSS_TEMPLATE % {
        "border": border,
        "btn": btn,
        "btn_hover": btn_hover,
        "btn_pressed": btn_pressed,
        "tab_selected": tab_sel,
        "menubar": menubar,
    }


def _qss_dark_blue() -> str:
    border = "#4a5a70"
    btn = "#2a3848"
    btn_hover = "#32445a"
    btn_pressed = "#223040"
    tab_sel = "#2c3d52"
    menubar = "#1e2a3a"
    return _QSS_TEMPLATE % {
        "border": border,
        "btn": btn,
        "btn_hover": btn_hover,
        "btn_pressed": btn_pressed,
        "tab_selected": tab_sel,
        "menubar": menubar,
    }


_QSS_TEMPLATE = """
QGroupBox {
    font-weight: bold;
    border: 1px solid %(border)s;
    border-radius: 4px;
    margin-top: 10px;
    padding-top: 8px;
}
QGroupBox::title {
    subcontrol-origin: margin;
    left: 8px;
    padding: 0 4px;
}
QPushButton {
    border: 1px solid %(border)s;
    border-radius: 3px;
    padding: 4px 12px;
    background: %(btn)s;
    min-height: 18px;
}
QPushButton:hover {
    background: %(btn_hover)s;
}
QPushButton:pressed {
    background: %(btn_pressed)s;
}
QPlainTextEdit, QLineEdit, QTextEdit {
    border: 1px solid %(border)s;
    border-radius: 3px;
    padding: 4px;
}
QTabWidget::pane {
    border: 1px solid %(border)s;
    border-radius: 3px;
    top: -1px;
}
QTabBar::tab:selected {
    background: %(tab_selected)s;
}
QMenuBar {
    background: %(menubar)s;
    border-bottom: 1px solid %(border)s;
}
QStatusBar {
    border-top: 1px solid %(border)s;
}
"""


def apply_theme(app_or_widget: QApplication | QWidget, theme_id: str) -> str:
    """Applies a built-in theme to the shared :class:`QApplication`.

    For ``system``, restores the platform default style and palette and clears
    application-level QSS. For other ids, forces the Fusion style, sets a
    high-contrast palette, and applies a small QSS polish for common widgets.

    Args:
        app_or_widget: Running application or any widget under that application.
        theme_id: One of :func:`available_themes` ids.

    Returns:
        The theme id that was actually applied (falls back to ``system`` when
        ``theme_id`` is unknown).
    """
    valid = {e.theme_id for e in THEME_ENTRIES}
    resolved = theme_id if theme_id in valid else "system"
    app = _resolve_application(app_or_widget)

    if resolved == "system":
        _apply_system_theme(app)
        return resolved

    fusion = QStyleFactory.create("Fusion")
    if fusion is not None:
        app.setStyle(fusion)

    if resolved == "light_high_contrast":
        app.setPalette(_palette_light_high_contrast())
        app.setStyleSheet(_qss_light_high_contrast())
    elif resolved == "dark_high_contrast":
        app.setPalette(_palette_dark_high_contrast())
        app.setStyleSheet(_qss_dark_high_contrast())
    elif resolved == "dark_blue":
        app.setPalette(_palette_dark_blue())
        app.setStyleSheet(_qss_dark_blue())
    else:
        _apply_system_theme(app)
        return "system"

    return resolved
