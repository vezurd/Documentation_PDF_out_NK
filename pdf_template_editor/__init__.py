"""
Редактор шаблонов штампа (PySide6). Этап F — отдельный процесс, не блокирует CTk GUI.

Запуск: ``python -m pdf_template_editor``
"""

from __future__ import annotations

import sys


def launch_template_editor(pdf_path: str | None = None) -> None:
    """Launch the PySide6 template editor as a standalone Qt application."""
    from PySide6.QtWidgets import QApplication
    from pdf_parsing_v2.v2_config import load_v2_config, resolve_templates_dir

    app = QApplication.instance()
    if app is None:
        app = QApplication(sys.argv)

    from .main_window import TemplateEditorWindow, load_template_editor_app_icon

    _icon = load_template_editor_app_icon()
    if _icon is not None:
        app.setWindowIcon(_icon)

    cfg = load_v2_config()
    cfg["templates_dir"] = resolve_templates_dir(cfg)

    win = TemplateEditorWindow(cfg=cfg)
    win.show()
    if pdf_path and hasattr(win, "_open_pdf"):
        try:
            win._open_pdf(pdf_path)
        except Exception:
            pass

    sys.exit(app.exec())
