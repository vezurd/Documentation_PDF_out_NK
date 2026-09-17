"""Entry point for the PDF v2 control center."""

from __future__ import annotations

import argparse
import os
import sys


def _ensure_project_root() -> None:
    root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    if root not in sys.path:
        sys.path.insert(0, root)


def main() -> int:
    _ensure_project_root()

    parser = argparse.ArgumentParser(description="PDF v2 Monitor")
    parser.add_argument("pdf_path", nargs="?", default=None, help="Папка с PDF файлами")
    parser.add_argument("--project", default=None, help="Имя проекта (catalog filter)")
    parser.add_argument(
        "--tab",
        default="run",
        choices=[
            "run",
            "settings",
            "monitor",
            "reports",
            "normcontrol",
            "nk",
            "od",
            "tags",
            "tag_analysis",
        ],
        help="Какая вкладка открывается первой (normcontrol / nk = НК; od = ОД; tags / tag_analysis = Анализ тегов)",
    )
    args = parser.parse_args()

    from PySide6.QtWidgets import QApplication

    app = QApplication.instance() or QApplication(sys.argv)

    from pdf_parsing_v2.v2_config import load_v2_config
    from pdf_v2_monitor.monitor_window import MonitorWindow, load_monitor_app_icon

    _icon = load_monitor_app_icon()
    if _icon is not None:
        app.setWindowIcon(_icon)

    cfg = load_v2_config()
    pdf_path = args.pdf_path
    project = args.project or cfg.get("project") or None

    if pdf_path and not os.path.isdir(pdf_path):
        pdf_path = ""

    win = MonitorWindow(pdf_path or "", cfg, project=project, initial_tab=args.tab)
    win.show()

    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
