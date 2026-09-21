"""Entry point for the DS / RFP control center."""

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

    parser = argparse.ArgumentParser(description="DS / RFP / MTO Control Center")
    parser.add_argument(
        "--tab",
        default="misc",
        choices=[
            "run",
            "mto_paths",
            "settings",
            "columns",
            "packing",
            "tsd",
            "tsd_help",
            "packing_help",
            "upd",
            "upd_load",
            "rfp",
            "rfp_run",
            "rfp_settings",
            "rfp_parts",
            "parts",
            "rfp_ds_id",
            "ds_id",
            "rfp_ul",
            "rfp_ds_mp",
            "ds_mp",
            "rfp_mp",
            "rfp_managers",
            "misc",
            "mto_run",
            "bbb",
            "mto_settings",
            "vpn",
            "cursor_vpn",
        ],
        help=(
            "Initial tab (default: misc / Прочее · Запуск). DS: run, mto_paths, "
            "settings, columns, packing/tsd, tsd_help/packing_help, upd/upd_load; RFP: rfp/rfp_run, "
            "rfp_settings, rfp_parts/parts, rfp_ds_id/ds_id/rfp_ul, "
            "rfp_ds_mp/ds_mp/rfp_mp/rfp_managers; leftover main.py: "
            "misc/mto_run, bbb/mto_settings, vpn/cursor_vpn"
        ),
    )
    args = parser.parse_args()

    from PySide6.QtWidgets import QApplication

    app = QApplication.instance() or QApplication(sys.argv)

    from ds_compare_center.theme import apply_ds_compare_ui_theme

    apply_ds_compare_ui_theme(app)

    from ds_compare_center.center_window import CenterWindow

    win = CenterWindow(initial_tab=args.tab)
    win.show()
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
