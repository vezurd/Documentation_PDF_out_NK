"""Entry point: ``python -m main_v2``."""

from __future__ import annotations

import sys

from PySide6.QtWidgets import QApplication

from main_v2.window import MainWindow


def main() -> None:
    """Starts the Qt event loop with :class:`MainWindow`."""
    app = QApplication(sys.argv)
    win = MainWindow()
    win.show()
    raise SystemExit(app.exec())


if __name__ == "__main__":
    main()
