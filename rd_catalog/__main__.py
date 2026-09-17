"""Standalone entry point for the RD catalog application."""

from __future__ import annotations

import argparse
import sqlite3
import sys
from pathlib import Path

from PySide6.QtCore import QLockFile
from PySide6.QtWidgets import QApplication, QMessageBox

from rd_catalog.config import load_config
from rd_catalog.db import CatalogDatabase
from rd_catalog.theme import apply_catalog_theme
from rd_catalog.window import CatalogWindow


def _arguments(argv: list[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Каталог версий РД AGCC")
    parser.add_argument(
        "--config",
        type=Path,
        default=None,
        help="Пользовательский JSON override конфигурации",
    )
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    """Initialize local runtime state and start the standalone Qt window."""

    args = _arguments(list(sys.argv[1:] if argv is None else argv))
    app = QApplication.instance() or QApplication(sys.argv)
    app.setOrganizationName("Documentation_PDF_out_NK")
    app.setApplicationName("rd_catalog")
    apply_catalog_theme(app)

    try:
        config = load_config(args.config)
        config.runtime_dir.mkdir(parents=True, exist_ok=True)
    except (OSError, ValueError) as exc:
        QMessageBox.critical(None, "Каталог РД", f"Ошибка конфигурации:\n{exc}")
        return 2

    lock = QLockFile(str(config.runtime_dir / "rd_catalog.lock"))
    if not lock.tryLock(100):
        QMessageBox.warning(
            None,
            "Каталог РД",
            "Каталог РД уже запущен. Закройте существующее окно перед повторным запуском.",
        )
        return 1

    database = CatalogDatabase(config.db_path)
    try:
        database.initialize()
    except (OSError, RuntimeError, sqlite3.Error) as exc:
        QMessageBox.critical(None, "Каталог РД", f"Не удалось открыть локальную БД:\n{exc}")
        return 2

    window = CatalogWindow(config, database)
    window.show()
    exit_code = app.exec()
    lock.unlock()
    return exit_code


if __name__ == "__main__":
    raise SystemExit(main())
