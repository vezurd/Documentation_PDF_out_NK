"""CLI: ``python -m rd_catalog_web --host 0.0.0.0 --port 8765``."""

from __future__ import annotations

import argparse
import sqlite3
import sys
from pathlib import Path

from rd_catalog_web.urls import (
    DEFAULT_HOST,
    DEFAULT_PORT,
    lan_share_url,
    listen_urls,
    local_open_url,
)


def _ensure_project_root() -> None:
    root = str(Path(__file__).resolve().parents[1])
    if root not in sys.path:
        sys.path.insert(0, root)


def _parse_args(argv: list[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Read-only LAN web monitor for the RD catalog. "
            "Same PC as the Qt writer; sqlite is local, not UNC."
        )
    )
    parser.add_argument(
        "--host",
        default=DEFAULT_HOST,
        help="Bind address (default 0.0.0.0, all interfaces; not a browser URL)",
    )
    parser.add_argument(
        "--port",
        type=int,
        default=DEFAULT_PORT,
        help="TCP port (default 8765)",
    )
    parser.add_argument(
        "--config",
        type=Path,
        default=None,
        help="Optional JSON override for load_config",
    )
    return parser.parse_args(argv)


def _is_unc(path: Path) -> bool:
    text = str(path)
    return text.startswith("\\\\") or text.startswith("//")


def main(argv: list[str] | None = None) -> int:
    """Load config, initialize sqlite, print LAN URLs, serve uvicorn.

    Args:
        argv: CLI arguments without the program name.

    Returns:
        Process exit code.
    """

    _ensure_project_root()
    args = _parse_args(list(sys.argv[1:] if argv is None else argv))
    try:
        import uvicorn
    except ImportError:
        print("pip install fastapi uvicorn", file=sys.stderr)
        return 2

    from rd_catalog_web.app import create_app

    try:
        app = create_app(args.config)
    except (OSError, ValueError, RuntimeError, sqlite3.Error) as exc:
        print(f"Не удалось запустить монитор: {exc}", file=sys.stderr)
        return 2

    config = app.state.config
    print("RD Catalog Web — только чтение, локальная сеть")
    print("Авторизации нет. Не выставляйте порт в интернет.")
    print(f"SQLite: {config.db_path}")
    if _is_unc(config.db_path):
        print(
            "Предупреждение: db_path похож на UNC/SMB. "
            "Нужен локальный файл %LOCALAPPDATA%\\...\\rd_catalog.sqlite.",
            file=sys.stderr,
        )
    print("Та же машина, что Qt-писатель каталога.")
    print(f"Этот ПК: {local_open_url(args.host, args.port)}")
    share = lan_share_url(args.host, args.port)
    print(f"Коллеги в LAN (не 0.0.0.0): {share}")
    extra = [
        url
        for url in listen_urls(args.host, args.port)
        if url not in {local_open_url(args.host, args.port), share}
    ]
    for url in extra:
        print(f"  также: {url}")
    try:
        uvicorn.run(app, host=args.host, port=args.port, log_level="info")
    except KeyboardInterrupt:
        return 0
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
