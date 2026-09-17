"""FastAPI application: read-only RD catalog monitor for the LAN."""

from __future__ import annotations

import threading
import time
from datetime import datetime, timezone
from pathlib import Path
from typing import Any
from urllib.parse import quote

from fastapi import Body, FastAPI, HTTPException, Query
from fastapi.responses import FileResponse, Response
from fastapi.staticfiles import StaticFiles

from rd_catalog.config import CatalogConfig, load_config
from rd_catalog.db import CatalogDatabase
from rd_catalog.kits_table_layout import resolve_kits_table_layout
from rd_catalog.monitor_views import (
    AN_HEADERS,
    CARD_PACKAGE_HEADERS,
    KITS_HEADERS,
    KITS_PAINT_LEGEND_BUTTON,
    KITS_PAINT_LEGEND_INTRO,
    KITS_PAINT_LEGEND_TITLE,
    KITS_PROGRESS_VISIBLE_SEP,
    WEB_MONITOR_SHEETS,
    CatalogMonitor,
    format_kits_progress_stats,
    kits_paint_legend,
    kits_progress_stats,
    load_catalog_monitor,
    jsonify_value,
)
from rd_catalog.table_xlsx import (
    KITS_XLSX_BUTTON,
    KITS_XLSX_MAX_KEYS,
    KITS_XLSX_SHEET,
    columns_from_client,
    dated_xlsx_filename,
    exported_table_from_kit_cells,
    exported_table_xlsx_bytes,
)

STATIC_DIR = Path(__file__).resolve().parent / "static"
WEB_BUSY_TIMEOUT_MS = 60_000
_FORCE_REFRESH_DEBOUNCE_S = 0.5


class MonitorCache:
    """One in-memory ``CatalogMonitor``; rebuilds on sqlite mtime or refresh.

    ``CatalogDatabase`` opens a short-lived connection per call and must not
    share a ``sqlite3.Connection`` across threads. The cached object is the
    join DTO plus a ``CatalogDatabase`` handle, not an open connection.
    """

    def __init__(self, database: CatalogDatabase, config: CatalogConfig) -> None:
        self._database = database
        self._config = config
        self._lock = threading.Lock()
        self._monitor: CatalogMonitor | None = None
        self._mtime: int = -1
        self._generated_at: str = ""
        self._built_mono: float = 0.0

    def get(self, *, refresh: bool = False) -> tuple[CatalogMonitor, str]:
        """Return the cached monitor, rebuilding when needed.

        Args:
            refresh: Force a rebuild even if the sqlite stamp is unchanged.
                Parallel force-refresh calls within 0.5 s share one rebuild.

        Returns:
            Monitor DTO and UTC ``generated_at`` timestamp.
        """

        with self._lock:
            stamp = _sqlite_stamp(self._database.path)
            stale = self._monitor is None or self._mtime != stamp
            recent_force = (
                not stale
                and refresh
                and (time.monotonic() - self._built_mono) < _FORCE_REFRESH_DEBOUNCE_S
            )
            if stale or (refresh and not recent_force):
                self._monitor = load_catalog_monitor(
                    self._database,
                    self._config,
                    sheets=WEB_MONITOR_SHEETS,
                )
                self._mtime = _sqlite_stamp(self._database.path)
                self._generated_at = datetime.now(timezone.utc).isoformat()
                self._built_mono = time.monotonic()
            assert self._monitor is not None
            return self._monitor, self._generated_at


def _sqlite_stamp(path: Path) -> int:
    """Max mtime of the sqlite file and its WAL/SHM sidecars."""

    stamp = 0
    for candidate in (path, Path(f"{path}-wal"), Path(f"{path}-shm")):
        try:
            stamp = max(stamp, candidate.stat().st_mtime_ns)
        except OSError:
            continue
    return stamp


def _meta_payload(
    monitor: CatalogMonitor,
    generated_at: str,
    config: CatalogConfig,
) -> dict[str, Any]:
    kits_layout = resolve_kits_table_layout(config.runtime_dir, KITS_HEADERS)
    progress = kits_progress_stats(monitor.kits)
    return {
        "last_scan": jsonify_value(monitor.last_scan),
        "warning": monitor.warning,
        "palette": dict(monitor.palette),
        "generated_at": generated_at,
        "kits_headers": list(kits_layout.order),
        "kits_header_widths": dict(kits_layout.widths),
        "kits_paint_legend": jsonify_value(kits_paint_legend(monitor.palette)),
        "kits_paint_legend_title": KITS_PAINT_LEGEND_TITLE,
        "kits_paint_legend_intro": KITS_PAINT_LEGEND_INTRO,
        "kits_paint_legend_button": KITS_PAINT_LEGEND_BUTTON,
        "kits_xlsx_button": KITS_XLSX_BUTTON,
        "kits_progress_text": format_kits_progress_stats(progress),
        "kits_progress_visible_sep": KITS_PROGRESS_VISIBLE_SEP,
        "an_headers": list(AN_HEADERS),
        "card_package_headers": list(CARD_PACKAGE_HEADERS),
    }


def create_app(config_path: str | Path | None = None) -> FastAPI:
    """Build the read-only catalog app.

    Loads ``CatalogConfig``, opens the local sqlite with a long busy timeout,
    and runs ``initialize`` (safe on an existing schema 10). Does not scan,
    ingest Google, or rebuild the pipeline.

    Args:
        config_path: Optional JSON override for ``load_config``.

    Returns:
        FastAPI application serving ``/`` and ``/api/*``.
    """

    config = load_config(config_path)
    database = CatalogDatabase(config.db_path, busy_timeout_ms=WEB_BUSY_TIMEOUT_MS)
    database.initialize()
    cache = MonitorCache(database, config)

    app = FastAPI(
        title="RD Catalog Web",
        description="Read-only LAN monitor for the AGCC RD catalog.",
        docs_url=None,
        redoc_url=None,
        openapi_url=None,
    )
    app.state.config = config
    app.state.database = database
    app.state.cache = cache

    def _load(refresh: bool) -> tuple[CatalogMonitor, str]:
        return cache.get(refresh=refresh)

    @app.get("/")
    def index() -> FileResponse:
        return FileResponse(STATIC_DIR / "index.html", media_type="text/html; charset=utf-8")

    @app.get("/api/meta")
    def api_meta(refresh: bool = Query(False)) -> dict[str, Any]:
        monitor, generated_at = _load(refresh)
        return _meta_payload(monitor, generated_at, config)

    @app.post("/api/refresh")
    def api_refresh() -> dict[str, Any]:
        monitor, generated_at = _load(True)
        return _meta_payload(monitor, generated_at, config)

    @app.get("/api/kits")
    def api_kits(refresh: bool = Query(False)) -> list[Any]:
        monitor, _generated = _load(refresh)
        payload = monitor.as_json()
        return list(payload.get("kits") or [])

    @app.get("/api/kits/{title}/{mark}/card")
    def api_kit_card(
        title: str,
        mark: str,
        refresh: bool = Query(False),
    ) -> dict[str, Any]:
        monitor, _generated = _load(refresh)
        card = monitor.kit_card(title, mark)
        if card is None:
            raise HTTPException(status_code=404, detail="kit not found")
        return jsonify_value(card)

    @app.post("/api/kits/xlsx")
    def api_kits_xlsx(
        payload: dict[str, Any] | None = Body(default=None),
        refresh: bool = Query(False),
    ) -> Response:
        """Download Комплекты xlsx for the caller's visible row order.

        Body ``keys`` is ``[{title, mark}, ...]`` in Tabulator order.
        Optional ``columns`` is ``[{header, width_px}, ...]`` from the live
        table; otherwise ``kits_table_layout.json``. Paint comes from the
        cached ``MonitorCell`` join, not from the client.
        """

        raw_keys = payload.get("keys") if isinstance(payload, dict) else None
        if raw_keys is None:
            raw_keys = []
        if not isinstance(raw_keys, list):
            raise HTTPException(status_code=400, detail="keys must be a list")
        if len(raw_keys) > KITS_XLSX_MAX_KEYS:
            raise HTTPException(status_code=400, detail="too many keys")
        keys: list[tuple[str, str]] = []
        for item in raw_keys:
            if not isinstance(item, dict):
                continue
            title = str(item.get("title") or "").strip()
            mark = str(item.get("mark") or "").strip()
            if title and mark:
                keys.append((title, mark))

        monitor, _generated = _load(refresh)
        layout = resolve_kits_table_layout(config.runtime_dir, KITS_HEADERS)
        raw_columns = payload.get("columns") if isinstance(payload, dict) else None
        if raw_columns is not None and not isinstance(raw_columns, list):
            raise HTTPException(status_code=400, detail="columns must be a list")
        columns = columns_from_client(
            raw_columns,
            known_names=KITS_HEADERS,
            fallback_order=layout.order,
            fallback_widths=layout.widths,
        )
        cells_by_key = {row.kit_key: row.cells for row in monitor.kits}
        table = exported_table_from_kit_cells(
            cells_by_key,
            keys,
            columns,
            sheet_name=KITS_XLSX_SHEET,
        )
        try:
            data = exported_table_xlsx_bytes(table)
        except OSError as exc:
            raise HTTPException(
                status_code=500,
                detail=f"xlsx write failed: {exc}",
            ) from exc
        filename = dated_xlsx_filename(KITS_XLSX_SHEET)
        return Response(
            content=data,
            media_type=(
                "application/vnd.openxmlformats-officedocument"
                ".spreadsheetml.sheet"
            ),
            headers={
                "Content-Disposition": (
                    'attachment; filename="Komplekty.xlsx"; '
                    f"filename*=UTF-8''{quote(filename)}"
                ),
                "Cache-Control": "no-store",
            },
        )

    @app.get("/api/an")
    def api_an(refresh: bool = Query(False)) -> list[Any]:
        monitor, _generated = _load(refresh)
        payload = monitor.as_json()
        return list(payload.get("an_rows") or [])

    app.mount("/static", StaticFiles(directory=str(STATIC_DIR)), name="static")
    return app
