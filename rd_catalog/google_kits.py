"""Public CSV export + local cache for RD catalog Google kit sheets.

Loads two spreadsheets:

- KSB ИД expected kits (``google_kits_*``)
- «Выдача РД ПД» issuance log (``google_issuance_*``)

Independent of ``base.base_google`` / СПО_1. Default path is anonymous HTTPS
export; service-account Sheets API is not required while the spreadsheets stay
shared as «anyone with the link — viewer».
"""

from __future__ import annotations

import csv
import io
import json
import os
import ssl
import time
import urllib.error
import urllib.parse
import urllib.request
from dataclasses import dataclass
from datetime import datetime, timezone
from email.utils import parsedate_to_datetime
from pathlib import Path
from typing import Any, Callable

from rd_catalog.config import CatalogConfig
from rd_catalog.kits import (
    GoogleKit,
    GoogleParseStats,
    IssuanceKit,
    parse_google_matrix,
    parse_issuance_matrix,
    parse_issuance_sends,
)

DEFAULT_KITS_SPREADSHEET_ID = "1iKF-5tf0LAewib-5_E0R7AurGKChk4-QPtbgI1DN4zk"
DEFAULT_ISSUANCE_SPREADSHEET_ID = "1NMRZLWdAYmVlWyExtC9PYcWS7cr3TuQzMwu59HiSYEU"
DEFAULT_ISSUANCE_SHEET_NAME = "Выдача РД ПД"

_CACHE_KITS_DATA = "google_kits_data.json"
_CACHE_KITS_META = "google_kits_meta.json"
_CACHE_ISSUANCE_DATA = "google_issuance_data.json"
_CACHE_ISSUANCE_META = "google_issuance_meta.json"

_CHROME_LIKE_UA = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36"
)


@dataclass(frozen=True, slots=True)
class GoogleKitsLoadResult:
    """Outcome of fetching both Google kit sources."""

    kits: tuple[GoogleKit, ...]
    issuance_kits: tuple[IssuanceKit, ...]
    stats: GoogleParseStats
    issuance_stats: GoogleParseStats
    source: str
    fetched_at: str
    warning: str | None = None
    error: str | None = None
    issuance_sends: tuple[IssuanceKit, ...] = ()


def kits_cache_paths(runtime_dir: str | Path) -> tuple[Path, Path]:
    """Return KSB ИД ``(data_json, meta_json)`` cache paths."""

    root = Path(runtime_dir)
    return root / _CACHE_KITS_DATA, root / _CACHE_KITS_META


def issuance_cache_paths(runtime_dir: str | Path) -> tuple[Path, Path]:
    """Return «Выдача РД ПД» ``(data_json, meta_json)`` cache paths."""

    root = Path(runtime_dir)
    return root / _CACHE_ISSUANCE_DATA, root / _CACHE_ISSUANCE_META


def _timeout_sec() -> float:
    return float(os.environ.get("RD_KITS_EXPORT_TIMEOUT_SEC", "8"))


def _retries() -> int:
    return max(0, int(os.environ.get("RD_KITS_EXPORT_RETRIES", "1")))


def _retry_delay_sec() -> float:
    return float(os.environ.get("RD_KITS_EXPORT_RETRY_DELAY_SEC", "0.35"))


def kits_tls_relaxed() -> bool:
    """Return whether Google HTTPS should skip certificate verification.

    Default is relaxed: certifi does not see the BCC SSL-inspection CA, so
    urllib/requests fail with ``CERTIFICATE_VERIFY_FAILED``. Set
    ``RD_KITS_SSL_STRICT=1`` to verify against the default CA bundle.

    Returns:
        True when callers should disable TLS verification.
    """

    if os.environ.get("RD_KITS_SSL_STRICT", "").strip().lower() in (
        "1",
        "true",
        "yes",
    ):
        return False
    return True


def _export_csv_url(spreadsheet_id: str, sheet_name: str) -> str:
    query: dict[str, str] = {"format": "csv"}
    if sheet_name.strip():
        query["sheet"] = sheet_name.strip()
    encoded = urllib.parse.urlencode(query, safe="")
    return f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}/export?{encoded}"


def _http_headers() -> dict[str, str]:
    return {
        "User-Agent": _CHROME_LIKE_UA,
        "Accept": "text/csv,*/*;q=0.8",
        "Accept-Language": "ru-RU,ru;q=0.9,en-US;q=0.8,en;q=0.7",
    }


def _opener() -> urllib.request.OpenerDirector:
    proxies = urllib.request.getproxies()
    handlers: list[Any] = [urllib.request.ProxyHandler(proxies)]
    if kits_tls_relaxed():
        context = ssl.SSLContext(ssl.PROTOCOL_TLS_CLIENT)
        context.check_hostname = False
        context.verify_mode = ssl.CERT_NONE
        handlers.append(urllib.request.HTTPSHandler(context=context))
    return urllib.request.build_opener(*handlers)


def _http_request(url: str, *, method: str, timeout: float) -> tuple[bytes, dict[str, str]]:
    request = urllib.request.Request(url, method=method, headers=_http_headers())
    with _opener().open(request, timeout=timeout) as response:
        body = response.read()
        headers = {key.lower(): value for key, value in response.headers.items()}
    return body, headers


def _http_request_resilient(url: str, *, method: str, timeout: float) -> tuple[bytes, dict[str, str]]:
    attempts = 1 + _retries()
    last_timeout: TimeoutError | None = None
    for index in range(attempts):
        try:
            return _http_request(url, method=method, timeout=timeout)
        except TimeoutError as exc:
            last_timeout = exc
            if index + 1 >= attempts:
                break
            time.sleep(_retry_delay_sec())
    assert last_timeout is not None
    raise last_timeout


def _export_error_hint(body: bytes) -> str | None:
    snippet = body[:8000].decode("utf-8", errors="replace").lstrip()
    low = snippet.lower()
    if not snippet:
        return "Пустой ответ от Google export."
    if snippet.startswith("<") and (
        "sign in" in low
        or "accounts.google.com" in low
        or "service login" in low
        or "access denied" in low
        or "войдите" in low
    ):
        return (
            "Google вернул страницу входа вместо CSV. Нужен доступ "
            "«Все, у кого есть ссылка — читатель»."
        )
    if snippet.startswith("<!") or snippet.startswith("<html"):
        return (
            "Ответ похож на HTML, а не CSV. Проверьте публичный доступ к таблице "
            "и HTTPS_PROXY / NO_PROXY."
        )
    return None


def _parse_csv(body: bytes) -> list[list[str]]:
    text = body.decode("utf-8-sig")
    return [list(row) for row in csv.reader(io.StringIO(text))]


def _meta_timestamp(last_modified: str | None) -> str:
    if last_modified:
        try:
            parsed = parsedate_to_datetime(last_modified)
            if parsed.tzinfo is None:
                return parsed.replace(tzinfo=timezone.utc).isoformat()
            return parsed.isoformat()
        except (TypeError, ValueError, OverflowError):
            pass
    return datetime.now(timezone.utc).isoformat()


def _write_cache(
    data_path: Path,
    meta_path: Path,
    rows: list[list[str]],
    headers: dict[str, str],
) -> None:
    data_path.parent.mkdir(parents=True, exist_ok=True)
    last_mod = headers.get("last-modified") or ""
    data_path.write_text(
        json.dumps({"rows": rows}, ensure_ascii=False),
        encoding="utf-8",
    )
    meta = {
        "modified_time": _meta_timestamp(last_mod),
        "export_last_modified": last_mod,
        "source": "export_csv",
        "row_count": len(rows),
    }
    meta_path.write_text(json.dumps(meta, ensure_ascii=False, indent=0), encoding="utf-8")


def _read_cached_rows(data_path: Path) -> list[list[str]] | None:
    if not data_path.is_file():
        return None
    try:
        payload = json.loads(data_path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return None
    rows = payload.get("rows") if isinstance(payload, dict) else None
    if not isinstance(rows, list):
        return None
    return [[str(cell) if cell is not None else "" for cell in row] for row in rows]


def _read_meta_time(meta_path: Path) -> str:
    if not meta_path.is_file():
        return ""
    try:
        meta = json.loads(meta_path.read_text(encoding="utf-8"))
        return str(meta.get("modified_time") or "")
    except (OSError, json.JSONDecodeError):
        return ""


def _empty_stats() -> GoogleParseStats:
    return GoogleParseStats(kept=0, skipped=0, skipped_reasons=())


def _parse_issuance_payload(
    rows: list[list[str]] | None,
) -> tuple[tuple[IssuanceKit, ...], tuple[IssuanceKit, ...], GoogleParseStats]:
    """Parse issuance CSV into latest-wins kits and every send row.

    Args:
        rows: Cached or exported sheet rows, or ``None`` when missing.

    Returns:
        ``(latest_wins, all_sends, stats)``. ``stats.kept`` counts latest-wins.
    """

    if not rows:
        return (), (), _empty_stats()
    latest, stats = parse_issuance_matrix(rows)
    return latest, parse_issuance_sends(rows), stats


def _export_one_sheet(
    spreadsheet_id: str,
    sheet_name: str,
    data_path: Path,
    meta_path: Path,
) -> tuple[list[list[str]], dict[str, str]]:
    url = _export_csv_url(spreadsheet_id, sheet_name)
    body, headers = _http_request_resilient(
        url, method="GET", timeout=_timeout_sec()
    )
    hint = _export_error_hint(body)
    if hint:
        raise ValueError(hint)
    rows = _parse_csv(body)
    _write_cache(data_path, meta_path, rows, headers)
    return rows, headers


def load_cached_google_kits(runtime_dir: str | Path) -> GoogleKitsLoadResult | None:
    """Parse both kit sources from local export caches when available.

    Args:
        runtime_dir: Catalog runtime directory.

    Returns:
        Cached result when at least one cache exists, else ``None``.
    """

    kits_data, kits_meta = kits_cache_paths(runtime_dir)
    iss_data, iss_meta = issuance_cache_paths(runtime_dir)
    kits_rows = _read_cached_rows(kits_data)
    iss_rows = _read_cached_rows(iss_data)
    if kits_rows is None and iss_rows is None:
        return None
    kits, stats = (
        parse_google_matrix(kits_rows) if kits_rows is not None else ((), _empty_stats())
    )
    issuance, issuance_sends, iss_stats = _parse_issuance_payload(iss_rows)
    fetched = _read_meta_time(kits_meta) or _read_meta_time(iss_meta)
    return GoogleKitsLoadResult(
        kits=kits,
        issuance_kits=issuance,
        stats=stats,
        issuance_stats=iss_stats,
        source="cache",
        fetched_at=fetched,
        issuance_sends=issuance_sends,
    )


def fetch_google_kits(
    config: CatalogConfig,
    *,
    cache_only: bool = False,
    include_issuance: bool = True,
) -> GoogleKitsLoadResult:
    """Download KSB ИД + «Выдача РД ПД» sheets or fall back to cache.

    After F/D/E writes the thread passes ``include_issuance=False`` so only
    КСБ ИД is re-exported; issuance stays on the existing JSON cache. Full
    Google load (toolbar) keeps the default ``True``.

    Args:
        config: Resolved catalog configuration.
        cache_only: Skip the network and read caches only. Ignores
            ``include_issuance`` and returns both sides from cache.
        include_issuance: When False (and not ``cache_only``), skip issuance
            HTTP and fill issuance from cache. Default True fetches both sheets.

    Returns:
        Combined parse result. Network failures set ``warning``/``error``.
    """

    if cache_only:
        cached = load_cached_google_kits(config.runtime_dir)
        if cached is not None:
            return cached
        return GoogleKitsLoadResult(
            kits=(),
            issuance_kits=(),
            stats=_empty_stats(),
            issuance_stats=_empty_stats(),
            source="cache",
            fetched_at="",
            error="Локальный кэш комплектов Google отсутствует.",
        )

    kits_data, kits_meta = kits_cache_paths(config.runtime_dir)
    iss_data, iss_meta = issuance_cache_paths(config.runtime_dir)
    warnings: list[str] = []
    kits: tuple[GoogleKit, ...] = ()
    issuance: tuple[IssuanceKit, ...] = ()
    issuance_sends: tuple[IssuanceKit, ...] = ()
    stats = _empty_stats()
    iss_stats = _empty_stats()
    sources: list[str] = []
    fetched_at = ""

    def _load_cached_side(
        label: str,
        data_path: Path,
        meta_path: Path,
        parser: Callable[[list[list[str]]], tuple[Any, GoogleParseStats]],
    ) -> tuple[Any, GoogleParseStats, str]:
        rows = _read_cached_rows(data_path)
        if rows is None:
            return (), _empty_stats(), ""
        parsed, parsed_stats = parser(rows)
        return parsed, parsed_stats, _read_meta_time(meta_path)

    try:
        rows, headers = _export_one_sheet(
            config.google_kits_spreadsheet_id,
            config.google_kits_sheet_name,
            kits_data,
            kits_meta,
        )
        kits, stats = parse_google_matrix(rows)
        sources.append("ksb_id_export")
        fetched_at = _meta_timestamp(headers.get("last-modified"))
    except (
        urllib.error.HTTPError,
        urllib.error.URLError,
        TimeoutError,
        OSError,
        ValueError,
        csv.Error,
        UnicodeError,
    ) as exc:
        cached_kits, stats, kits_time = _load_cached_side(
            "КСБ ИД", kits_data, kits_meta, parse_google_matrix
        )
        kits = cached_kits
        if kits:
            sources.append("ksb_id_cache")
            fetched_at = kits_time or fetched_at
            warnings.append(f"Google: export недоступен ({type(exc).__name__}: {exc})")
        else:
            warnings.append(f"Google: {type(exc).__name__}: {exc}")

    if include_issuance:
        try:
            rows, headers = _export_one_sheet(
                config.google_issuance_spreadsheet_id,
                config.google_issuance_sheet_name,
                iss_data,
                iss_meta,
            )
            issuance, issuance_sends, iss_stats = _parse_issuance_payload(rows)
            sources.append("issuance_export")
            fetched_at = fetched_at or _meta_timestamp(headers.get("last-modified"))
        except (
            urllib.error.HTTPError,
            urllib.error.URLError,
            TimeoutError,
            OSError,
            ValueError,
            csv.Error,
            UnicodeError,
        ) as exc:
            cached_rows = _read_cached_rows(iss_data)
            issuance, issuance_sends, iss_stats = _parse_issuance_payload(cached_rows)
            iss_time = _read_meta_time(iss_meta)
            if issuance:
                sources.append("issuance_cache")
                fetched_at = fetched_at or iss_time
                warnings.append(
                    f"Выдача РД ПД: export недоступен ({type(exc).__name__}: {exc})"
                )
            else:
                warnings.append(f"Выдача РД ПД: {type(exc).__name__}: {exc}")
    else:
        cached_rows = _read_cached_rows(iss_data)
        issuance, issuance_sends, iss_stats = _parse_issuance_payload(cached_rows)
        iss_time = _read_meta_time(iss_meta)
        if issuance:
            sources.append("issuance_cache")
            fetched_at = fetched_at or iss_time

    if not kits and not issuance:
        return GoogleKitsLoadResult(
            kits=(),
            issuance_kits=(),
            stats=stats,
            issuance_stats=iss_stats,
            source="error",
            fetched_at="",
            error="; ".join(warnings) or "Не удалось загрузить Google-таблицы.",
            issuance_sends=(),
        )

    source = "+".join(sources) if sources else "partial"
    return GoogleKitsLoadResult(
        kits=kits,
        issuance_kits=issuance,
        stats=stats,
        issuance_stats=iss_stats,
        source=source,
        fetched_at=fetched_at,
        warning="; ".join(warnings) if warnings else None,
        issuance_sends=issuance_sends,
    )
