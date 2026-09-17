import sys
from pathlib import Path

# Запуск как «python .../base/base_google.py»: родитель каталога base — корень проекта
if __name__ == "__main__" and not __package__:
    _proj_root = str(Path(__file__).resolve().parents[1])
    if _proj_root not in sys.path:
        sys.path.insert(0, _proj_root)

import concurrent.futures
import csv
import io
import json
import os
import ssl
import time
import urllib.error
import urllib.parse
import urllib.request
from datetime import datetime, timezone
from email.utils import parsedate_to_datetime

import pygsheets

from base.base_utils import check_code
from base.base_classes import *
from base.tables_columns import ColNames, CODE
from utils.colors import Color
import base.t_comm_initial_classes as t_com_init_cls

serv_file = "base_check/client_secret.json"
json_file_name_code_base = "base_check/code_base_data.json"
_CACHE_META_FILE = "base_check/code_base_meta.json"

_SPREADSHEET_ID = "1P_9LcZ2LGqWRUcjvun5G1r3cd78vGe6fhxPo7tasqDc"
_SPREADSHEET_URL = (
    f"https://docs.google.com/spreadsheets/d/{_SPREADSHEET_ID}/edit?usp=sharing"
)
_WORKSHEET_NAME = os.environ.get("CODE_BASE_SHEET_NAME", "СПО_1")
# Таймаут одного HTTP-запроса export (GET/HEAD): по умолчанию 6 с.
# Доп. попытки — CODE_BASE_EXPORT_RETRIES (каждая не дольше этого лимита).
_EXPORT_TIMEOUT_SEC = int(os.environ.get("CODE_BASE_EXPORT_TIMEOUT_SEC", "6"))
# Сколько повторить export GET/HEAD при только TimeoutError (0 = одна попытка).
_EXPORT_RETRIES = max(0, int(os.environ.get("CODE_BASE_EXPORT_RETRIES", "0")))
# Пауза перед повтором (сек).
_EXPORT_RETRY_DELAY_SEC = float(
    os.environ.get("CODE_BASE_EXPORT_RETRY_DELAY_SEC", "0.35")
)
# Не дергать Google повторно в том же процессе чаще, чем раз в N сек (MTO и BBB оба вызывают load_base).
# 0 = всегда пытаться синхронизировать.
_SYNC_COOLDOWN_SEC = float(os.environ.get("CODE_BASE_SYNC_COOLDOWN_SEC", "300"))
# По умолчанию только public export CSV. Этап Sheets API: ``CODE_BASE_SYNC_API_FIRST=1``.
_api_first_env = os.environ.get("CODE_BASE_SYNC_API_FIRST", "0").strip().lower()
_SYNC_API_FIRST = _api_first_env in ("1", "true", "yes", "on")
# Таймаут всего блока authorize + drive.get_update_time + open + get_values (сек).
# По умолчанию 6 с — быстрый отказ API и переход на export; не зависит от export timeout.
_API_SYNC_TIMEOUT_SEC = float(os.environ.get("CODE_BASE_API_TIMEOUT_SEC", "6"))

_last_export_ok_monotonic: float | None = None
_TLS_RELAX_WARNED = False
# Один Opener на процесс: повторные запросы к docs.google.com могут переиспользовать соединение.
_urllib_opener_singleton: urllib.request.OpenerDirector | None = None


def _urllib_tls_relaxed_from_env() -> bool:
    """urllib export: по умолчанию ослабленная проверка TLS (на контурах с CRL/OCSP).

    Строгая проверка: CODE_BASE_SSL_STRICT=1, или CODE_BASE_PROBE_STRICT_TLS=1,
    или CODE_BASE_SSL_NO_REVOKE=0 / false / off.
    """
    if os.environ.get("CODE_BASE_SSL_STRICT", "").strip().lower() in (
        "1",
        "true",
        "yes",
    ):
        return False
    if os.environ.get("CODE_BASE_PROBE_STRICT_TLS", "").strip().lower() in (
        "1",
        "true",
        "yes",
    ):
        return False
    sn = os.environ.get("CODE_BASE_SSL_NO_REVOKE", "").strip().lower()
    if sn in ("0", "false", "no", "off"):
        return False
    return True


# urllib для export CSV: по умолчанию CERT_NONE (не затрагивает pygsheets/httplib2).
_HTTP_URLLIB_TLS_RELAXED: bool = _urllib_tls_relaxed_from_env()


def _sync_debug_enabled() -> bool:
    """Краткий лог пути синка code base: по умолчанию включён (CODE_BASE_SYNC_DEBUG=0 — выкл.)."""
    v = os.environ.get("CODE_BASE_SYNC_DEBUG", "1").strip().lower()
    return v not in ("0", "false", "no", "off")


def _sync_log(msg: str) -> None:
    if _sync_debug_enabled():
        print(msg, flush=True)


def _warn_tls_relaxed_once() -> None:
    global _TLS_RELAX_WARNED
    if _TLS_RELAX_WARNED or not _HTTP_URLLIB_TLS_RELAXED:
        return
    _TLS_RELAX_WARNED = True
    print(
        "code_base: проверка TLS сертификата для urllib отключена (CERT_NONE); "
        "строго: CODE_BASE_SSL_STRICT=1 или CODE_BASE_SSL_NO_REVOKE=0. "
        "Таймаут чтения CSV — CODE_BASE_EXPORT_TIMEOUT_SEC (не связан с SSL).",
        flush=True,
    )


# Минимум колонок в строке (индексы до 27 используются в ColNames.GoogleBase)
_MIN_CODE_BASE_COLS = 28

# Как у обычного Chrome: иначе Google часто отдаёт не CSV, а редирект/страницу для «ботов».
_CHROME_LIKE_UA = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) "
    "Chrome/131.0.0.0 Safari/537.36"
)


def _export_csv_url() -> str:
    """Тот же spreadsheetId, что в ссылке /edit; /export?format=csv&sheet=… — выгрузка листа в CSV (не HTML-страница)."""
    q = urllib.parse.urlencode(
        {"format": "csv", "sheet": _WORKSHEET_NAME},
        safe="",
    )
    return f"https://docs.google.com/spreadsheets/d/{_SPREADSHEET_ID}/export?{q}"


def _http_browser_headers() -> dict[str, str]:
    return {
        "User-Agent": _CHROME_LIKE_UA,
        "Accept": "text/csv,*/*;q=0.8",
        "Accept-Language": "ru-RU,ru;q=0.9,en-US;q=0.8,en;q=0.7",
    }


def _http_opener():
    """Opener с прокси из окружения (HTTPS_PROXY и т.д.) — как у многих CLI, не как у Chrome."""
    if _HTTP_URLLIB_TLS_RELAXED:
        _warn_tls_relaxed_once()
    proxies = urllib.request.getproxies()
    handlers: list = [urllib.request.ProxyHandler(proxies)]
    if _HTTP_URLLIB_TLS_RELAXED:
        _ctx = ssl.SSLContext(ssl.PROTOCOL_TLS_CLIENT)
        _ctx.check_hostname = False
        _ctx.verify_mode = ssl.CERT_NONE
        handlers.append(urllib.request.HTTPSHandler(context=_ctx))
    return urllib.request.build_opener(*handlers)


def _get_urllib_opener_cached() -> urllib.request.OpenerDirector:
    """Переиспользует opener (один на процесс) — лучше для keep-alive к docs.google.com."""
    global _urllib_opener_singleton
    if _urllib_opener_singleton is None:
        _urllib_opener_singleton = _http_opener()
    return _urllib_opener_singleton


def _http_request(
    url: str,
    *,
    method: str = "GET",
    timeout: float | None = None,
) -> tuple[bytes, dict[str, str]]:
    req = urllib.request.Request(
        url,
        method=method,
        headers=_http_browser_headers(),
    )
    to = timeout if timeout is not None else float(_EXPORT_TIMEOUT_SEC)
    opener = _get_urllib_opener_cached()
    if os.environ.get("CODE_BASE_EXPORT_DEBUG", "").strip() in ("1", "true", "yes"):
        px = urllib.request.getproxies()
        print(f"[code_base export] прокси из getproxies(): {px or '(пусто — прямой выход)'}")
    with opener.open(req, timeout=to) as resp:
        body = resp.read()
        headers = {k.lower(): v for k, v in resp.headers.items()}
    return body, headers


def _http_request_export_resilient(
    url: str,
    *,
    method: str,
    timeout: float,
) -> tuple[bytes, dict[str, str]]:
    """Повторяет запрос при только ``TimeoutError`` (лимит времени одной попытки не меняется)."""
    attempts = 1 + _EXPORT_RETRIES
    last: TimeoutError | None = None
    for i in range(attempts):
        try:
            return _http_request(url, method=method, timeout=timeout)
        except TimeoutError as e:
            last = e
            if i + 1 >= attempts:
                break
            _sync_log(
                f"[code_base] export {method}: TimeoutError — повтор {i + 2}/{attempts} "
                f"через {_EXPORT_RETRY_DELAY_SEC:g} с"
            )
            time.sleep(_EXPORT_RETRY_DELAY_SEC)
    assert last is not None
    raise last


def _export_response_error_hint(body: bytes) -> str | None:
    """Если вместо CSV пришла HTML (вход в Google, прокси, блокировка) — пояснение для пользователя."""
    snippet = body[:8000].decode("utf-8", errors="replace").lstrip()
    low = snippet.lower()
    if not snippet:
        return "Пустой ответ от сервера."
    if snippet.startswith("<") and (
        "sign in" in low
        or "accounts.google.com" in low
        or "service login" in low
        or "access denied" in low
        or "войдите" in low
    ):
        return (
            "Google вернул страницу входа (HTML), а не CSV. В Chrome вы уже в аккаунте — "
            "таблица открывается с cookie; Python запрашивает URL без входа. Нужно: "
            "общий доступ «Все, у кого есть ссылка — читатель» (без авторизации) "
            "или отдельный способ с токеном/API."
        )
    if snippet.startswith("<!") or snippet.startswith("<html"):
        return (
            "Ответ похож на HTML, а не CSV (блокировка, прокси, редирект). "
            "Проверьте публичный доступ к таблице; для корп. сети задайте HTTPS_PROXY / NO_PROXY "
            "(Chrome часто использует системный прокси, Python — только из переменных окружения)."
        )
    return None


def _normalize_matrix(rows: list[list[str]]) -> list[list[str]]:
    if not rows:
        return []
    width = max(_MIN_CODE_BASE_COLS, max(len(r) for r in rows))
    out: list[list[str]] = []
    for r in rows:
        rr = [str(c) if c is not None else "" for c in r]
        if len(rr) < width:
            rr.extend([""] * (width - len(rr)))
        out.append(rr)
    return out


def _parse_export_csv(body: bytes) -> list[list[str]]:
    text = body.decode("utf-8-sig")
    reader = csv.reader(io.StringIO(text))
    return [list(row) for row in reader]


def _meta_timestamp_from_last_modified(last_modified: str | None) -> str:
    if last_modified:
        try:
            dt = parsedate_to_datetime(last_modified)
            if dt.tzinfo is None:
                return dt.replace(tzinfo=timezone.utc).isoformat()
            return dt.isoformat()
        except (TypeError, ValueError, OverflowError):
            pass
    return datetime.now(timezone.utc).isoformat()


def _export_unchanged_by_head(meta: dict) -> bool:
    """Пробуем без тела ответа: HEAD + сравнение Last-Modified с прошлой выгрузкой."""
    old_lm = (meta.get("export_last_modified") or "").strip()
    if not old_lm:
        return False
    head_timeout = float(_EXPORT_TIMEOUT_SEC)
    try:
        _, headers = _http_request_export_resilient(
            _export_csv_url(), method="HEAD", timeout=head_timeout
        )
    except (urllib.error.HTTPError, urllib.error.URLError, TimeoutError, OSError):
        return False
    new_lm = (headers.get("last-modified") or "").strip()
    if not new_lm:
        return False
    if new_lm == old_lm:
        _sync_log(
            "[code_base] путь: HEAD export Last-Modified без изменений → локальный кэш"
        )
        _sync_log("[code_base] результат: OK (кэш актуален)")
        return True
    return False


def _save_cache_meta_dict(meta: dict) -> None:
    with open(_CACHE_META_FILE, "w", encoding="utf-8") as f:
        json.dump(meta, f, ensure_ascii=False, indent=0)


def _fetch_and_cache_via_export() -> None:
    """Загрузка листа через публичный export CSV (без pygsheets / Drive API)."""
    global _last_export_ok_monotonic
    url = _export_csv_url()
    body, headers = _http_request_export_resilient(
        url, method="GET", timeout=float(_EXPORT_TIMEOUT_SEC)
    )
    hint = _export_response_error_hint(body)
    if hint:
        raise ValueError(hint)
    rows = _parse_export_csv(body)
    data = _normalize_matrix(rows)
    last_mod = headers.get("last-modified")
    mt = _meta_timestamp_from_last_modified(last_mod)
    meta = {
        "modified_time": mt,
        "export_last_modified": last_mod or "",
        "source": "export_csv",
    }
    with open(json_file_name_code_base, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False)
    _save_cache_meta_dict(meta)
    lm_note = last_mod or "нет заголовка Last-Modified"
    _sync_log("[code_base] путь: HTTPS export CSV (docs.google.com)")
    _sync_log(
        f"[code_base] результат: OK — сохранено строк: {len(data)}, Last-Modified: {lm_note}"
    )
    _last_export_ok_monotonic = time.monotonic()


def load_base():
    """Загрузка Google-базы: синхронизация кэша ``code_base_data.json``.

    По умолчанию **только** публичный HTTPS export CSV (urllib) — тот же путь, что стабилен
    в проде. Опционально **сначала** Sheets API: задайте ``CODE_BASE_SYNC_API_FIRST=1``
    (нужен ``base_check/client_secret.json`` и доступ таблицы на ``client_email``):
    Drive ``modifiedTime``, при необходимости ``get_values``; при ошибке/таймауте — fallback на export.

    Таблица для export должна быть доступна по ссылке (читатель); для API — расшарена на
    ``client_email`` сервисного аккаунта.

    Повторные вызовы в одном процессе ограничены ``CODE_BASE_SYNC_COOLDOWN_SEC``.

    Таймауты одного запроса: ``CODE_BASE_API_TIMEOUT_SEC`` (API, **6** с), ``CODE_BASE_EXPORT_TIMEOUT_SEC`` (export, **6** с). Повторы export при таймауте: ``CODE_BASE_EXPORT_RETRIES`` (по умолчанию 0), пауза ``CODE_BASE_EXPORT_RETRY_DELAY_SEC``.

    ENV: ``CODE_BASE_SYNC_DEBUG`` — краткий лог (по умолчанию вкл.). TLS urllib: см. README кода.
    """
    _auto_sync()

    code_base_data_raw = load_json(json_file_name_code_base)
    t_com = TableComments(tabel_class=t_com_init_cls.GoogleBase)
    code_base_data_std = code_base_formating(code_base_data_raw, t_com)
    _sync_log(
        f"[code_base] итог load_base: записей после фильтрации по коду: {len(code_base_data_std)}"
    )
    return code_base_data_std


def _api_credentials_available() -> bool:
    return os.path.isfile(serv_file)


def _sync_via_api_inner() -> str:
    """Синхронизация через Drive modifiedTime + Sheets API get_values.

    Returns:
        ``skip_unchanged`` — файл на Drive не новее кэша, тело листа не качали.
        ``downloaded`` — матрица сохранена с диска из API.

    Raises:
        Любая ошибка pygsheets/Drive — вызывающий код перейдёт на export CSV.
    """
    global _last_export_ok_monotonic
    gc = pygsheets.authorize(service_file=serv_file)
    remote_mt = gc.drive.get_update_time(_SPREADSHEET_ID)
    meta = _load_cache_meta()
    cache_ok = os.path.exists(json_file_name_code_base)
    cached_drive = (meta.get("drive_mtime") or "").strip()
    if cache_ok and cached_drive and cached_drive == remote_mt:
        _sync_log(
            "[code_base] путь: Drive modifiedTime без изменений (проверка до get_values)"
        )
        _sync_log(f"[code_base] деталь: modifiedTime={remote_mt}")
        _sync_log("[code_base] результат: OK (кэш актуален, Sheets API)")
        return "skip_unchanged"

    _sync_log(
        "[code_base] путь: Sheets API — spreadsheets.get / values "
        f"(лист {_WORKSHEET_NAME!r})"
    )
    sh = gc.open_by_url(_SPREADSHEET_URL)
    wks = sh.worksheet_by_title(_WORKSHEET_NAME)
    rng = wks.get_values(None, None, returnas="matrix")
    data = _normalize_matrix(rng)
    prev = meta or {}
    meta_out = {
        "modified_time": remote_mt,
        "drive_mtime": remote_mt,
        "export_last_modified": prev.get("export_last_modified") or "",
        "source": "sheets_api",
    }
    with open(json_file_name_code_base, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False)
    _save_cache_meta_dict(meta_out)
    _sync_log(
        f"[code_base] результат: OK — Sheets API, строк матрицы: {len(data)}, "
        f"modifiedTime={remote_mt}"
    )
    _last_export_ok_monotonic = time.monotonic()
    return "downloaded"


def _sync_via_api() -> str:
    """Оборачивает API-синк в таймаут (отдельный поток).

    Без ожидания воркера при выходе: иначе ``ThreadPoolExecutor`` в режиме
    ``shutdown(wait=True)`` (контекст ``with``) блокирует процесс, если поток
    ещё висит на сети после срабатывания ``fut.result(timeout=...)``.
    """
    pool = concurrent.futures.ThreadPoolExecutor(max_workers=1)
    try:
        fut = pool.submit(_sync_via_api_inner)
        return fut.result(timeout=_API_SYNC_TIMEOUT_SEC)
    finally:
        pool.shutdown(wait=False, cancel_futures=True)


def _auto_sync():
    global _last_export_ok_monotonic
    cache_ok = os.path.exists(json_file_name_code_base)
    meta = _load_cache_meta()

    try:
        if (
            _SYNC_COOLDOWN_SEC > 0
            and _last_export_ok_monotonic is not None
            and cache_ok
            and (time.monotonic() - _last_export_ok_monotonic) < _SYNC_COOLDOWN_SEC
        ):
            age = time.monotonic() - _last_export_ok_monotonic
            _sync_log(
                "[code_base] путь: кулдаун после успешной выгрузки — запрос к Google не повторяем"
            )
            _sync_log(
                f"[code_base] деталь: последняя выгрузка {age:.0f} с назад, "
                f"лимит CODE_BASE_SYNC_COOLDOWN_SEC={_SYNC_COOLDOWN_SEC:.0f} с"
            )
            _sync_log("[code_base] результат: OK (локальный кэш без сети)")
            return

        if _SYNC_API_FIRST and _api_credentials_available():
            try:
                outcome = _sync_via_api()
                if outcome in ("skip_unchanged", "downloaded"):
                    return
            except TimeoutError as e:
                print(
                    f"code_base: Sheets API — таймаут {_API_SYNC_TIMEOUT_SEC:g} с ({e!r}); "
                    f"см. CODE_BASE_API_TIMEOUT_SEC.",
                    flush=True,
                )
                _sync_log(
                    f"[code_base] результат: Ошибка — API sync TimeoutError {_API_SYNC_TIMEOUT_SEC:g} с"
                )
                _sync_log(
                    "[code_base] переход: загрузка через публичный export CSV (urllib)…"
                )
            except Exception as e:
                print(f"code_base: Sheets API недоступен ({type(e).__name__}: {e})", flush=True)
                _sync_log(
                    f"[code_base] результат: Ошибка Sheets API — {type(e).__name__}: {e}"
                )
                _sync_log(
                    "[code_base] переход: загрузка через публичный export CSV (urllib)…"
                )
        else:
            if not _SYNC_API_FIRST:
                _sync_log(
                    "[code_base] путь: Sheets API выключен по умолчанию — только export "
                    "(включить: CODE_BASE_SYNC_API_FIRST=1)"
                )
            elif not _api_credentials_available():
                _sync_log(
                    "[code_base] путь: нет файла ключа SA — только export CSV "
                    f"({serv_file})"
                )

        if cache_ok and meta and _export_unchanged_by_head(meta):
            return
        _fetch_and_cache_via_export()
    except urllib.error.HTTPError as e:
        msg = f"HTTP {e.code} при загрузке export CSV"
        if e.code == 403:
            msg += (
                " — проверьте, что у таблицы включён доступ «Все, у кого есть ссылка: читатель» "
                "(или эквивалент для анонимного просмотра)."
            )
        print(f"{msg}: {e.reason}")
        _sync_log(f"[code_base] результат: Ошибка — {msg}")
        if not cache_ok:
            raise RuntimeError(
                f"Нет локального кэша и не удалось загрузить таблицу: {msg}"
            ) from e
        _sync_log("[code_base] переход: используем локальный кэш (CSV недоступен)")
        print("Используем локальный кэш")
    except ValueError as e:
        # Часто HTML-вход вместо CSV — см. текст в исключении
        print(f"Не удалось загрузить GoogleТаблицу (export CSV): {e}")
        _sync_log(f"[code_base] результат: Ошибка — {e}")
        if not cache_ok:
            raise RuntimeError(
                f"Нет локального кэша и не удалось загрузить GoogleТаблицу: {e}"
            ) from e
        _sync_log("[code_base] переход: используем локальный кэш (CSV недоступен)")
        print("Используем локальный кэш")
    except TimeoutError as e:
        lim = float(_EXPORT_TIMEOUT_SEC)
        print(
            f"Экспорт Google CSV: таймаут {lim:g} с ({e!r}) — "
            "прервано, используется локальный кэш base_check/code_base_data.json "
            f"(лимит: CODE_BASE_EXPORT_TIMEOUT_SEC, сейчас {int(lim)})."
        )
        _sync_log(f"[code_base] результат: Ошибка — таймаут export CSV {lim:g} с")
        if not cache_ok:
            raise RuntimeError(
                f"Нет локального кэша и истёк таймаут export CSV ({lim:g} с): {e}"
            ) from e
        _sync_log("[code_base] переход: используем локальный кэш (CSV недоступен)")
        print("Используем локальный кэш")
    except (urllib.error.URLError, OSError, UnicodeDecodeError, csv.Error) as e:
        print(f"Не удалось загрузить GoogleТаблицу (export CSV): {e}")
        _sync_log(f"[code_base] результат: Ошибка — {type(e).__name__}: {e}")
        if isinstance(e, urllib.error.URLError):
            print(
                "Подсказка: Chrome часто использует системный/корпоративный прокси, "
                "а Python urllib — только переменные окружения HTTPS_PROXY / HTTP_PROXY (и при необходимости NO_PROXY). "
                "CODE_BASE_EXPORT_DEBUG=1 выведет словарь прокси из getproxies()."
            )
        if not cache_ok:
            raise RuntimeError(
                f"Нет локального кэша и не удалось загрузить GoogleТаблицу: {e}"
            ) from e
        _sync_log("[code_base] переход: используем локальный кэш (CSV недоступен)")
        print("Используем локальный кэш")


def _load_cache_meta():
    if not os.path.exists(_CACHE_META_FILE):
        return None
    try:
        with open(_CACHE_META_FILE, encoding="utf-8") as f:
            return json.load(f)
    except (json.JSONDecodeError, OSError):
        return None


def _save_cache_meta(modified_time):
    """Совместимость: сохранить только modified_time (старый формат вызовов)."""
    _save_cache_meta_dict(
        {
            "modified_time": modified_time,
            "export_last_modified": "",
            "source": "legacy",
        }
    )


def google_base_vs_base(base_std):
    pass


def google_get_data(url, worksheet):
    gc = pygsheets.authorize(service_file=serv_file)
    # Open spreadsheet and then worksheet
    sh = gc.open_by_url(url)
    # Select worksheet by id, index, title.
    wks = sh.worksheet_by_title(worksheet)
    # # Get a list of all worksheets
    # wks_list = sh.worksheets()
    rng = wks.get_values(None, None, returnas="matrix")
    print(f"Получение данных из ГуглТаблицы листа <{worksheet}> прошло успешно")
    return rng


def google_get_worksheet(url, worksheet):
    gc = pygsheets.authorize(service_file=serv_file)
    # Open spreadsheet and then worksheet
    sh = gc.open_by_url(url)
    # Select worksheet by id, index, title.
    wks = sh.worksheet_by_title(worksheet)
    return wks


def google_update_cell(wks):
    pass


def load_json(file):
    with open(file, encoding="utf-8") as f:
        data = json.load(f)
    return data


def code_base_formating(base, t_com: TableComments):
    base_raw_obj = []
    for row in base:
        code = str_remove_n_x000D_(row[4])  # code
        if check_code(code) == 1:  # Проверяем значение кода продукции по маске
            base_row = RowStd()  # Создаем переменную для внесения значений из строк Google
            base_row.t_com = t_com
            for i in range(len(row)):
                if i in ColNames.GoogleBase.column_dict:  # Проверяем что нужный столбец, есть в словаре МТО колонок
                    key = ColNames.GoogleBase.column_dict[i]  # Из индекса колонки получаем имя столбца: "name" и т.п.
                    # if key == G_BASE_CABLE_LAYING_TYPE and code== "BCC0001706":
                    #     print(key,row[i] )
                    v = row[i]  # Получаем значение из ячейки
                    base_row.el[key] = CheckElement(v, Color.no)  # заносим результат
            base_raw_obj.append(base_row)  # Заносим ряд RowStd в список
    base_std = base_raw_to_std(base_raw_obj)
    return base_std


def base_raw_to_std(base_raw_obj) -> list[RowStd()]:
    base_std = []
    for row in base_raw_obj:  # row -> obj RowStd()
        if check_code(row.el[CODE].value) == 1:
            dic = {}
            for k, v in ColNames.GoogleBase.column_dict.items():
                dic[v] = row.el[v].value
            # if row.el[CODE].value == "BCC0001706":
            #     pass
            temp_row = RowStd.get_std_check_row(dic, row.t_com)
            base_std.append(temp_row)
            # print(row)
    return base_std


if __name__ == "__main__":
    """Smoke test: same chain as the app (`load_base`)."""
    try:
        rows = load_base()
    except Exception as e:
        print(f"load_base: ERROR — {type(e).__name__}: {e}", flush=True)
        sys.exit(1)
    n = len(rows)
    if n == 0:
        print(
            "load_base: WARN — 0 rows after filtering (sheet/cache/codes).",
            flush=True,
        )
        sys.exit(2)
    print(f"load_base: OK — rows after filtering: {n}", flush=True)
    sys.exit(0)
