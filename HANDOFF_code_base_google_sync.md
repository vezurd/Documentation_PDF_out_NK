# Handoff: `base/base_google.py` — синхронизация code base и диагностика

Контекст: **`load_base()`** по умолчанию — **только** публичный HTTPS **export CSV**; этап **Sheets API** включается флагом **`CODE_BASE_SYNC_API_FIRST=1`** (pygsheets + Drive `modifiedTime`, fallback на export).

---

## 1. Как сейчас в **проде** (основной сценарий)

`load_base()` → `_auto_sync()`:

- **По умолчанию** (`CODE_BASE_SYNC_API_FIRST` не задан или не `1`): синк **только** через `_fetch_and_cache_via_export()` — публичный CSV (`docs.google.com/.../export?...`). **Sheets API не вызывается.**
- **Если задано** `CODE_BASE_SYNC_API_FIRST=1` (и есть `base_check/client_secret.json`): сначала API (`_sync_via_api` — Drive `modifiedTime`, при необходимости `get_values`), при ошибке/таймауте — тот же export.

Кэш: `base_check/code_base_data.json`, мета: `base_check/code_base_meta.json`. Остальные env: см. `.cursor/rules/AI_code_base_google_sync.mdc`.

**Pygsheets** в других сценариях: `google_get_data()` / `google_get_worksheet()`.

---

## 2. Запуск `python base/base_google.py` / `python -m base.base_google`

Короткий **smoke-тест** той же цепочки, что в приложении: вызов `load_base()`, проверка числа строк. Без старой «сводки сети» / pygsheets по шагам (удалено).

---

## 3. Последние наблюдения по тестам (релевантно внедрению)
1. **На проблемном контуре** ранее: `curl` к `sheets.googleapis.com` с Schannel — `0x80092012` (отзыв); с `--ssl-no-revoke` — TLS и отправка GET, но иногда таймаут **0 bytes** на чтении ответа → не только OCSP/CRL.
2. **Python**: официальный CPython на Windows для `ssl`/`urllib` обычно **OpenSSL**, не Schannel — симптомы могут **отличаться** от `curl.exe`.
3. **Диагностический прогон** с дефолтным **CERT_NONE** для urllib: в сводке сети вместо таймаутов появились **HTTP 403** на анонимный GET к Sheets API, **Discovery ~378k байт OK**, **OAuth корень 404** — то есть **urllib доходит до HTTP-ответа** с `sheets.googleapis.com`. В том же прогоне **pygsheets прошёл все 4 этапа** — это **не** следствие того же флага по коду (httplib2 отдельный стек); возможна корреляция с сетью/временем, нужны повторы **со `--strict-tls`** для сравнения.
4. **Обход без API**: публичный **export с `docs.google.com`** (как в `load_base()`) на том же ПК в `curl` давал **307** на `googleusercontent.com` — отдельная цепочка хостов, часто живёт, когда `googleapis.com` «тяжёлый» для корпоративного контура.

---

## 4. Что проверить перед / при внедрении «API → потом CSV»

### 4.1 Сеть и TLS (чтобы API-путь был предсказуем)

- Повторить прогон **со `--strict-tls`** (urllib строго): совпадают ли снова таймауты на Sheets в сводке и падает ли pygsheets на шаге 2.
- На целевых ПК: доступ **OCSP/CRL** для цепочек Google (для **curl/Schannel** и при необходимости для корпоративного прокси).
- **SSL inspection** / прокси: исключения для `*.googleapis.com`, `oauth2.googleapis.com` (POST token + GET spreadsheets).
- Таймауты API: отдельный лимит для `spreadsheets.get` / первого ответа, не только 6 с как у export.

### 4.2 Доступ и данные

- Таблица **расшарена** на `client_email` из service account JSON (иначе быстрый 403, не «тихий» таймаут).
- Квоты / включённый **Sheets API** в GCP проекте ключа.
- Семантика **одинаковых данных**: API (`get_values`/матрица) vs CSV (кодировка, пустые хвосты строк) — уже есть `_normalize_matrix` / `code_base_formating`; при смене источника сверить краевые случаи.

### 4.3 Архитектура кода (предложение направления)

1. В `_auto_sync()` (или рядом): **попытка A** — синк через pygsheets (или прямой REST с токеном SA + явный HTTP timeout), запись в тот же `code_base_data.json` + meta с `source: sheets_api` (или аналог).
2. **При ошибке/таймауте** и при наличии условий (например, флаг env или конфиг): **попытка B** — текущий `_fetch_and_cache_via_export()` (как сейчас).
3. Не убрать **кулдаун** и **HEAD Last-Modified** без продумывания: для API другой механизм инкремента (Drive `modifiedTime` с отдельным таймаутом — см. предупреждения в правиле проекта).
4. **Не включать** `CERT_NONE` в прод для `load_base()` без крайней необходимости; если понадобится — только за флагом окружения и с явным логом риска.

### 4.4 Регрессия

- Сценарии **без кэша** / с устаревшим кэшем / **MTO + BBB** в одном процессе (кулдаун).
- Поведение при **HTML входа** вместо CSV (публичный доступ).

---

## 5. Ключевые файлы

| Файл | Назначение |
|------|------------|
| `base/base_google.py` | `load_base()`, `_auto_sync()`, export, опциональный API-этап, `google_get_*`, `if __name__` — smoke `load_base()`. |
| `.cursor/rules/AI_code_base_google_sync.mdc` | История решения (export vs pygsheets, таймауты, кэш). |

---

## 6. Краткий итог для следующего агента

**По умолчанию в проде:** синк code base = **только CSV export** (стабильный путь).  
**Опция:** **`CODE_BASE_SYNC_API_FIRST=1`** — сначала Sheets API (Drive `modifiedTime` + при необходимости `get_values`), при неудаче — **тот же** export. Таймауты API и urllib — в коде и `AI_code_base_google_sync.mdc`.
