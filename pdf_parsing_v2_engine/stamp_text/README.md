# `stamp_text` — очистка текста полей штампа (v2)

Пакет заменяет использование `utils.string_parsing` внутри движка v2. Сюда переносится только то, что нужно для ключей `clean` в JSON-шаблонах.

Общий контракт с остальным движком см. в [`../INTERFACES.md`](../INTERFACES.md) (раздел про `stamp_text`, уровни диагностики page / field / UI).

---

## Структура файлов

| Файл | Роль |
|------|------|
| `context.py` | `StampTextContext` — `doc_type`, `page_num`, опционально `field_id` (передаётся из `stamp_extractor`). |
| `primitives.py` | Нормализация без «семантического мусора»: переносы, даты, NBSP, дефисы, `list_flatter`. |
| `junk.py` | Отбрасывание строк по меткам (бывший `string_Junk_Cleaner`), режимы `JunkMode`. |
| `stamp_field_cleaners.py` | Все доменные клинеры полей штампа (заголовки, BBB 6.x, ревизии/18.x). Секции в файле — навигация. Паттерн **A**: приватное `_* (text, ctx) -> str` + публичное `pipeline_*` → `CleanOutcome` (через `strict_clean_outcome`). Паттерн **B**: `pipeline_doc_title` — свой `CleanOutcome` (strict/relaxed + `ParseWarning`). |
| `outcome.py` | `ParseWarning`, `CleanOutcome`, **`strict_clean_outcome`** — однопутевой strict без предупреждений. |
| `pipelines.py` | Реестр `FIELD_CLEAN_PIPELINES`: ключ `clean` → функция `(text, ctx) → CleanOutcome`. |

Снаружи пакета:

- `field_cleaners.py` — реестр `CLEANERS` (только `str`) и `CLEANER_DESCRIPTIONS` для подсказок в редакторе.
- `stamp_extractor` вызывает `run_field_clean` / `unknown_cleaner_outcome` / `cleaner_exception_outcome`.

---

## Как течёт данные

1. Из PDF читается `raw_text` поля.
2. Если в шаблоне задан `clean`, вызывается **`run_field_clean(clean_key, raw_text, StampTextContext(...))`**.
3. Пайплайн возвращает **`CleanOutcome`**: наружу в `FieldResult` попадает только **`value`** как строка (`cleaned_value`), плюс **`parse_warnings`** и **`clean_tier`**.
4. Для downstream-слоёв предупреждения обходятся через публичный **`V2Document.iter_field_diagnostics()`**; именно этот мост используется, чтобы включать field-level diagnostics в финальный checklist normcontrol без чтения приватного `document._v2_results` из `rules`.
5. В `stamp_field_cleaners.py` тело разбора — приватные **`_* -> str`**; публичные **`pipeline_*`** возвращают **`CleanOutcome`** (`strict_clean_outcome` или ручная сборка). **`pipeline_doc_title`** — особый случай strict/relaxed + предупреждения (паттерн B).

Импортировать **`utils.string_parsing`** из кода в `pdf_parsing_v2/` не нужно и не стоит — дублирование логики переносится сюда осознанно.

---

## Добавить новый ключ `clean`

1. **Домен** — в `stamp_field_cleaners.py`: **`_my_key (text, ctx) -> str`** и **`pipeline_my_key`** → `strict_clean_outcome(..., via="my_key")` либо полный `CleanOutcome` при strict/relaxed. Тяжёлую логику не оставлять в `field_cleaners.py`.
2. **Реестр** — в `pipelines.py` добавить **`"my_key": sf.pipeline_my_key`** в `FIELD_CLEAN_PIPELINES`.
3. **Редактор** — строку в **`CLEANER_DESCRIPTIONS`** в `field_cleaners.py` (tooltip в combo Clean).
4. **Шаблон** — в JSON поля указать тот же ключ в `clean`.

Имена ключей должны совпадать везде: JSON, `FIELD_CLEAN_PIPELINES`, `CLEANER_DESCRIPTIONS`.

---

## Типовые шаблоны (как делать правильно)

Вся доменная логика для ключей `clean` живёт в **`stamp_field_cleaners.py`** (секции по смыслу). Реестр **`pipelines.py`** только подключает функции.

### Паттерн A — один путь, без предупреждений

Подходит, если достаточно одной ветки и **`parse_warnings`** не нужны.

1. В `stamp_field_cleaners.py` в нужной секции:

```python
def _my_simple_field(text: str, ctx: StampTextContext) -> str:
    """Field clean ``my_simple_field`` (body)."""
    del ctx
    out = remove_newlines(text)
    cleaned = junk_clean_lines(out, JunkMode.DEFAULT)
    return " ".join(cleaned).strip() if cleaned else ""


def pipeline_my_simple_field(text: str, ctx: StampTextContext) -> CleanOutcome:
    """Field clean ``my_simple_field``."""
    return strict_clean_outcome(text, _my_simple_field(text, ctx), via="my_simple_field")
```

2. В `pipelines.py`:

```python
"my_simple_field": sf.pipeline_my_simple_field,
```

3. В `field_cleaners.py` → `CLEANER_DESCRIPTIONS` одна строка-описание для tooltip в редакторе.

4. В JSON поля: `"clean": "my_simple_field"`.

### Паттерн B — strict / relaxed и `ParseWarning`

Нужен, когда важно зафиксировать запасной разбор (например «жёсткий» срез текста дал слишком много символов) и передать диагностику в **`FieldResult.parse_warnings`**, отладочный Excel и финальный checklist normcontrol.

1. В `stamp_field_cleaners.py` рядом с хелперами:

```python
def pipeline_my_field(text: str, ctx: StampTextContext) -> CleanOutcome:
    """Field clean ``my_field`` (strict slice; relaxed + warning if suspicious)."""
    del ctx
    strict_raw = _my_field_strict_raw(text)
    value_strict = _my_field_finalize(strict_raw)
    if not _my_field_strict_suspicious(value_strict, strict_raw):
        return CleanOutcome(
            value=value_strict,
            raw_input=text,
            warnings=[],
            via="my_field",
            clean_tier="strict",
        )
    relaxed_raw = _my_field_relaxed_raw(text)
    value_relaxed = _my_field_finalize(relaxed_raw)
    warn = ParseWarning(
        code="MY_FIELD_RELAXED",
        message="Strict parse looked wrong; used relaxed boundaries.",
        tier="relaxed",
    )
    return CleanOutcome(
        value=value_relaxed,
        raw_input=text,
        warnings=[warn],
        via="my_field",
        clean_tier="relaxed",
    )
```

2. В `pipelines.py`:

```python
"my_field": sf.pipeline_my_field,
```

3. Стабильные **`code`** у предупреждений (для фильтров и отчётов), **`message`** — коротко по-русски или по-английски, как принято в проекте для UI.

Эталон в коде: **`pipeline_doc_title`** в `stamp_field_cleaners.py`.

### Чеклист нового ключа

| Шаг | Где |
|-----|-----|
| Реализация | `stamp_field_cleaners.py` (секция) |
| Регистрация | `pipelines.py` — `"key": sf.pipeline_key` |
| Подсказка в UI | `field_cleaners.CLEANER_DESCRIPTIONS` |
| Шаблон JSON | поле `clean` |

### Проверка в редакторе шаблонов

После **«Тестировать»** (извлечение по F5 / кнопка теста) в таблице **«Список полей»** отображаются колонки **`Score`**, **`Значение`**, **`Предупр.`** — в **`Предупр.`** выводится краткий текст по **`FieldResult.parse_warnings`** (полный текст — во всплывающей подсказке ячейки). Так удобно отлаживать клинеры с паттерном B без запуска batch-pipeline.

---

## «Жёсткие» и мягкие проверки

- **Строгий путь** — основная логика парсинга; при успехе **`clean_tier="strict"`** (как у `strict_clean_outcome` / явный `CleanOutcome`).
- **Мягкий путь (relaxed)** — запасной разбор, если строгий дал пусто / не сошлось; при этом нужно:
  - добавить в `CleanOutcome.warnings` элемент **`ParseWarning(..., tier="relaxed")`** с понятным `code` и `message`;
  - выставить **`clean_tier="relaxed"`** на итоговом `CleanOutcome`, если итоговое значение получено именно из мягкой ветки.

Шаблон для кастомного пайплайна без `strict_clean_outcome` (несколько веток):

```python
def pipeline_my_field(text: str, ctx: StampTextContext) -> CleanOutcome:
    warnings: list[ParseWarning] = []
    strict_val = my_strict_parse(text, ctx)
    if strict_val:
        return CleanOutcome(
            value=strict_val,
            raw_input=text,
            warnings=warnings,
            clean_tier="strict",
        )
    relaxed_val = my_relaxed_parse(text, ctx)
    warnings.append(
        ParseWarning(
            code="MY_FIELD_RELAXED",
            message="Short explanation for Excel/UI",
            tier="relaxed",
        )
    )
    return CleanOutcome(
        value=relaxed_val,
        raw_input=text,
        warnings=warnings,
        clean_tier="relaxed",
    )
```

Затем в `FIELD_CLEAN_PIPELINES` зарегистрировать `"my_field": sf.pipeline_my_field` (или локальный `pipeline_my_field`, если живёт в `pipelines.py` — не рекомендуется).

Важно: **`FieldResult.cleaned_value` остаётся строкой** — наружу не отдаётся сам `CleanOutcome`. Предупреждения живут в `FieldResult.parse_warnings`, затем доступны в отладочном Excel (комментарий к ячейке поля) и в финальном checklist normcontrol через `V2Document.iter_field_diagnostics()`.

---

## Юнит-тесты

Имеет смысл тестировать слои отдельно: `primitives`, `junk`, затем доменные функции на коротких фикстурах. Примеры — `pdf_parsing_v2/test_stamp_text.py`.

---

## Не путать с другими уровнями диагностики

- **`V2PageResult.warnings`** — страница (шаблон, score, сетка). Сюда предупреждения полей **не** смешиваются.
- **`FieldResult.parse_warnings`** — парсинг текста поля (`UNKNOWN_CLEANER`, relaxed, исключения в cleaner); источник для debug Excel, редактора и финального checklist normcontrol.
- **UI редактора** — неизвестный ключ в combo `clean` (жёлтая подсветка), это проверка до извлечения, а не замена `parse_warnings`.
