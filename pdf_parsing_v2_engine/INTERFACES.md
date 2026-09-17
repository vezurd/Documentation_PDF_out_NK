# pdf_parsing_v2_engine — interfaces

Краткий слой документации движка (извлечение). Расширенная карта модулей и pipeline по-прежнему в `pdf_parsing_v2/INTERFACES.md`, пока не завершён полный split.

## grid_matcher: statefulness audit

Аудит `pdf_parsing_v2_engine/grid_matcher.py` (09.04.2026) для **ProcessPoolExecutor (PPE)** и границы pickle: модуль ~2900 строк, публичные пути `adapt_all_fields_indexed`, `adapt_by_cell_assignment`, `adapt_all_fields`.

### Сводная таблица (проблема → статус → рекомендация)

| Проблема | Статус | Рекомендация |
|----------|--------|--------------|
| **`_CACHED_V2_CFG`** + **`_load_v2_config()`** — ленивый импорт **`pdf_parsing_v2.v2_config.load_v2_config`** (обратная зависимость **engine → оркестратор**). Срабатывает при **`cfg is None`**: `adapt_all_fields_indexed` (~строка 662), `adapt_by_cell_assignment` (~строка 1114). В **pipeline** `cfg` обычно передан — для PPE безвредно (отдельный процесс = свой кэш при spawn). | **tech-debt** (граница пакетов + скрытый I/O); для PPE при явном `cfg` — **safe** | В параллельных воркерах всегда передавать явный `cfg`; при желании убрать кэш/импорт из движка в нейтральный слой конфигурации. |
| **`_snap_4_boundaries._max_shape_change`** — запись `template.max_shape_change_ratio` в атрибут функции перед снапом (~строка 698), чтение через `getattr` (~строка 523). | **risk** в одном процессе с **потоками**; для **PPE** — **safe** (типично один поток на процесс) | Рефакторинг: передавать порог **параметром** в `_snap_4_boundaries` (или замыкание), убрать мутабельный атрибут. |
| **Мутация входов** `StampTemplate`, `FrameInfo`, `cfg`, `fitz_page` в `adapt_*`. | **safe** для шаблона/рамки/`cfg`/`fitz_page` (нет in-place правок; `outside_stamp` — `replace(field, …)` во временный `FieldDef`). **risk** косвенный: список **`results`** мутируется in-place в **`expand_fields_by_bindings`**, **`close_fragmented_gaps`** (`AdaptResult.bbox`, `status`). | Не шарить один и тот же список `results` между параллельными задачами в одном процессе, если нужна неизменность. |
| **I/O:** файлы, `stdout`, логирование. | **safe** | Прямых `open` / записи в модуле нет. Текст alignment — в памяти (`log_lines` → `AlignmentReport.log`). Неявный диск только при **`cfg is None`** через `_load_v2_config()`. |
| **Pickle:** аргументы и возвраты. | **risk** на границе воркера | **`fitz.Page` не pickle-safe** — в воркере открывать PDF по пути локально. В **результатах** нет `fitz`-объектов: `list[tuple[str, AdaptResult]]`, **`CellAssignmentInfo`** — примитивы/dataclass (при сериализуемых типах в полях шаблона). Кэш `_CACHED_V2_CFG` в pickle задачи не входит; в новом процессе кэш снова `None` до первого вызова с `cfg=None`. |
| **Контракт `cfg`:** `merge_cfg_for_find_tables` делает **`dict(cfg or {})`** (shallow). | **safe** для текущего кода; **tech-debt** при вложенных dict | Текущий путь читает скалярные ключи и дописывает top-level поля шаблона в **копию**; вложённые объекты **разделяются** по ссылке — если когда-либо появится in-place мутация вложенностей, возможна утечка побочного эффекта в вызывающий `cfg`. |
| **Детерминизм:** SciPy `linear_sum_assignment` vs greedy без `scipy`. | **tech-debt** (воспроизводимость) | Фиксировать SciPy в CI или тестировать обе ветки. |

Локальные структуры (`dist_cache` и т.д.) — только внутри вызовов. Числовые константы модуля (`_EPS`, `_MIN_FIELD_PTS`, …) — без побочных эффектов.

### Verification

- Корректная проверка объявлений кэша: **`global _CACHED_V2_CFG`** — ровно **2** вхождения (`adapt_all_fields_indexed`, `adapt_by_cell_assignment`).
- Поиск по подстроке **`global `** (с пробелом) даёт **ложные совпадения** в комментариях («global transform», путь к `grid_matcher` в тексте лога и т.д.); для регресса использовать паттерн выше.

### Рекомендации для будущего `parallel.py` (без реализации здесь)

1. Передавать **явный `cfg`** (как pipeline/F5), чтобы воркеры не тянули `load_v2_config` через `grid_matcher`.
2. DTO в parent-процесс без живых объектов MuPDF.
3. При **ThreadPoolExecutor** вокруг `grid_matcher` — устранить `_snap_4_boundaries._max_shape_change` как общий mutable.

## `grid_matcher.py`: statefulness audit (09.04.2026)

Three issues relevant to parallel processing (ProcessPoolExecutor):

### 1. `_CACHED_V2_CFG` (line 31)

Module-level `dict | None`, lazy-loaded via `_load_v2_config()` which imports
from `pdf_parsing_v2.v2_config` — **reverse dependency** engine → orchestrator.

- In pipeline: `cfg` is always passed → fallback not used.
- In PPE workers: `cfg` passed via `initargs` → fallback not used.
- **Risk:** circular import at module level (guarded by lazy import inside function).
- **Recommendation:** document; in the future remove fallback, make `cfg` mandatory
  in `adapt_by_cell_assignment` and `_align_grid_lines_ordered`.

### 2. `_snap_4_boundaries._max_shape_change` (line 698 set, line 523 get)

`max_shape_change_ratio` from template written as a **function attribute** —
**thread-unsafe** (shared mutable state within a process).

- **ProcessPoolExecutor:** safe (per-process copy of module).
- **ThreadPoolExecutor:** NOT safe (do NOT use threads for extraction).
- **Recommendation:** refactor to pass as parameter in a future session.

### 3. `cfg` flat dict contract

`extract_all_pages_for_file` does `dict(cfg)` (shallow copy).
**Contract:** `cfg` must be a flat dict with primitive values only
(str, int, float, bool, None). Nested mutable objects would share
references across calls.

- Enforced by convention, not by code.
- PPE workers receive a deep copy (pickle/unpickle).
