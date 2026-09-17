---
name: v2 decoupling from v1 + parallel extraction + timing
overview: "Отвязка v2-пакетов от v1 (pdf_parsing.*) по группам A–D; реализация параллельной обработки PDF (фаза 4 AI_v2_parallel.mdc); инфраструктура таймингов v2_timing.py. Продолжение v2_packages_split."
todos:
  - id: stamp-fields-copy
    content: "[C2] Скопировать 30+ c_* констант из pdf_parsing.shtamp_extract_classes в pdf_parsing_v2_engine/stamp_fields.py. Обновить импорты в checks.py (28 констант) и compat.py (6 констант). Оставить doc_ATTRIBUTES/PageStampAttributes как v1-импорт. Verification: python -c import + rg 'from pdf_parsing.shtamp_extract_classes import c_'"
    status: completed
  - id: audit-grid-matcher
    content: "[C2+V] Аудит grid_matcher: (1) _CACHED_V2_CFG — убрать или задокументировать обратную зависимость engine→orchestrator; (2) _snap_4_boundaries._max_shape_change — задокументировать thread-unsafe / process-safe; (3) cfg flat dict — контракт в INTERFACES.md. Результат: заметка в INTERFACES.md engine."
    status: completed
  - id: tag-parser-analyze-param
    content: "[C2] Добавить параметр source_dict_override=None в tag_parser.analyze(). Если None — текущее поведение (глобальный source_dict). Если передан — использовать его. Обратная совместимость. Verification: v2_pipeline smoke test."
    status: completed
  - id: parallel-impl
    content: "[C2] Реализовать pdf_parsing_v2/parallel.py по спеке SPEC_parallel.md: V2FileExtractResult, V2ParallelResult, ParallelCallback, _worker_init/_worker_extract, run_parallel, run_parallel_extraction. Verification: python -c 'from pdf_parsing_v2.parallel import run_parallel'"
    status: completed
  - id: timing-impl
    content: "[C2] Реализовать pdf_parsing_v2/v2_timing.py по спеке SPEC_parallel.md §5: TimingEntry, TimingCollector (record, record_file, summary, to_excel). Verification: unit test создания xlsx."
    status: completed
  - id: pipeline-parallel-integration
    content: "[C2] Интегрировать parallel + timing в v2_pipeline.py: cfg['parallel'], cfg['max_workers'], cfg['timing_log'] → run_parallel_extraction вместо run_v2_extraction при parallel=True и len>1. Тайминги для всех фаз. Verification: smoke test parallel=false (регрессия), parallel=true на 3+ PDF."
    status: completed
  - id: pipeline-timing-integration
    content: "[C2] v2_pipeline.py: TimingCollector вокруг find_files, extraction, tags, postprocess, total. to_excel в out_result_dir/v2_timing_log.xlsx при timing_log=true. Verification: наличие xlsx после прогона с timing_log=true."
    status: completed
  - id: syspath-consolidate
    content: "[C2] Консолидировать sys.path manipulations: проверить mto_grid_diagnostic.py и pdf_template_editor/__main__.py — убедиться что используют тот же паттерн что v2_pipeline._ensure_project_root_on_path. Задокументировать в INTERFACES.md. Не трогать v1."
    status: completed
  - id: docs-update-decoupling
    content: "[C2] Обновить AI_memory.mdc (§14 ловушки: parallel, stamp_fields), AI_v2_parallel.mdc (статус → реализовано 4A, ссылка на SPEC), pdf_parsing_v2/__init__.py (реэкспорт run_parallel)."
    status: completed
  - id: v2-domain-future
    content: "[BACKLOG] Создать stamp_domain/ или pdf_shared/: перенести doc_ATTRIBUTES, PageStampAttributes, c_* в общий пакет. Обновить utils.path, v1, v2. Ввести V2Document dataclass. Отложено — требует ревью границ v1."
    status: pending
isProject: false
---

# План: отвязка v2 от v1 + параллель + тайминги

## Контекст

Продолжение `v2_packages_split`. Все пакеты v2 (`pdf_parsing_v2/`, `_engine/`, `_od/`,
`_tags/`, `_rules/`) созданы и работают. Остались:
1. Зависимости от v1 (`pdf_parsing.*`, `utils.*`, `tags.*`)
2. Параллельная обработка (фаза 4 из `AI_v2_parallel.mdc`)
3. Инфраструктура таймингов

Архитектурный анализ зависимостей — в чате PREMIUM-ревью (09.04.2026).
Спека parallel.py — `pdf_parsing_v2/SPEC_parallel.md`.

## Группа A: pdf_parsing.shtamp_extract_classes

### Стратегия

**Гибрид Copy-and-own + conscious coupling:**

| Объект | Решение | Обоснование |
|--------|---------|-------------|
| c_* константы (30+) | **Копировать** в `stamp_fields.py` | Строковые литералы, тривиальный copy; устраняет 34 импорта |
| `PageStampAttributes` | **Оставить** в v1, импорт в `compat.py` | Это назначение compat; перенос тянет весь v1 runtime |
| `doc_ATTRIBUTES` | **Оставить** в v1, TYPE_CHECKING | Инстанцирование в `utils.path`; для PPE не pickle'ится |

### Шаг A1: `stamp_fields.py` [C2]

1. Создать `pdf_parsing_v2_engine/stamp_fields.py`:
   - Скопировать ВСЕ `c_*` константы из `pdf_parsing/shtamp_extract_classes.py` (строки 13–65).
   - Docstring: назначение + примечание о синхронизации с v1.

2. Обновить импорты:
   - `pdf_parsing_v2_rules/checks.py`: 28 импортов → `from pdf_parsing_v2_engine.stamp_fields import ...`
   - `pdf_parsing_v2_engine/compat.py`: 6 импортов → `from pdf_parsing_v2_engine.stamp_fields import ...`
   - `compat.py`: оставить импорт `PageStampAttributes` из v1 (единственный оставшийся).

3. `v2_pipeline.py`: TYPE_CHECKING импорт `doc_ATTRIBUTES` — оставить (не c_*).

**Verification:**
```
python -c "from pdf_parsing_v2_engine.stamp_fields import c_1_DOC_TITLE"
python -c "from pdf_parsing_v2_rules.checks import run_all_checks"
python -c "from pdf_parsing_v2_engine.compat import to_page_stamp_attributes"
rg "from pdf_parsing.shtamp_extract_classes import c_" --type py
# Ожидание: 0 результатов в pdf_parsing_v2* (кроме compat для PageStampAttributes)
```

### Будущее: stamp_domain/ [BACKLOG]

Создать `stamp_domain/` (или `pdf_shared/`):
- Перенести `doc_ATTRIBUTES`, `PageStampAttributes`, `c_*`, `dict_coordinate`
- v1 и v2 оба импортируют из domain
- `utils.path` переключается на domain
- Требует ревью границ v1 → отложено

---

## Группа B: utils.* — решение

**Оставить `utils/` как shared** — это проектная инфраструктура, не v1-специфичная.

Конкретные зависимости:
| Файл v2 | Модуль utils | Функции | Действие |
|----------|-------------|---------|----------|
| `checks.py` | `string_parsing` | ~10 функций | Оставить |
| `checks.py` | `file_name_converts` | `ProjectFileName` | Оставить |
| `od_parsing.py` | `string_parsing`, `file_name_converts`, `formats_A` | 3 модуля | Оставить |
| `extractors.py` | `formats_A` | `get_Real_Page_Format` | Оставить |
| `output.py` | `path` | `make_dir` | Оставить |
| `v2_pipeline.py` | `path` | `get_files_single`, `get_path_out_dir` | Оставить |
| `main_window.py` | `string_parsing` | `getDocTypeFromFile` | Оставить |

Никаких действий не требуется. `utils/` — общая библиотека.

---

## Группа C: tags.tag_parser

### Текущее состояние

- ✅ `add_context(target_dict=)` — per-file usage готов
- ✅ `merge_dicts(base, local)` — merge готов
- ❌ `analyze(out_dir)` — читает глобальный `source_dict`, нет параметра

### Шаг C1: параметр source_dict [C2]

Файл: `tags/tag_parser.py`, функция `analyze()` (строка 70):

```python
# Было:
def analyze(out_dir, main_doc_title="МТО", ...):
    ...
    for i, v in source_dict.items():

# Стало:
def analyze(out_dir, main_doc_title="МТО", ..., *, source_dict_override=None):
    ...
    src = source_dict_override if source_dict_override is not None else source_dict
    for i, v in src.items():
```

**Обратная совместимость:** `source_dict_override=None` → текущее поведение.

**Verification:** v2_pipeline smoke test (tags phase без изменений вызова).

---

## Группа D: sys.path

### Текущее состояние

| Файл | Строки | Паттерн |
|------|--------|---------|
| `v2_pipeline.py:50-53` | `_ensure_project_root_on_path()` | Хорошо изолировано |
| `mto_grid_diagnostic.py:17-20` | `sys.path.insert(0, _ROOT)` | Скрипт-утилита |
| `pdf_template_editor/__main__.py:6` | `sys.path.insert(...)` | Entry point |

### Шаг D1: консолидация [C2]

1. Проверить что `mto_grid_diagnostic.py` и `__main__.py` используют тот же подход.
2. Задокументировать в `INTERFACES.md` оркестратора.
3. Будущее: `pyproject.toml` + `pip install -e .` (отдельная задача build infra).

---

## Параллель: фаза 4A+4B

### Пререквизит: аудит grid_matcher [C2+V]

Из ревью v2_packages_split (09.04.2026) — 3 проблемы:

1. **`_CACHED_V2_CFG`** (module-level): lazy load → обратная зависимость engine→orchestrator.
   - В pipeline cfg всегда передан → fallback не используется.
   - В worker: cfg через initargs → тоже не используется.
   - Рекомендация: задокументировать; в перспективе удалить fallback.

2. **`_snap_4_boundaries._max_shape_change`**: атрибут функции, thread-unsafe.
   - PPE (ProcessPool): per-process copy → **безопасно**.
   - ThreadPool: **НЕ безопасно** (не планируется).
   - Рекомендация: задокументировать; рефактор в параметр — отложить.

3. **`cfg` flat dict**: `extract_all_pages_for_file` делает `dict(cfg)` (shallow copy).
   - Зафиксировать контракт: cfg — flat dict, примитивные значения.

### Шаг P1: parallel.py [C2] по SPEC_parallel.md

Содержимое:
- DTOs: `V2FileExtractResult`, `V2ParallelResult`
- `ParallelCallback` Protocol
- `_worker_init`, `_worker_extract`
- `run_parallel(tasks, templates, cfg, ...)`
- `run_parallel_extraction(curr_proj, templates, cfg, ...)` — drop-in для `run_v2_extraction`

### Шаг P2: v2_timing.py [C2] по SPEC_parallel.md §5

Содержимое:
- `TimingEntry`, `TimingCollector`
- `record`, `record_file`, `summary`, `to_excel`

### Шаг P3: интеграция в v2_pipeline [C2]

- `cfg["parallel"]`, `cfg["max_workers"]`, `cfg["timing_log"]`
- При `parallel=True` и `len(curr_proj) > 1` → `run_parallel_extraction`
- `TimingCollector` вокруг всех фаз
- `to_excel` в `out_result_dir/v2_timing_log.xlsx`
- `pdf_parsing_v2/__init__.py`: реэкспорт `run_parallel`

**Verification:**
- `python -c "from pdf_parsing_v2.parallel import run_parallel, run_parallel_extraction"`
- Smoke test sequential (parallel=false) — тот же набор выходных файлов
- Smoke test parallel (parallel=true) на 3+ PDF — тот же набор выходных файлов
- `v2_timing_log.xlsx` при `timing_log=true` — 2 листа (Phases, Files)

---

## Порядок сессий

| Сессия | Содержание | Модель | ~Шагов |
|--------|-----------|--------|--------|
| **S1** | A1: stamp_fields.py + обновление импортов checks+compat | [C2] | ~3 |
| **S2** | Аудит grid_matcher (заметка INTERFACES.md) | [C2+V] | 1 |
| **S3** | C1: tag_parser.analyze(source_dict_override=) | [C2] | 1 |
| **S4** | P1+P2: parallel.py + v2_timing.py | [C2] | ~5 |
| **S5** | P3: интеграция в pipeline + smoke tests | [C2] | ~4 |
| **S6** | D1 + docs-update | [C2] | ~3 |

**Зависимости:** S1 → S4 (stamp_fields нужен для воркера).
S2 → S4 (аудит нужен до реализации parallel). S3 независим.

---

## Риски

| Риск | Вероятность | Смягчение |
|------|-------------|-----------|
| c_* drift (v1 изменит значение, v2 нет) | Низкая | CI тест; значения = ГОСТ-идентификаторы, стабильны |
| I/O contention при PPE | Средняя | max_workers ≤ 8; SSD снижает проблему |
| _CACHED_V2_CFG в worker | Низкая | cfg передан через initargs; fallback не сработает |
| Нарушение детерминизма | Низкая | sort by index; тесты на порядок |
| pickle failure нового типа | Низкая | Чеклист в SPEC_parallel.md §8 |
