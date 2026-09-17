# Спека: `pdf_parsing_v2/parallel.py` + `v2_timing.py`

> **Фаза 4A [PREMIUM]** из `AI_v2_parallel.mdc`. Для реализации фазой 4B [C2].

## Цель

Параллельная обработка нескольких PDF файлов в `v2_pipeline` через
`ProcessPoolExecutor`. Единица параллелизма — **один PDF файл** (не страница).

## Файлы

| Файл | Пакет | Назначение |
|------|-------|------------|
| `parallel.py` | `pdf_parsing_v2` | PPE для extraction фазы |
| `v2_timing.py` | `pdf_parsing_v2` | Сбор таймингов pipeline (extraction + tags + postprocess) |
| `v2_callbacks.py` | `pdf_parsing_v2` | Протоколы callback-ов (извлечены из `v2_pipeline.py`) |

---

## 1. DTOs (pickle-safe)

### `V2FileExtractResult`

```python
@dataclass
class V2FileExtractResult:
    """Extraction result for one PDF file (pickle-safe, returned by worker)."""

    file_path: str
    doc_type: str
    index: int                     # original index in curr_proj
    pages: list[V2PageResult]
    elapsed_sec: float
    error: str | None = None

    @property
    def ok(self) -> bool:
        return self.error is None
```

Все поля — примитивы или pickle-safe dataclass-ы (см. `SPEC_pipeline_phases.md §4`).
`fitz` объекты **не** попадают в `V2PageResult`.

### `V2ParallelResult`

```python
@dataclass
class V2ParallelResult:
    """Aggregate result from parallel extraction."""

    file_results: list[V2FileExtractResult]   # sorted by original index
    total_elapsed_sec: float
    n_workers: int

    @property
    def n_files(self) -> int: ...

    @property
    def n_pages_total(self) -> int: ...

    @property
    def n_errors(self) -> int: ...
```

---

## 2. Callback-протокол

```python
class ParallelCallback(Protocol):
    """Progress reporting for parallel extraction.

    All methods called from the **main process** (not from workers).
    Safe to update GUI via signals from these methods.
    """

    def on_file_start(self, file_path: str, index: int, total: int) -> None: ...

    def on_file_done(
        self, file_path: str, index: int, n_pages: int, elapsed_sec: float
    ) -> None: ...

    def on_file_error(
        self, file_path: str, index: int, error: Exception
    ) -> None: ...

    def on_batch_progress(self, completed: int, total: int) -> None: ...
```

**Совместимость с `ExtractionCallback`:**
`ParallelCallback` — надмножество `ExtractionCallback` (добавляет `on_batch_progress`).
Реализация callback может поддерживать оба протокола (duck typing).

---

## 3. Worker

### Инициализация (per-process, однократно)

```python
_WORKER_TEMPLATES: list[StampTemplate] | None = None
_WORKER_CFG: dict | None = None

def _worker_init(templates_bytes: bytes, cfg_bytes: bytes) -> None:
    """Unpickle shared data once per worker process."""
    global _WORKER_TEMPLATES, _WORKER_CFG
    _WORKER_TEMPLATES = pickle.loads(templates_bytes)
    _WORKER_CFG = pickle.loads(cfg_bytes)
```

Шаблоны и cfg pickle-ются **один раз** в main и передаются через `initargs`,
а не через каждый `submit`. Это экономит ~MB трафика per-file.

### Рабочая функция

```python
def _worker_extract(
    file_path: str, doc_type: str, index: int
) -> V2FileExtractResult:
    """Process one PDF in a worker process.

    Imports engine lazily (first call per worker).
    Returns V2FileExtractResult (always, never raises).
    """
    from pdf_parsing_v2_engine.stamp_extractor import extract_all_pages_for_file

    t0 = time.perf_counter()
    try:
        pages = extract_all_pages_for_file(
            file_path, doc_type, _WORKER_TEMPLATES, _WORKER_CFG
        )
        return V2FileExtractResult(
            file_path=file_path,
            doc_type=doc_type,
            index=index,
            pages=pages,
            elapsed_sec=time.perf_counter() - t0,
        )
    except Exception as e:
        return V2FileExtractResult(
            file_path=file_path,
            doc_type=doc_type,
            index=index,
            pages=[],
            elapsed_sec=time.perf_counter() - t0,
            error=f"{type(e).__name__}: {e}",
        )
```

**Collect-and-continue:** исключение ловится внутри worker, записывается в
`V2FileExtractResult.error`. Один сбойный PDF не роняет batch.

**Thread-safety `grid_matcher`:** `_snap_4_boundaries._max_shape_change` —
атрибут функции (thread-unsafe), но в `ProcessPoolExecutor` каждый worker —
**отдельный процесс**, потому безопасно (per-process copy модуля).

### Максимальное число воркеров

```python
def _default_max_workers() -> int:
    cpus = os.cpu_count() or 4
    return max(1, min(cpus - 1, 8))
```

- `cpu_count - 1`: оставить ядро для main process (GUI, postprocess).
- Cap 8: ограничение по I/O (все воркеры читают PDF с одного диска).
- Конфигурируется через `cfg["max_workers"]` или аргумент.

---

## 4. Public API

### `run_parallel` — низкоуровневый

```python
def run_parallel(
    tasks: list[tuple[str, str, int]],
    templates: list[StampTemplate],
    cfg: dict,
    *,
    max_workers: int | None = None,
    callback: ParallelCallback | None = None,
    timing: TimingCollector | None = None,
) -> V2ParallelResult:
    """Run extraction for multiple PDF files in parallel.

    Args:
        tasks: list of ``(file_path, doc_type, original_index)``.
        templates: loaded stamp templates (pickle-safe).
        cfg: flat v2 config dict (shallow-copy-safe).
        max_workers: PPE workers; ``None`` → auto.
        callback: progress reporting (main process).
        timing: optional timing collector.

    Returns:
        V2ParallelResult with ``file_results`` sorted by original index.
    """
```

**Алгоритм:**
1. `pickle.dumps(templates)`, `pickle.dumps(cfg)` — однократно.
2. `ProcessPoolExecutor(max_workers, initializer=_worker_init, initargs=(...))`.
3. `executor.submit(_worker_extract, file_path, doc_type, idx)` для каждой задачи.
4. `as_completed` loop: собирает результаты, вызывает callback, записывает в timing.
5. `results.sort(key=lambda r: r.index)` — детерминизм.
6. Возврат `V2ParallelResult`.

### `run_parallel_extraction` — drop-in замена `run_v2_extraction`

```python
def run_parallel_extraction(
    curr_proj: list[doc_ATTRIBUTES],
    templates: list[StampTemplate],
    cfg: dict,
    *,
    max_workers: int | None = None,
    callback: ParallelCallback | None = None,
    timing: TimingCollector | None = None,
) -> None:
    """Parallel replacement for ``run_v2_extraction``.

    Mutates *curr_proj*: fills ``document.pages`` and ``document._v2_results``
    in the same way as the sequential version (deterministic order).
    """
```

**Алгоритм:**
1. Если `len(curr_proj) < 2` → fallback на sequential `run_v2_extraction`.
2. Строит tasks из `curr_proj`: `(doc.file_full_path, doc.doc_Type, idx)`.
3. Вызывает `run_parallel(tasks, ...)`.
4. Для каждого `V2FileExtractResult`:
   - Если `ok`: `to_page_stamp_attributes(v2r)` → `document.pages.append(psa)`,
     заполняет `page_marka`, `page_type`, `page_title`.
     Устанавливает `document._v2_results = fr.pages`.
   - Если error: `print(...)` (collect-and-continue).
5. Print summary (файлов, страниц, воркеров, время, ошибки).

---

## 5. Timing infrastructure (`v2_timing.py`)

### `TimingEntry`

```python
@dataclass
class TimingEntry:
    label: str
    elapsed_sec: float
    detail: dict[str, Any] = field(default_factory=dict)
```

### `TimingCollector`

```python
class TimingCollector:
    """Accumulate timing data across pipeline phases.

    Usage::

        timing = TimingCollector()
        timing.record("find_files", 0.12)
        timing.record_file("WIR-0005.pdf", 91.0, n_pages=5)
        timing.to_excel("v2_timing_log.xlsx")
    """

    def __init__(self) -> None: ...

    def record(self, label: str, elapsed_sec: float, **detail: Any) -> None:
        """Record a pipeline phase timing."""
        ...

    def record_file(self, file_name: str, elapsed_sec: float, **detail: Any) -> None:
        """Record per-file extraction timing."""
        ...

    @property
    def entries(self) -> list[TimingEntry]: ...

    @property
    def file_timings(self) -> dict[str, dict[str, Any]]: ...

    def summary(self) -> dict[str, Any]:
        """Return JSON-serializable summary."""
        ...

    def to_excel(self, path: str) -> None:
        """Write timing data to xlsx (2 sheets: Phases, Files)."""
        ...
```

**Двойная компоновка:**
- `record(label, elapsed_sec)` — для фаз pipeline: `find_files`, `extraction`,
  `tags`, `od_parsing`, `tag_analyze`, `rules_check`, `debug_report`, `total`.
- `record_file(file_name, elapsed_sec, n_pages=, error=)` — per-file (sequential
  или parallel).

**Excel вывод (2 листа):**
- **Phases**: label, elapsed_sec, detail.
- **Files**: file_name, elapsed_sec, n_pages, error.

Файл: `{out_result_dir}/v2_timing_log.xlsx`.

---

## 6. Интеграция в `run_v2_pipeline`

```python
def run_v2_pipeline(pdf_path, cfg, project=None):
    timing = TimingCollector() if cfg.get("timing_log", False) else None
    ...
    use_parallel = cfg.get("parallel", False) and len(curr_proj) > 1

    # Phase 1: Extraction
    t_ext = time.perf_counter()
    if use_parallel:
        from pdf_parsing_v2.parallel import run_parallel_extraction
        run_parallel_extraction(
            curr_proj, templates, cfg,
            max_workers=cfg.get("max_workers"),
            timing=timing,
        )
    else:
        run_v2_extraction(curr_proj, templates, cfg)
    if timing:
        timing.record("extraction", time.perf_counter() - t_ext)

    # Phase 2: Tags (sequential — pdfminer, separate pass)
    t_tags = time.perf_counter()
    run_v2_tags(curr_proj)
    if timing:
        timing.record("tags", time.perf_counter() - t_tags)

    # Phase 3: Postprocess
    t_post = time.perf_counter()
    run_v2_postprocess(...)
    if timing:
        timing.record("postprocess", time.perf_counter() - t_post)

    if timing:
        timing.record("total_pipeline", time.perf_counter() - t_total_start)
        timing.to_excel(os.path.join(out_result_dir, "v2_timing_log.xlsx"))
```

### Конфигурация (`pdf_v2_config.json`)

```json
{
    "parallel": false,
    "max_workers": null,
    "timing_log": false
}
```

- `parallel: true` — включить PPE для extraction.
- `max_workers: null` — авто (cpu_count - 1, max 8). Или число.
- `timing_log: true` — собирать тайминги и писать xlsx.

### GUI (будущее, фаза 5)

Чекбоксы в `pdf_v2_settings_gui.py`:
- ☐ Параллельная обработка
- Количество воркеров: [авто ▾]
- ☐ Лог таймингов

---

## 7. Что НЕ вводить в этой фазе

| Что | Когда |
|-----|-------|
| PPE для `parse_tags` (pdfminer) | Отдельная задача (PPE тегов) |
| `tag_parser.analyze(source_dict=)` | Перед PPE тегов |
| GUI мониторинг воркеров (`pdf_v2_monitor/`) | Фаза 5 `AI_v2_parallel.mdc` |
| `shared_memory` для шаблонов | Только если профилирование покажет bottleneck |
| Кэш `V2PageResult` на диск | Отдельная задача (offline postprocess) |

---

## 8. Pickle-safety checklist

| Объект | Pickle-safe? | Примечание |
|--------|-------------|------------|
| `StampTemplate` | ✅ | dataclass, примитивы + list[FieldDef] |
| `FieldDef` | ✅ | dataclass, примитивы |
| `V2PageResult` | ✅ | см. SPEC_pipeline_phases.md §4 |
| `FieldResult` | ✅ | str, tuple, bool, list |
| `cfg` dict | ✅ | flat dict, примитивные значения |
| `doc_ATTRIBUTES` | ❌ НЕ pickle'ится | Остаётся в main process |
| `fitz.Document` | ❌ | Не попадает в возвращаемое значение |

---

## 9. Риски и ограничения

1. **I/O contention:** все воркеры читают PDF с одного диска.
   Лимит `max_workers ≤ 8` смягчает. Для SSD — не проблема.

2. **`_CACHED_V2_CFG`** в `grid_matcher.py`: module-level dict, lazy load.
   В worker'е: `cfg` передан через initargs → fallback не используется.
   Но обратная зависимость engine→orchestrator (`import v2_config`) может
   сработать при первом вызове. Рекомендация: убрать fallback, сделать `cfg`
   обязательным. Отслеживается в `audit-grid-matcher` todo.

3. **Windows `spawn`:** `ProcessPoolExecutor` на Windows — `spawn` (не `fork`).
   Все аргументы worker'а ДОЛЖНЫ быть pickle-safe. Проверено: ✅.

4. **Детерминизм порядка:** `as_completed` возвращает в произвольном порядке.
   `results.sort(key=lambda r: r.index)` восстанавливает исходный порядок.
   `curr_proj[fr.index]` = правильный документ.

5. **Порог параллелизма:** при `len(curr_proj) < 2` — fallback на sequential.
   Overhead PPE (spawn + pickle) не оправдан для одного файла.

---

## 10. `per_page.doc_types` (ключи)

Список ключей совпадает с `PER_PAGE_DOC_TYPES` в `pdf_parsing_v2/v2_config.py`.
Типы **GA** и **PL** — чертёжный кластер вместе с **WIR**/**LAY**/**CAE**/**DW**;
флаги selective per-page для них задаются в `pdf_v2_config.json` так же, как для **LAY**/**CAE**.
