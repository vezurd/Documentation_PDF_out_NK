# Спека: разрез `v2_pipeline` на 3 фазы

> **Шаг 1A [PREMIUM]** из плана `v2_packages_split`. Для реализации шагом 1D [C2].

## Текущее состояние

`run_v2_pipeline` — монолитная функция с 9 этапами:
setup → file discovery → templates → per-file extraction → tags → OD → tag_parser.analyze → rules → debug report/visual.

## Целевая структура

```
run_v2_pipeline(pdf_path, cfg, project) -> str:
    # --- Setup (остаётся в compose) ---
    _ensure_project_root_on_path()
    curr_proj = get_files_single(pdf_path)
    out_result_dir = get_path_out_dir(pdf_path)
    templates = load_[project_]templates(templates_dir, ...)
    effective_project = project or cfg.get("project") or None

    # --- Phase 1: Extraction ---
    run_v2_extraction(curr_proj, templates, cfg)

    # --- Phase 2: Tags ---
    run_v2_tags(curr_proj)

    # --- Phase 3: Postprocess ---
    run_v2_postprocess(
        curr_proj, cfg, pdf_path, out_result_dir,
        effective_project=effective_project,
        templates=templates,
    )
    return out_result_dir
```

---

## 1. `run_v2_extraction`

```python
def run_v2_extraction(
    curr_proj: list[doc_ATTRIBUTES],
    templates: list[StampTemplate],
    cfg: dict[str, Any],
    *,
    callback: ExtractionCallback | None = None,
) -> None:
```

**Что делает:** для каждого `document` в `curr_proj` вызывает
`extract_all_pages_for_file` (engine) → `to_page_stamp_attributes` (compat) →
мутирует `document.pages` и `document._v2_results`.

**Мутация `curr_proj`:**
- `document.pages.append(psa)` — v1-совместимые `PageStampAttributes`
- `document._v2_results = v2_results` — сырые `list[V2PageResult]` (для отчётов)

**Обработка ошибок:** per-file try/except; при ошибке файла — print + `callback.on_file_error` (если есть).
Файл **не прерывает** обработку остальных (collect-and-continue).

**Логирование:** при `cfg.get("debug_verbose_log")` — `_log_v2_result` per page.

### PPE-замена (будущее, `parallel.py`)

```python
run_parallel_extraction(curr_proj, templates, cfg, *, max_workers, callback)
```

Воркер вызывает `extract_all_pages_for_file(file_path, doc_type, templates, cfg)`.
Главный процесс получает `list[V2PageResult]` per file, вызывает `to_page_stamp_attributes`,
заполняет `curr_proj` **в порядке исходного индекса файла** (детерминизм).

---

## 2. `run_v2_tags`

```python
def run_v2_tags(
    curr_proj: list[doc_ATTRIBUTES],
    *,
    tag_dict: dict | None = None,
) -> dict:
```

**Что делает:** вызывает `parse_tags(curr_proj, tag_dict=tag_dict)`.

**Возвращает:**
- Если `tag_dict=None` — `tag_parser.source_dict` (глобальный, сброшен и заполнен).
- Если `tag_dict` передан — тот же dict, заполненный.

**Побочный эффект:** при `tag_dict=None` вызывается `tag_parser.reset()` внутри `parse_tags`.

### PPE-замена (будущее)

Per-file: `local = {}` → `parse_tags([doc], tag_dict=local)` → `tag_parser.merge_dicts(merged, local)`.
После merge: `tag_parser.source_dict = merged` (для `analyze`).

### Tech debt (зафиксировать, не реализовывать сейчас)

`tag_parser.analyze()` читает глобальный `source_dict` — нет параметра override.
Добавить `analyze(out_dir, *, source_dict=None)` **перед** PPE тегов.

---

## 3. `run_v2_postprocess`

```python
def run_v2_postprocess(
    curr_proj: list[doc_ATTRIBUTES],
    cfg: dict[str, Any],
    pdf_path: str,
    out_result_dir: str,
    *,
    effective_project: str | None = None,
    templates: list[StampTemplate] | None = None,
) -> None:
```

**Что делает (последовательно):**
1. OD: `od_table_parsing_f(pdf_path, 1)` → `proj_od_list`
2. Tag analysis: `tag_parser.analyze(out_result_dir)`
3. Rules: `rules_check_start(curr_proj, proj_od_list, pdf_path, out_result_dir)`
4. Debug Excel: `collect_results_from_curr_proj(curr_proj)` → `save_v2_debug_report(...)` (если `export_debug_excel`)
5. Debug visual: `render_debug_for_file(...)` (если `debug_visual`)

**Зависимости:** шаги 1–3 используют v1 imports (lazy);
`templates` нужен только для debug visual (шаг 5); если `None` — visual пропускается.

**Ошибки:** per-step try/except (как сейчас), не прерывают pipeline.

---

## 4. `extract_all_pages_for_file` (engine, публичная)

```python
def extract_all_pages_for_file(
    file_path: str,
    doc_type: str,
    templates: list[StampTemplate],
    cfg: dict[str, Any],
) -> list[V2PageResult]:
    """Process one PDF file — pure function, safe for PPE.

    Opens and closes fitz document internally.
    Sets ``cfg["source_pdf_basename"]`` from *file_path*.
    No global state, no side effects beyond file I/O (read-only).

    All arguments and return values are pickle-safe
    (no fitz objects in V2PageResult).
    """
```

Это текущая `_process_single_file`, переименованная и перенесённая в engine
(`pdf_parsing_v2_engine/stamp_extractor.py` или отдельный `extraction.py`).

### Pickle-safety возвращаемого значения

| Тип | Поля | Pickle |
|-----|------|--------|
| `V2PageResult` | int, str, float, dict[str,FieldResult], dict[str,Any], list[str] | OK |
| `FieldResult` | str, str\|None, tuple[float×4], bool\|None, list[ParseWarning], str\|None | OK |
| `FrameInfo` | float×6, int, float×3 | OK |
| `ParseWarning` | str, str, str\|None | OK |

Ни один dataclass не содержит `fitz.*` объектов. **Pickle-safe.**

---

## 5. Callback-протокол (опционально в этой волне)

```python
from typing import Protocol

class ExtractionCallback(Protocol):
    """Optional progress reporting for extraction phase."""

    def on_file_start(self, file_path: str, index: int, total: int) -> None: ...
    def on_file_done(self, file_path: str, index: int, n_pages: int, elapsed_sec: float) -> None: ...
    def on_file_error(self, file_path: str, index: int, error: Exception) -> None: ...
```

**В этой волне:** `callback=None` → текущее поведение (print to stdout).
Callback-параметр добавить в сигнатуру `run_v2_extraction` сразу,
чтобы не менять её при parallel. Реализация callback = `None`-guard + print.

`run_v2_tags` и `run_v2_postprocess`: callback пока **не нужен**
(tags — быстрый шаг; postprocess — несколько коротких try/except). Добавить при необходимости.

---

## Что НЕ вводить в этой волне

| Что | Когда |
|-----|-------|
| `V2FileExtractResult`, `V2ExtractionBatch` DTOs | При реализации `parallel.py` |
| `tag_parser.analyze(source_dict=...)` параметр | Перед PPE тегов |
| `v2_timing.py` / per-file timing | При реализации `parallel.py` |
| Сам `parallel.py` | Фаза 4B `AI_v2_parallel.mdc` |

---

## Инструкция для C2 (шаг 1D)

1. В `v2_pipeline.py` создать три функции: `run_v2_extraction`, `run_v2_tags`, `run_v2_postprocess` по сигнатурам выше.
2. Перенести тело текущих шагов 3–4 → `run_v2_extraction`, шаг 5 → `run_v2_tags`, шаги 6–10 → `run_v2_postprocess`.
3. `run_v2_pipeline` = setup + compose трёх фаз (обратная совместимость, та же сигнатура и возвращаемое значение).
4. `_process_single_file` → `extract_all_pages_for_file`, сделать публичной (убрать `_`). На этапе 1B переместить в engine.
5. `ExtractionCallback` Protocol — в `v2_pipeline.py` или отдельный `v2_callbacks.py`; `callback=None` в сигнатуре `run_v2_extraction`; при `None` — print (текущее поведение).
6. **Verification:** `run_v2_pipeline` даёт тот же набор выходных файлов. `run_v2_extraction` вызывается отдельно без ошибок.
