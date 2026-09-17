# pdf_parsing_v2 — интерфейсы модулей

Краткая карта входных точек для скриптов и тестов. Актуальные имена реэкспорта см. `pdf_parsing_v2/__init__.py`.

## Pipeline

- **`v2_pipeline.run_v2_pipeline`** — основной batch-запуск по конфигу (пути, шаблоны, отчёты). См. параметры в модуле и `pdf_v2_config.json`. В итоговом **`summary`** на верхнем уровне: **`normcontrol`**, **`od_table`**, **`tag_analysis`** (анализ тегов МТО / `*_tag_analyze.txt` + структурированный payload — **`.cursor/rules/AI_v2_tag_analysis.mdc`**). Если тип PDF нет ни в одном шаблоне — **`uncovered_doc_types`** / **`uncovered_n_files`**; control center показывает предупреждение после **`load_templates`**.

## Извлечение одной страницы

- **`stamp_extractor.extract_page`** — извлечение полей по выбранному шаблону; `cfg` задаёт режим текста (`char_center` и т.д.) и пробрасывается в `adapt_by_cell_assignment` (в т.ч. `find_tables`); опционально `_pdf_path`, `_page_num` для заголовка alignment.
- **`stamp_extractor.select_templates`** — выбор подходящих шаблонов для страницы.

## Конфигурация

- **`v2_config.load_v2_config`**, **`v2_config.get_v2_config_path`** — загрузка и расположение `pdf_v2_config.json`.

## Шаблоны

- **`template_loader.load_all_templates`**, **`load_project_templates`**, **`load_projects`**, **`load_catalog_for_project`** — загрузка JSON и структур проектов/наборов; каталог первого совпадения по имени проекта (см. докстринг).

## Модели и результаты

- **`models`** — `FieldDef`, `StampTemplate`, `TemplateSet`, `V2PageResult`, `FrameInfo`, типы скоринга и каталога полей.
- **`FieldResult`** — `cleaned_value` и `raw_value` остаются строками (`str | None`). Диагностика парсинга: **`parse_warnings: list[ParseWarning]`** (`code`, `message`, опц. `tier`), **`clean_tier`** (`"strict"` / `"relaxed"` или `None`). Предупреждения не подмешиваются в `V2PageResult.warnings` (см. уровни ниже).

## Очистка текста полей (`stamp_text`, `field_cleaners`)

- Движок v2 **не** импортирует `utils.string_parsing`. Логика перенесена в пакет **`pdf_parsing_v2/stamp_text/`**:
  - **`primitives`** — нормализация строк (переносы, даты, NBSP/дефисы).
  - **`junk`** — отбрасывание строк по меткам (legacy `string_Junk_Cleaner`), режимы **`JunkMode`**.
  - **`stamp_field_cleaners`** — для каждого ключа `clean`: приватное тело `_* (text, ctx) -> str` и публичный **`pipeline_* -> CleanOutcome`** (эталон strict/relaxed — `pipeline_doc_title`).
  - **`pipelines.run_field_clean(key, text, ctx)`** → **`CleanOutcome`** (`value`, `warnings`, …). Реестр ключей = **`FIELD_CLEAN_PIPELINES`** (совпадает с шаблоном `clean`).
- **`field_cleaners.CLEANERS`** — тонкие обёртки `(text, doc_type, page_num) -> str` для обратной совместимости; источник истины — **`stamp_text.pipelines`**.
- **`field_cleaners.CLEANER_DESCRIPTIONS`** — краткие описания для tooltips в редакторе (панель свойств).

### Уровни диагностики (не смешивать)

| Уровень | Где | Назначение |
|--------|-----|------------|
| **Page** | `V2PageResult.warnings` | Шаблон, score, сетка / alignment |
| **Field** | `FieldResult.parse_warnings` | Парсинг текста, `UNKNOWN_CLEANER`, `CLEANER_EXCEPTION`, regex (через `is_valid`) |
| **UI** | Редактор | Неизвестный ключ `clean` — жёлтая подсветка combo + tooltip; не дублировать в Excel-колонку страницы |

- Неизвестный **`clean`** в шаблоне: **`extract_page`** задаёт `cleaned_value = raw_value` и **`ParseWarning(code=UNKNOWN_CLEANER, …)`**.

### Граница `compat` (v2 → v1 rules_check)

- **`compat.to_page_stamp_attributes`**: обычные ключи штампа — из **`FieldResult`** (`cleaned_value` / `raw_value`). Ключи **`61_Page_Layers`**, **`62_Page_Bookmarks`**, **`63_Page_Width`**, **`64_Page_Height`**, **`65_Page_Real_Format`**, **`file_name`** берутся **только** из **`V2PageResult.metadata`** (источник истины для `rules_check`); для **61–64** элементы списков остаются **`int`**, не строки. **`parse_warnings`** в v1-форму страницы **не** пробрасываются.
- **Геометрия сетки штампа (после ``adapt_by_cell_assignment``)** — дополнительные ключи в **`V2PageResult.metadata`** (не пересекаются с 61–65 / ``file_name``): **`stamp_effective_bbox`** — ``tuple[float, float, float, float]`` в **displayed** / ``find_tables``-пространстве; **`stamp_proposed_bbox_applied`** — ``bool``; опционально **`stamp_search_bbox`** — ``tuple`` расширенной области поиска. Пустой словарь, если шаблон без ``grid_adapt`` или адаптация не дала ``CellAssignmentInfo``.

### Свойства PDF / служебные поля шаблона

- Пакет **`document_properties/`** — расчёт метаданных страницы/файла (слои, текстовые аннотации, размеры мм, формат листа, basename) без импорта v1.
- **`FieldDef.document_property`**: если задано (и совпадает с `id`), поле **не** читается по bbox; значение — из того же словаря, что и `metadata`; **зарезервированные id** без этого флага в шаблоне → ошибка конфигурации в **`extract_page`**.
- **`grid_matcher`**, **`frame_detector` (drawing_union ROI)**: такие поля исключены из геометрии штампа (как невидимые для union/Hungarian).
- Редактор: строки в списке полей **[PDF]**, на холсте не рисуются; при первом открытии шаблона недостающие обязательные свойства добавляются через **`ensure_document_property_field_defs`**. В **`cfg`** пайплайна / F5: **`source_pdf_basename`** для **`file_name`**.

## Сетка и адаптация

- **`grid_matcher`** — поиск ячеек, привязка полей, alignment (в связке с `find_tables_settings`, настройками шаблона).
- **Префильтр островов ячеек** (`pdf_parsing_v2_engine/grid_detected_prefilter.py`) вызывается из **`adapt_by_cell_assignment`** после отсечения oversized и до extent / Hungarian / **`_align_grid_lines_ordered`**. Управление через **`pdf_v2_config.json`** (все опциональны, см. дефолты в коде): **`grid_detected_prefilter_enabled`**, **`grid_detected_prefilter_min_cells`**, **`grid_detected_prefilter_y_pad_frac`**, **`grid_detected_prefilter_y_pad_floor_pt`**, **`grid_detected_prefilter_max_removal_frac`**, **`grid_detected_prefilter_min_component_cells`**, **`grid_detected_prefilter_min_component_area_frac`**, **`grid_detected_prefilter_adjacency_tol_floor_pt`**. Диагностика попадает в **`CellAssignmentInfo.detected_prefilter`** и JSON **`prealign.detected_prefilter`** бандла **`stamp_prealign_debug`**; слои панели «Границы» — `prefilter_*` в **`prealign_overlay_layers.py`** (PNG **`adapt_debug`** для этих слоёв не рендерится).

### Поля `outside_stamp`

- В **`adapt_by_cell_assignment`**: поля с **`FieldDef.outside_stamp=True`** не участвуют в Hungarian и получают **`status="excluded"`**. Их **`bbox`** — rect из **`field_to_fitz_rect`** (origin / stretch) **без** глобального transform штампа **`(transform_scale_*, transform_d*)`**. Перед этим вызывается **`snap_outside_stamp_vertical_to_frame`**: внутренняя горизонталь прижимается к линии рамки (`frame.y0` для origin снизу, `frame.y1` для сверху), вертикальная высота сохраняется; при одновременных stretch **top**+**bottom** snap не делается. Поля с **`document_property`** исключены из matching и получают нулевой placeholder-bbox (не используются при извлечении).
- В **редакторе** при открытом PDF тот же вертикальный snap применяется к **`bbox_mm`** в памяти при загрузке шаблона и при F5 «Тестировать» (сохранение в JSON только по явному Save).
- **`bound_*`** для таких полей в pipeline **не используются** (`expand_fields_by_bindings` их пропускает). Старые значения в JSON можно игнорировать.
- Если детектированных ячеек нет, глобальный трансформ не вычисляется — для внутренних полей остаётся ветка «только шаблон»; для **outside_stamp** по-прежнему rect из шаблона + snap по рамке.

### find_tables snap: конфиг и шаблон

- В `pdf_v2_config.json`: `find_tables_snap_x_tolerance`, `find_tables_snap_y_tolerance` — базовые значения для PyMuPDF `find_tables()` (см. `find_tables_settings.get_find_tables_kwargs`).
- В JSON шаблона (корень, опционально): те же ключи. Если поле задано (`float`), оно **перекрывает** значение из `cfg` при вызове `find_tables` внутри `get_detected_stamp_cells`, `adapt_by_cell_assignment` и legacy `adapt_all_fields*`.
- Слияние: **`find_tables_settings.merge_cfg_for_find_tables(cfg, template)`** (shallow copy `cfg` + не-`None` поля шаблона).
- Редактор: **`pdf_template_editor.auto_detect.detect_cells(..., template=...)`** — при переданном шаблоне те же переопределения, что и у адаптации (мета-панель «Свои snap для find_tables»).
- **`grid_lines_utils`** — линии каркаса шаблона (генерация, refs); используется движком и редактором.

## Трансформы координат

- **`coord_transform`** — переводы мм ↔ fitz, displayed ↔ unrotated (обязательны при rotation 90°/270°).

## Рамка чертежа

- **`frame_detector.find_frame(fitz_page, debug=False, *, frame_mode="gost", template=None) -> tuple[FrameInfo, dict | None]`** — детекция рамки по `get_drawings()`. По умолчанию **`frame_mode="gost"`** — прежняя эвристика (max span X/Y, порог 20% стороны страницы, иначе поля ГОСТ); поведение без именованных аргументов не менялось. **`frame_mode="drawing_union"`** (opt-in в JSON шаблона: **`frame_mode`**) — union AABB отфильтрованных rect, пересекающих ROI: пересечение печатного поля ГОСТ с раздутым bbox полей штампа (не `outside_stamp`, не `document_property`) относительно GOST-frame, ±25 мм; если после ROI не осталось rect → снова ГОСТ. Для union-режима в **`find_frame`** передаётся тот же **`StampTemplate`**, что и в извлечение (поля для ROI). При `debug=True` в диагностике есть **`frame_mode`**, для gost — прежние поля; для drawing_union — **`roi_pdfminer`**, **`roi_source`**, **`rect_count_after_roi`**, **`union_bbox_pdfminer`**, … В **`pdf_template_editor/adapt_debug.collect_adapt_debug_snapshot`** вызывается с `debug=True` и **`frame_mode`/`template` из шаблона**; в snapshot добавляются `frame_detection` (включая `stamp_search_bbox_pts`), `origin_points_pts` / `origin_points_mm`, `template_origin`.

## Редактор (отдельный пакет)

- Запуск: **`python -m pdf_template_editor`**. Импортирует движок из `pdf_parsing_v2.*`; не импортируйте редактор внутрь `v2_pipeline` без необходимости.

## Совместимость и отладка

- **`compat`**, **`debug_visual`**, **`v2_report`** — вспомогательные слои; см. импорты в pipeline.

## Отладочный Excel v2

- **`v2_report.save_v2_debug_report(all_results, result_dir, catalog)`** — третий аргумент **`FieldCatalog` обязателен** (порядок колонок `debug_report_order`, `include_in_debug_report`, подписи и `debug_report_header_note` из каталога). Поля вне каталога — в хвосте отчёта. Комментарий к ячейке поля: при `raw ≠ cleaned` — блок **RAW**; при непустом **`FieldResult.parse_warnings`** — блок **Field warnings** (колонка «Предупреждения» на страницу по-прежнему только `V2PageResult.warnings`).
- **`v2_pipeline.run_v2_pipeline`**: при **`export_debug_excel: true`** в конфиге нужны непустой **`project`** (или аргумент `project`) и успешный **`load_catalog_for_project`**; иначе выбрасывается **`ValueError`** (без legacy-порядка колонок). При `export_debug_excel: false` каталог для отчёта не загружается.
