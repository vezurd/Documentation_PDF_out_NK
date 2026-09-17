---
description: Память проекта — структура, ключевые классы, колонки, известные ловушки. Читай перед любой работой с кодом.
alwaysApply: false
---

> **Архивный снимок** (31.03.2026). Актуальная компактная версия: `AI_memory.mdc`. Индекс архивов: `AI_pdf_v2_template_editor_roadmap_INDEX.md`.

# AI Memory — Documentation_PDF_out_NK

> **Правило для чатов**: по итогу работы, если обнаружены новые важные факты о проекте,
> добавляй их в соответствующий раздел этого файла. При необходимости переструктурируй
> и консолидируй — файл должен оставаться компактным и актуальным.

---

## 1. Назначение проекта

Инструмент сравнения тегов оборудования из RFP (запрос предложений) с актуальными МТО
(материально-техническое обеспечение) и вендорской документацией VO (Variation Orders).

Точка входа: `main.py` → вызывает `RFQ/tags_rfp_compare/agregate_tags.py::main()`.

---

## 2. Структура модулей (tags_rfp_compare)

```
agregate_tags.py          — главный оркестратор, вызывает step1..step4
step1_load_rfp.py         — загрузка RFP (Excel), split строк по тегам/VALUES
step2_load_mto.py         — загрузка MTO (Excel, кэш)
step3_load_vo.py          — загрузка VO (Excel, кэш)
step4_analyze_and_match.py — основной анализ, вызывает sub-steps:
  step4/step4_1_check_mto_data.py  — валидация, split MTO/VO строк по тегам
  step4/step4_2_match_rfp_with_mto.py — сопоставление RFP↔MTO и RFP↔VO
  step4/step4_3_assign_position_status.py — статусы позиций
  step4/step4_4_add_unmatched_mto_rows.py — добавление несопоставленных MTO/VO;
   no-tag VO→MTO: `_expand_no_tag_candidate_codes` + `_ref_cabinet_forbids_empty_mto_cabinet` в step4_2;
   `add_unmatched_vo_rows` / `match_added_vo_rows_with_mto_by_code` принимают `replacement_table` (проброс из step4_worker)
  step4/step4_5_check_values_sum.py — проверка сумм VALUES
  step4/step4_6_save_match_result_to_excel.py — сохранение результата (xlsxwriter)
  step4/step4_6_cell_colors.py — раскраска ячеек
  step4/step4_postmerge_ops.py — per-row post-merge ops (вызывается из worker)

rfp_tags_utils.py         — утилиты (пути, таймлоги, memory-логи)
rfp_tags_settings_gui.py  — GUI настроек
column_optimization.py    — оптимизация колонок (удаление ненужных)
compare_step4_files.py    — корректное сравнение результатов (multiset строк, без зависимости от порядка)
mto_file_filter.py        — фильтрация MTO файлов
```

---

## 3. Ключевые классы (base/)

### RowStd (`base/base_classes.py`)
Стандартная строка таблицы. Содержит:
- `el: Dict[str, CheckElement]` — значения по колонкам
- `row_type: RowType` — тип строки
- `t_com: TableComments` — метаданные таблицы/файла

**Важные методы:**
- `get_value(col)` — получение значения (для VALUES* автоконвертация в float)
- `get_tags_list()` → `list[str]` — теги из TAGS
- `get_row_copy(in_row)` — полная глубокая копия
- `get_row_copy_light(in_row, ...)` — ОБЛЕГЧЁННАЯ копия (см. раздел 5)

### CheckElement (`base/base_classes.py`)
- `value` — основное значение
- `struck_value` — зачеркнутый текст (из MTO)
- `color` — цвет для Excel

### RowType
`section_row`, `cabinet_title_row`, `position_row`, `empty_row`, `other_row`, `system_row`, `head_row`

---

## 4. Колонки (base/tables_columns.py)

### Базовые RFP колонки
| Константа | Описание |
|-----------|----------|
| `CODE` | Код оборудования |
| `TAGS` | Список тегов (list[str]) |
| `VALUES` | Количество |
| `NAME` | Наименование |
| `TYPE_MARK` | Тип/марка |
| `DS_TITLE` | Title system (ключ группировки RFP↔MTO↔VO) |
| `DS_NUMBER` | Номер ДС |
| `DS_NAME` | Наименование ДС |
| `DS_LOT` | Лот из реестра |
| `IN_CABINET` | Шкаф (может быть "A/B" для нескольких) |

### Колонки MTO (заполняются при match)
`TAG_MTO`, `CODE_MTO`, `MTO_CODE_STRUCK`, `NAME_MTO`, `TYPE_MARK_MTO`, `NUMBERS_MTO`, `VALUES_MTO`

### Колонки VO (заполняются при match)
`TAG_VO`, `CODE_VO`, `NAME_VO`, `VALUES_VO`

### Статусные колонки (заполняются при/после match)
| Константа | Когда заполняется |
|-----------|-------------------|
| `MATCH_STATUS` | match_rfp_with_mto |
| `MATCH_STATUS_VO` | match_rfp_with_vo |
| `HAS_TAGS_RFP` | match_rfp_with_mto |
| `POSITION_STATUS` | assign_position_status / match |
| `MATCH_STATUS_CODE_REPLACEMENT` | match (таблица замен) |
| `TAG_EFFECTIVE` | post-match: _set_effective_tags_before_export |
| `MATCH_STATUS_TAGS` | post-match: _set_tags_match_summary |
| `MATCH_STATUS_VO_MTO_TAGS` | post-match: _set_tags_pair_statuses |
| `MATCH_STATUS_RFP_MTO_TAGS` | post-match: _set_tags_pair_statuses |
| `MATCH_STATUS_RFP_MTO_STRUCK_CODE` | post-match |
| `EQUIPMENT_TYPE_STATUS` | post-match: _set_equipment_type_status |
| `MTO_CABINET_EQUIPMENT` | post-match |

---

## 5. КРИТИЧНО: get_row_copy_light и _COPY_COLS

`get_row_copy_light` создаёт облегчённые копии при split строк по тегам (step1, step4_1).
Копии **разделяют** один `CheckElement` для колонок НЕ в `_COPY_COLS`.
Запись в shared CheckElement перезаписывает значение для ВСЕХ копий из одного split.

### _COPY_COLS — колонки с СОБСТВЕННЫМ CheckElement в каждой копии

Все столбцы, в которые происходит запись при match или post-match, ОБЯЗАНЫ быть в `_COPY_COLS`:

```
# Записываются при match (add_mto_data_to_rfp_row / add_vo_data_to_rfp_row):
VALUES, VALUES_2, VALUES_MTO, VALUES_VO,
TAG_MTO, CODE_MTO, MTO_CODE_STRUCK, NAME_MTO, TYPE_MARK_MTO, NUMBERS_MTO,
TAG_VO, CODE_VO, NAME_VO,
MATCH_STATUS, MATCH_STATUS_VO, POSITION_STATUS, IN_CABINET, HAS_TAGS_RFP,

# Записываются при post-match (step4_analyze_and_match.py):
TAG_EFFECTIVE, MATCH_STATUS_TAGS, MATCH_STATUS_CODE_REPLACEMENT,
MATCH_STATUS_VO_MTO_TAGS, MATCH_STATUS_RFP_MTO_TAGS,
MATCH_STATUS_RFP_MTO_STRUCK_CODE, EQUIPMENT_TYPE_STATUS,
MTO_CABINET_EQUIPMENT, DS_LOT,
```

### Безопасные для sharing (НЕ записываются после split)

Всё остальное: CODE, NAME, TYPE_MARK, DS_TITLE, DS_NUMBER, DS_NAME, ANNOTATION,
NUMBERS, UNITS и пр. — только читаются, можно разделять.

### Баг от 25.02.2026 (исправлен)

Последние 9 колонок (TAG_EFFECTIVE и др.) отсутствовали в `_COPY_COLS`.
Все light-копии из одного split разделяли один CheckElement для TAG_EFFECTIVE.
При post-match обработке последняя запись перезаписывала значение для ВСЕХ копий →
массовые ложные дубликаты в Шаг4_TAG_EFFECTIVE_дубликаты.xlsx.
Особенно критично для VO (теги часто отличаются от RFP/MTO).

**Правило: при добавлении новой записываемой колонки — ВСЕГДА добавлять в _COPY_COLS.**

---

## 6. Поток данных step4

```
rfp_data (List[RowStd]) ──────────────────────────────────┐
mto_data (Dict[title_system, List[RowStd]]) ──┐           │
vo_data  (Dict[title_system, List[RowStd]]) ─┐│           │
                                              ││           │
step4_1: check_mto_data / check_vo_data ◄─────┘│  (split по тегам)
step4_2: match_rfp_with_mto ◄─────────────────────────────┤
         match_rfp_with_vo  ◄─────────────────────────────┤
step4_3: assign_position_status ◄─────────────────────────┤
step4_4: add_unmatched_mto_rows ◄─────────────────────────┤
         add_unmatched_vo_rows  ◄─────────────────────────┤
         match_added_vo_rows_with_mto_by_code ◄───────────┤
         match_rows_with_vo_and_rfp_code_missing_mto ◄────┤
         match_unmatched_vo_with_mto_by_reverse_replacement◄┤
         add_unmatched_mto_cabinet_rows ◄─────────────────┤
                                                          │
post-match (в step4_analyze_and_match.py):                │
  _set_equipment_type_status ◄────────────────────────────┤
  _set_tags_match_summary ◄───────────────────────────────┤
  _set_tags_pair_statuses ◄───────────────────────────────┤
  _collapse_rfp_rows_before_export ◄──────────────────────┤
  _apply_lot_registry_to_rows ◄───────────────────────────┤
  _set_effective_tags_before_export ◄─────────────────────┤
  _set_mto_cabinet_equipment_before_export ◄──────────────┤
  assign_row_colors ◄─────────────────────────────────────┤
  save_match_result_to_excel ◄────────────────────────────┘
```

---

## 7. TAG_EFFECTIVE — логика определения

Приоритет: VO тег → MTO тег → RFP тег (первый из списка).
Если есть RFP теги, но нет CODE_MTO и CODE_VO → строка получает POSITION_STATUS="Исключен".

---

## 8. Схлопывание (_collapse_rfp_rows_before_export)

Строки без тегов с одинаковыми кодами группируются и схлопываются (суммируются VALUES).
Ключ группировки: `(DS_TITLE, DS_NUMBER, CODE, CODE_MTO, CODE_VO, IN_CABINET)`.
Строки из разных шкафов (IN_CABINET) НЕ схлопываются.

---

## 9. Таблица замен (replacement_table)

Словарь `{old_code: [(new_code, description), ...]}`.
Используется в обе стороны: прямой поиск (RFP→MTO) и обратный (VO→MTO).
Загружается из файла через `load_replacement_table()`.

---

## 10. Кэширование

`utils/cache_utils.py` — кэширование загруженных Excel данных (MTO, VO).
Кэш инвалидируется при изменении файлов.
`get_row_copy_light` используется и в `cache_utils.py` — учитывать при изменениях.

### Авто-синхронизация Google-таблиц (12.03.2026)

Флаг `flag_get_google_data` полностью удалён из проекта. Вместо ручного управления —
автоматическая проверка `modifiedTime` через Drive API (`gc.drive.get_update_time(id)`).

**`base/base_google.py`** — `load_base()` (без аргументов):
- Кэш: `base_check/code_base_data.json` (данные) + `base_check/code_base_meta.json` (modifiedTime).
- При запуске: лёгкий API-вызов (~0.5с) проверяет `modifiedTime` таблицы `1P_9LcZ...`.
- Если таблица не менялась — читает из кэша. Если менялась — полная загрузка + обновление кэша.
- При ошибке сети — fallback на существующий кэш.

**`cable_mapping/mapping/load_map_from_google.py`** — та же схема для mapping-таблицы `1VyHiGjc...`:
- Кэш: `mapping_base_data.json`, `no_out_cables.json`, `json_open_laying.json` + `mapping_base_meta.json`.
- `_auto_sync_map()` проверяет modifiedTime один раз, при изменении обновляет все 3 листа.
- `load_map_google_base()`, `load_no_out_cables_google_base()`, `open_laying_google_base()` —
  без аргумента flag, авто-sync при первом вызове.

Удалённые артефакты: `google_sheets.flag_get_google_data`, чекбокс «Синхронизировать базу с ГуглТаблицей?»
в `main.py`, все импорты `from base.google_sheets import flag_get_google_data`.

**Обогащение RFP строк из Google-базы (16.03.2026)**

`step4_postmerge_ops.py::enrich_rfp_from_google_base()` — подставляет `NAME`, `TYPE_MARK`, `VENDOR`
из Google-базы по коду `CODE` в каждую `position_row` результата. Вызывается внутри
`apply_all_postmerge_per_row_ops` после `assign_row_colors` (per-TM, в worker-процессах).

Архитектура передачи:
- `agregate_tags.py`: вызывает `load_base()` параллельно со step2+step3 в `ThreadPoolExecutor(max_workers=3)`;
  строит `code_base_by_code = _build_code_base_by_code(rows)` → `Dict[str, tuple(name, type_mark, vendor)]`;
  передаёт в `step4_analyze_and_match(code_base_by_code=...)`.
- `step4_analyze_and_match.py`: принимает `code_base_by_code: Dict[str, tuple] | None = None`;
  **НЕ вызывает** `load_base()` внутри; передаёт dict в `tm_common_kwargs` → workers.
- Pickle-оптимизация: dict хранит только 3 скалярных значения на запись (не полные RowStd),
  что минимизирует pickle-объём при передаче через ProcessPoolExecutor initializer.
- `load_base_wall_clock` — метрика в `timing_log.xlsx` для диагностики времени загрузки Google-базы.

**Правило:** `_build_code_base_by_code` — module-level функция в `step4_analyze_and_match.py`,
импортируется в `agregate_tags.py`. При переносе модуля обновить импорт в обоих местах.

---

## 11. Конфигурация

JSON файл загружается через `rfp_tags_utils.load_config()`.
Секции: `paths`, `step1`, `step2`, `step3`, `step4`, `column_optimization`, `rfp_tags_utils`.
GUI: `rfp_tags_settings_gui.py`.
- `rfp_tags_utils.save_input_fingerprints` — флаг записи `input_fingerprints.json` (MTO/VO).
  При `False` шаги `step2_load_mto_data` и `step3_load_vo_data` не вызывают `save_input_fingerprint()`.
- `step2.export_load_results_excel` — флаг выгрузки `Шаг2_MTO_результаты_загрузки_*.xlsx`.
- `step2.export_positions_database_excel` — флаг выгрузки единой базы всех загруженных MTO `position_row`:
  файл `Шаг2_MTO_база_position_row_*.xlsx`, колонки `TITLE`, `SYSTEM` + MTO-колонки (`ColNames.MTO.column_dict`)
  с человеко-читаемыми заголовками для фильтрации и анализа.
  Перед записью выполняется отдельное схлопывание только для экспортной выборки
  по ключу `TITLE + SYSTEM + (все MTO поля кроме VALUES)`, с суммированием `VALUES`.
  Важно: схлопывание не должно менять `mto_data`, передаваемые дальше в pipeline.
  При `step2.export_positions_database_excel=True` загрузка `step2+step3` выполняется последовательно
  (без thread-buffered stdout), чтобы видеть live-прогресс подготовки/схлопывания/записи файла.
  Для этого файла отключена построчная стилизация ячеек (оставлены ширины колонок по config, фильтр и freeze panes),
  чтобы снизить время записи.
  Перед экспортным схлопыванием применяется раскрытие MTO-строк по тегам по логике Step4:
  для `position_row` с `len(TAGS)>1` (кроме `UNITS=="м"`) создаются отдельные строки по одному тегу,
  `VALUES` распределяется по 1 на тег (при дефиците — 0), затем выполняется схлопывание.
- `step4.export_title_systems_comparison_excel` — флаг выгрузки `Шаг4_RFP_MTO_VO_сравнение_title_system_*.xlsx`.
- Чтение Excel (`base/base_xlsx_load.py`) всегда использует `data_only=True` —
 формулы читаются как кэшированные значения (результат последнего пересчёта в Excel).
- Для объединённой кнопки MTO+BBB добавлена секция `folder_rules` в `base/bbb_config.py`:
 - `folder_rules.search_only_in_dwg` — ограничение запуска по имени выбранной папки (`DWG`).
 По умолчанию `True`; настраивается в `base/bbb_settings_gui.py` чекбоксом
 «Искать только в папке DWG».
- Для BBB-проверок по GoogleBase добавлены поля Google-таблицы:
 - `ColNames.GoogleBase.column_dict[26] -> BBB_WORK_CODE_LIST`
 - `ColNames.GoogleBase.column_dict[27] -> BBB_MTR_GROUP`
  `BBB_WORK_CODE_LIST` парсится через `tags/tag_parser.py::get_bbb_work_code`.
- В `base/bbb_config.py` добавлен флаг `bbb_vs_code_base.check_work_code_and_mtr_group` (default `True`):
 включает проверку `BBB_WORK_CODE` и `BBB_MTR_GROUP` в `bbb_checks.py::bbb_vs_code_base()`
 для BOE/BOM position-строк по 3-статусной схеме (green/red/yellow + комментарии).
- В `base/bbb_config.py` добавлен флаг `bbb_vs_mto.compare_title_marka_revision` (default `True`):
  включает проверку `TITLE` / `BBB_MARKA` / `BBB_REVISION` в `bbb_checks.py::bbb_vs_mto_combined()`.
  Эталон берётся из MTO имени файла через `pdf_parsing.shtamp_extract_classes.doc_ATTRIBUTES`
  (`doc_Title_4d`, `doc_Marka`, `doc_Revision`); mismatch красится в `Color.red` с комментарием
  «как указано в MTO», match — `Color.green`.
- В `base/bbb_config.py` добавлен флаг `bbb_vs_mto.compare_bom_annotation_with_mto` (default `True`):
  включает проверку `ANNOTATION` для BOM в `bbb_checks.py::bbb_vs_mto_combined()`.
  Логика: сначала из MTO по `CODE` вычитается `BOE_total` (sum-only), затем остаточные MTO-слагаемые
  распределяются по всем строкам BOM этого `CODE` (учёт всех вхождений кода в BOM).
  Для каждой BOM-строки `ANNOTATION` сверяется с назначенным набором слагаемых как мультимножество
  целых чисел (порядок не важен). mismatch красится в `Color.red` с комментарием (с позициями),
  а в ячейку дописывается `Сумма из МТО` без позиций. Для single-term строки:
  пустой `ANNOTATION` допустим, корректное единственное число в `ANNOTATION` — `Color.yellow` с
  комментарием «Не требуется, позиция встречается единично.».
  Важный кейс: один residual MTO-терм может распределяться на несколько BOM-строк того же `CODE`
  (например `30` в MTO и 5 строк BOM по `6`), поэтому распределение должно поддерживать частичное
  потребление терма, а не только exact-subset из целых MTO-слагаемых.
- В `base/bbb_config.py` добавлен флаг `bbb_vs_mto.check_missing_mto_codes` (default `True`):
  включает проверку покрытия кодов MTO в `bbb_checks.py::bbb_vs_mto_combined()`.
  Для кодов из MTO, отсутствующих ни в BOE ни в BOM, добавляются diff-строки цвета `Color.match_added_mto`.
  Целевой документ (BOE/BOM) определяется по `SECTION_TYPE` первой MTO-строки через конфиг секций.
  `get_section_target(section_types_cfg, name)` → "BOE" если секция не найдена.
  **section_types в JSON НЕ хранятся** — только в Excel на сетевом диске (см. ниже).
- `SECTION_TYPE = "section_type"` добавлена в `base/tables_columns.py` (после `IN_CABINET`) и в `ColNames.column_list`.
- `base/section_type_info.py` — `add_section_type_info(table_std, source_file)`:
  заполняет `SECTION_TYPE` для каждой `position_row` по ближайшему `section_row`.
  Вызывается из `get_std_from_excel_file` после `add_cabinet_composition_info`.
  Новые секции авто-добавляются в Excel через `bbb_section_types_excel.add_new_sections_if_missing()`.
- `base/google_sheets.py`: `column_dict[101] = SECTION_TYPE`.
- `base/bbb_output_config.py`: `SECTION_TYPE` добавлен в BOE и BOM конфиги («Секция MTO», width=35).
- `base/bbb_section_types_excel.py` — новый модуль, хранение section_types в общем Excel на сети:
  Файл: `\\bcc\eng\PrDoc\377_НИПИГАЗ\АГХК\КСБ\РД\Таблички_графики\08_НК\11_BBB\BBB_типы_секций_MTO.xlsx`
  `load_section_types()` — read_only, не блокирует; fallback → 13 базовых секций.
  `save_section_types(dict)` — `PermissionError` (файл открыт другим) → warning, non-critical.
  `add_new_sections_if_missing(dict)` — атомарное добавление новых секций.
  Визуально: verified строки зелёные (сверху алфавит), новые — жёлтые (снизу алфавит).
- `base/bbb_settings_gui.py`: кнопка «Открыть типы секций MTO» рядом с заголовком секции bbb_vs_mto;
  чекбокс `check_missing_mto_codes` в той же секции GUI.
- `base/base_mto.py::get_mto_std_from_file()` должен корректно обрабатывать отсутствие MTO-файла в папке:
  если среди `.xlsx` нет `doc_Type == "MTO"`, функция печатает понятное сообщение и возвращает `None`
  (без передачи `file_full_path=-1` в `openpyxl.load_workbook`, иначе падает с
  `'int' object has no attribute 'strip'`).
  Важно: в режиме проверки папки (`dir_path != "-1"`) поиск MTO должен выполняться ВСЕГДА по папке,
  независимо от `file_full_path`; привязка поиска к `file_full_path` приводит к ложному
  «MTO путь не задан», даже когда файл MTO в папке есть.

---

## 12. Производительность и multiprocessing

Подробности в **`.cursor/rules/AI_performance.mdc`** (подключается для файлов step4/).

Ключевые факты:
- `step4.use_multiprocessing` — флаг в конфиге, включает ProcessPoolExecutor для `match_added_vo_rows_with_mto_by_code`
- `RowStd.batch_copy_light()` — batch-создание копий с shared empty CE для пустых _COPY_COLS
- `pass_match_code` — pointer per key (`code_ptr` dict) вместо None-marking, O(N) вместо O(N^2)
- **_COPY_COLS дублируется** в `get_row_copy_light` и `batch_copy_light` — обновлять оба при изменениях
- **batch_copy_light shared CE безопасность**: CheckElement'ы MTO/VO копий read-only в pipeline
  (matching пишет в RFP rows, `create_rfp_row_from_mto` создаёт новые объекты)
- Для этапа B добавлен `step4/step4_worker.py::process_title_mark()`:
  per-TM pipeline (check_mto/check_vo/match/assign/add_unmatched + local collapse).
- `step4_analyze_and_match.py` теперь запускает последовательный цикл по `title_mark`,
  мерджит TM-результаты, логирует память внутри цикла (`step4_tm_loop_after_N`) и вызывает `gc.collect()`.
- Для этапа C добавлены флаги `step4.parallel_by_title_mark` и `step4.max_workers`:
  при `parallel_by_title_mark=True` `step4_analyze_and_match.py` запускает
  `ProcessPoolExecutor + as_completed` для `process_title_mark()` с fallback на sequential.
- Глобально после merge остаются post-merge шаги (`_set_effective_tags_before_export`,
  duplicate detection TAG_EFFECTIVE, lot registry, colors, Excel save, check_values_sum).
- `step1_load_rfp.py` поддерживает `skip_split`; при `skip_split=True` split RFP пропускается на step1
  и выполняется внутри worker через `split_rfp_rows()` только для текущего TM.
- Кэш step1 учитывает режим split через `extra_key` (`skip_split=0/1`), чтобы не смешивать кэш
  предраскрытых и нераскрытых RFP данных.
- Для сравнения baseline/new используется **только** `compare_step4_files.py` в multiset-логике:
  сравниваются мультимножества всех строк по общим колонкам (`only_new`, `only_old`), а не
  ключи вида (`Имя ДС`, `№ позиции`, `Теги RFP`), т.к. такие ключи неуникальны и дают ложные diff.
- Добавлен контроль идентичности входов для clean control pair:
  `input_fingerprints.json` в `result_dir`, формируется в step2/step3:
  - `MTO`: fingerprint по списку MTO файлов
  - `VO`: fingerprint по списку VO файлов
  Поля: relative path, file size, mtime, сводный md5 по набору файлов.
- Для восстановления порядка строк после per-TM merge при `split_rfp_in_worker=True`:
  - в `step1_load_rfp.py::split_rfp_rows()` split-копии наследуют `_orig_idx` и получают `_orig_split_seq`
  - в `step4_analyze_and_match.py::_restore_original_row_order()` сортировка должна учитывать
    `( _orig_idx, _orig_split_seq )`, иначе split-копии теряют позицию и уезжают в хвост результата.
- Валидация 26.02.2026 показала: даже при сортировке `( _orig_idx, _orig_split_seq )`
  baseline-order ещё не восстановлен (`order_mismatch_rows=26706/27195` против run `10.19`),
  при этом multiset строк идентичен (`only_new=0`, `only_old=0`).
- `input_fingerprints.json` использует ключ `fingerprint_md5` (не `set_md5`) для MTO/VO наборов.
- При `parallel_by_title_mark=True` диагностические отчёты из `check_mto_data`/`check_vo_data`
  (Шаг4_Дублирующиеся_теги_MTO, Шаг4_Несовпадения_тегов_VALUES, Шаг4_Дублирующиеся_теги_VO)
  НЕ пишутся в worker'ах. Вместо этого worker возвращает `mto_diagnostics`/`vo_diagnostics` через dict,
  в main-процессе данные агрегируются и сохраняются ЕДИНЫМ файлом через
  `save_mto_diagnostics_reports()`/`save_vo_diagnostics_reports()` после merge всех TM.
  Управляется параметром `defer_reports=True` (ставится автоматически при `use_parallel_tm`).
- Worker НЕ возвращает `mto_data_tm`/`vo_data_tm` (post-split RowStd) — pickle overhead.
  Вместо этого worker вычисляет `mto_summary`/`vo_summary` (dict/set, ~1-2 KB per TM):
  - `mto_summary`: `tagged_sum` (float), `tagged_values_by_code` (Dict[str,float]),
    `codes_with_tags` (Set[str])
  - `vo_summary`: `tagged_sum` (float), `codes` (Set[str])
  Main-процесс агрегирует в `agg_mto_tagged_sums`, `agg_mto_tagged_values_by_code`,
  `agg_vo_tagged_sums`, `agg_vo_codes`, `agg_mto_codes_with_tags`.
  `check_values_sum()` и `_save_code_ban_list()` работают с pre-computed summaries.
- `step1.skip_split` доступен в конфиге и GUI (checkbox в секции Step1).
  При `True` split RFP выполняется внутри worker per-TM (вместо step1).
- Добавлен флаг `step4.verbose_progress_messages` (config + GUI):
  включает/выключает служебные консольные сообщения из `step4_3_assign_position_status.py`
  и `step4_4_add_unmatched_mto_rows.py` (например «Определение статуса позиции...», «Вставка N несопоставленных строк...»).
- `load_base()` вынесен из `step4_analyze_and_match` в `agregate_tags.py` и запускается
  параллельно со step2+step3 (`ThreadPoolExecutor(3)`), чтобы скрыть сетевую задержку (~8-15с).
  `code_base_by_code: Dict[str, tuple]` передаётся параметром в `step4_analyze_and_match`.
  Pickle для worker initializer: 3k × 3 скалярных значения (не RowStd). Измерение: `load_base_wall_clock` в timing_log.

---

## 13. Экспорт проекта для коллег (ZIP)

- Добавлен скрипт `utils/release_zip.py` для сборки релизного ZIP архива проекта
  (без локальных кэшей/временных артефактов) в папку `_release/`.
- Скрипт читает исключения из `.release_zipignore` в корне проекта
  (помимо встроенных default-паттернов), что позволяет поддерживать состав архива без правок кода.
- Для быстрого запуска добавлены:
  - GUI-кнопка в `main.py`: «Собрать ZIP для коллег»
  - bat-скрипт `build_release_zip.bat`

---

## 14. Утилиты: известные ловушки

### `base/base_xlsx_load.py::load_from_xlsx_file` + `base_excel_out.check_color_out` — блокировка xlsx на Windows (20.03.2026)

`load_from_xlsx_file`: `try`/`finally` с обязательным `wb.close()` (в т.ч. при `KeyError` листа или сбое в `iter_rows`);
в `finally` после закрытия — `del wb` и `gc.collect()` (как у `remove_strikethrough_from_xlsx`), чтобы отпускать handle ZIP.

`check_color_out`: после `wb.save` вызывается `wb.close()` в `finally`, чтобы не держать шаблон открытым в длительно живущем GUI.

### `utils/path.py::get_files_single` — регистр расширений (13.03.2026)

Сравнение расширений файлов сделано **регистронезависимым** (`next_file_lower.endswith(e.lower())`).
До исправления файлы с расширением `.XLSX` (верхний регистр) не находились при поиске с `[".xlsx"]`.
Пример проблемного имени: `AGCC.287-3150-KSB.BOE-0001_01-AN01_RU.XLSX`.

**Правило: при добавлении новых вызовов `get_files_single` передавать расширения в нижнем регистре —
функция сама приведёт к нижнему регистру при сравнении.**

### `base/bbb_load.py::find_bbb_files` — диагностика нераспознанных файлов

Добавлен вывод нераспознанных `.xlsx` файлов (когда `doc_Type` не в `BOE/BOM/BOQ/MTO`),
чтобы сразу было видно несоответствие имён ожидаемому шаблону.

### BBB Analysis: схема классификации файлов по имени

`doc_ATTRIBUTES.doc_Type` определяется через `string_parsing.getDocTypeFromFile(file_name)`.
Ожидаемый шаблон имени: `AGCC.287-TTTT-MARKA.TYPE-NNNN_...`
(два сегмента до типа разделены точкой, тип — до третьего дефиса).
Если имя файла не соответствует — `doc_Type` будет пустым или некорректным,
файл не попадёт в BBB-анализ.

### Экспорт BOE/BOM/BOQ «Для 1C» (`main.py`)

Кнопка «Для 1C» → `open_prepare_bbb_for_1c`: для каждого файла
`backup_xlsx_to_old_subfolder`, затем `remove_strikethrough_from_xlsx` с
`create_backup=False`, всегда **`base/excel_recalc_xlwings.py::recalculate_workbook_save`**
(Excel COM, `CalculateFullRebuild`, два прохода: до и после замены формул; нужны Windows,
Excel, xlwings), затем
`replace_formulas_with_cached_values(..., sheet_name=doc_type)` — только лист BOE/BOM/BOQ,
скрытые справочные листы не обходятся при замене формул. Затем
`remove_hidden_sheets_for_1c` удаляет **все** листы со статусом hidden/veryHidden
(белого списка нет). Без установленного `xlwings` и Excel кнопка «Для 1C» не завершит
пересчёт. Значения формул при replace — из кэша в файле (`data_only`).

### PDF pipeline — subprocess + config + параллелизация (26.03.2026, roadmap Запуски 1–4)

- Кнопка «Открыть папку с PDF»: `main.py` — `wire_pdf_button`, `_pdf_subprocess_work` (`subprocess.run` + `capture_output`), `_pdf_folder_dialog` (имя папки **PDF**, без учёта регистра); таймер «Проверка PDF» через `gui_subprocess_job`.
- Отдельный процесс: `pdf_parsing/pdf_pipeline_main.py` — `run_pdf_pipeline(pdf_path, cfg)`;
  **последняя непустая строка stdout** = `result_dir` (GUI открывает папку при успехе).
- Настройки: `pdf_parsing/pdf_config.py` + JSON `pdf_pipeline_config.json` в корне проекта;
  окно `pdf_parsing/pdf_settings_gui.py` (⚙ рядом с кнопкой PDF).
  Поля `use_parallel`, `max_workers` (0=авто), `timing_log`, `find_functions_pdf_folder`
  (папка для отладочного `find_functions.py`; кнопка запуска в том же окне, путь передаётся в subprocess и сохраняется при «Сохранить»).
- Штамп стр.1: `normalize_19_13` возвращает `(grid, {"remap_15_to_13_cols", "stamp_second_remove_idx"})` при 15→13:
  после `remove_column(4)` если в строках 13–17 столбец **7** непустой (стадия «Р» в бывшем col8), второе удаление — **8**,
  иначе **7**; `map_stamp_col_15_to_13(col, second_remove_idx)`; для `7_Total_number_of_sheets` и `26_Document_Revision`
  при пустом чтении — fallback на исходный столбец **14** после того же маппинга (SOT/2879).
  `merge_row15_fragmented_quantity` — склейка «на»/число/«л.» для `6_2`.
- Синхронный путь сохранён: `open_pdf_folder`, `check_tags_button` (отладка) → `main.start()`.
- Pickle-диагностика B.4: `pdf_parsing/pdf_pickle_diagnostic.py`.

**Параллелизация PDF (Запуск 2, 26.03.2026)**

- `tag_parser.add_context(text, doc_name, page_num, target_dict=None)` — при `target_dict=None`
  пишет в глобальный `source_dict`; при передаче dict — в него (B.2).
- `tag_parser.merge_dicts(base_dict, local_dict)` — мердж локальных tag dicts.
- `find_tags.parse_tags(curr_proj, tag_dict=None)` — поддерживает запись в переданный dict.
- `pdf_parsing/pdf_worker.py` — per-file worker (B.1):
  - `process_single_pdf(file_full_path) -> dict` — pdfplumber + fitz + pdfminer для одного PDF;
    возвращает `{"pages", "tag_dict", "error"}`.
  - `worker_init(shared_kwargs)` / `process_single_pdf_parallel(fp)` для PPE.
- `pdf_pipeline_main.py::run_pdf_pipeline(pdf_path, cfg)` — оркестратор (B.3):
  - `OD_tab_parsing.od_table_parsing_f` **после** шага 1 (как `main.start()`), не параллельно с pdfplumber на том же файле — иначе на Windows возможен `PDFSyntaxError: No /Root object`.
  - `ProcessPoolExecutor(n_workers)` + `as_completed` при `use_parallel=True` и >1 файл.
  - LPT-сортировка (большие файлы первыми), `gc.disable/enable`, прогресс `[N/total]`.
  - **B.5:** при любой ошибке PPE — `_fallback_sequential`: `tag_parser.reset()`, очистка `doc.pages`, полный `files_attribute_collector_f(pdf_path, curr_proj=...)` (тот же `pdf_path`, что у pipeline).
  - Мердж: `tag_parser.source_dict = merged_tags` после успешного цикла → `analyze`.
- `OD_tab_parsing.od_table_parsing_f`: выбор файла ОД в папке — только `.pdf` и `getDocTypeFromFile(name) == "OD"` (не подстрока `"OD"` в имени; иначе ложные MOD/NODE).
- Проверка 1005 (ревизии АН в штампе): `rules_check.py::_an_revision_long_from_18_1_rows` — полный шифр `X-ANyy`: геометрия `18_1_1` верх / `18_1_2` середина / `18_1_3` низ; токен `AN\d+` — **первый** сверху; база X — **первая** сверху ячейка **без** `AN\d+` (сверху вниз до ревизии без АН); ячейка целиком `X-ANyy` задаёт шифр и блокирует разбор по строкам.

**Timing log (Запуск 5, 27.03.2026):** `pdf_parsing/pdf_timing_log.py` — буфер + `finalize_timing_log` → `pdf_timing_log.xlsx` в папке результатов. Этапы: `find_files`, `file_<basename>` (per-file, только parallel), `parallel/sequential_pdf_processing_total`, `od_table_parsing_f`, `tag_parser_analyze`, `rules_check_start`, `total_pipeline`. Флаг `timing_log` в `pdf_pipeline_config.json` + чекбокс в `pdf_settings_gui.py`. `finalize_timing_log` в `finally` там, где интегрирован timing (при наличии в коде).

**Детальное профилирование (Запуск 7, 27.03.2026):**
- Флаг `profiling_enabled` (default `false`) в `pdf_pipeline_config.json` + чекбокс GUI.
- `get_cur_page_1` принимает `prof_dict=None`; замыкания внутри пишут `p{N}_find_tables`, `p{N}_inner_extract_pages`, `p{N}_table_extract`.
- `_extract_page_attrs(fp, doc, profiling_enabled)` собирает `open_pdfplumber_fitz`, per-page prof + `get_cur_page_total`, `post_processing`; возвращает `(pages, profiling_data)`.
- `_run_parallel` в `pdf_pipeline_main.py` логирует все ключи из `profiling_data["pages"]` в timing log.

**Результаты профилирования WIR-0005 (27.03.2026):**
- Thread1 total 91s: `inner_extract_pages`=45.4s (50%) + неизмеренное (pdfplumber.close?)=37s (41%) + `find_tables`=8.4s (9%).
- `inner_extract_pages` — скрытый pdfminer в `get_crop_table_dwg`, вызывается для всех DWG-типов (WIR/LAY/DW/CAE) чтобы найти границу рамки через `BiggestElement.find_big_element`. Thread2 параллельно тоже запускает pdfminer (37s).
- **Рычаг 1 реализован (27.03.2026):** `BiggestElement.find_big_element_fitz` в `find_function_MTO_BBB.py` — замена pdfminer на `fitz_page.get_drawings()`. Исправлены две критических ошибки:
  1. `files_att_collector.py` не передавал `fitz_page` в `get_cur_page_1` → добавлен аргумент.
  2. **Ротация страниц (rotation=270°)**: `fitz.get_drawings()` возвращает координаты в НЕПОВЁРНУТОМ пространстве PDF (y-up), тогда как `fitz_page.rect` — в повёрнутом. Для rotation=90°/270°: `pdfminer_x = fitz.y`, `pdfminer_y = fitz.x` (прямое отображение, без инверсии). Для rotation=0°: `pdfminer_y = page_h - fitz.y` (инверсия y). Без этого исправления для всех DWG LAY/CAE/WIR (rotation=270°) координаты рамки были некорректными.
- **Рычаг 2 реализован (27.03.2026):** замер `close_pdfplumber` в `pdf_worker.py` + вывод в timing log.
- **Валидация D.1 пройдена:** `compare_pdf_pipeline_outputs.py` — xlsx `only_sequential=0, only_parallel=0`. Tag-файлы совпадают по мультимножеству строк (sorted_equal=True), разница только в порядке строк из-за параллельной обработки.
- Детали, архитектура, roadmap: `.cursor/rules/AI_pdf_parsing_optimization.mdc`.

**Валидация D (Запуск 4, 26.03.2026):** `pdf_parsing/compare_pdf_pipeline_outputs.py` (D.1: seq vs par,
`--pdf-folder` или `--dirs`); `pdf_parsing/pdf_perf_benchmark.py` (D.2: время, опционально `--measure-pickle`, `--cpu-sample` + psutil).

**PDF v2 — шаблонный движок штампа (28.03.2026, roadmap Запуски 1–5):** пакет `pdf_parsing_v2/`
(ядро и шаблоны вне `pdf_parsing/`; `compat.py` импортирует `PageStampAttributes` из v1). Конфиг `pdf_v2_config.json` в корне; `pdf_parsing_v2/v2_config.py`.
Модели: `models.py`. Шаблоны: `templates/dwg_page1.json`, `dwg_page2.json`, `bbb_mto_page1.json`, `bbb_mto_page2.json`
(пересборка: `python pdf_parsing_v2/_build_stamp_templates.py`). Набор `templates/sets/agcc_287.json` ссылается на все четыре.
Очистка текста: `field_cleaners.py` + `CLEANERS`. В `main.py` — кнопка «Проверка PDF v2 (шаблоны)»: `wire_pdf_v2_button` + subprocess `v2_pipeline.py` (как v1 PDF), при наличии файла передаётся путь к `pdf_v2_config.json`.
Рядом с кнопкой — ⚙ → `show_pdf_v2_settings(parent)` (`pdf_parsing_v2/pdf_v2_settings_gui.py`).

**Фаза 0 диагностики сетки штампа (30.03.2026 15:19):**
- Новый модуль `pdf_parsing_v2/grid_diagnostic.py` (CLI: `python -m pdf_parsing_v2.grid_diagnostic`).
- Анализирует стр.1 PDF по всем типам (BBB/MTO/OD/CJ/DWG): `find_frame` + `fitz.find_tables`,
 фильтр ячеек в область штампа (bbox шаблона + 20мм), кластеризация X/Y линий (`0.5мм`),
 статистика по сетке, кандидаты якорей (крупные ячейки >5% площади штампа), OD split-check по полю `6_2_Quantity_of_sheets`.
- Экспорт: `grid_diagnostic_report.xlsx` (листы `Per-file`, `X-lines`, `Y-lines`, `Cross-type summary`,
 `Anchors`, `OD_splits`) + PNG overlay в `overlay_png/`.
- Прогон на `pdf_parsing_v2/templates/test_pdf` (70 файлов) завершён успешно; артефакты:
 `pdf_parsing_v2/grid_diagnostic_output/grid_diagnostic_2026.03.30_1516/`.

**Фаза 1 (часть) — расширение модели данных (30.03.2026 15:30):**
- `pdf_parsing_v2/models.py::FieldDef` расширен:
 `field_type: "data"|"empty"|"static"`, `expected_text: str|None`, `is_anchor: bool`.
- `pdf_parsing_v2/models.py::StampTemplate` расширен параметрами grid adaptation:
 `grid_adapt`, `grid_tolerance_template_mm`, `grid_tolerance_detected_mm`,
 `anchor_min_area_fraction`, `snap_size_tolerance`, `snap_max_distance_mm`, `snap_iou_threshold`.
- Backward-compatible загрузка/сохранение JSON сохранена:
 старые шаблоны читаются через defaults; новые поля сериализуются в `to_dict`.
- Каталог `pdf_parsing_v2/templates/catalogs/fields_agcc.json` дополнен static-полями:
 `static_razrabotal`, `static_proveril`, `static_n_kontrol`, `static_utverdil`,
 `static_izm`, `static_kol_uch`, `static_list`, `static_n_doc`.

**Фаза 2a/2b — grid matcher (30.03.2026 15:37):**
- Добавлен `pdf_parsing_v2/grid_matcher.py` как единый модуль движка для editor/pipeline.
- Реализованы dataclass'ы: `CellBbox`, `AnchorMatch`, `SnapResult`, `AdaptResult`.
- Реализованы функции:
 `cluster_lines`, `get_detected_stamp_cells`, `match_anchor_cells`,
 `approximate_field_bbox`, `snap_to_detected`, `adapt_all_fields`.
- Логика якорей: сначала ручные `FieldDef.is_anchor`, при отсутствии — авто-якоря
 по `template.anchor_min_area_fraction`.
- Логика snap: статус `matched` / `split_merged` / `no_match` по IoU, ratio площади
 и дистанции (`snap_max_distance_mm`), merge split-ячеек через outer bbox.

**Фаза 2c/3a — интеграция адаптации и UI свойств (30.03.2026 15:44):**
- `stamp_extractor.py::_extract_with_template()` использует `adapt_all_fields()` при
 `template.grid_adapt=True`; для каждого поля берётся адаптированный bbox, с fallback
 на старый `field_to_fitz_rect`.
- `template_editor/properties_panel.py` расширен:
 `field_type` (`data|empty|static`) + `expected_text` (только для static).
- `template_editor/cell_items.py`: цвета по типу поля (data=blue, static=purple, empty=gray dashed).
- `template_editor/template_meta_panel.py`: добавлен чекбокс
 «Адаптировать под сетку PDF» (`grid_adapt`); `main_window.py` пробрасывает `grid_adapt`
 в тестовый шаблон F5 и в сохранение JSON.
- `main_window.py::_collect_field_defs_from_scene()` теперь сохраняет новые атрибуты
 `field_type`, `expected_text`, `is_anchor` при пересборке FieldDef из geometry scene.

**Фаза 3b (часть) — консолидация сетки в редакторе (30.03.2026 15:41):**
- `template_editor/main_window.py`: добавлены действия `Консолидировать сетку` и `Применить`.
 Консолидация: сбор границ ячеек на canvas → `grid_matcher.cluster_lines()` → overlay
 (синие вертикали, оранжевые горизонтали, красные проблемные ячейки).
- Критерий проблемной ячейки: хотя бы одна граница дальше `tolerance` от ближайшей логической линии.
 В консоль пишется: `N вертикальных, M горизонтальных, K проблемных`.
- `Применить` делает snap всех границ ячеек к ближайшим логическим линиям и помечает шаблон modified.
- `template_editor/template_meta_panel.py`: добавлено поле `Grid tol, мм`
 (`get_grid_tolerance_template_mm` / `set_grid_tolerance_template_mm`).
- `template_editor/cells_panel.py`: статус проблемных ячеек маркируется `⚠`.
- `grid_tolerance_template_mm` теперь пробрасывается в тестовый `StampTemplate` (F5) и в JSON при сохранении.
- Overlay консолидации снимается при следующем ручном действии пользователя
 (selection change, drag/resize, смена страницы, правка свойств).

**Фаза 3c — предпросмотр подгонки под сетку (30.03.2026 15:44):**
- `template_editor/main_window.py`: добавлена кнопка `Подогнать под сетку`
 (toolbar + меню), доступная только при `grid_adapt=True` и открытом PDF.
- Подгонка использует единый движок `grid_matcher.adapt_all_fields()`:
 временный шаблон строится из текущих назначенных ячеек, затем `FieldRectItem` переносятся
 в адаптированные bbox (без автосохранения в JSON).
- Диагностический overlay: dotted-стрелки `approx -> final`, синие matched detected ячейки,
 жёлтый контур `split_merged`, красный контур `no_match`.
- В лог выводятся агрегаты качества подгонки:
 `anchor/matched/split_merged/no_match`, `max_snap_mm`, `moved`.

**Template editor UX — добавление нового поля рисованием (30.03.2026 16:47):**
- `template_editor/main_window.py`: действие `Добавить поле` (checkable) в toolbar + меню.
- `template_editor/graphics_view.py`: draw-mode с preview-рамкой и сигналом `rect_drawn(QRectF)`.
- После рисования создаётся новый `FieldRectItem` с новым `cell_index`, режим рисования
 автоматически выключается (one-shot), фокус переводится на созданное поле
 (выделение + `centerOn` + загрузка в `PropertiesPanel`), список ячеек обновляется.

**Template editor UX — `static_empty` и копирование Ctrl+drag (30.03.2026 18:19):**
- `pdf_parsing_v2/models.py`: `FieldDef.field_type` использует канонический `static_empty`
 (без alias для старой опечатки).
- `template_editor/properties_panel.py`: тип `static_empty` добавлен в combo «Тип поля».
- `template_editor/cell_items.py`: `Ctrl + drag` дублирует поле:
 создаётся копия с теми же свойствами (deepcopy `FieldDef`), оригинал остаётся на старом месте,
 перетаскиваемый экземпляр уходит в новую позицию.
- `template_editor/cell_items.py`: `static` и `static_empty` имеют различимые оттенки
 (пустой static визуально контрастнее относительно static с текстом).
- `template_editor/main_window.py`: подключён `scene.changed`-хук с контролем числа `FieldRectItem`,
 чтобы список ячеек автоматически обновлялся после создания дубликатов.

**Template editor UX — якоря/snap-параметры/preview-only и поля вне штампа (30.03.2026 20:30):**
- `pdf_parsing_v2/models.py`:
 - `FieldDef` расширен `outside_stamp: bool = False`;
 - `StampTemplate` расширен `preview_only_adapt: bool = False`;
 - JSON backward-compatible (старые шаблоны загружаются через defaults).
- `template_editor/properties_panel.py`:
 - добавлены чекбоксы `Anchor` (`is_anchor`) и `Вне штампа` (`outside_stamp`).
- `template_editor/template_meta_panel.py`:
 - вынесены параметры `grid_tolerance_detected_mm`, `snap_size_tolerance`,
 `snap_max_distance_mm`, `snap_iou_threshold`, `preview_only_adapt`.
- `template_editor/main_window.py`:
 - `Подогнать под сетку` использует параметры из meta panel (без hardcode);
 - в `preview-only` поля не перемещаются;
 - `no_match` поля не перемещаются даже в apply-режиме.
- `grid_matcher.py`:
 - `outside_stamp=True` исключает поле из bbox штампа, auto-anchor и snap;
 - добавлен статус `excluded` для таких полей.
- `template_editor/cell_items.py` + `cells_panel.py`:
 - поля вне штампа имеют отдельный оранжевый цвет и dashed-визуализацию.

**Глобальные допуски `find_tables` для v2 (30.03.2026 20:43):**
- Новый модуль `pdf_parsing_v2/find_tables_settings.py`:
 `get_find_tables_kwargs(cfg)` + `call_find_tables(page, cfg)` с fallback на вызов
 без kwargs для совместимости версий PyMuPDF.
- `pdf_v2_config.json` / `v2_config.py`:
 добавлены глобальные ключи `find_tables_snap_x_tolerance`, `find_tables_snap_y_tolerance`
 (дефолты `2.2` и `2.0`, как в v1 `find_functions.py`).
- Эти настройки применяются в трёх местах v2:
 - `grid_matcher.get_detected_stamp_cells()` (pipeline + editor adapt),
 - `template_editor/auto_detect.detect_cells()` (авто-разметка),
 - `grid_diagnostic._extract_cells()` (диагностика сетки).

**Диагностика `Подогнать под сетку`: визуализация search bbox (30.03.2026 20:51):**
- `template_editor/main_window.py`: в overlay режима адаптации добавлен фиолетовый
 dashed-контур рамки поиска (`search_bbox`), совпадающий с bbox фильтрации
 detected ячеек в `grid_matcher.adapt_all_fields`.
- В консольный лог добавлены координаты `search_bbox=(x0,y0,x1,y1)`, чтобы
 выявлять обрезание верхних/крайних ячеек из-за слишком узкой области поиска.

**Стабилизация matching в `grid_matcher` (30.03.2026 21:00):**
- Авто-якоря по крупным полям отключены: трансформация от якорей строится
 только при явно выставленных `FieldDef.is_anchor=True`.
- `snap_to_detected` больше не делает дальние «прыжки»:
 при отсутствии nearby-кандидатов в `snap_max_distance_mm` возвращается `no_match`.
- Для `split_merged` добавлена дополнительная IoU-проверка merged bbox.
- В лог `Подгонка под сетку` добавлен счётчик `manual_anchors`.

**Визуальная маркировка anchor-полей (30.03.2026 21:07):**
- `template_editor/cell_items.py`: для `FieldDef.is_anchor=True` рамка ячейки
 отображается отдельным anchor-цветом (`_COLOR_ANCHOR`).
- При активном `test_status` (ok/warning/error) тестовая раскраска приоритетнее anchor-цвета.

**Исправление list-ошибок (Запуск 5, 28.03.2026):**
- `field_cleaners.py::_ensure_str(value)` — гарантирует str: список→join, None→"".
  Все cleaners, вызывающие `string_parsing.list_*`, обёрнуты в `_ensure_str`.
- `stamp_extractor.py` + `compat.py`: защитные `isinstance(x, list)` проверки перед `.strip()`.
- **Правило:** при добавлении нового cleaner, вызывающего `string_parsing.list_*`,
  обязательно оборачивать в `_ensure_str`.

**Excel debug-отчёт v2 (Запуск 5, 28.03.2026):**
- `pdf_parsing_v2/v2_report.py` — `save_v2_debug_report(all_results, result_dir)`:
  файл `Отладка_v2_штамп_<timestamp>.xlsx`; строка = страница PDF, столбцы = cleaned+raw per field.
  Цветовая индикация по `is_valid`/`expected`. `ColumnDef` dataclass (паттерн из `step4_6_save_match_result_to_excel.py`).
  `collect_results_from_curr_proj(curr_proj)` — сбор `doc._v2_results` из обработанных документов.
- `v2_pipeline.py` сохраняет список `V2PageResult` в `document._v2_results` после per-file обработки.
- Управляется флагом `export_debug_excel` (bool, default `true`) в `pdf_v2_config.json`.

**GUI настроек v2 (`pdf_parsing_v2/pdf_v2_settings_gui.py`, Запуск 5, 28.03.2026):**
- `show_pdf_v2_settings(parent)` — CTkToplevel, паттерн как `pdf_settings_gui.py`.
- Секции: «Общие» (`enabled`, `templates_dir`) и «Отладка» (`export_debug_excel`, `debug_visual`, `debug_visual_overlay`, `debug_visual_dir`).

**Правило для всех новых Excel-отчётов:** использовать `ColumnDef`-паттерн из `v2_report.py`
(или `step4_6_save_match_result_to_excel.py`) — модульная структура для удобства переиспользования в чатах.

**Правило для дат в roadmap-файлах:** всегда указывать дату **и время до минуты** в формате
`ДД.ММ.ГГГГ ЧЧ:ММ` (например: `29.03.2026 14:35`). Применяется ко всем событиям во всех
`.mdc`-файлах: создание файла, статус этапа, результат чата, заметки.

Ядро движка (Запуск 2, 28.03.2026 — ЗАВЕРШЁН):
- `frame_detector.py::find_frame(fitz_page) → FrameInfo` — порт `BiggestElement.find_big_element_fitz`,
  чистая функция без глобалов, fallback на ГОСТ-отступы; rotation 0°/270° handled.
- `coord_transform.py` — `field_bbox_to_absolute` (mm→pts), `pdfminer_to_fitz` (y-flip для
  PyMuPDF ≥ 1.19.0, одинаков для всех rotation), `field_to_fitz_rect` (pipeline + padding + clamp).
- `stamp_extractor.py::extract_page(fitz_page, doc_type, page_num, templates) → V2PageResult` —
  единый путь; select_templates → extract per field via get_textbox → clean → validate regex →
  score → pick best. `borders_margin` из v1 НЕ используется; вместо него `padding_mm` (+0.5мм default).

Pipeline v2 (Запуск 4 D.1–D.3, 28.03.2026):
- `pdf_parsing_v2/v2_pipeline.py` — `run_v2_pipeline(pdf_path, cfg) → result_dir`:
  цикл по `get_files_single(pdf_path)`, per-file `extract_page` → `compat.to_page_stamp_attributes` →
  `document.pages.append(psa)`, затем `parse_tags` / `OD_tab_parsing` / `tag_parser.analyze` / `rules_check`.
  CLI: `sys.argv[1]` = pdf_folder; опционально `sys.argv[2]` = `pdf_v2_config.json`; последняя строка stdout = result_dir.
  `_process_single_file` — чистая функция, готова к PPE (параллельность — одно изменение).
- `pdf_parsing_v2/compat.py` — `to_page_stamp_attributes`: поля штампа по совпадению `field_id` с ключами `dict_attributes`; `metadata` для вторичных атрибутов (пока обычно пусто в `extract_page`).

GUI редактор шаблонов (Запуск 6 F.1–F.6, 28.03.2026 — ЗАВЕРШЁН):
- `pdf_parsing_v2/template_editor/` — PySide6 (LGPL), отдельный процесс от CTk GUI.
  Запуск: `python -m pdf_parsing_v2.template_editor`.
  Модули: `main_window.py` (QMainWindow + меню + toolbar + splitter), `graphics_view.py`
  (QGraphicsView, zoom-to-cursor AnchorUnderMouse, pan Space/MiddleMouse, PDF via fitz pixmap dpi=150),
  `cell_items.py` (FieldRectItem — drag/select/handles/цвета/highlight, 8 resize handles),
  `auto_detect.py` (fitz.find_tables → bbox ячеек), `cells_panel.py` (заголовок «Список полей»;
  QTreeWidget #/Поле/Статус/Значение; клик→centerOn; при выделении ячейки на canvas — подсветка строки
  + `scrollToItem` к видимой области; фильтр «только неназначенные»), `properties_panel.py` (QFormLayout: id из каталога,
  label, clean из CLEANERS, regex, expected, padding_mm), `catalog_editor.py` (QDialog для FieldCatalog),
  `set_editor.py` (QDialog для TemplateSet), `templates_browser.py` (левая панель: шаблоны/каталоги/наборы;
  у набора дочерние узлы = `template_files`; DnD из «Шаблоны» в набор добавляет шаблон, из набора на «Шаблоны» —
  убирает; автосохранение `TemplateSet.to_json`), `_save_dialog.py` (метаданные при сохранении шаблона).
  Тестирование (F5): собирает assigned → StampTemplate → extract_page → раскраска ячеек.
  Save: пересчитывает scene→pdfminer→мм от рамки → JSON. Load: JSON → field_to_fitz_rect → Canvas.

Валидация v1 vs v2 (Запуск 6 E.1–E.2, 28.03.2026):
- `pdf_parsing_v2/compare_v1_v2.py` — сравнение результатов v1 vs v2:
  `--pdf-folder` (два прогона + сравнение) или `--dirs V1_DIR V2_DIR`.
  Tag-файлы: побайтово + sorted-строки; xlsx: multiset строк по общим колонкам.
  Exit code: 0 = совпадение, 2 = расхождения.
- `pdf_parsing_v2/debug_visual.py` — PNG-аннотации для визуальной отладки:
  `render_debug_page` (рамка + цветные bbox + подписи), `render_debug_page_overlay`
  (все шаблоны-кандидаты). Интегрирован в `v2_pipeline.py` (шаг 9) при `debug_visual=True`.
  `render_debug_for_file` — helper для pipeline, автоматический overlay при score < 0.7.

**Boundary-snap алгоритм адаптации (30.03.2026):**
- `grid_matcher.py::adapt_all_fields()` переработан: вместо IoU-snap используется
  boundary-snap — каждая из 4 границ поля привязывается к ближайшей grid-линии.
- Новые функции: `_extract_grid_lines`, `_build_adjacency_graph`, `_snap_4_boundaries`,
  `_shape_score`, `_score_candidate`, `_boundary_consistency`.
- `AdaptResult` расширен: `shape_score`, `snapped_boundaries` (0-4), `confidence`, `iteration`.
- Новые статусы: `derived` (fallback на approx_bbox при неудачном snap), остальные сохранены.
- `StampTemplate` расширен: `max_shape_change_ratio` (default 2.5), `cascade_score_threshold` (default 0.4).
- Граф смежности: два поля считаются соседними если расстояние между границами < tol
  И проекция перекрытия по перпендикулярной оси > 50% меньшей из двух проекций.
- Scoring нормализован: при отсутствии matched соседей `boundary_consistency` не штрафует.
- UI: `cells_panel` — колонки Score/Iter; `template_meta_panel` — скрыты `snap_size_tol` и `snap_iou`,
  добавлены `max_shape_change` и `min_score`; overlay — grid-линии (X синие, Y оранжевые) + score метки.
- План: `pdf_parsing_v2/PLAN_cascade_adaptation.md`. Фаза B (каскад) — при необходимости.

**Редактор шаблонов — адаптация и зазоры (31.03.2026 09:45):**
- «Подогнать под сетку» в `template_editor` использует `grid_matcher.adapt_by_cell_assignment`
  (cell-assignment + `scipy.optimize.linear_sum_assignment`, fallback жадный по порядку полей).
- Пост-проход **«Заполнить зазоры до соседей»** (`pdf_parsing_v2/stamp_fill_gaps.py`, замена удалённого
  «Слепить поля»): расширение rect до соседей **только в режиме сэндвича** (оба гориз./верт. соседа);
  **не** в общем v2 pipeline. Сообщено **некорректное** поведение — на доработке; детали и файлы: **T5.2 / T5.3**
  в `.cursor/rules/AI_template_editor_roadmap.mdc`.

Дорожные карты PDF v2 (**компактные**): `.cursor/rules/AI_pdf_v2_roadmap.mdc`, `.cursor/rules/AI_template_editor_roadmap.mdc`. Полные снимки старых версий и индекс — `.cursor/rules/AI_pdf_v2_template_editor_roadmap_INDEX.md`.

Подробности — `.cursor/rules/AI_pdf_parallel_roadmap.mdc`.

---

## 15. Roadmap переноса GUI в веб (черновик, 24.03.2026)

Черновик дорожной карты (LAN, сервер на Windows, UNC): `.cursor/rules/AI_web_migration_roadmap.mdc` (`alwaysApply: false`).
Исходная оценка — план Cursor `web_migration_assessment_ca214a9a`; процесс переноса будет уточняться.
