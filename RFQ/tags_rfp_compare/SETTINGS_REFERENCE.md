# Настройки RFP Tags Compare

> Актуально на 2026-09-04. Основной интерфейс настроек — нативная вкладка
> **RFP · Настройки** в `python -m ds_compare_center`.

## Файлы и профили

- `RFQ/tags_rfp_compare/rfp_tags_utils.py` — defaults, deep merge, load/save.
- `RFQ/tags_rfp_compare/rfp_tags_compare_config.json` — профиль **Основной RFP**.
- `RFQ/tags_rfp_compare/rfp_tags_compare_config_asbuild.json` — профиль **As-build**.
- `ds_compare_center/rfp_settings_panel.py` — основной редактор PySide6.
- `RFQ/tags_rfp_compare/rfp_tags_settings_gui.py` — legacy CTk только для старых
  entry points (`main.py` / `main_v2.settings_launchers`).
- `RFQ/tags_rfp_compare/agregate_tags.py` — загрузка конфига, preflight и pipeline.

`get_default_config()` и `get_default_asbuild_config()` — источники defaults. При
загрузке пользовательский JSON рекурсивно накладывается поверх defaults. Оба GUI
при сохранении начинают со свежезагруженного профиля и меняют только показанные
поля: неизвестные ключи и полный `column_optimization` не теряются.

## Полная схема верхнего уровня

Ниже перечислены все runtime-секции и их текущие ключи. Значения показаны для
основного профиля defaults; пользовательский JSON может их перекрывать.

```json
{
  "memory_log": true,
  "load_tags": true,
  "paths": {
    "rfp_path": "...",
    "rfp_registr_lot_path": "...",
    "mto_path": "...",
    "vo_path": "...",
    "code_ban_file": "...",
    "units_convert_matrix": "...",
    "ds_manager_matrix": "...",
    "gem_supply_codes": "...",
    "replacement_table_file": "...",
    "result_dir_base": "..."
  },
  "rfp_parts": {
    "auto_update_checklist": true,
    "use_latest_net": true,
    "summary_path": ".../Сводная таблица ДС по вед.договорам.xlsx",
    "checklist_path": ".../ДС_увеличение_уменьшение.xlsx",
    "increase_dir": ".../RFP на увеличение",
    "decrease_dir": ".../RFP на уменьшение"
  },
  "step1": {
    "debug": false,
    "skip_split": false,
    "use_rfp_code_ban": true,
    "rfp_pipeline_mode": "standard"
  },
  "units_split_ban": {
    "enabled": true,
    "units": ["м", "м2", "бухта", "кг", "т"]
  },
  "step2": {
    "debug": false,
    "export_load_results_excel": true,
    "export_positions_database_excel": false,
    "flat_mto_structure": false
  },
  "step3": {
    "debug": false,
    "export_to_excel": false
  },
  "step4": {
    "debug": true,
    "debug_step4_1": true,
    "debug_step4_2": true,
    "debug_step4_3": true,
    "debug_step4_4": true,
    "collapse_debug": false,
    "unified_debug": true,
    "timing_log": true,
    "use_multiprocessing": false,
    "parallel_by_title_mark": false,
    "max_workers": 4,
    "verbose_progress_messages": true,
    "memory_top_stats": false,
    "memory_top_n": 15,
    "debug_tag": [],
    "debug_code": [],
    "debug_title_system": [],
    "print_title_systems_table": false,
    "export_title_systems_comparison_excel": true,
    "assign_rfp_mto_code_compare_colors": true,
    "filter_to_mto_titles": false,
    "include_mto_vo_without_rfp_anchor": false,
    "include_packing_lists": true,
    "export_bcc_accum_matrix": true,
    "ul_match_use_mto_tags": true
  },
  "rfp_tags_utils": {
    "finalize_timing_top_n": 5,
    "save_input_fingerprints": true
  },
  "column_optimization": {
    "rfp": {},
    "mto": {}
  }
}
```

Секция `rfp_parts` содержит штатные ключи:
`auto_update_checklist`, `use_latest_net`, `input_mode`, `ds_source_dir`,
`ds_registry_path`, `summary_path`, `checklist_path`, `increase_dir`,
`decrease_dir`. `use_latest_net=true` (default основного профиля) и
`input_mode=legacy_net` заставляют Step1 читать последний `rfp_parts_net.xlsx`
из `RFP сводный файл\YYYY.MM.DD_HH.MM` (при `load_tags=false` —
`rfp_parts_net_no_tags.xlsx`) и перед этим пересобрать свод, если
в `RFP_Зиновьев` появился новый файл, файл новее свода, или для режима
без тегов нет sibling-файла; `use_latest_net=false` —
`paths.rfp_path` (as-build). `input_mode=ds_only` читает
`Свод ДС для запуска.xlsx`; `hybrid` — `Свод ДС-RFP для запуска.xlsx`.
Пустой `ds_source_dir` берёт `last_ds_trusted_folder` с вкладки
RFP · Сбор частей. Нет нужного свода — явная ошибка, без silent fallback
на parts net.
Ключи `auto_update_checklist` / `summary_path` / папки increase/decrease
в JSON сохраняются, но Запуск их не использует.
Контракт ДС/hybrid: [AI_ds_rfp_hybrid.mdc](.cursor/rules/AI_ds_rfp_hybrid.mdc).

### `paths.units_convert_matrix`

- **Default (общий для основного и as-build профилей):**
  `\\bcc\eng\PrDoc\377_НИПИГАЗ\АГХК\КСБ\RFP_MTO_VO\_RFP\матрица_ед_изм.xlsx`
- **Назначение:** xlsx-матрица коэффициентов преобразования единиц измерения для
  **строгого** RFP/MTO/УЛ gate: строка source unit → target unit (Google MTO) с
  коэффициентом. Не используется в grouped DS compare (`ds_units_normalize` /
  `highlight_compared_units_columns`).
- As-build **не** переопределяет путь — наследует тот же default из
  `get_default_config()`.

### `paths.ds_manager_matrix`

- **Default (общий для основного и as-build профилей):**
  `\\bcc\eng\PrDoc\377_НИПИГАЗ\АГХК\КСБ\RFP_MTO_VO\_RFP\Список ДС - Фамилии МП.xlsx`
- **Назначение:** UNC xlsx, лист-свод с колонками **`Имя ДС`** / **`Фамилия`**
  (допустимы и старые **`№ ДС`** / **`МП`**). Не читать лист позиций
  (`Имя ДС` + `№ позиции` без фамилии). Ключи вида `ДС01` / `ДС47_13`
  сопоставляются с `ДС1` и `ДС47_13А`. Preflight сверяет строки матрицы
  с файлами в `RFP_Зиновьев` (не fatal).
- **Лист «Фактический и порядковый»:** собирается автоматически при каждом
  запуске RFP (`sync_ds_roster_sheet`) и вручную с вкладки **RFP · ДС ↔ МП**
  (`python -m ds_compare_center --tab rfp_ds_mp`). Строка 1 = шапка Лист2: **`Имя ДС`** /
  **`Кол-во RFP`** / **`Фамилия`**, далее фактический, порядковый, статус.
  Таблицу можно копировать на Лист2 целиком. При расхождении с файлами
  `RFP_Зиновьев` этот лист переписывается; **значения «Имя ДС» / «Фамилия»
  на Лист2 не затираются**. Справа дописываются столбцы **Статус**,
  **Дубликат**, **Что делать** (робот читает только Имя ДС и Фамилия). Если xlsx
  занят Excel — рядом пишется копия с датой/временем в имени
  (`…_ГГГГ.ММ.ДД_ЧЧ.ММ.СС.xlsx`), прогон не падает.

### `paths.gem_supply_codes`

- **Default (общий для основного и as-build профилей):**
  `\\bcc\eng\PrDoc\377_НИПИГАЗ\АГХК\КСБ\RFP_MTO_VO\_RFP\Коды_Поставок_ГЭМ_8950.xlsx`
- **Назначение:** UNC xlsx кодов поставки ГЭМ. Читается первый лист, первая
  строка — заголовок (пропускается), дальше коды. После успешного
  сопоставления УЛ (`compare_rfp_rows_with_packing`) overlay заменяет
  только статус **«Неотгружено по УЛ»** на позициях титулов `8950*`.
  Отсутствующий или нечитаемый файл **не fatal** (лог + overlay no-op).

`column_optimization` загружается из
`RFQ/tags_rfp_compare/column_optimization_default.json`. В `rfp` поддерживаются
`DS_NAME`, `DS_NUMBER`, `DS_TITLE`, `DS_SPECIFICATION`, `tags`, `DS_CODE_1C`,
`code`, `name`, `type_mark`, `values`, `units`, `name_2`, `values_2`; в `mto` —
`tags`, `numbers`, `name`, `type_mark`, `code`, `vendor`, `units`, `values`,
`mass`, `annotation`. У каждого элемента есть `load`, у текстовых полей может
быть `truncate`.

## Зависимости настроек

- `load_tags=true` (default): теги читаются из сырых RFP/MTO/VO/УЛ как сейчас.
  `false` — preflight/Step1 берут соседний `rfp_parts_net_no_tags.xlsx` (лот
  как в файле ДС, без раскрытия на теги; боевой `rfp_parts_net.xlsx` не
  затирается; нет файла — пересборка в новый штамп). После загрузки (до
  units_gate и split) обнуляет колонку `TAGS` в памяти у RFP/MTO/VO/УЛ.
  Match Step4 и посадка УЛ идут по коду. Кэш pickle УЛ не переписывается.
  Не путать с `column_optimization.*.tags.load=false`: тот флаг применяется
  **после** finalize/split и сам по себе раскладку не отменяет.
- `step1.rfp_pipeline_mode`: `standard`, `with_orphan_mto_vo` или `mto_vo_only`.
  В двух не-стандартных режимах
  `step4.include_mto_vo_without_rfp_anchor=true` принудительно; в
  `mto_vo_only` Step1/RFP load пропускается.
- `units_split_ban.enabled=false` отключает запрет единиц; пустой список при
  включённом флаге заменяется default-списком.
- `step1.skip_split` переносит split RFP в worker и обычно используется вместе
  с `step4.parallel_by_title_mark`.
- `step4.max_workers=0` означает автоматический выбор.
- `step4.use_multiprocessing` — дополнительный MP внутри отдельных match-функций.
  Его default в коде — `false`, хотя рабочий пользовательский JSON вправе
  перекрыть значение на `true`.
- `step4.parallel_by_title_mark` — другой механизм: параллельная обработка
  независимых title/mark. Не считать эти два флага синонимами.
- `step4.include_packing_lists=true` включает общий кэш УЛ. Ошибка
  `tags_count > quantity` останавливает RFP export до финального Excel.
- `step4.ul_match_use_mto_tags=true` (default): после тегов RFP и безтеговых
  слотов УЛ сначала искать совпадение с `TAG_MTO` той же строки, затем добор
  по коду. Выключено — сразу слепой добор (`Собрано по несопоставленным тегам`).

Списки `debug_tag`, `debug_code`, `debug_title_system` хранятся JSON-массивами,
а в GUI редактируются строкой через запятую.

## Профиль As-build

As-build использует собственные пути RFP/MTO/result/code-ban и merge с
`get_default_asbuild_config()`. Native GUI фиксирует и при сохранении
восстанавливает следующие различия:

- `step2.flat_mto_structure=true`;
- `step4.filter_to_mto_titles=true`;
- `step4.include_packing_lists=false`;
- `step4.assign_rfp_mto_code_compare_colors=false`;
- `rfp_parts.auto_update_checklist=false`;
- `rfp_parts.use_latest_net=false`.

Оба флага `rfp_parts.auto_update_checklist` и `rfp_parts.use_latest_net`
принудительно восстанавливаются backend load/save и CLI; остальные четыре —
as-build defaults и disabled-поля native редактора.

Запуск `python -m RFQ.tags_rfp_compare.agregate_tags <config.json>` определяет
as-build по штатному пути as-build-конфига, мержит файл с as-build defaults и
также принудительно выключает «последний свод частей» (`use_latest_net=false`). Ошибка чтения или разбора
явно переданного файла завершает процесс с ненулевым кодом.

## Native UI и preflight

Вкладка PySide6 переключает профили **Основной RFP / As-build**, показывает все
поля legacy-формы плюс `rfp_parts`, зависимости pipeline и фиксированные
as-build-различия. Блок УЛ показывает quality/status общего кэша и позволяет
перейти на вкладку упаковочных листов.

Для основного профиля `rfp_parts.use_latest_net=true` запускает preflight
в `agregate_tags` **до Step1**:

1. сравниваются даты/состав файлов `RFP_Зиновьев` с последним
   `rfp_parts_net.xlsx` (снимок `rfp_parts_sources.json`); при
   `load_tags=false` ещё проверяется соседний `rfp_parts_net_no_tags.xlsx`;
2. при новых или более новых файлах (или нет no-tags при выключенных тегах)
   вызывается тот же сбор частей, что вкладка «RFP · Сбор частей»
   (`run_rfp_parts_analyze`, без списка ДС) — пишет оба net в новый штамп;
3. актуальный теговый свод не пересобирается, если папка не менялась и
   нужный файл режима уже есть.

Ошибка папки частей или сборки свода закрывает запуск до Step1.
`mto_vo_only` и as-build (`use_latest_net=false`) пропускают этот шаг.
Сверка `ДС_увеличение_уменьшение.xlsx` остаётся только на вкладке сбора.

При `step4.include_packing_lists=true` сразу после свода частей тот же
запуск проверяет свежесть кэша УЛ (`ul_preflight`): fingerprint папки ТСД
vs `tsd_packing_rows.cache`, при расхождении — `load_and_cache_tsd_packing`
как вкладка «ДС · Упаковочные листы». As-build (`include_packing_lists=false`)
пропускает шаг.

## Прогресс запуска

Кнопка основного запуска динамически показывает `↔ УЛ` только при
`step4.include_packing_lists=true`. Справа отображаются пятнадцать live milestones:
`prepare`, `ds_id_check`, `ds_mp_check`, `parts_preflight`, `ul_preflight`, `rfp_load`, `mto_google_load`, `vo_load`, `units_gate`,
`step4_match`, `global_checks`, `packing_lists`, `save_excel`, `bcc_accum_matrix`, `complete`.
`ds_id_check` — сверка имён (номера ДС RFP ↔ папки УЛ), non-fatal; `detail` это `ОК: …` или `Замечания: …`.
`ds_mp_check` — сверка файлов RFP ↔ фамилии МП (`sync_ds_roster_sheet`), non-fatal; `detail` это
`ОК: все файлы RFP с фамилией…` / `Надо дозаполнить: …` / `Ошибка: …`.
`ul_preflight` — свежесть свода УЛ; `unchanged` / `rebuilt` или Skipped при выключенных УЛ.

Канонический stdout-протокол:

```text
@@RFP_MILESTONE {"milestone_id":"prepare","state":"Running","detail":""}
```

Допустимые состояния: `Waiting`, `Running`, `Done`, `Skipped`, `Error`.
`global_checks` и `packing_lists` в `detail` несут краткий баланс (вход при Running, `OK …→… (Δ±0.00)` при Done).
Chunk-parser корректно собирает marker/JSON/UTF-8, разделённые между chunks.
При ненулевом exit code GUI помечает `Error` только текущий active milestone,
не переписывая уже завершённые строки.

Вкладка **RFP · Сбор частей** остаётся report-only: кнопка называется
«Проверить части RFP и сформировать отчёт», production `Сводная RFP.xlsx` она
не формирует. Какой файл читает **RFP · Запуск**, задаёт галка
`rfp_parts.use_latest_net` в «RFP · Настройки».

## Добавление нового параметра

1. Добавить default в `get_default_config()`; для фиксированной разницы
   as-build — также в `get_default_asbuild_config()` и правила профиля.
2. Добавить поле в `ds_compare_center/rfp_settings_panel.py`; при необходимости
   синхронизировать legacy `rfp_tags_settings_gui.py`.
3. Пробросить effective value из `agregate_tags.py` до потребителя.
4. Сохранять merge-on-save: нельзя строить новый урезанный JSON только из
   виджетов.
5. Обновить эту справку и regression smoke.

## Verification

```powershell
python tmp/test_rfp_checklist_updater_smoke.py
python tmp/test_rfp_progress_markers_smoke.py
python tmp/test_rfp_native_settings_smoke.py
python tmp/test_rfp_load_tags_smoke.py
python tmp/test_rfp_effective_path_smoke.py
python tmp/test_rfp_packing_compare_smoke.py
```

Связанный контекст: `.cursor/rules/AI_rfq_context.mdc`,
`.cursor/rules/AI_rfp_parts.mdc`, `.cursor/rules/AI_ds_compare.mdc`,
`.cursor/rules/AI_performance.mdc`.
