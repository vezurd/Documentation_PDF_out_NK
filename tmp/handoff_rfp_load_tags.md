# Handoff: `load_tags=false` не убирает раскрытие ДС45

Дата: **2026-09-04**. **Статус: сделано 2026-09-04** (соседний `rfp_parts_net_no_tags.xlsx`; УЛ отдельно не пересобирается). Живой UNC `RFP_Зиновьев` у агента не открывался (`exists: False`).

Правила: [AI_rfq_context.mdc](.cursor/rules/AI_rfq_context.mdc), [AI_rfp_step4_excel.mdc](.cursor/rules/AI_rfp_step4_excel.mdc), [AI_rfp_parts.mdc](.cursor/rules/AI_rfp_parts.mdc), [SETTINGS_REFERENCE.md](RFQ/tags_rfp_compare/SETTINGS_REFERENCE.md).

---

## Зачем

Пользователь хочет финальный Step4 **без учёта тегов**: строка лот=`1` + два тега **не** должна становиться двумя строками. Галка в **RFP · Настройки** снята, прогон сделан — в Excel ДС45 / `BCC0002659` всё равно ~10 строк по `VALUES=1`.

---

## Сделано (флаг работает, но поздно)

- Конфиг **`load_tags`** (default **`true`**). Профиль: [rfp_tags_compare_config.json](RFQ/tags_rfp_compare/rfp_tags_compare_config.json) уже **`false`**.
- GUI: [rfp_settings_panel.py](ds_compare_center/rfp_settings_panel.py), legacy [rfp_tags_settings_gui.py](RFQ/tags_rfp_compare/rfp_tags_settings_gui.py).
- После raw RFP/MTO/VO + `load_packing_dataset`, **до** `run_units_gate` / `step1_finalize`: [column_optimization.py](RFQ/tags_rfp_compare/column_optimization.py) `apply_load_tags_mode` / `strip_loaded_tags` — только in-memory, кэш pickle не пишется.
- Wrapper [step1_load_rfp.py](RFQ/tags_rfp_compare/step1_load_rfp.py) `step1_load_rfp_data(..., load_tags=True)`.
- Оркестратор: [agregate_tags.py](RFQ/tags_rfp_compare/agregate_tags.py).
- Смоук: [tmp/test_rfp_load_tags_smoke.py](tmp/test_rfp_load_tags_smoke.py); плюс [tmp/test_rfp_native_settings_smoke.py](tmp/test_rfp_native_settings_smoke.py).

`column_optimization.*.tags.load=false` **не** решать эту задачу: он после finalize/split.

---

## Почему ДС45 всё равно раскрылась

Запуск: `rfp_parts.use_latest_net=true` → Step1 читает **`rfp_parts_net.xlsx`**, не файл ДС45.

В [analyze_rfp_parts.py](RFQ/rfp_parts/analyze_rfp_parts.py):

1. `_build_unit_counters`: на **каждый тег** `unit_counter[key] += 1`, даже если лот `1` (`tags > lot`). Остаток отрицательный не пишется.
2. `_append_summary_unit_rows`: для тегированного ключа — цикл `VALUES=1` на единицу счётчика; **`DS_NUMBER` = глобальный seq**, не № п/п исходной ДС.

Итог: 5 исходных комплектов × 2 тега (RX+TX) → **10 строк net по 1**, номера вроде 22430–22439. Step1-split тут ни при чём: он дал бы пары **`1` и `0`**. На скрине везде **`1`**.

`load_tags=false` обнуляет `TAGS` **уже после** net. Collapse Step4 склеивает безтеговые только с **одним** `DS_NUMBER` → 10 строк остаются.

Схлопывание после strip **суммированием** 1+1 даст **`VALUES=2`** — это тоже неверно (лот был 1). Наивный collapse net не чинит qty.

Исторический след (агент UNC не видел): [tmp/rfp_сбор_отчеты_2026_08_12_15_28/rfp_parts_report.txt](tmp/rfp_сбор_отчеты_2026_08_12_15_28/rfp_parts_report.txt) — ДС45, лот 1, два тега; [tmp/tsd_id_fix_list.txt](tmp/tsd_id_fix_list.txt) — УЛ ДС45 `BCC0002659` qty 1, теги RX+TX.

Проверочный код на скрине: `BCC0002659`, титул `8630-KSB3`, порядковый `ДС45_100`, Kramer / `PT-2UT/R-KIT`.

---

## Не делать снова

- Не считать багом `split_rfp_rows` / worker split, пока вход — net с `VALUES=1` на тег.
- Не чинить через `column_optimization.tags.load`.
- Не переписывать боевой `rfp_parts_net.xlsx` флагом `load_tags` (теговый прогон должен видеть прежний свод).
- Не склеивать все строки одного кода в одну `VALUES=Σ` без решения пользователя: 5 позиций лот=1 ≠ одна строка `5`, и ≠ десять строк `1`.

---

## Сделано 2026-09-04 (смысл галки)

- Сбор частей всегда пишет соседний **`rfp_parts_net_no_tags.xlsx`**: N исходных строк → N строк, лот как в ДС, TAGS пустые (`_build_no_tags_summary_rows`). Боевой `rfp_parts_net.xlsx` не затирается.
- `load_tags=false` + `use_latest_net`: preflight требует no-tags файл; нет → новый штамп. `resolve_effective_rfp_path` читает no-tags.
- **УЛ:** отдельный кэш/свод не нужен (qty не раскрывается по тегам; strip в памяти).

Смоуки: [tmp/test_rfp_parts_tag_lot_remainder_smoke.py](tmp/test_rfp_parts_tag_lot_remainder_smoke.py), [tmp/test_rfp_parts_net_preflight_smoke.py](tmp/test_rfp_parts_net_preflight_smoke.py), [tmp/test_rfp_effective_path_smoke.py](tmp/test_rfp_effective_path_smoke.py), [tmp/test_rfp_load_tags_smoke.py](tmp/test_rfp_load_tags_smoke.py).

Цель ДС45 закрыта: лот=`1` + два тега → **одна** строка `VALUES=1` в no-tags net (N комплектов → N строк, не сумма). УЛ без отдельного артефакта.
