# Handoff: остаток лота без тега в своде частей + сверка MTO/УЛ/отчётов

**Статус: сделано 2026-09-03** (`ALGORITHM_VERSION` v5; remainder в `_build_unit_counters`; отчёты переименованы). Не делать фикс повторно.

Дата: **2026-09-03**. Исходный чат — анализ + handoff; реализация — следующий чат.

Правила: [AI_rfp_parts.mdc](.cursor/rules/AI_rfp_parts.mdc), [AI_rfp_step4_excel.mdc](.cursor/rules/AI_rfp_step4_excel.mdc), [AI_project_standards.mdc](.cursor/rules/AI_project_standards.mdc) (`[C2]` / `[C2+V]` / `[PREMIUM]`).

## Зачем

Инцидент **2265-KSB / BCC0000642 / ДС67**: в части RFP лот **3 + 1 тег** и **9 без тега** (=12). УЛ и МТО тоже 12. В `rfp_parts_net.xlsx` и Step4 по коду осталось **10**. Две штуки УЛ ушли в leftover **«Только в УЛ» (`RFP − УЛ = −2`)**, две МТО — в **«Только в МТО (Недопоставка)»**.

Причина не в посадке УЛ, а в `_build_unit_counters`: строка **с тегами** даёт **+1 на тег**, остаток лота **выбрасывается** (есть только WARN / строка в `rfp_parts_duplicate_tags.xlsx`).

Цель следующего чата:

1. В своде: **лот > тегов** → не терять qty: **по 1 на каждый тег + остаток на безтеговый ключ** (пример: 3 → 1 тегированная + 2 безтеговые).
2. Сверить, как то же раскрывается в **MTO** и **УЛ**; не ломать контуры без нужды.
3. Для **тегов > qty**: оставить **полное раскрытие по тегам**; явно указать, **в каком xlsx** смотреть проблемные места (и при необходимости чуть улучшить отчёт).

## Инцидент (факты, не трогать как «баг посадки»)

Источники (прогон `_результат_проверки_2026.09.02.13.49`, тот же net `RFP сводный файл/2026.09.02_13.49/`):

| Контур | Файл / место | Qty |
|--------|----------------|-----|
| Часть RFP | `RFP_Зиновьев/ДС67. AGCC.287-0000-12.4.1-RFP-0032_0_RU.xlsx` стр. **142** / **150** | 9 без тега + **3** с `2265-S-FV-0711` |
| Свод | `rfp_parts_net.xlsx` поз. **25207** / **25206** | **9** + **1** (trace ещё `qty=3.000`) |
| УЛ | `согл УЛ ДС67/Packing list PL_2076961.32_3705.01 - ред.220526.xlsx` Single 1 стр. **112** / **113** | 9 без тега + 3 с тем же тегом |
| MTO | `2265/AGCC.287-2265-KSB.MTO-0001_01-AN02_RU.xlsx` | 4 строки × 3 = 12 |
| Step4 | `Шаг4_Сопоставление_RFP_MTO_20260902_140222.xlsx` | RFP 10; leftover УЛ 2; MTO-only 2 |

Книга сумм `Шаг4_Проверка_сумм_VALUES_*.xlsx` по **2265-KSB** = OK: снимок и результат уже считают усечённые 10.

Сопутствующее (на qty не влияет): дубль тега `2265-S-FV-0711` на **ДС26_72 / BCC0002772** (УЗИП) и **ДС67 / BCC0000642** (муфта). В `rfp_parts_duplicate_tags.xlsx` тип «Дубль тега».

Папка `…10.39` на момент разбора не содержала итоговый Step4 (после VALUES идёт УЛ + запись Excel). Поведение по коду то же, пока net не пересобран.

Диагностика (read-only, можно не коммитить): `tmp/diag_2265_ksb_bcc0000642.py`, `tmp/diag_2265_sources.py`, `tmp/diag_2265_rfp_lot_cols.py`, `tmp/diag_2265_find_2pcs.py`, `tmp/diag_2265_ds67_all.py`, `tmp/diag_2265_yesterday_net.py`.

## Целевая семантика свода (предложение для плана)

Ключ суммирования **не менять**: `ds_name + title + code + tag` / `<NO_TAG>` (`AGGREGATION_KEY`, `_unit_key`).

Для **целого** лота и непустого списка тегов:

| Случай | Сейчас | Нужно |
|--------|--------|--------|
| `tags == lot` | +1 на тег | без изменения |
| `lot > tags` (инцидент) | +1 на тег, **остаток 0** | +1 на тег, затем **`lot - tags`** на ключ `_unit_key(record, "")` (`NO_TAG_KEY`) |
| `tags > lot` | +1 на **каждый** тег (net qty > лота) | **оставить полное раскрытие по тегам**; не добавлять отрицательный остаток |
| дробный лот + теги | сейчас тоже только +1/тег, дробь теряется | **не в этом шаге**, если нет явного кейса; зафиксировать в правиле как открытый край |

Эталон поведения — очередь УЛ в `_build_unit_queues`: целое qty раскладывается в слоты `qty=1`, тег берётся по индексу, **лишние слоты без тега**; `tags > quantity` → **не** класть в очередь.

После фикса инцидент: 1 тегированная единица + 2 безтеговые. Безтеговые **сольются** с соседней строкой того же ДС/титула/кода (девятка) → в net **1 + 11**, не «9 и отдельная 2». Для Step4 это правильно (тот же collapse-ключ без тега).

Обязательно **bump `ALGORITHM_VERSION`** (`v4` → `v5` в `RFQ/units_convert/models.py`): preflight пересоберёт все net. Менять строку `AGGREGATION_KEY` не нужно, если состав ключа тот же.

## Как раскрывается сейчас в других контурах

Не путать три оси: **свод частей**, **split Step1/MTO/VO**, **очередь УЛ**.

### RFP части → net — [analyze_rfp_parts.py](RFQ/rfp_parts/analyze_rfp_parts.py)

`_build_unit_counters` (tagged branch ~1408–1428): `continue` сразу после `+1` на тег.  
`_append_summary_unit_rows`: тегированный ключ → N строк `VALUES=1`; `<NO_TAG>` → одна строка на всю сумму.

### RFP Step1 split — [step1_load_rfp.py](RFQ/tags_rfp_compare/step1_load_rfp.py) `split_rfp_rows`

- Несколько тегов: по строке на тег, `VALUES=1` пока есть qty, дальше **`VALUES=0`** (полное раскрытие тегов).
- **Один тег + VALUES>1: не сплитуется**, строка живёт как есть.
- Без тегов + VALUES>1: сплит на единицы (если не ban).

После фикса net тегированная строка уже будет `VALUES=1`, split её не тронет. As-build / `use_latest_net=false` (сырой RFP с 3+1 тег) — **другой контур**; выравнивать `split_rfp_rows` под remainder **не смешивать** с этим шагом, только пометить follow-up.

Проверка тегов≠VALUES (не меняет qty): `_check_rfp_tags_and_values` → Excel.

### MTO — [step4_1_check_mto_data.py](RFQ/tags_rfp_compare/step4/step4_1_check_mto_data.py) `check_mto_data`

- Без тегов: split по VALUES (`batch_copy_light`).
- Теги есть: mismatch **всегда логируется**, если `tags_count != values_int`.
- Split по тегам **только если `tags_count > 1`**: как RFP, extra tags → `VALUES=0`, leftover qty → warning в консоль («Осталось N нераспределенных VALUES»), **строка остатка не создаётся**.
- **Один тег + VALUES=3: строка не сплитуется** (как Step1).

VO (`check_vo_data`): split нескольких тегов; отдельного Excel «теги≠VALUES» у VO нет (дубли тегов VO — свой файл).

Для инцидента после фикса net: RFP 12 vs MTO 12, MTO-only 2 уйдут. Менять MTO-split в этом шаге **не требуется**, если ревью подтвердит.

### УЛ — [step4_packing_compare.py](RFQ/tags_rfp_compare/step4/step4_packing_compare.py) `_build_unit_queues`

Целое qty: слот `qty=1` на индекс; `tag = tags[i] if i < len(tags) else ""`. Это **уже** «1 тег + 2 пустых слота» для qty=3. Дробный хвост — отдельный слот без тега.

`len(tags) > quantity` → issue `packing_tags_exceed_quantity`, строка **не в очередь**. В Step4 это **fatal** (`RfpPackingFatalError`), итоговый match Excel **не пишется**. Loader ТСД ([tsd_packing_load.py](RFQ/ds_compare/tsd_packing_load.py) ~814) пишет тот же код в issues кэша, Excel УЛ не валит. Grouped DS matcher ([ds_packing_grouped_compare.py](RFQ/ds_compare/ds_packing_grouped_compare.py)) при tags>qty продолжает с quality partial.

**Не менять УЛ**, если только свод догоняет эту модель для `lot > tags`.

## Где смотреть «теги ≠ qty» (xlsx)

Фильтр «тегов больше, чем qty» = колонка тегов > колонка qty (имена ниже).

| Контур | Файл | Лист / тип | Колонки | Меняет qty? |
|--------|------|------------|---------|-------------|
| **Сбор частей** | `RFP сводный файл/YYYY.MM.DD_HH.MM/rfp_parts_duplicate_tags.xlsx` | типы **«Число тегов ≠ лот»** и «Дубль тега» | **Кол-во лота**, **Число тегов**, файл, лист, строка, код, титул | нет (сейчас); после фикса mismatch всё равно писать |
| Части, текст | `rfp_parts_diagnostics.xlsx` / txt WARN | `количество тегов (N) не совпадает с количеством лота (M)` | — | нет |
| GUI частей | таблица файлов | дубли и TAG≠лот **не** показываются; ссылка на xlsx над таблицей | — | — |
| **Step1 RFP** (после net) | `Шаг1_RFP_несовпадение_количества*.xlsx` | строки net, где теги ≠ VALUES | теги / VALUES | нет |
| Step1 дубли тегов | `Шаг1_RFP_дублирующиеся_теги*.xlsx` | кросс-ДС дубли (как FV-0711) | — | нет |
| **MTO** | `Шаг4_Несовпадения_тегов_VALUES_*.xlsx` | «Несовпадения тегов и VALUES» | **Кол-во тегов**, **VALUES**, CODE, title_system, список тегов | нет; extra tags всё равно раскрываются с 0 |
| VO | `Шаг4_Дублирующиеся_теги_VO_*.xlsx` | только дубли, не count mismatch | — | — |
| **УЛ Step4** | при tags>qty Excel match **нет**; `ul_compare_report_*.txt` + issues | `packing_tags_exceed_quantity` | — | fatal |
| УЛ загрузка ТСД | `…/УЛ сводный файл/tsd_packing_critical.txt` и cache issues | тот же код | — | partial, свод УЛ пишется |
| Units gate | `дробные значения после конвертации {RFP\|MTO\|УЛ}.xlsx` | не про count tags vs lot | `Теги` = tags_count | — |

Для **tags > lot** в частях уже есть перечень: фильтр типа **«Число тегов ≠ лот»** и условие `Число тегов > Кол-во лота`. Имеет смысл **[C2]** разнести kind на два ярлыка (`Лот > тегов` / `Тегов > лота`), чтобы не фильтровать вручную — опционально, не блокер.

## План шагов (для следующего чата)

Маркировка по [AI_project_standards.mdc](.cursor/rules/AI_project_standards.mdc). Одна сессия ≈ до ~10 шагов.

### 1. [PREMIUM] Зафиксировать контракт (короткая спека в ответе или в этом файле)

- Подтвердить: remainder только для **целого** `lot > tags`; ключ остатка = `<NO_TAG>` (слияние с соседней безтеговой строкой того же ДС/титула/кода).
- `tags > lot`: по-прежнему +1/тег; отчёт не терять.
- MTO/УЛ/Step1 split **не менять** в этом PR, только сверка + абзац в правиле.
- Bump **`ALGORITHM_VERSION` v4→v5** (forced rebuild net).
- Край: отрицательный лот уже ERROR; дробь+теги — out of scope.

Результат шага = 10–20 строк спеки, чтобы [C2] не импровизировал.

### 2. [C2] `_build_unit_counters`

Файл: [analyze_rfp_parts.py](RFQ/rfp_parts/analyze_rfp_parts.py).

После цикла `for tag in tags: unit_counter[key] += 1` для целого `record.values`:

- `extra = int(record.values) - len(tags)`
- если `extra > 0`: `unit_counter[_unit_key(record, "")] += extra`, examples/trace merge на этот ключ (как у безтеговой ветки).
- `continue` как сейчас.

Не добавлять extra, если `values` не целое. WARN mismatch оставить (оба направления).

### 3. [C2] `ALGORITHM_VERSION = "v5"`

[models.py](RFQ/units_convert/models.py). Preflight: [AI_rfp_parts.mdc](.cursor/rules/AI_rfp_parts.mdc) — stored v1–v4 → rebuild. Проверить смоуки, которые хардкодят `v4`.

### 4. [C2] Тесты

Новые кейсы (предпочтительно расширить [tmp/test_rfp_parts_unit_conversion_smoke.py](tmp/test_rfp_parts_unit_conversion_smoke.py) или отдельный `tmp/test_rfp_parts_tag_lot_remainder_smoke.py`):

1. 1 тег, lot 3, тот же ds/title/code + безтеговая 9 → units: tagged key **1**, `NO_TAG` **11**.
2. 2 тега, lot 2 → только два tagged, extra 0.
3. 3 тега, lot 1 → три tagged (**tags > lot**), extra нет; WARN/remark.
4. Регрессия конвертации: существующие 8 кейсов unit conversion smoke.

`_build_summary_rows` / net row shape: 1 строка VALUES=1 с тегом + 1 строка VALUES=11 без тега (не 11 строк, безтеговый путь пишет одну сумму).

### 5. [C2] Отчёт `rfp_parts_duplicate_tags.xlsx` (желательно)

`_TAG_REMARK_KIND_LABEL` / `_collect_tag_remarks`: два kind вместо одного `count_mismatch`, колонки те же. GUI по-прежнему не показывает эти строки.

### 6. [C2] Правила

Обновить [AI_rfp_parts.mdc](.cursor/rules/AI_rfp_parts.mdc): семантика tagged counters; v5; таблица отчётов tags≠lot; ссылка на инцидент 2265 кратко. Не раздувать [AI_memory.mdc](.cursor/rules/AI_memory.mdc).

### 7. [C2+V] Verification

```powershell
$env:PYTHONUTF8='1'
python tmp/test_rfp_parts_unit_conversion_smoke.py
python tmp/test_rfp_parts_net_preflight_smoke.py
python tmp/test_rfp_parts_tag_lot_remainder_smoke.py
```

Ручная сверка после пересборки net (preflight Запуска или `python -X utf8 -m RFQ.rfp_parts`): в свежем `rfp_parts_net.xlsx` ДС67 / 2265-KSB / BCC0000642 сумма **12** (1 с тегом + 11 без). Повторный Step4: leftover УЛ по этому коду **0**, MTO-only 2 по этому коду **нет**. Не гонять полный Запуск, пока смоуки зелёные, если нет доступа к UNC.

## Риски `[RISK:]`

- **Слияние remainder с соседней безтеговой строкой** — ожидаемо; в Excel пользователь увидит 11, не 9+2. Написать в комментарии net/правиле.
- **Дубль одного тега на двух ДС** — ключ включает `ds_name`, remainder ДС67 не уезжает на ДС26.
- **Forced rebuild net (v5)** — первый Запуск после выкладки пересоберёт свод (долго на UNC). Это нужно.
- **Не чинить MTO leftover-qty** в том же диффе: у MTO extra tags → VALUES=0, extra qty только warning. Иначе расползётся match.
- **Не ослаблять fatal УЛ** `packing_tags_exceed_quantity`.
- Минимальный дифф: не рефакторить `_append_summary_unit_rows`.

## Что не делать в том чате

- Правки посадки Step4 / `split_rfp_rows` / `check_mto_data`, пока свод не проверен.
- Менять чтение лота (колонки 16–17).
- «Восстанавливать» 2 шт. в Step4 в обход net.

## Точки входа в код

| Что | Где |
|-----|-----|
| Баг/фикс | `_build_unit_counters` в [analyze_rfp_parts.py](RFQ/rfp_parts/analyze_rfp_parts.py) ~1392 |
| Запись net-строк | `_append_summary_unit_rows` ~1990 |
| Отчёт TAG≠лот | `_collect_tag_remarks`, `_write_duplicate_tags_xlsx` |
| Версия свода | `ALGORITHM_VERSION` в [models.py](RFQ/units_convert/models.py); snapshot в `_write_build_deps` |
| Preflight rebuild | [parts_net_preflight.py](RFQ/rfp_parts/parts_net_preflight.py) |
| Эталон УЛ | `_build_unit_queues` в [step4_packing_compare.py](RFQ/tags_rfp_compare/step4/step4_packing_compare.py) ~569–624 |
| Отчёт MTO | `_save_count_mismatches_to_excel` → `Шаг4_Несовпадения_тегов_VALUES_*.xlsx` |
