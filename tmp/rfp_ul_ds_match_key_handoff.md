# Задание: ключ посадки УЛ + фактический ДС в Step4 Excel

> Для нового чата. Не дублировать покрытие имён (`ds_id_check`) — оно уже в проде.
> Дата: 2026-09-02.

## Цель

1. **Посадка УЛ (Step4)** — жёсткий ключ **`(actual_ds, title, system, code)`**.
2. Строки **без совпадения по фактическому ДС не сажаются** (нет fallback на ключ без ДС).
3. В финальном Excel: **новый столбец фактического ДС**; текущий **`DS_NAME` («Имя ДС»)** остаётся **порядковым** идентификатором.

Не делать в этом задании: grouped DS↔УЛ, overlay ГЭМ/ЗИП, смену четырёх глобальных проходов тегов, `parse_ds_name_from_file_name`.

## Зачем

Сейчас очередь УЛ — `(title, system, code)` ([AI_rfq_context.mdc](.cursor/rules/AI_rfq_context.mdc) §11, [AI_rfp_step4_excel.mdc](.cursor/rules/AI_rfp_step4_excel.mdc)). Имя ДС и папка ТСД **не** в ключе: ДС29 и ДС82 с одним `8445-SOT`+кодом делят очередь; слот одной ДС может уйти другой (пример в step4-правиле: `…-FV-0001` / `…-FV-0011`). Пользователь это подтвердил письмом коллегам; coverage-парсер уже живой, посадка ещё нет.

Живой прогон покрытия 2026-09-02: пары по **фактическому** номеру ок (в т.ч. `ДС92_24Б` ↔ `согл УЛ ДС24`, `ДС4905`, `ДС4`/`ДС6` после переименования). Вне ключа посадки: ГФ (2 папки), RFP-only ДС47/95/101.

## Контракт номеров ДС (уже есть)

Парсер: [ds_identity.py](RFQ/rfp_parts/ds_identity.py) — `parse_rfp_ds_identity`, `parse_ul_folder_ds_identity`, `DsIdentity.actual` / `.sequential` / `.label`.

| Источник | Пример | sequential | actual |
|----------|--------|------------|--------|
| RFP обычный | `ДС23. …` / `DS_NAME=ДС23` | 23 | 23 |
| RFP корректировка | `ДС92_24Б. …` / `DS_NAME=ДС92_24Б` | 92 | **24** |
| Папка УЛ | `согл УЛ ДС24` | 24 | **24** |
| Копия файла | `ДС4905_1.xlsx` | 4905 | 4905 (не compound) |
| ГФ с `ДСn` | `согл УЛ ДС1 ГФ 5титулов`, `согл УЛ ДС7 ГФ 2 тит` | 1 / 7 | **есть** (`ДСn` важнее пометки ГФ) |
| ГФ без номера | `согл УЛ ГФ 5титулов` | — | **нет** (`kind=gf`) |

**Join посадки только по `actual` (int).** Не использовать `match_keys` покрытия для составных папок УЛ (`ДС4_11` → `[4, 11]`) — это было только для таблицы имён.

Legacy [parse_ds_name_from_file_name](RFQ/rfp_parts/ds_checklist.py) по-прежнему отдаёт префикс `ДС92_24Б` в `DS_NAME` свода частей. **Не заменять.** Для actual: `parse_rfp_ds_identity(ds_name).actual`.

### Не путать колонки

| Поле | Заголовок сейчас | Смысл |
|------|------------------|--------|
| **`DS_NAME`** | «Имя ДС» | порядковый ярлык файла (`ДС23`, `ДС92_24Б`) — **оставить** |
| **`DS_NUMBER`** | «№ позиции» | номер строки в спецификации (1, 2, 3…) — **не трогать** |
| **новое** | «Фактический ДС» | `ДС24` / `24` из `DsIdentity.actual` |

## Жёсткий ключ (без fallback по ДС)

Новый ключ очереди / snapshot / leftover:

`(normalize(actual_ds), title, system, code)`

`actual_ds` — нормализованное число или каноническая метка `ДС24`; главное — **одинаково** на RFP и УЛ.

**Запрещено в этой итерации:**

- сажать УЛ, если у RFP нет `actual`, или у слота УЛ нет `actual`;
- если actual разный (29 vs 82) — **разные очереди**, даже при том же title+code;
- fallback «как раньше, без ДС», если actual не распарсился;
- матчить порядковый 92 с папкой `ДС92`, если actual RFP = 24.

**Оставить как есть (это не fallback по ДС):**

- четыре глобальных прохода тегов **внутри уже суженной очереди** (pass 1–4);
- `_candidate_key`: пустой RFP `CODE` → `CODE_MTO` / `CODE_VO` (H1) — по-прежнему про **код**, не про ДС. После смены ключа этот перебор должен искать ключ **с тем же actual**.

Поведение при отсутствии actual:

- RFP без actual → УЛ не сажается (статус присутствия как сейчас при пустой очереди: «Неотгружено…» / «Только в RFP…» и т.д., **не** «Проблема данных УЛ» только из-за пустого ДС — уточнить в реализации; допустим отдельный comment «нет фактического ДС»).
- УЛ без actual (`kind=gf` без `ДСn`, unparsed) → leftover **«Только в УЛ»**, в RFP не идут.
  Папка `согл УЛ ДС1 ГФ …` **имеет** actual=1.

## Откуда брать actual

**RFP:** `parse_rfp_ds_identity(row.DS_NAME)` (значение уже на строке из частей / net). Пустой `DS_NAME` (MTO-only / VO-only без ярлыка) → нет actual → нет посадки.

**УЛ:** папка первого уровня из `ANNOTATION` (относительный путь файла). Кэш v6 уже хранит `ANNOTATION` ([PACKING_CACHE_COLUMNS](RFQ/packing_list_provider.py)). Парсить **в момент посадки**, без bump `PACKING_CACHE_VERSION`, без обязательного «Прочитать ТСД». Хелпер: parent folder → `parse_ul_folder_ds_identity`.

Не писать actual УЛ в pickle, пока не понадобится свод xlsx.

## Карта файлов (менять)

| Файл | Что |
|------|-----|
| [step4_packing_compare.py](RFQ/tags_rfp_compare/step4/step4_packing_compare.py) | `build_ordered_snapshot`, `_build_unit_queues`, `_candidate_key`, leftover, `compare_rfp_rows_with_packing`; ключ 4-tuple |
| [packing_list_provider.py](RFQ/packing_list_provider.py) | **не ломать** `packing_match_key(title, system, code)` — им пользуется grouped DS. Новый хелпер Step4, напр. `rfp_packing_match_key(actual, title, system, code)` |
| [ds_packing_grouped_compare.py](RFQ/ds_compare/ds_packing_grouped_compare.py) | **вне скоупа**, пока 3-tuple API жив |
| [tables_columns.py](base/tables_columns.py) | константа + слот `ColNames.column_list` |
| [step4_6_save_match_result_to_excel.py](RFQ/tags_rfp_compare/step4/step4_6_save_match_result_to_excel.py) | `ColumnDef`: фактический рядом с `DS_NAME`; header `DS_NAME` → «Порядковый ДС» (или оставить «Имя ДС», если так читаемее — **видимое имя согласовать**, смысл порядковый) |
| [base_classes.py](base/base_classes.py) | `_COPY_COLS` / `_ALWAYS_COPY_COLS` **только если** новое поле пишут после split. Если заполняют на каждой result-строке в packing (после collapse) — в `_COPY_COLS` |
| [step4_6_cell_colors.py](RFQ/tags_rfp_compare/step4/step4_6_cell_colors.py) | если красят блок RFP identity |
| [step4_6_quality_sheets.py](RFQ/tags_rfp_compare/step4/step4_6_quality_sheets.py) | если листы Сводка/Посадка показывают имя ДС |
| [tmp/test_rfp_packing_compare_smoke.py](tmp/test_rfp_packing_compare_smoke.py) | обязательный регресс |

Чеклист нового столбца: [AI_rfp_step4_excel.mdc](.cursor/rules/AI_rfp_step4_excel.mdc) «Новый / скрытый столбец».

`build_ordered_snapshot` сегодня суммирует qty по 3-tuple — **обязательно** включить actual, иначе две ДС с одним кодом снова делят `UL_ORDERED_VALUES`.

## Excel (раскладка)

Сейчас в [OUTPUT_COLUMNS_CONFIG](RFQ/tags_rfp_compare/step4/step4_6_save_match_result_to_excel.py):

`DS_NAME` («Имя ДС», collapsed) → `DS_LOT` hidden → `DS_NUMBER` («№ позиции») → `DS_MANAGER`.

Предлагаемый порядок:

**Порядковый ДС** (`DS_NAME`) → **Фактический ДС** (новое) → № позиции → Менеджер ДС.

Значения: порядковый = как сейчас (`ДС92_24Б`); фактический = `ДС{actual}` (как в coverage). Не класть actual в `DS_NUMBER`.

Матрица BCC / `DS_MANAGER` / collapse: ключ collapse — `(DS_TITLE, DS_NUMBER, CODE, …)`, не `DS_NAME`. Не ломать, если actual только display+packing.

## Смоук (минимум)

В [test_rfp_packing_compare_smoke.py](tmp/test_rfp_packing_compare_smoke.py):

1. Два RFP: `DS_NAME=ДС29` и `ДС82`, один `8445-SOT` + один код; УЛ только из папки `…/согл УЛ ДС82/file.xlsx` (`ANNOTATION`). Слот **не** должен сесть на ДС29; ДС82 получает поставку; leftover не отдаётся ДС29.
2. Корректировка: RFP `DS_NAME=ДС92_24Б`, папка `согл УЛ ДС24` — **матч**; папка `согл УЛ ДС92` — **не матч**.
3. ГФ-папка без actual — слоты не сажаются на RFP с ДС1/2/7.
4. Регресс: четыре прохода тегов **внутри одной ДС** (сосед не крадёт RFP-тег) — существующие тесты очереди не сломать.
5. Excel: оба столбца на месте, заголовки.

Дополнительно: `python tmp/test_ds_identity_smoke.py` (парсер не регрессировать).

## Правила обновить после кода

- [AI_rfp_step4_excel.mdc](.cursor/rules/AI_rfp_step4_excel.mdc) — ключ `(actual, title, system, code)`; столбцы; убрать «имя ДС не в ключе».
- [AI_rfq_context.mdc](.cursor/rules/AI_rfq_context.mdc) §11 — то же; ловушка 8445 закрыта ключом ДС.
- [AI_ds_compare.mdc](.cursor/rules/AI_ds_compare.mdc) — одна строка: grouped matcher **пока** 3-tuple.
- Ссылки только относительные (`RFQ/...`), без `C:\`.

## Стиль / скоуп

Google Python, `from __future__ import annotations`, минимальный дифф. Маркировка: логика ключа **[PREMIUM]** / реализация по спеке **[C2]** — [AI_project_standards.mdc](.cursor/rules/AI_project_standards.mdc).

Не менять: `units_gate`, TM-loop RFP↔MTO, `ds_id_check`, GUI вкладки покрытия.

## Критерий готовности

- Прогон смоуков выше зелёный.
- На одинаковом title+code две ДС не делят очередь УЛ.
- В Шаг4 видны порядковый и фактический ДС.
- `python -m RFQ.rfp_parts.ds_id_coverage` / вкладка **RFP · ДС ↔ УЛ** не обязаны меняться.
