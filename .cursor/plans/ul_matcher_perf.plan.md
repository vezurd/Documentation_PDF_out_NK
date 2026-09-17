# План: производительность шага «Сопоставление с УЛ и audit» (milestone `packing_lists`)

> Создан 09.09.2026. Область: `RFQ/tags_rfp_compare/step4/step4_packing_compare.py` + тонкая обвязка в
> `RFQ/tags_rfp_compare/step4_analyze_and_match.py`.
> Правила маркировки шагов и промптов: `.cursor/rules/AI_project_standards.mdc`.
> Контекст матчера УЛ (статусы, ключи, ловушки): `.cursor/rules/AI_rfp_step4_excel.mdc`.
> Итоги замеров писать в `.cursor/rules/AI_optimization_roadmap.mdc` и `.cursor/rules/AI_performance.mdc`.

## Цель

Убрать квадратичные участки в `compare_rfp_rows_with_packing` **без единого изменения результата**.
Критерий приёмки по данным — жёсткий: итоговый `Шаг4_Сопоставление_RFP_MTO_*.xlsx` должен быть
идентичен baseline по мультимножеству строк (`RFQ/tags_rfp_compare/compare_step4_files.py`:
`only_old=0`, `only_new=0`), а `ul_compare_report_*.txt` — совпадать по всем COUNTERS.

## Baseline (измерено 09.09.2026, read-only диагностика)

Скрипты: `tmp/diag_ul_matcher_cost.py`, `tmp/diag_ul_matcher_hot.py`, `tmp/diag_ul_matcher_gc.py`.

| Факт | Значение |
|---|---|
| строк УЛ в кэше | 30 023 |
| поштучных слотов после `_build_unit_queues` | **1 255 135** |
| различных ключей A `(actual, title, system, code)` | 14 691 |
| топ длин очередей | 23 532 / 22 000 / 20 908 / 19 250 / 13 519 / 12 758 |
| `sum(len(queue)^2)` по всем ключам | 4,00·10⁹ |
| `_allocate_queue_phase(blind)` на очереди 22 000 в одну строку | **49,67 с** |
| `_delivered_qty` на 22 000 слотов, один вызов | 2,04 мс |
| `_merge_packing_units_trace` на 20 000 слотов | 3,62 с (≈181 мкс/слот) |
| `RowStd.get_std_check_row({}, TableComments())` | 53 мкс (146 колонок → 146 `CheckElement`) |
| один проход по 20 000 siblings в `_reserved_tags_for_state` | 0,87 мс (44 нс/итерация) |
| `_row_leftover_code_key` (regex + нормализация) | 1,80 мкс |
| вклад GC | посадка 0%, trace-merge ≈35% |

## Baseline по фазам (прогон 09.09.2026 12:57, шаг 1.3 выполнен)

Папка: `_РЕЗУЛЬТАТА_ПРОВЕРКИ/_результат_проверки_2026.09.09.12.57`.
Источник цифр — секция `PHASES (seconds):` в `ul_compare_report_2026.09.09_13.19.30.txt`
(`timing_log.xlsx` не писался, `step4.timing_log=false` — отчёт оказался надёжнее канала логов).

**Baseline xlsx для сверки идентичности в шаге 5.1:**
`Шаг4_Сопоставление_RFP_MTO_20260909_131931.xlsx` (6 950 256 B).

| Фаза | Секунд | Доля | Чем лечится |
|---|---:|---:|---|
| `a2_untagged` | **679,4** | 67,4% | сессия 2 |
| `fill_rows` | **162,9** | 16,2% | сессия 4 |
| `a3_mto_tags` | **108,2** | 10,7% | сессия 3 (блок A) |
| `a4_blind` | **49,3** | 4,9% | сессия 3 (блок A) + сессия 2 |
| `build_unit_queues` | 4,0 | 0,4% | — |
| `a_states_build` | 2,3 | 0,2% | — |
| `prepare_rows` | 1,2 | 0,1% | — |
| `block_b` | **0,8** | 0,1% | — |
| `leftover_dump` | 0,4 | — | — |
| `a1_rfp_tags` / `block_c` / `presence_states` / `drain_leftover` / `invalid_key_rows` | < 0,1 каждая | — | — |
| **итого шаг УЛ** | **≈ 1008 с (16,8 мин)** | | |

Ключевые COUNTERS прогона (сверять с ними в шаге 5.1): `result_position_rows: 36879`,
`packing_rows: 30017`, `packing_units: 1234072`, `allocated_units: 1229400`,
`matched_rows: 31369`, `not_in_packing: 5434`, `packing_only_added: 391`, `complete: 31061`,
`shortfall: 167`, `overdelivery: 391`, `unit_mismatches: 0`, `invalid_key_rows: 0`,
`queue_keys: 14691`, `max_queue_len: 23532`, `ISSUES (4)`.

### Расхождения замера с предварительными оценками

1. **`block_b` / `block_c` — 0,8 с и 0,07 с, а не «от 90 с и выше».** Оценка была построена на
   квадрате числа leftover-участников, но их реально мало (`packing_only_added: 391`).
   Квадратичность `_reserved_leftover_tags` с regex внутри существует, но на текущих данных
   не проявляется. **Вывод:** часть сессии 3 про блоки B/C — гигиена на будущее, а не
   оптимизация; приоритет только у блока A.
2. **`a2_untagged` 679 с против расчётных ~370 с.** Модель `sum(len^2) × 92,7 нс` считала только
   двойной пересчёт `remaining_qty()`. Недостающие ~300 с — либо промахи кэша на реальных
   `allocated` (слоты разбросаны по памяти, в синтетике лежали подряд), либо скан и
   `deque.rotate` по теговому префиксу очереди: `_build_unit_queues` кладёт теговые слоты
   первыми (`tags[index] if index < len(tags) else ""`), поэтому в фазе `untagged`
   каждый забор платит O(число тегов) на `rotate` туда и обратно.
   **[RISK: сессия 2 гарантированно снимает только пересчёт `remaining_qty()`]** — после неё
   обязательно перезамерить `a2_untagged` и решить, нужен ли отказ от поштучных слотов
   (см. «Что не входит в этот план»).
3. Запись файлов внутри milestone подтверждена как пренебрежимая: один txt на 4,8 КБ.

## Целевые показатели

| Метрика | Baseline | Цель |
|---|---|---|
| `step4::compare_packing_lists` | ~10–13 мин (оценка) | **< 90 с** |
| результат Excel | — | идентичен (multiset) |
| `ul_compare_report_*.txt` COUNTERS | — | идентичны |

---

## Сессия 1 — измеримость (пункт 5)

Делается первой: без пофазных таймингов эффект последующих шагов нечем подтвердить.

### Шаг 1.1 [PREMIUM] Дизайн канала таймингов

**Исполнитель:** дорогая модель (куратор). **Результат:** спека для 1.2, зафиксирована ниже.

Ограничение: `compare_rfp_rows_with_packing` не имеет `result_dir` и не должна его получать
(матчер остаётся чистой функцией над строками и dataset). Тянуть `append_timing_log` внутрь
`step4_packing_compare.py` нельзя — модуль не знает про `result_dir` и вызывается из смоуков без него.

**Решение:** тайминги возвращаются наружу в `RfpPackingAudit`, логирует их вызывающий.

1. В `RfpPackingAudit` добавить поле `phase_timings: dict[str, float] = field(default_factory=dict)`.
   **Не** кладём в `RfpPackingStats`: `save_rfp_packing_report` печатает `vars(audit.stats)` как
   COUNTERS, и секунды туда попадать не должны (иначе сломается сверка COUNTERS между прогонами).
2. Внутри `compare_rfp_rows_with_packing` — локальный контекст-менеджер или хелпер
   `_phase(name)`, пишущий в `audit.phase_timings[name] += elapsed`.
3. Фазы (имена — точно эти, чтобы отчёты разных прогонов сравнивались):
   `prepare_rows`, `build_unit_queues`, `a_states_build`, `a1_rfp_tags`, `a2_untagged`,
   `a3_mto_tags`, `a4_blind`, `drain_leftover`, `block_b`, `block_c`, `presence_states`,
   `fill_rows`, `leftover_dump`, `invalid_key_rows`.
4. Диагностические счётчики формы очередей — два **int**-поля в конец `RfpPackingStats`:
   `queue_keys`, `max_queue_len`. Float-поля в `RfpPackingStats` не добавлять никогда.
   Это единственное допустимое изменение состава COUNTERS во всём плане: baseline снимается
   шагом 1.3, то есть уже после сессии 1, поэтому сверка COUNTERS в шаге 5.1 идёт между двумя
   версиями с этими полями. Существующие поля не переименовывать и не переупорядочивать —
   `vars()` сохраняет порядок объявления, новые строки должны просто добавиться в конец.
   Сессии 2–4 состав COUNTERS не меняют вообще.
5. В `save_rfp_packing_report` — отдельная секция `PHASES (seconds):` после `COUNTERS:`.
6. В `step4_analyze_and_match.py`, в блоке `include_packing_lists`, сразу после успешного
   `compare_rfp_rows_with_packing`: `for name, sec in packing_audit.phase_timings.items():
   append_timing_log(result_dir, f"step4::ul::{name}: {sec:.3f}s")`. Существующий
   `log_timing("compare_packing_lists", step_start)` остаётся как есть.
   На пути `except RfpPackingFatalError` тайминги не логируем (там уже пишется фатальный отчёт).

### Шаг 1.2 [C2] Реализация таймингов

**Файлы:** `RFQ/tags_rfp_compare/step4/step4_packing_compare.py`,
`RFQ/tags_rfp_compare/step4_analyze_and_match.py`.

**Verification:**

```powershell
$env:PYTHONUTF8='1'
python tmp/test_rfp_packing_compare_smoke.py
python tmp/test_step4_units_gate_order_smoke.py
python tmp/test_rfp_supply_status_smoke.py
```

Готово, когда: смоуки зелёные; `rg "phase_timings" RFQ | rg -c .` > 0; в
`save_rfp_packing_report` секция `COUNTERS:` содержит ровно те же имена, что до правки
(сверить `git diff` по функции глазами — только добавленная секция PHASES).

### Шаг 1.3 [C2+V] Baseline-прогон

Запуск «Сравнить RFP ↔ MTO ↔ РКД ↔ УЛ» с `step4.timing_log=true`.
Зафиксировать: `step4::ul::*`, `step4::compare_packing_lists`, `step4_total`, `total_runtime`,
плюс сохранить копию итогового xlsx как baseline для `compare_step4_files.py`.
**[RISK: прогон ~20 мин и пишет в боевую сетевую папку результатов]** — согласовать окно с
пользователем, не запускать параллельно со вторым прогоном.

---

## Сессия 2 — инкрементальный `delivered` (пункт 1)

Основной выигрыш. Ожидание: минус до ~370 с.

### Шаг 2.1 [PREMIUM] Спека (готова, реализовывать по ней буквально)

**Обоснование побитовой точности.** `_delivered_qty(units)` считает
`float(sum(_unit_qty(u) for u in units))`, то есть сложение слева направо в порядке списка.
Слоты добавляются в `allocated` в том же порядке, поэтому running-total, накапливаемый при каждом
append начиная с `0.0`, даёт **тот же самый** IEEE-754 результат (в т.ч. для дробных остатков
после `_extract_unit_at`). Значит ни одно сравнение с `_EPS` не может изменить решение.

**Правки в `step4_packing_compare.py`:**

1. В `_RowAllocationState` добавить поле `delivered: float = 0.0` и метод:

```python
    def take(self, unit: _PackingUnit) -> None:
        """Append an allocated unit and keep the running delivered total in sync."""
        self.allocated.append(unit)
        self.delivered += _unit_qty(unit)

    def remaining_qty(self) -> float:
        return self.ordered - self.delivered
```

2. Заменить **все четыре** `state.allocated.append(taken)` в `_allocate_queue_phase`
   (фазы `rfp_tags`, `untagged`, `mto_tags`, `blind`) на `state.take(taken)`.
3. **Не трогать** `_delivered_qty` и его вызовы в `_allocation_comment`, `_fill_result_row`,
   `_create_packing_only_row` — там на входе список `units`, а не state; значения в Excel
   продолжают считаться ровно как сегодня.
4. Поле `delivered` называть без подчёркивания (dataclass, не приватный протокол), в docstring
   класса указать инвариант: `delivered == _delivered_qty(allocated)`.

**[RISK: любой новый `allocated.append` в обход `take()` тихо ломает инвариант]** — verification
обязателен.

### Шаг 2.2 [C2] Реализация + регрессионный тест инварианта

Добавить в `tmp/test_rfp_packing_compare_smoke.py`, класс `RfpPackingMatcherSmokeTest`,
тест `test_running_delivered_matches_full_sum`: собрать `_RowAllocationState`, скормить
`take()` смесь целых слотов и дробного (`qty=0.3`, `qty=0.7`), после каждого `take`
проверить `assertEqual(state.remaining_qty(), state.ordered - _delivered_qty(state.allocated))`
(именно `assertEqual`, не `assertAlmostEqual` — проверяем побитовое равенство).

**Verification:**

```powershell
$env:PYTHONUTF8='1'
python tmp/test_rfp_packing_compare_smoke.py
python tmp/test_rfp_supply_status_smoke.py
python tmp/test_step4_ds_status_stats_smoke.py
rg -n "allocated\.append" RFQ/
python tmp/diag_ul_matcher_hot.py
```

Готово, когда: `rg` находит `allocated.append` **только** внутри `_RowAllocationState.take`;
смоуки зелёные; в `diag_ul_matcher_hot.py` посадка на очереди 22 000 падает с ~50 с до
единиц секунд (ожидание — линейный рост вместо квадратичного).

---

## Сессия 3 — резервирование тегов (пункт 3)

**Приоритет по факту замера:** обязателен только **блок A** (`a3_mto_tags` 108,2 с +
`a4_blind` 49,3 с). Правка блоков B/C оставлена в спеке, но её выигрыш на текущих данных
< 1 с — делать вместе, потому что дёшево, и не расширять её объём.

### Шаг 3.1 [PREMIUM] Спека (готова; ключевое ограничение — читать внимательно)

**[RISK: главная ловушка этого пункта]** Резервирование **динамическое**: и
`_reserved_tags_for_state`, и `_reserved_leftover_tags` фильтруют соседей по
`other.remaining_qty() > _EPS` в момент хода **текущей** строки. За время прохода соседи,
обработанные раньше, могут стать неактивными. Поэтому **запрещено** считать сами reserved-множества
один раз на старте прохода: это изменит посадку и вернёт баги 8525 (кража тега соседа) и
8445 ДС29/ДС82. Предвычисляется **только группировка**, живое чтение `remaining_qty()` остаётся.

Почему группировка даёт побитово тот же результат: обе функции уже отбрасывают всех, кто не
совпал по ключу (`other.key != state.key` / `code_key_fn(other) != code_key`). Итерация по
заранее собранному bucket с тем же относительным порядком проходит по тому же подмножеству,
а результат — `set`, то есть от порядка не зависит вовсе.

**Блок A.** В `compare_rfp_rows_with_packing` после сборки `a_states`:

```python
    states_by_key: dict[tuple[str, str, str, str], list[_RowAllocationState]] = defaultdict(list)
    for state in a_states:
        states_by_key[state.key].append(state)
```

В проходах A3 (`mto_tags`) и A4 (`blind`) передавать
`_reserved_tags_for_state(state, states_by_key[state.key])` вместо `a_states`.
Тело `_reserved_tags_for_state` не менять (проверка `other.key != state.key` остаётся как
дешёвая страховка).

**Блоки B/C.** В `_allocate_leftover_block`:

1. Один раз посчитать ключи и bucket'ы:

```python
    code_keys: dict[int, tuple[str, str, str] | None] = {
        id(state): _row_leftover_code_key(state.row, code_column) for state in participants
    }
    by_code_key: dict[tuple[str, str, str], list[_RowAllocationState]] = defaultdict(list)
    for state in participants:
        code_key = code_keys[id(state)]
        if code_key is not None:
            by_code_key[code_key].append(state)
```

2. Во всех трёх пассах (`rfp_tags` с own_tags, `untagged`, `blind`) брать
   `code_keys[id(state)]` вместо повторного `code_key_fn(state)`.
3. Сигнатуру `_reserved_leftover_tags` упростить до
   `(state, siblings: list[_RowAllocationState], tags_fn)`; тело — тот же цикл с живым
   `other.remaining_qty() <= _EPS`, но по `siblings`. Вызов:
   `frozenset()` если `code_keys[id(state)] is None`, иначе
   `_reserved_leftover_tags(state, by_code_key[code_key], tags_fn)`.
   Ветка «свой code_key = None → пустое множество» сохраняется по смыслу.
4. Локальную функцию `code_key_fn` в `_allocate_leftover_block` удалить (её заменяет `code_keys`).
   Параметр `code_column` остаётся — он нужен для построения `code_keys`.
5. `id()`-ключи здесь допустимы: `participants` живут до конца блока (тот же идиом, что
   существующий `row_index_by_id`).

### Шаг 3.2 [C2+V] Реализация

**Verification:**

```powershell
$env:PYTHONUTF8='1'
python tmp/test_rfp_packing_compare_smoke.py
python tmp/test_rfp_supply_status_smoke.py
rg -n "_reserved_tags_for_state\(state, a_states\)|code_key_fn" RFQ/
```

Готово, когда: `rg` не находит ни `_reserved_tags_for_state(state, a_states)`, ни `code_key_fn`;
зелёные тесты, в первую очередь эти (они прямо покрывают ловушки резервирования):

- `test_sibling_rfp_tag_not_stolen_when_wrong_title_has_other_tag` (инцидент 8525)
- `test_sibling_on_same_title_does_not_steal_other_rfp_tag`
- `test_two_ds_same_title_code_do_not_share_ul_queue` (8445 ДС29/ДС82)
- `test_leftover_other_ds_sits_on_mto_not_rfp_8445`
- `test_two_mto_only_same_code_keep_own_tags_in_b1` (B1 own-tags)
- `test_mto_tag_match_before_blind_fill`, `test_mismatch_priority_over_via_mto_in_same_allocation`

**Ревью (обязательно, дорогая модель):** убедиться, что в диффе нет ни одного предвычисленного
`reserved`-множества и ни одного снапшота `remaining_qty()` вне цикла.

---

## Сессия 4 — trace без псевдо-`RowStd` (пункт 2)

### Шаг 4.1 [PREMIUM] Спека (готова)

Сегодня `_merge_packing_units_trace` создаёт по одному `RowStd` (146 `CheckElement`) на каждый
слот, только чтобы передать две строки в `merge_units_status_trace`. Порядок и дедуп там такие:
`join_unique_units_text([значение цели, *значения источников])`, дедуп по точному
`str(v or "").strip()` с сохранением первого вхождения, разделитель `"; "`, пустой результат
пишется как `None`. Источники идут в порядке `units`. Значит эквивалентная прямая реализация:

```python
def _merge_packing_units_trace(row: RowStd, units: list[_PackingUnit]) -> None:
    """Attach UL source status/trace without overwriting existing RFP/MTO traces."""
    for column in (UNITS_CHECK_STATUS, UNITS_CONVERSION_TRACE):
        if column not in row.el:
            row.el[column] = CheckElement(None)
    if not any(
        unit.units_check_status or unit.units_conversion_trace for unit in units
    ):
        return
    for column, attribute in (
        (UNITS_CHECK_STATUS, "units_check_status"),
        (UNITS_CONVERSION_TRACE, "units_conversion_trace"),
    ):
        merged = join_unique_units_text(
            [row.get_value(column), *(getattr(unit, attribute) for unit in units)]
        )
        row.el[column].value = merged or None
```

**[RISK: ранний выход обязателен]** Сегодня, если ни у одного слота нет status/trace,
`pseudo_rows` пуст и `merge_units_status_trace` **не вызывается** — существующее значение цели
остаётся нетронутым (в частности пустая строка не превращается в `None`, а значение с пробелами
не обрезается). Без `if not any(...): return` поведение на этой ветке изменится.

Импорт `join_unique_units_text` из `RFQ.ds_compare.ds_units_normalize` добавить рядом с
существующими `merge_units_status_trace` / `normalize_units_text`. Сам
`merge_units_status_trace` в этом модуле после правки может остаться неиспользованным — тогда
убрать его из импортов (проверить `rg`).

Дедуп внутри `join_unique_units_text` уже по set, а строк status всего три
(`identity` / `converted` / `no_google`) и trace сильно повторяются, так что уникальных значений
на строку — единицы. Дополнительный pre-dedupe не нужен.

### Шаг 4.2 [C2+V] Реализация + новый тест

Прямого теста на `_merge_packing_units_trace` в репозитории **нет** — добавить в
`tmp/test_rfp_packing_compare_smoke.py` (класс `RfpPackingIntegrationSmokeTest`)
`test_packing_units_trace_merge_order_and_dedupe`, который фиксирует:

1. порядок «сначала существующее значение цели, потом слоты в порядке списка»;
2. дедуп повторяющихся status/trace до одного вхождения, разделитель `"; "`;
3. ранний выход: слоты с пустыми status/trace **не** меняют существующее значение цели
   (в т.ч. цель со значением `""` остаётся `""`, а не `None`);
4. слот, у которого заполнен только trace, не добавляет пустой status.

**Verification:**

```powershell
$env:PYTHONUTF8='1'
python tmp/test_rfp_packing_compare_smoke.py
python tmp/test_step4_units_gate_order_smoke.py
python tmp/test_units_convert_matrix_smoke.py
python tmp/diag_ul_matcher_hot.py
```

Готово, когда: смоуки зелёные (особенно `test_mto_units_and_trace_transfer` и
`test_converted_units_comment_only_for_matrix_conversion` — они читают эти же колонки);
`rg -n "get_std_check_row" RFQ/tags_rfp_compare/step4/step4_packing_compare.py` показывает
только `_create_packing_only_row`; в `diag_ul_matcher_hot.py` trace-merge на 20 000 слотов
уходит с 3,6 с в миллисекунды.

---

## Итоги сессий 2–4 (выполнены, ждут контрольного прогона)

Всё в `RFQ/tags_rfp_compare/step4/step4_packing_compare.py`; тесты — в
`tmp/test_rfp_packing_compare_smoke.py` (было 74 теста, стало 76).

| Сессия | Что сделано | Микрозамер `tmp/diag_ul_matcher_hot.py` |
|---|---|---|
| 2 | `_RowAllocationState.delivered` + `take()`, `remaining_qty()` без полного суммирования | посадка 22 000 слотов: **49,67 с → 0,02 с**, рост линейный (0,8–0,9 мкс/слот) |
| 4 | `_merge_packing_units_trace` без псевдо-`RowStd`, прямой `join_unique_units_text` | trace на 20 000 слотов: **3,62 с → 3,4 мс**; `_fill_unit_details` целиком: 3,86 с → 25,5 мс |
| 3 | `states_by_key` для A3/A4; `code_keys` / `by_code_key` в `_allocate_leftover_block` | сравнений в A3+A4: ~1,97·10⁹ → ~1,3·10⁵ |

Ревью куратора по каждой сессии:

- **Сессия 2.** Все четыре `state.allocated.append(taken)` заменены на `take()`; `rg` находит
  `allocated.append` только внутри `take`. `_delivered_qty` и три его потребителя
  (`_allocation_comment`, `_fill_result_row`, `_create_packing_only_row`) не изменены —
  значения в Excel по-прежнему считаются полной суммой.
- **Сессия 4.** Новое тело — точный инлайн семантики `merge_units_status_trace`, включая
  guard `column not in row.el` и запись `merged or None`. Обе колонки входят в
  `ColNames.column_list`, а `RowStd.__init__` заводит `CheckElement` на каждую колонку, поэтому
  guard на боевых строках — no-op, и создание колонок до раннего выхода ничего не меняет.
  `merge_units_status_trace` в `ds_units_normalize.py` не тронут (нужен `copy_mto_units_fields`).
- **Сессия 3 (обязательное ревью).** Предвычисленных reserved-множеств и снапшотов
  `remaining_qty()` вне цикла в диффе нет: тела `_reserved_tags_for_state` и
  `_reserved_leftover_tags` читают `other.remaining_qty()` живьём на каждом ходу.
  Проверено отдельно, что группировка не устаревает: `a_states.append` есть только в
  `a_states_build` до построения `states_by_key`, а A3/A4 отрабатывают до `block_b`;
  `b_participants` / `c_participants` формируются и упорядочиваются целиком до вызова
  `_allocate_leftover_block` и внутри него не мутируются, поэтому `id()`-ключи валидны.
  В третьем пассе ветка `frozenset()` при `code_key is None` недостижима из-за
  `if not queue: continue` выше — оставлена как защита, семантику не меняет.

Полный смоук-набор Step4/RFP/units: **22 файла, 281 тест, все зелёные**, включая семь тестов
на ловушки резервирования (инциденты 8525 и 8445 ДС29/ДС82).

**Расчётный прогноз на боевых данных:** 1008 с → ориентировочно 60–100 с. Точность прогноза
ограничена неизвестной природой ~300 с в `a2_untagged` сверх модели `remaining_qty()`
(см. расхождение 2 выше) — контрольный прогон должен это разрешить.

## Результат контрольного прогона (шаг 5.1 пройден)

Прогон `_результат_проверки_2026.09.09.13.40`, отчёт `ul_compare_report_2026.09.09_13.42.39.txt`.

| Фаза | Baseline | Контроль | Дельта |
|---|---:|---:|---:|
| `a2_untagged` | 679,362 | **1,122** | −678,240 |
| `fill_rows` | 162,896 | **3,721** | −159,175 |
| `a3_mto_tags` | 108,163 | **0,140** | −108,023 |
| `a4_blind` | 49,335 | **0,098** | −49,237 |
| `build_unit_queues` | 3,965 | 2,230 | −1,735 |
| `a_states_build` | 2,284 | 1,090 | −1,194 |
| `prepare_rows` | 1,206 | 0,593 | −0,613 |
| `block_b` | 0,825 | 0,101 | −0,724 |
| `leftover_dump` | 0,392 | 0,060 | −0,332 |
| остальные | 0,188 | 0,264 | +0,076 |
| **ИТОГО** | **1008,616** | **9,319** | **−999,3 (108,2×)** |

**Идентичность подтверждена (шаг 5.1 принят):**

- `compare_step4_files.py`: `rows_new=37270`, `rows_old=37270`, `common_columns=43`,
  **`only_new=0`, `only_old=0`**.
- Все `COUNTERS` и `LOADER COUNTERS` в отчёте совпали до единицы (сверка
  `tmp/read_run_ul_timings.py`).
- Комментарии к ячейкам как мультимножество: 654 против 654, `only_old=0`, `only_new=0`
  (значения `_allocation_comment` в Excel `compare_step4_files.py` не проверяет — сверено
  отдельно через `tmp/diff_step4_row_order.py`).
- Листы «Сводка», «Статистика ДС», «Посадка», «Статусы», «Проблемы УЛ» — позиционно
  идентичны, 0 расхождений.

**[ВАЖНО: порядок строк Step4 нестабилен между прогонами и это не регрессия]**
Позиционный дифф двух файлов даёт ~710 тыс. «расхождений» просто из-за другого порядка строк:
в baseline лист начинается с ДС23, в контроле — с ДС6, а комментарий из `AH11083` старого
файла лежит в `AH10256` нового с тем же текстом. Проверено по 25 архивным прогонам: порядок
плавал **задолго до** этих правок — например 27.08 11.48 (ДС23) и 27.08 12.16 (ДС29) при
одинаковых 37 108 строках, и сегодняшние 09.09 10.47 (ДС23) против 09.09 11.48 (ДС6) при
одинаковых 37 277 строках, оба до оптимизаций. Ведущий ДС ротируется между
ДС23 / ДС29 / ДС6 / ДС1_ГОСФИН — похоже на порядок завершения параллельной обработки ДС.
**Вывод для будущих сверок:** идентичность Step4 проверять только мультимножеством
(`compare_step4_files.py`), позиционный дифф для этого файла бессмыслен.

**Гипотеза про ~300 с в `a2_untagged` закрыта.** Фаза упала до 1,12 с, то есть все 679 с были
пересчётом `remaining_qty()` (модель недооценила из-за промахов кэша на реальных данных), а не
сканом теговых слотов. Отказ от поштучных слотов из «Что не входит в этот план» **не нужен** —
цель перевыполнена.

---

## Сессия 5 — контрольный прогон и документация

### Шаг 5.1 [C2+V] Контрольный прогон

Тот же профиль, что baseline (`step4.timing_log=true`). Зафиксировать `step4::ul::*`.

**Сверка идентичности (обязательна, без неё сессии 2–4 не считаются принятыми):**

```powershell
$env:PYTHONUTF8='1'
python RFQ/tags_rfp_compare/compare_step4_files.py <baseline.xlsx> <new.xlsx>
```

Критерий: `only_old=0`, `only_new=0`. Плюс сверить COUNTERS в `ul_compare_report_*.txt`
baseline vs new — все счётчики (`allocated_units`, `not_in_packing`, `packing_only_added`,
`complete`, `shortfall`, `overdelivery`, `unit_mismatches`, …) должны совпасть до единицы.
**[RISK: расхождение хотя бы одного счётчика = регрессия семантики, а не «шум»]** —
в этом случае откатывать сессию 3 первой (она единственная касается решений о посадке).

### Шаг 5.2 [C2] Документация

- `.cursor/rules/AI_optimization_roadmap.mdc` — новый «Этап D: шаг УЛ», baseline/после, заметки.
- `.cursor/rules/AI_performance.mdc` — раздел по УЛ: цифры до/после, ссылка на `step4::ul::*`.
- `.cursor/rules/AI_rfp_step4_excel.mdc` — в разделе «УЛ matcher» дописать: инвариант
  `delivered == _delivered_qty(allocated)` и запрет на статический снапшот reserved-тегов.

---

## Что не входит в этот план

- **Отказ от поштучных слотов** (держать слот как `(tag, qty)` вместо 1,25 млн объектов) —
  отдельный **[PREMIUM]** дизайн, меняет семантику проходов A1–A4/B/C и дампа «Только в УЛ».
  Рассматривать только если после сессий 2–4 шаг не уложился в цель.
- `gc.disable()` вокруг шага УЛ — ~35% только на trace-merge, который сессия 4 и так убирает.
  Смысла нет.
- Оптимизация `save_match_result_to_excel` — другой milestone (`save_excel`).

## Распределение по исполнителям

Политика моделей субагентов — `.cursor/rules/AI_project_standards.mdc`, «Модели субагентов»:
дорогая модель только на критичные шаги.

| Шаг | Метка | Кто | Модель субагента |
|---|---|---|---|
| 1.1, 2.1, 3.1, 4.1 | [PREMIUM] | куратор (спеки в этом файле, уже готовы) | — |
| 1.2 (сделан), 2.2, 5.2 | [C2] | субагент `generalPurpose` по спеке шага | Grok 4.6 |
| 4.2 | [C2+V] | субагент + ревью куратора по диффу | Grok 4.6 |
| 3.2 | [C2+V] | субагент + **обязательное** ревью куратора | дорогая: единственный шаг, где отклонение от спеки тихо меняет посадку |
| 1.3, 5.1 | [C2+V] | пользователь запускает прогон, куратор сверяет | — |

Одна сессия ≈ один чат исполнителя. Сессии 2, 3, 4 независимы по коду и могут идти
последовательно в любом порядке, но сессия 1 — строго первой, а сессия 5 — строго последней.
