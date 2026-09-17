# План реализации: улучшенная адаптация шаблона к сетке PDF

**Создан:** 30.03.2026  
**Реализовано:** 30.03.2026 (Фаза A — boundary-snap)  
**Статус:** ФАЗА A ЗАВЕРШЕНА, Фаза B (каскад) — при необходимости  

**Решение по архитектуре (обсуждение перед реализацией):**

Вместо полного каскадного алгоритма реализована **Фаза A — boundary-snap**:
1. Заменён IoU-snap на `_snap_4_boundaries` (привязка каждой из 4 границ к grid-линиям)
2. Добавлен `shape_score` для диагностики деформации
3. Добавлен граф смежности + `boundary_consistency` как компонент скоринга
4. Нормализация score по доступным компонентам (без штрафа за отсутствие соседей)

**Причина двухфазного подхода:**
- Каскад несёт риск propagation of errors (неверная привязка → смещение соседей)
- Boundary-snap решает 80% проблем при 30% сложности каскада
- Фаза B (каскад) может быть добавлена если Фаза A недостаточна

**Ключевые решения из обсуждения:**
- P1 (sx/sy от маленьких полей): не актуально — Фаза A не использует local transform от соседей
- P2 (граф смежности): реализован с проверкой проекции перекрытия > 50%
- P3 (инверсия после snap): нормализация + min_size + max_shape_change проверка
- Scoring weights нормализуются при отсутствии matched соседей
- Статусы: anchor / matched / derived / excluded (derived вместо no_match для fallback)

**Связанные файлы:** `grid_matcher.py`, `models.py`, `template_editor/main_window.py`,
`template_editor/cells_panel.py`, `template_editor/template_meta_panel.py`

---

## Контекст (обязательно прочитать перед реализацией)

- `.cursor/rules/AI_v2_editor.mdc` — бэклог задач редактора
- `.cursor/rules/AI_v2_engine.mdc` — архитектура v2
- `pdf_parsing_v2/grid_matcher.py` — текущий алгоритм (заменяется)
- `pdf_parsing_v2/models.py` — `FieldDef`, `StampTemplate`, `AdaptResult`
- `pdf_parsing_v2/coord_transform.py` — `field_to_fitz_rect`, `SCALE = 2.83444`

---

## Что не так с текущим алгоритмом (`adapt_all_fields`)

1. **Один глобальный трансформ** от всех якорей → неточен для полей далеко от якорей
2. **Один проход** без обратной связи → нет возможности уточнить позицию
3. **Snap к ячейке целиком** через IoU → хрупкий (±2мм → `no_match`)
4. **Пропорциональное масштабирование размера** от якоря → деформирует узкие поля (6_1, 6_2)
5. **Нет скора изменения формы** → неизвестно, насколько поле деформировано
6. **5+ непонятных параметров** в UI

---

## Новый алгоритм: `adapt_fields_cascade`

### Общая схема

```
Phase 0: Извлечь grid lines из detected cells
  x_lines = cluster_lines([x0,x1 for each detected], tol)
  y_lines = cluster_lines([y0,y1 for each detected], tol)

Phase 1: Привязать явные якоря (is_anchor=True)
  Для каждого якоря: best_cell = argmin anchor_score(template_bbox, detected)

Phase N (cascade, до конвергенции):
  Отсортировать непривязанные поля по matchability_key (большие уникальные — первые)
  Для каждого поля:
    neighbors = k_nearest_matched(field, matched, k=4)
    local_tf  = compute_local_transform(neighbors)
    approx    = apply_local_transform(local_tf, template_bbox)
    candidates = [c for c in detected if intersects(expand(approx, margin))]
    best, score = best_candidate(field, candidates, matched, x_lines, y_lines)
    if score > threshold:
      snapped = snap_4_boundaries(best, x_lines, y_lines)
      matched[field.id] = snapped, confidence=score

По завершении:
  Для оставшихся непривязанных: interpolate_from_neighbors → status="derived"
```

---

## Детальная спецификация функций

### 1. `adapt_fields_cascade()` — главная точка входа

```python
def adapt_fields_cascade(
    template: StampTemplate,
    frame: FrameInfo,
    fitz_page: fitz.Page,
    cfg: dict | None = None,
) -> dict[str, AdaptResult]:
```

Параметры из `StampTemplate`:
- `template.snap_max_distance_mm` — радиус поиска кандидатов (main knob, default 5.0)
- `template.grid_tolerance_detected_mm` — кластеризация X/Y линий (advanced, default 0.5)
- `template.snap_size_tolerance` — допустимое изменение пропорций (переименовать в `max_shape_change_ratio`, default 2.5)

Возвращает `dict[field_id → AdaptResult]` с расширенными полями (см. ниже).

---

### 2. `_extract_grid_lines()` — извлечение линий сетки

```python
def _extract_grid_lines(
    detected: list[CellBbox],
    tolerance_mm: float,
) -> tuple[list[float], list[float]]:
    """Returns (x_lines, y_lines) as sorted lists of pts."""
```

Берёт `x0, x1` всех ячеек → `cluster_lines` → X-линии.  
Берёт `y0, y1` всех ячеек → `cluster_lines` → Y-линии.

---

### 3. `_matchability_key()` — порядок обработки полей

```python
def _matchability_key(
    field: FieldDef,
    matched: dict[str, CellBbox],
    frame: FrameInfo,
    detected: list[CellBbox],
) -> tuple[float, float]:
```

Возвращает `(priority, dist_to_nearest_matched)` — меньше = обработать раньше.

```
priority = -area_pts² / (n_similar_detected + 1)
  где n_similar_detected = число detected cells с похожим aspect ratio (±20%)

dist_to_nearest_matched = min distance от template_bbox.center до matched_bbox.center
  если matched пустой → infinity
```

**Итоговый порядок:** большие уникальные поля рядом с уже привязанными — первые.

---

### 4. `_compute_local_transform()` — локальный трансформ

```python
@dataclass
class LocalTransform:
    dx: float      # смещение по X (pts)
    dy: float      # смещение по Y (pts)
    sx: float      # масштаб по X (ratio)
    sy: float      # масштаб по Y (ratio)

def _compute_local_transform(
    neighbors: list[tuple[CellBbox, CellBbox, float]],  # (template, detected, confidence)
) -> LocalTransform:
```

Вычисляет взвешенное среднее смещений и масштабов:
```
weight[i] = neighbors[i].confidence / max(distance_to_field, 4.0)

dx = sum(weight[i] * (detected[i].cx - template[i].cx)) / sum(weights)
dy = sum(weight[i] * (detected[i].cy - template[i].cy)) / sum(weights)
sx = sum(weight[i] * (detected[i].w  / template[i].w))  / sum(weights)
sy = sum(weight[i] * (detected[i].h  / template[i].h))  / sum(weights)
```

Если `neighbors` пустой → `LocalTransform(0, 0, 1.0, 1.0)` (identity).

> ⚠️ **ПРОБЛЕМА P1**: `sx` и `sy` могут быть сильно искажены если у соседа маленькая ширина/высота.
> Возможное решение: clamp sx/sy в диапазон [0.7, 1.5] и/или взвешивать по площади поля.

---

### 5. `_apply_local_transform()` — применить трансформ

```python
def _apply_local_transform(
    tf: LocalTransform,
    template_bbox: CellBbox,
) -> CellBbox:
    """Apply local transform: shift center + scale size."""
    cx, cy = template_bbox.center
    new_cx = cx + tf.dx
    new_cy = cy + tf.dy
    new_w  = template_bbox.width  * tf.sx
    new_h  = template_bbox.height * tf.sy
    return CellBbox(new_cx - new_w/2, new_cy - new_h/2,
                    new_cx + new_w/2, new_cy + new_h/2)
```

---

### 6. `_score_candidate()` — скор кандидата

```python
def _score_candidate(
    field: FieldDef,
    frame: FrameInfo,
    candidate: CellBbox,
    approx: CellBbox,
    matched: dict[str, tuple[CellBbox, float]],  # field_id → (bbox, confidence)
    template: StampTemplate,
) -> float:
```

Три компонента:

**a) shape_score** (изменение формы):
```
w_ratio = candidate.w / max(template_bbox.w, EPS)
h_ratio = candidate.h / max(template_bbox.h, EPS)
shape_score = min(w_ratio, 1/w_ratio) * min(h_ratio, 1/h_ratio)
```
Диапазон [0, 1]. 1.0 = форма не изменилась.

**b) position_score** (близость к аппроксимации):
```
dist_mm = center_distance(candidate, approx) / SCALE
position_score = 1.0 / (1.0 + dist_mm)
```

**c) boundary_consistency** (совместимость общих границ):
```
Для каждого уже привязанного соседнего поля, которое должно делить границу с текущим:
  expected_shared_boundary = compute_shared_boundary(field, neighbor, frame)
  actual_boundary = candidate boundary
  error_mm = abs(actual - expected) / SCALE
  score += 1.0 / (1.0 + error_mm)
boundary_consistency = avg(score per shared neighbor boundary)
```

**Итого:**
```
total = shape_score * W_SHAPE + position_score * W_POS + boundary_consistency * W_BOUND
  где W_SHAPE=0.3, W_POS=0.3, W_BOUND=0.4
```

> ⚠️ **ПРОБЛЕМА P2**: `boundary_consistency` требует знать, какие поля делят границу.
> Варианты:
> a) Вычислять на лету: два поля делят границу если в шаблоне расстояние между границами < tolerance
> b) Предвычислить граф смежности при старте алгоритма
> Рекомендация: вариант (b) — `_build_adjacency_graph(template, frame, tol_mm=1.0)`

---

### 7. `_snap_4_boundaries()` — привязка 4 границ к grid lines

```python
def _snap_4_boundaries(
    bbox: CellBbox,
    x_lines: list[float],
    y_lines: list[float],
    max_snap_pts: float,
) -> tuple[CellBbox, int]:
    """
    Returns (snapped_bbox, n_snapped).
    n_snapped = 0..4 (сколько из 4 границ удалось привязать к линии).
    """
```

Для каждой из 4 границ:
```
new_x0 = nearest_line(bbox.x0, x_lines, max_snap_pts) or bbox.x0
new_x1 = nearest_line(bbox.x1, x_lines, max_snap_pts) or bbox.x1
new_y0 = nearest_line(bbox.y0, y_lines, max_snap_pts) or bbox.y0
new_y1 = nearest_line(bbox.y1, y_lines, max_snap_pts) or bbox.y1
```

> ⚠️ **ПРОБЛЕМА P3**: после snap_4_boundaries поле может инвертироваться (x0 > x1) если
> ближайшие линии оказались не те. Нужна нормализация и проверка min_size.

---

### 8. `_shape_score()` — финальный скор изменения формы

```python
def _shape_score(original: CellBbox, adapted: CellBbox) -> float:
    w_ratio = adapted.width  / max(original.width,  _EPS)
    h_ratio = adapted.height / max(original.height, _EPS)
    return min(w_ratio, 1/w_ratio) * min(h_ratio, 1/h_ratio)
```

Добавляется в `AdaptResult` (см. ниже).

---

### 9. `_interpolate_from_neighbors()` — fallback для исключений

```python
def _interpolate_from_neighbors(
    field: FieldDef,
    frame: FrameInfo,
    matched: dict[str, tuple[CellBbox, float]],
    k: int = 4,
) -> CellBbox:
```

Берёт K ближайших привязанных полей, строит локальный трансформ, применяет к template_bbox.
**Без snap** — чистая интерполяция. Status = `"derived"`.

> ⚠️ **ПРОБЛЕМА P4**: если нет ни одного привязанного соседа (все failed) — вернуть template_bbox.
> Это нормальная деградация (лучше чем crash), но нужно явно логировать.

---

## Расширение `AdaptResult`

```python
@dataclass
class AdaptResult:
    field_id: str
    bbox: CellBbox
    status: str  # "anchor"|"cascade"|"derived"|"excluded"
    approx_bbox: CellBbox
    matched_detected: list[CellBbox]
    snap_distance_mm: float
    # НОВЫЕ ПОЛЯ:
    shape_score: float = 1.0       # 1.0=без изменений, 0=сильная деформация
    snapped_boundaries: int = 0    # 0-4: сколько из 4 границ привязано к grid line
    iteration: int = 0             # на какой итерации каскада привязано
    confidence: float = 0.0        # итоговый скор (total из _score_candidate)
```

> ⚠️ **ПРОБЛЕМА P5**: старый код в `main_window.py` читает `ar.status` и проверяет
> `ar.status in ("matched", "split_merged", "anchor")`. Нужно обновить эти проверки
> на новые статусы: `"anchor"`, `"cascade"`, `"derived"`, `"excluded"`.

---

## Расширение `StampTemplate` (упрощение параметров)

### Что убрать из UI (оставить в модели для backward compat):
- `snap_size_tolerance` — заменяется `max_shape_change_ratio`
- `snap_iou_threshold` — не нужен при boundary-based подходе
- `anchor_min_area_fraction` — auto-anchors отключены, параметр бессмысленен

### Что добавить:
```python
max_shape_change_ratio: float = 2.5  # max(w_ratio, 1/w_ratio) <= this value → accept
cascade_max_iterations: int = 10     # max итераций каскада
cascade_score_threshold: float = 0.4 # min total score для принятия привязки
```

> ⚠️ **ПРОБЛЕМА P6**: backward compat JSON. Старые шаблоны не имеют новых полей.
> Решение: defaults в `StampTemplate.from_dict()` (уже есть паттерн).

---

## Изменения в UI

### `template_meta_panel.py`

Убрать из UI (но оставить в модели):
- `Snap size tol` (`snap_size_tolerance`)
- `Snap IoU` (`snap_iou_threshold`)

Добавить в UI:
- `Max shape change` (`max_shape_change_ratio`, default 2.5, spinbox 1.0–5.0)
- `Cascade iterations` (`cascade_max_iterations`, default 10, spinbox 1–20) — advanced
- `Min score` (`cascade_score_threshold`, default 0.4, spinbox 0.1–1.0) — advanced

---

### `cells_panel.py`

Добавить колонку **Score** (ширина ~60px) после существующих:
- `> 0.7` → зелёный
- `0.4–0.7` → жёлтый
- `< 0.4` → красный
- `derived` → серый + текст "~"

Добавить колонку **Iter** (итерация каскада, 1–N или "A" для anchor, "D" для derived).

> ⚠️ **ПРОБЛЕМА P7**: QTreeWidget с 5+ колонками — нужно проверить что заголовки
> корректно resize и не обрезаются. Рекомендация: setResizeMode(ResizeToContents) для Score/Iter.

---

### `main_window.py` — оверлей `_on_adapt_to_grid`

Текущий оверлей: стрелки approx→final, detected cells, no_match/split_merged контуры.

Нужно добавить:
1. **Сетку линий** (X-линии = вертикальные синие пунктиры, Y-линии = оранжевые горизонтальные)
   — по аналогии с «Консолидировать сетку»
2. **Итерационную метку** рядом с каждым полем: "it.1", "it.2", "D" (derived)
3. **Shape score** рядом с полем (мелкий текст): "0.87"
4. **Легенда** в углу canvas

Убрать из оверлея:
- Стрелки approx→final (загромождают, мало информативны)

> ⚠️ **ПРОБЛЕМА P8**: текстовые метки (`QGraphicsTextItem`) на canvas медленны при >50 полях.
> Если лагает — использовать `QGraphicsSimpleTextItem` или рисовать через `painter.drawText`.

---

## Граф смежности (предвычисление)

```python
def _build_adjacency_graph(
    template: StampTemplate,
    frame: FrameInfo,
    tol_mm: float = 1.0,
) -> dict[str, list[tuple[str, str]]]:
    """
    Returns: {field_id: [(neighbor_id, shared_edge), ...]}
    shared_edge: "left"|"right"|"top"|"bottom"
    
    Два поля смежны если расстояние между одной из их границ < tol_mm * SCALE.
    Например: field_A.x1 ≈ field_B.x0 → field_A("right") смежно с field_B("left")
    """
```

Строится один раз в начале `adapt_fields_cascade`.

> ⚠️ **ПРОБЛЕМА P9**: для шаблонов где поля НЕ идеально выровнены (после ручной правки
> в редакторе), tol_mm=1.0 может не найти реальных соседей. Нужна настройка или авто-tol
> как `grid_tolerance_template_mm`.

---

## Переключение в `main_window.py`

Заменить вызов:
```python
# было:
from pdf_parsing_v2.grid_matcher import adapt_all_fields
adapted = adapt_all_fields(template=template, frame=frame, fitz_page=fitz_page, cfg=self._cfg)

# стало:
from pdf_parsing_v2.grid_matcher import adapt_fields_cascade
adapted = adapt_fields_cascade(template=template, frame=frame, fitz_page=fitz_page, cfg=self._cfg)
```

Старую `adapt_all_fields` **не удалять** — оставить как fallback и для `stamp_extractor.py`.

---

## Порядок реализации (для чата-исполнителя)

1. `grid_matcher.py`:
   - Добавить новые dataclass'ы (расширить `AdaptResult`)
   - `_extract_grid_lines()`
   - `_build_adjacency_graph()`
   - `_matchability_key()`
   - `_compute_local_transform()`
   - `_apply_local_transform()`
   - `_snap_4_boundaries()`
   - `_score_candidate()` с boundary_consistency
   - `_shape_score()`
   - `_interpolate_from_neighbors()`
   - `adapt_fields_cascade()` — main orchestrator

2. `models.py`:
   - Добавить `max_shape_change_ratio`, `cascade_max_iterations`, `cascade_score_threshold`
   - Расширить `AdaptResult` новыми полями
   - Backward-compat в `from_dict`

3. `template_editor/template_meta_panel.py`:
   - Убрать из UI: `snap_size_tolerance`, `snap_iou_threshold`
   - Добавить: `max_shape_change_ratio`, `cascade_score_threshold` (в секцию Advanced)

4. `template_editor/cells_panel.py`:
   - Добавить колонки Score + Iter

5. `template_editor/main_window.py`:
   - Переключить `_on_adapt_to_grid` на `adapt_fields_cascade`
   - Обновить оверлей: добавить grid lines, убрать стрелки, добавить метки

---

## Открытые проблемы (требуют решения)

### P1 — Деформация `sx`/`sy` от маленьких соседей
**Проблема:** если сосед имеет маленькую высоту (например, `6_1` высотой 5мм), то `sy = detected.h / template.h` даст нестабильный результат при любом шуме.  
**Варианты решения:**
- a) Clamp: `sx = clamp(sx, 0.7, 1.5)` per neighbor до усреднения
- b) Взвешивать по площади поля-соседа (мелкие соседи дают малый вес)
- c) Использовать только трансляцию (`dx, dy`) от маленьких соседей, масштаб — только от больших

### P2 — Граф смежности: как определить "общую границу"
**Проблема:** два поля в шаблоне могут казаться смежными (близкие границы) но на самом деле разделены другим полем.  
**Вариант решения:** строить граф только для пар, где нет другого поля между ними (проверка через bbox overlap вдоль общей границы).

### P3 — Инверсия bbox после snap
**Проблема:** если `x0` snaps к правой линии, а `x1` к левой — bbox инвертируется.  
**Решение:** после snap применить `normalize(new_bbox)` + если `width < min_size` → вернуть `approx_bbox` без snap (status="snap_failed").

### P4 — Все соседи не привязаны (cold start)
**Проблема:** поле далеко от якорей, все соседи тоже ещё не привязаны → `_interpolate_from_neighbors` получает пустой `matched`.  
**Решение:** вернуть `template_bbox` (identity) + status="unresolved". Это поле будет подхвачено на следующей итерации после того, как соседи привяжутся.

### P5 — Обратная совместимость статусов в `main_window.py`
**Проблема:** `can_move = ar.status in ("matched", "split_merged", "anchor")`.  
**Решение:** заменить на `ar.status in ("anchor", "cascade")` и добавить `"derived"` в список "не перемещать" (только показывать overlay с предупреждением).

### P6 — JSON backward compat
**Решение:** в `StampTemplate.from_dict()` добавить defaults:
```python
max_shape_change_ratio=float(data.get("max_shape_change_ratio", 2.5)),
cascade_max_iterations=int(data.get("cascade_max_iterations", 10)),
cascade_score_threshold=float(data.get("cascade_score_threshold", 0.4)),
```

### P7 — QTreeWidget с 5 колонками
**Решение:** использовать `header.setSectionResizeMode(col, QHeaderView.ResizeToContents)` для Score и Iter.

### P8 — Производительность текстовых меток на canvas
**Решение:** рисовать метки только при zoom > 50% (проверить `self._gfx_view.transform().m11() > 0.5`); при малом zoom — только цвет рамки.

### P9 — Граф смежности при неидеальных шаблонах
**Решение:** использовать `grid_tolerance_template_mm` (уже есть в `StampTemplate`) как tolerance для обнаружения смежности. По умолчанию 2.0мм — достаточно для ручных правок.

---

## Тест: что должно работать после реализации

- [ ] `adapt_fields_cascade` возвращает `AdaptResult` для всех полей шаблона
- [ ] Поля `6_1` и `6_2` (смежные) получают общую привязанную X-границу
- [ ] `shape_score` в `AdaptResult` отражает реальное изменение пропорций (проверить вручную)
- [ ] `iteration` растёт от 0 (якоря) до N (дальние поля)
- [ ] `derived` статус получают поля без кандидатов (не `no_match`)
- [ ] Без якорей (`is_anchor=False` на всех) → все поля получают identity transform → snap к grid lines
- [ ] С 2 якорями на углах → cascade правильно распространяется
- [ ] `cells_panel.py` показывает Score и Iter без ошибок
- [ ] Overlay в редакторе показывает X/Y grid lines
- [ ] Старый `adapt_all_fields` не сломан (существующие шаблоны работают)

---

## Исправление: дедупликация field ID (30.03.2026 22:10)

**Проблема:** при наличии дублирующихся `field.id` в шаблоне (типично для `static_proveril`,
`static_utverdil`, `static_empty` в DWG-шаблонах) `adapt_all_fields` использовал `dict[str, AdaptResult]`
с ключом `field.id`, что приводило к:
- Перезаписи результатов (108 полей → 39 записей)
- Якоря с одинаковым ID потребляли detected cells, но результат перезаписывался последним
- Corrupted anchor transform → каскадное искажение approx_bbox для всех полей

**Решение:**
- `_make_unique_keys(fields)` — уникальные позиционные ключи: `id` при уникальности, `id#N` при дубликатах
- `adapt_all_fields_indexed()` — новая основная функция, возвращает `list[tuple[str, AdaptResult]]`,
  aligned 1-к-1 с `template.fields`
- `adapt_all_fields()` — обёртка (backward compat), `dict(adapt_all_fields_indexed(...))`
- `_match_anchor_cells_indexed()` — якоря с позиционными ключами
- `_build_adjacency_graph_keyed()` — граф смежности с позиционными ключами
- `main_window._on_adapt_to_grid` и `_on_adapt_debug_dump` используют indexed API
- `cells_panel.update_adapt_results` — ключ `id(cell_item)` вместо `field.id`
- Pipeline (`stamp_extractor.py`) — без изменений, т.к. JSON-шаблоны имеют уникальные ID

---

## Cell-assignment вместо boundary-snap (30.03.2026 22:25)

**Проблема boundary-snap:** при плотной сетке (ячейки ~5мм) ошибка интерполяции от якорей
(~5-6мм) больше половины высоты ячейки → snap прыгает на соседнюю строку. 108 полей
snap'ают 4 границы независимо к 15×21 grid-линиям → массовые наложения.

**Решение — `adapt_by_cell_assignment()`:**
1. Проверка прямоугольности detected cells (`_check_rectangularity`)
2. Глобальный transform по top-left bbox (template stamp → detected stamp) + масштаб
3. Greedy matching: для каждого template field находим ближайшую unassigned detected cell
4. Cost = center distance (мм) + 5 × |log(area_ratio)|
5. Max distance = 10мм (конфигурируемо)

**Преимущества перед boundary-snap:**
- Наложения невозможны (detected cells не пересекаются, one cell per field)
- Устойчив к ошибке приближения: достаточно попасть центром в правильную ячейку
- Нет проблемы «сдвиг на строку»

**Дополнительно: кнопка «Слепить поля»:**
- Для ручного режима (не зависит от adaptation)
- Для каждой пары соседних полей (общая граница ± tolerance)
  snap'ает общую границу к средней точке → устраняет зазоры и наложения
- Не двигает поля в абсолютном смысле, только выравнивает смежные границы
