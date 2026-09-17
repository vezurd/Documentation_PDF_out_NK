---
description: Дорожная карта доработки GUI редактора шаблонов (pdf_parsing_v2/template_editor/). Бэклог улучшений с задачами и статусами.
globs: pdf_parsing_v2/template_editor/**
alwaysApply: false
---

# AI Дорожная карта: Редактор шаблонов v2

> Создан 29.03.2026 14:35. Связан с общим roadmap: `.cursor/rules/AI_pdf_v2_roadmap.mdc`.
> Архитектура движка, форматы JSON, модели — **не дублируются**, читать там.

## ПРАВИЛА ДЛЯ ВСЕХ ЧАТОВ (редактор шаблонов)

> **ОБЯЗАТЕЛЬНО** для каждого чата, работающего над `pdf_parsing_v2/template_editor/`:
>
> 1. **Прочитай этот файл целиком** перед началом работы
> 2. **Прочитай `AI_pdf_v2_roadmap.mdc`** — этап F (F.1–F.6), архитектурные решения,
>    UX-требования, формат JSON-шаблонов, модели dataclass
> 3. **Прочитай все файлы `template_editor/`** перед изменением — они связаны
> 4. **После завершения задачи** — обнови статус в бэклоге ниже
> 5. **Не ломай** уже работающие механизмы: zoom/pan, rubber band, highlight
6. **Даты в дорожной карте** — указывать дату **и время до минуты**: `ДД.ММ.ГГГГ ЧЧ:ММ`
   (например: `29.03.2026 14:35`). Это касается всех событий: создание файла, статус этапа,
   результат чата. Так же поступать и в других roadmap-файлах проекта.

---

## Бэклог улучшений

> Статусы: `[ ]` НЕ НАЧАТ · `[~]` В РАБОТЕ · `[x]` ЗАВЕРШЁН · `[!]` ЗАБЛОКИРОВАН
> Приоритеты: `P1` критично · `P2` важно · `P3` желательно

### Группа 1: Диагностика и консоль

#### T1.1 — Вывод в консоль (P1) `[x]`

**Завершено 29.03.2026 ~18:00.** Реализован `self._log(msg, level)` в `main_window.py`
(module-level `_log` + метод экземпляра). Логируется: запуск редактора (templates_dir),
открытие PDF (имя файла, страниц, rotation, время), авто-разметка, загрузка/сохранение
шаблона, активация каталога, тест F5 (score, поля с данными, пустые raw_value).
Ошибки с полным traceback при исключениях в `extract_page` и загрузке файлов.

---

#### T1.2 — F5 «Тестировать» не работает молча (P1) `[x]`

**Завершено 29.03.2026 ~18:00.** `_on_test()` обёрнут в `try/except` с выводом
полного traceback и `QMessageBox.critical`. После успешного теста statusbar показывает
`«Тест завершён: score=0.72, полей с данными: 12/15»`. Предупреждения для пустых
`raw_value` логируются в консоль с уровнем `warn`.

---

### Группа 2: Управление шаблонами и каталогами

#### T2.1 — Единая папка шаблонов из конфига (P1) `[x]`

**Завершено 29.03.2026 ~18:00.** `__init__.py` загружает `load_v2_config()`,
резолвит `templates_dir` до абсолютного пути, передаёт в `TemplateEditorWindow(cfg=...)`.
`main_window.py` хранит `self._templates_dir` и использует во всех диалогах.

---

#### T2.2 — Браузер шаблонов (IDE-style левая панель) (P1) `[x]`

**Завершено 29.03.2026 ~18:00.** Новый файл `templates_browser.py` —
`TemplatesBrowserPanel(QWidget)`. Расположен слева от canvas (как в IDE),
QSplitter 3-панельный: browser | canvas | cells+props.
- QTreeWidget с группами «Шаблоны», «Каталоги», «Наборы»
- Кнопка ↻ для пересканирования templates_dir
- Двойной клик по шаблону → `_on_template_browser_load` (спрашивает при несохр. изм.)
- Двойной клик по каталогу → `_on_catalog_browser_activate` (устанавливает активный)
- Двойной клик по набору → `_on_set_browser_edit` (открывает SetEditorDialog)
- Активный шаблон/каталог выделен bold + синим цветом

---

#### T2.3 — Индикатор текущего активного шаблона (P2) `[x]`

**Завершено 29.03.2026 ~18:00.**
- Заголовок: `«Редактор шаблонов v2 — dwg_page1.json[*]»` (Qt `[*]` placeholder)
- `setWindowModified(True/False)` — `[*]` появляется при несохранённых изменениях
- QLabel «Шаблон: …» добавлен в toolbar справа
- `_set_modified(True)` вызывается при: авто-разметке, удалении ячеек,
  изменении свойств поля (`_on_field_changed`), перетаскивании/ресайзе ячейки (через `_geo_cb`)

---

#### T2.4 — Меню «Набор шаблонов» починено (P2) `[x]`

**Завершено 29.03.2026 ~18:00.** `_on_set_editor()` передаёт `templates_dir`
в `SetEditorDialog(templates_dir=...)`. Диалоги «Добавить шаблон», «Загрузить набор»,
«Сохранить набор» открываются в `templates_dir/sets/` или `templates_dir`.
После закрытия — `_templates_browser.refresh()`.

---

#### T2.5 — Меню «Каталог полей» починено (P2) `[x]`

**Завершено 29.03.2026 ~18:00.** `_on_catalog_editor()` передаёт `templates_dir`
в `CatalogEditorDialog(templates_dir=...)`. Диалоги загрузки/сохранения открываются
в `templates_dir/catalogs/`.

---

#### T2.6 — Наборы: дерево шаблонов + DnD + автосохранение (P2) `[x]`

**Завершено 29.03.2026 21:05.** В `templates_browser.py` у каждого JSON набора под узлом
отображаются шаблоны из `template_files` (имя файла, tooltip — разрешённый путь; `⚠` если файл не найден).
Перетаскивание из «Шаблоны» на узел набора или на строку-шаблон внутри набора **добавляет** шаблон в конец
(дубликат по абсолютному пути — предупреждение в лог, без изменений). Перетаскивание строки шаблона
из набора на заголовок «Шаблоны» или на любой файл в списке «Шаблоны» **убирает** шаблон из набора.
После операции набор сразу сохраняется через `TemplateSet.from_json` → обновление `template_files` →
`to_json` (поля `name`, `projects`, `field_catalog` сохраняются из файла). Класс `_TemplatesTree`:
MIME `application/x-pdf-v2-template-tree-drag`, `startDrag` / `dropEvent`. Сигнал `message(level, text)`
→ `main_window._on_browser_message` → `_log`. Двойной клик по дочернему шаблону набора —
`template_load_requested` (открыть в редакторе).
**Замечание PySide6:** `QMimeData` импортировать из `PySide6.QtCore`, не из `QtGui`.

---

### Группа 3: UX и удобство

#### T3.1 — Кнопка «Редактор шаблонов v2» в main.py (P3) `[x]`

**Проверено 29.03.2026:** кнопка реализована в `main.py` строки 1042–1053:
`_launch_template_editor()` → `subprocess.Popen([sys.executable, "-m", "pdf_parsing_v2.template_editor"])`.
Запуск не блокирует основное GUI. **Задача закрыта.**

---

#### T3.2 — Resize handles: доработка drag (P2) `[x]`

**Проверено 29.03.2026:** `_Handle` в `cell_items.py` полностью реализован.
`mousePressEvent` сохраняет `_orig_scene_rect`, `mouseMoveEvent` вычисляет новый
QRectF по маскам `_MOVES_LEFT/RIGHT/TOP/BOTTOM` и вызывает `parent.setPos(r.topLeft())`
+ `parent.setRect(...)`. Нормализация и минимальный размер (`_MIN_ITEM_SIZE`) присутствуют.
**Задача закрыта.**

---

#### T3.3 — bbox_mm: обнуление в _apply и отсутствие live-обновления (P2) `[x]`

**Завершено 29.03.2026 ~18:00.**

1. **`_apply()` теперь вычисляет реальный bbox:** `_compute_bbox_mm(item)` берёт
   `sceneBoundingRect()` и конвертирует через `gfx_view.scene_to_mm_from_frame`.
   `PropertiesPanel.__init__` принимает `gfx_view` как 2-й аргумент (после catalog).
   `_wire_panels()` передаёт `self._gfx_view`.

2. **Live update при drag/resize:** `FieldRectItem` получил атрибут `_geo_cb: object = None`.
   Вызывается из `itemChange(ItemPositionHasChanged)` и `setRect()`.
   `main_window._on_scene_selection_changed` устанавливает/очищает `_geo_cb` на
   выбранном элементе → `_on_item_geometry_changed` → `props_panel.refresh_bbox_from_item`.
   Также через `_geo_cb` → `_set_modified(True)` при перетаскивании.

---

#### T3.7 — Список полей: заголовок + автопрокрутка при выделении на canvas (P3) `[x]`

**Завершено 29.03.2026 22:35.** `cells_panel.py`: над деревом ячеек заголовок **«Список полей»** (жирный);
в `highlight_items` после подсветки строк по выделению на сцене — `setCurrentItem` +
`scrollToItem(..., PositionAtCenter)` для первого видимого (не скрытого фильтром) поля.

#### T3.8 — Консолидация сетки шаблона (preview + apply) (P1) `[x]`

**Завершено 30.03.2026 15:41.**

- `main_window.py`: добавлены действия **«Консолидировать сетку»** и **«Применить»** в toolbar/меню.
- Консолидация использует `grid_matcher.cluster_lines()` (единый движок с pipeline),
  строит логические X/Y-линии по границам ячеек и показывает overlay:
  - вертикали (синий), горизонтали (оранжевый),
  - проблемные ячейки (красный dashed) при отклонении границы > tolerance.
- Добавлен `Применить`: snap границ всех ячеек к ближайшим логическим линиям, обновление геометрии на canvas.
- `template_meta_panel.py`: поле **Grid tol, мм** (`grid_tolerance_template_mm`) + get/set.
- `cells_panel.py`: для проблемных ячеек в колонке статуса добавляется маркер `⚠`.
- Overlay и предупреждения сбрасываются при следующем пользовательском действии
  (выделение/перемещение/изменение свойств/переключение страницы).

#### T3.9 — Кнопка «Подогнать под сетку» (P1) `[x]`

**Завершено 30.03.2026 15:44.**

- `main_window.py`: добавлено действие **«Подогнать под сетку»** (toolbar + меню).
- Кнопка активируется только при `grid_adapt=True` и открытом PDF (`_update_adapt_action_state`).
- Реализация использует общий движок `grid_matcher.adapt_all_fields()`:
  собирается временный шаблон из текущей геометрии полей, затем поля сцены смещаются
  в адаптированные координаты.
- Overlay диагностики:
  - dotted-стрелки от `approx` к `final`,
  - синие контуры matched detected ячеек,
  - жёлтый контур для `split_merged`,
  - красный контур для `no_match`.
- Лог в консоль: `anchor/matched/split_merged/no_match`, `max_snap_mm`, `moved`.
- Подгонка не выполняет автосохранение — изменения остаются как обычные unsaved edits.

#### T3.10 — Кнопка «Добавить поле» (рисование прямоугольника) (P1) `[x]`

**Завершено 30.03.2026 16:47.**

- `main_window.py`: добавлено действие `Добавить поле` (checkable) в toolbar и меню «Правка».
- `graphics_view.py`: добавлен draw-mode и сигнал `rect_drawn(QRectF)`:
  - в режиме рисования курсор `CrossCursor`,
  - drag ЛКМ рисует preview-рамку,
  - отпускание ЛКМ создаёт финальный `QRectF`.
- `main_window.py::_on_rect_drawn_add_field`:
  - создаёт новый `FieldRectItem` с очередным `cell_index`,
  - автоматически снимает режим рисования (one-shot),
  - переводит фокус на новое поле (выделение + centerOn + загрузка в `PropertiesPanel`),
  - обновляет `CellsPanel` и помечает шаблон modified.

#### T3.11 — Тип `static_empty` + Ctrl-drag copy (P1) `[x]`

**Завершено 30.03.2026 18:19.**

- `models.py`: поддерживается канонический `field_type="static_empty"` (без alias/опечаток).
- `properties_panel.py`: в dropdown типа поля добавлен `static_empty`.
- `cell_items.py`:
  - визуализация `static_empty` как отдельный стиль: dashed-контур и отдельный цвет, заметно отличающийся от `static`;
  - реализовано копирование поля при `Ctrl + drag`:
    при начале перетаскивания создаётся дубликат на исходной позиции с теми же свойствами
    (через deepcopy `FieldDef`), а перетаскиваемый элемент уходит в новую позицию.
- `main_window.py`: добавлен lightweight контроль количества `FieldRectItem` через
  `scene.changed` для автоматического обновления списка при появлении дубликатов.

#### T3.12 — Якоря/снап/preview-only и поля вне штампа (P1) `[x]`

**Завершено 30.03.2026 20:30.**

- `properties_panel.py`: добавлены чекбоксы:
  - `Anchor` (проброс в `FieldDef.is_anchor`);
  - `Вне штампа` (проброс в `FieldDef.outside_stamp`).
- `template_meta_panel.py`: добавлены параметры адаптации:
  - `Detected tol, мм` (`grid_tolerance_detected_mm`);
  - `Snap size tol` (`snap_size_tolerance`);
  - `Snap max dist, мм` (`snap_max_distance_mm`);
  - `Snap IoU` (`snap_iou_threshold`);
  - `Preview only` (`preview_only_adapt`).
- `main_window.py`: кнопка `Подогнать под сетку` теперь:
  - использует `snap_*` и `grid_tolerance_detected_mm` из мета-панели (вместо hardcode);
  - в `preview-only` режиме не двигает поля, а только рисует overlay;
  - не двигает поля со статусом `no_match`;
  - учитывает новый статус `excluded`.
- `grid_matcher.py`: поля `outside_stamp=True` исключены из:
  - расчёта bbox штампа,
  - auto-anchor подбора,
  - snap-подгонки (статус `excluded`).
- `cell_items.py` и `cells_panel.py`: для полей вне штампа добавлен отдельный цвет
  (оранжевый оттенок) и dashed-стиль.
- `models.py`: расширены модели/JSON:
  - `FieldDef.outside_stamp: bool = False`;
  - `StampTemplate.preview_only_adapt: bool = False`.

#### T3.13 — Глобальные допуски `find_tables` из v2 config (P1) `[x]`

**Завершено 30.03.2026 20:43.**

- Добавлен общий модуль `pdf_parsing_v2/find_tables_settings.py`:
  - `get_find_tables_kwargs(cfg)` и `call_find_tables(page, cfg)` с fallback на
    `find_tables()` без kwargs при несовместимой версии PyMuPDF.
- `pdf_v2_config.json` / `v2_config.py`:
  - добавлены глобальные ключи `find_tables_snap_x_tolerance`, `find_tables_snap_y_tolerance`
    (дефолты `2.2` и `2.0`, как в v1).
- `pdf_v2_settings_gui.py`:
  - добавлен раздел `Поиск таблиц (find_tables)` с полями `snap_x_tolerance` / `snap_y_tolerance`,
    сохраняемыми в config.
- Подключение к реальным вызовам:
  - `grid_matcher.get_detected_stamp_cells()` (pipeline + editor adapt),
  - `template_editor/auto_detect.detect_cells()` (авто-разметка),
  - `grid_diagnostic._extract_cells()` (диагностика сетки).
- Для согласованности дефолтов редактора:
  - `main_window._wire_panels()` подхватывает глобальные `snap/grid` дефолты из `self._cfg`.

#### T3.14 — Показ рамки поиска при «Подогнать под сетку» (P1) `[x]`

**Завершено 30.03.2026 20:51.**

- `template_editor/main_window.py`: в overlay `Подогнать под сетку` добавлен
  явный прямоугольник **рамки поиска** (фиолетовый dashed), который соответствует
  bbox фильтрации `find_tables` в `adapt_all_fields` (объединение всех `FieldDef`,
  кроме `outside_stamp=True`).
- В консольный лог `Подгонка под сетку` добавлены координаты `search_bbox=(x0,y0,x1,y1)`,
  чтобы диагностировать случаи, когда верхние ячейки оказываются за пределами области поиска.

#### T3.15 — Деагрессивный matching без ручных якорей (P1) `[x]`

**Завершено 30.03.2026 21:00.**

- `grid_matcher.py`:
  - авто-якоря по «крупным полям» отключены: при отсутствии `is_anchor=True`
    transform от якорей не строится (более предсказуемо для editor UX);
  - snap больше не выбирает дальние ячейки из всего массива, если нет локальных кандидатов
    в `snap_max_distance_mm` (возвращает `no_match`);
  - `split_merged` теперь дополнительно проверяется по IoU merged bbox.
- `main_window.py`: в лог `Подгонка под сетку` добавлен счётчик `manual_anchors`,
  чтобы сразу видеть, почему адаптация работает «без якорной геометрии».

#### T3.16 — Отдельный цвет рамки для anchor-полей (P2) `[x]`

**Завершено 30.03.2026 21:07.**

- `template_editor/cell_items.py`: для `FieldDef.is_anchor=True` добавлен отдельный цвет
  рамки (`_COLOR_ANCHOR`), чтобы якорные поля визуально отличались от обычных назначенных.
- Приоритет цвета: `test_status` (ok/warn/error) выше anchor-цвета.

---

### Группа 3б: Каталог, метаданные шаблона, сохранение

#### T3.4 — Каталог полей пустой + авто-загрузка (P1) `[x]`

**Завершено 29.03.2026 ~23:30.**

1. **`fields_agcc.json` заполнен** — 35 записей со всеми полями из шаблонов
   (1_DOC_TITLE … 51_Page_Format), с `default_clean`, `default_expected`, `short_num`.

2. **Авто-загрузка каталога** при старте редактора (`showEvent` → `_try_auto_load_catalog`):
   сканирует `templates_dir/catalogs/`, загружает первый непустой `.json`.
   Если каталог найден — `PropertiesPanel` получает его, браузер помечает активным,
   статусбар показывает имя файла. Повторная загрузка не происходит если `_catalog` уже задан.

---

#### T3.5 — Параметры шаблона в постоянной панели (P1) `[x]`

**Завершено 29.03.2026 ~23:30.**

Новый файл `template_meta_panel.py` — `TemplateMetaPanel(QWidget)`:
- Группа «Параметры шаблона» в верхней части правой панели (над списком ячеек)
- Поля: имя шаблона, doc_types (2 ряда чекбоксов), page_selector, priority, padding_mm
- `load_from_template(tmpl)` — заполняется при загрузке шаблона из файла
- Сигнал `meta_changed` → `_set_modified(True)`
- Полностью заменяет `SaveTemplateDialog` — значения берутся из панели при сохранении

Изменения в `main_window.py`:
- `_meta_panel` инициализируется в `_wire_panels()` (первый виджет в правой панели)
- `_load_template_from_file` вызывает `_meta_panel.load_from_template(tmpl)`

---

#### T3.6 — Сохранение без диалога + кнопка в блоке параметров (P1) `[x]`

**Завершено 29.03.2026 ~23:30 (кнопка и индикатор добавлены ~23:50).**

- **Ctrl+S** (`act_save_template`) — «Сохранить»: если `_active_template_path` задан,
  сохраняет сразу туда без диалога выбора файла; иначе — ведёт себя как «Сохранить как…»
- **Ctrl+Shift+S** (`act_save_template_as`) — «Сохранить как…»: всегда показывает диалог
- Общая логика вынесена в `_do_save(path)`, `SaveTemplateDialog` больше не используется
- Метаданные (имя, doc_types и т.д.) берутся из `TemplateMetaPanel`, а не из диалога
- Кнопка **«💾 Сохранить»** прямо в блоке «Параметры шаблона» (`save_requested` → `_on_save_template`)
- Индикатор **«● не сохранён»** (оранжевый) рядом с кнопкой: показывается при изменениях,
  скрывается после сохранения; управляется через `_set_modified` → `meta_panel.set_modified`

---

### Группа 4: Исправление координат rotation=90/270

#### T4.1 — get_textbox / get_text используют unrotated-координаты (P1) `[x]`

**Завершено 29.03.2026 ~20:30.** Исправлен фундаментальный баг: `fitz_page.get_text()`
и `fitz_page.get_textbox()` в данной версии PyMuPDF возвращают/принимают координаты
в **неповёрнутом (unrotated)** пространстве оригинального PDF для страниц с rotation=90/270.
При этом `fitz_page.rect` и весь rendering-pipeline работают в **отображаемом (displayed)**
пространстве. Из-за этого все 35 полей шаблона возвращали пустую строку:
поля с displayed x∈[4229,5054] передавались как clip в `get_textbox`, но unrotated
страница имеет ширину 2384 — координаты полностью за пределами.

**Диагноз выявлен через:**
- `find_frame(debug=True)` — вывод всех кандидатов drawings
- Вывод `FrameInfo` (x0/x1/y0/y1/page_w/page_h) в `_on_test`
- Вывод fitz-rect первых 5 полей + пометка «ВНЕ СТРАНИЦЫ»
- Зонд `fitz_page.get_text("text")` + первые 8 непустых спанов с `unrot/disp` bbox

**Исправления в трёх файлах:**

`coord_transform.py` — добавлены две функции конвертации:
- `fitz_displayed_to_unrotated(rect, frame)`:
  для rotation=90/270 преобразует displayed→unrotated перед `get_textbox`;
  формула: `Rect(page_h - disp_y1, disp_x0, page_h - disp_y0, disp_x1)`
- `fitz_unrotated_to_displayed(rect, frame)`:
  обратное преобразование для отрисовки span-bbox в debug PNG;
  формула: `Rect(unrot_y0, page_h - unrot_x1, unrot_y1, page_h - unrot_x0)`
- Обновлён docstring `pdfminer_to_fitz` — явная заметка о различии displayed vs unrotated.
- Обновлён общий docstring модуля — добавлена координатная система 4 (fitz unrotated).

`stamp_extractor.py` — в `_extract_with_template` перед `get_textbox(fitz_rect)`:
```python
query_rect = fitz_displayed_to_unrotated(fitz_rect, frame)
raw_text = fitz_page.get_textbox(query_rect)
```
`fitz_rect` (displayed) по-прежнему сохраняется в `bbox_pts` результата для debug PNG.

`debug_visual.py` — в `_draw_pdf_text_spans`:
- Переключён на `get_text("dict")` (в `rawdict` span-текст в `chars`, не в `text`).
- Функция принимает `frame: FrameInfo | None = None`.
- Каждый span bbox конвертируется `fitz_unrotated_to_displayed(span_bbox, frame)`
  перед рисованием на PNG (иначе спаны рисовались бы в unrotated-координатах на
  displayed-canvas → всё в левом нижнем углу).
- `render_debug_page` передаёт `frame=v2_result.frame` в `_draw_pdf_text_spans`.

**Правило (критично для будущих чатов):**
> Для страниц с `rotation=90` или `rotation=270`:
> - `fitz_page.rect` — displayed пространство (правильно для canvas/pixmap)
> - `fitz_page.get_text()` / `fitz_page.get_textbox()` — unrotated пространство
> - Перед `get_textbox(rect)` ОБЯЗАТЕЛЬНО применять `fitz_displayed_to_unrotated(rect, frame)`
> - Для рисования span-bbox на PNG — ОБЯЗАТЕЛЬНО `fitz_unrotated_to_displayed(bbox, frame)`
> - `get_drawings()` тоже unrotated — это уже учтено в `find_frame`

---

## Архитектурный контекст (кратко)

**Папки:**
```
pdf_parsing_v2/template_editor/   — GUI редактор (PySide6)
pdf_parsing_v2/templates/         — JSON шаблоны
pdf_parsing_v2/templates/catalogs/ — каталоги полей
pdf_parsing_v2/templates/sets/     — наборы шаблонов (TemplateSet)
pdf_v2_config.json                 — общий конфиг (templates_dir и др.)
```

**Ключевые классы:**
- `TemplateEditorWindow(QMainWindow)` — `main_window.py` (координатор)
- `StampGraphicsView(QGraphicsView)` — `graphics_view.py` (PDF + ячейки)
- `FieldRectItem(QGraphicsRectItem)` — `cell_items.py` (ячейка на Canvas)
- `TemplatesBrowserPanel(QWidget)` — `templates_browser.py` (IDE-style левая панель; наборы — дерево
  `template_files`, DnD add/remove, автосохранение JSON набора)
- `CellsPanel(QWidget)` — `cells_panel.py` (заголовок «Список полей», список ячеек, автопрокрутка к строке при выборе на canvas)
- `PropertiesPanel(QWidget)` — `properties_panel.py` (свойства ячейки)
- `CatalogEditorDialog(QDialog)` — `catalog_editor.py`
- `SetEditorDialog(QDialog)` — `set_editor.py`

**Координаты:** scene (pixels при dpi=150) ↔ pts (fitz) ↔ mm от рамки (шаблон).
Scale: `72/150 ≈ 0.48` (scene→pts), `SCALE = 2.83444` pts/mm.
Конвертер: `view.scene_to_mm_from_frame(scene_pos)` → `(mm_from_right, mm_from_bottom)`.

**ВАЖНО — rotation=90/270:**
- `fitz_page.rect` / canvas / pixmap → **displayed** пространство
- `fitz_page.get_text()` / `get_textbox()` / `get_drawings()` → **unrotated** пространство
- Перед `get_textbox` применять `fitz_displayed_to_unrotated(rect, frame)` из `coord_transform.py`
- Для рисования span-bbox применять `fitz_unrotated_to_displayed(bbox, frame)`
- Подробно: T4.1 выше

**Конфиг:** `pdf_v2_config.load_v2_config()` → dict с ключом `templates_dir`.
Единый для редактора и v2 pipeline. Не дублировать путь к шаблонам.

**Layout (29.03.2026):** QSplitter горизонтальный: browser (0) | canvas (1) | right panel (2).
Размеры при запуске: 200 | 900 | 400 px. `PropertiesPanel` принимает `gfx_view` 2-м аргументом.

### Группа 5: Grid adaptation (boundary-snap)

#### T5.1 — Boundary-snap алгоритм вместо IoU-snap (P1) `[x]`

**Завершено 30.03.2026 22:00.**

`grid_matcher.py::adapt_all_fields()` переработан — Фаза A плана каскадной адаптации:

**Новые функции:**
- `_extract_grid_lines(detected, tol)` — консолидированные X/Y grid-линии из detected ячеек.
- `_build_adjacency_graph(fields, frame, tol_mm)` — граф смежности шаблонных полей:
  два поля смежны если расстояние между парой границ < tol И проекция перекрытия по
  перпендикулярной оси > 50% меньшей из двух проекций (защита от ложной смежности через поле между ними).
- `_snap_4_boundaries(bbox, x_lines, y_lines, max_snap_pts)` — привязка каждой из 4 границ
  к ближайшей grid-линии. Нормализация (x0<x1), min_size, max_shape_change проверка.
- `_shape_score(original, adapted)` — 1.0 = без изменений, 0.0 = сильная деформация.
- `_boundary_consistency(field_id, candidate, matched, adjacency)` — score совместимости
  с уже привязанными соседями (1.0 = идеальное совпадение границ).
- `_score_candidate(...)` — комбинированный score: shape×0.3 + position×0.3 + boundary×0.4;
  при отсутствии matched соседей нормализуется до (shape+position)/2.

**Переработан `adapt_all_fields`:**
- Phase 0: extract grid lines
- Phase 1: match anchors → snap к grid lines
- Phase 2: для каждого поля — approx (от якорей) → snap_4_boundaries → score → matched/derived

**Расширен `AdaptResult`:** `shape_score`, `snapped_boundaries` (0-4), `confidence`, `iteration`.
**Новый статус:** `derived` — fallback на approx_bbox при неудачном snap (вместо `no_match` для привязанных).

**`models.py`:** `StampTemplate` + `max_shape_change_ratio` (default 2.5), `cascade_score_threshold` (default 0.4).
Backward-compat: `from_dict` с defaults, `to_dict` сериализует.

**`template_editor/cells_panel.py`:** колонки **Score** и **Iter** (ResizeToContents);
`update_adapt_results(adapted)` — цвета: зелёный > 0.7, жёлтый 0.4-0.7, красный < 0.4, серый derived.

**`template_editor/template_meta_panel.py`:** скрыты из UI `Snap size tol` и `Snap IoU`
(hidden spinboxes для backward compat); добавлены `Max shape change` (1.0-5.0) и `Min score` (0.1-1.0).

**`template_editor/main_window.py`:**
- `can_move = ar.status in ("matched", "anchor")` — `derived` не перемещает.
- Overlay: grid-линии (X синие dotted, Y оранжевые dotted), score-метки рядом с полями
  (A=anchor, D=derived, число=confidence), убраны стрелки approx→final.
- Передаются `max_shape_change_ratio`, `cascade_score_threshold` во все 3 StampTemplate.

**План:** `pdf_parsing_v2/PLAN_cascade_adaptation.md`. Фаза B (каскад) — при необходимости.

---

#### T5.2 — «Подогнать под сетку»: cell-assignment + Hungarian (P1) `[x]`

**Завершено 30.03.2026 23:30 (уточнено 31.03.2026 09:45).**

Редактор вызывает `grid_matcher.adapt_by_cell_assignment()` (не `adapt_all_fields` boundary-snap):
- Фильтр **oversized** detected-ячеек (площадь > 50% площади штампа).
- Глобальный transform: по умолчанию **только сдвиг**; масштаб если отличие от 1.0 > 5%.
- **Назначение поле↔ячейка:** `scipy.optimize.linear_sum_assignment` на матрице стоимости
  `_cell_match_cost` + отсев столбцов без ни одного конечного ребра; при отсутствии **scipy**
  — жадный обход **в порядке полей шаблона** (не sorted-greedy по глобальной стоимости — он ухудшал результат).
- Дубликаты `field.id` устраняются через `_make_unique_keys` / индексированные ключи в смежных структурах.

**Файлы:** `pdf_parsing_v2/grid_matcher.py` (`adapt_by_cell_assignment`, `_solve_cell_assignment_linear_sum`,
`_solve_cell_assignment_order_greedy`).

---

#### T5.3 — «Заполнить зазоры до соседей» вместо «Слепить поля» (P2) `[!]`

**Добавлено 31.03.2026 09:45.** Кнопка в toolbar/меню **«Правка»**; старое действие **«Слепить поля»** удалено.

**Назначение:** после cell-assignment, когда PDF **разбил** одну логическую ячейку на несколько,
шаблонное поле могло сжаться до одного фрагмента — пост-проход **растягивает** прямоугольник до
границ **ближайших соседей** внутри габарита штампа (для визуального контроля в редакторе).

**Файлы:**
- `pdf_parsing_v2/stamp_fill_gaps.py` — `fill_stamp_gaps_for_rects(items, stamp, ...)`: чистая геометрия (tuple rects), без Qt.
- `template_editor/main_window.py` — `_stamp_scene_rect_for_fill()` (габарит штампа = `_compute_stamp_search_bbox` → сцена),
  `_on_fill_stamp_gaps()`: только назначенные поля, **не** `outside_stamp`; допуск `eps_scene` из **Tol мм** в toolbar.

**Алгоритм (текущий):**
- **Горизонталь:** расширение **только если** у поля есть **и левый, и правый** сосед
  (вертикальное перекрытие ≥ `overlap_frac` × меньшая высота); тогда
  `x0 = max(stamp, left.x1)`, `x1 = min(stamp, right.x0)` при `left.x1 < right.x0`.
- **Вертикаль:** аналогично — **только если** есть **верхний и нижний** сосед (горизонтальное перекрытие).
- Одна проходка по **исходной** геометрии (без каскада по уже расширенным rect).

**Ограничение (явно):** **не** подключено к основному `v2_pipeline` / `stamp_extractor` — только редактор.

**Статус `[!]` — известная проблема:** пользователь сообщил **некорректную работу** в реальных шаблонах
(ложные соседи, лишние растяжения, крайние случаи строк/колонок). **Следующие чаты:** доработать логику
соседства и/или ограничения (по `field_id`, по строке подписей, по порядку слоёв), затем обновить статус.
До стабилизации **не** включать в общий цикл v2.

---

## Правила для чатов

1. **Перед изменением `main_window.py`** — прочитать все файлы `template_editor/`,
   т.к. виджеты связаны сигналами между собой
2. **Логирование** — добавлять `self._log(...)` во все ключевые точки, не `print` напрямую
3. **Путь к шаблонам** — всегда брать из `self._cfg["templates_dir"]`, не хардкодить
4. **После задачи** — обновить статус `[ ]` → `[x]` в бэклоге и дату
5. **PySide6 специфика** — не смешивать потоки (QTimer вместо sleep), не вызывать
   Qt-методы из не-Qt потоков
6. **Адаптация / зазоры (31.03.2026):** см. **T5.2** (Hungarian + `adapt_by_cell_assignment`) и **T5.3**
   (`stamp_fill_gaps` + кнопка «Заполнить зазоры»). T5.3 помечен `[!]` — поведение на доработке;
   перед правками перечитать `stamp_fill_gaps.py` и связанный код в `main_window.py`.
