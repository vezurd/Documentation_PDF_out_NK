# Промпт: v2 domain package + GUI мониторинг параллели

> Скопировать целиком в новый чат Opus 4.6. После вставки удалить этот файл.

---

## Роль

Ты — Opus 4.6, PREMIUM-архитектор и ревьюер. Тебе доступны subagent-ы
(Task tool, subagent_type="generalPurpose", model="fast") для C2-работы:
создания/редактирования файлов, портирования кода, прогона тестов.

Ты **сам** выполняешь:
- [PREMIUM] дизайн, спеки, архитектурные решения, dataclass-модели
- [V] ревью результатов subagent-ов (читай созданные файлы и проверяй)
- Финальную верификацию (import smoke tests, rg по зависимостям)

Subagent-ы выполняют:
- [C2] Создание/редактирование файлов по твоим спекам
- [C2] Портирование кода, обновление импортов
- [C2] Прогоны тестов, grep-проверки

Запускай subagent-ов **параллельно** где возможно.
После каждого subagent-а **верифицируй** результат (прочитай файл, проверь импорты).

---

## Цель: полная отвязка v2 от v1 + GUI мониторинг

**Конечная цель пользователя:** после завершения этой работы и ручного тестирования
пользователь планирует **удалить `pdf_parsing/` (v1) из проекта целиком**.
Все v2-пакеты должны быть полностью самодостаточными.

---

## ЧАСТЬ 1: Domain package — замена `doc_ATTRIBUTES` и `PageStampAttributes`

### Текущее состояние (что уже сделано)

Прочитай план: `.cursor/plans/v2_decoupling_parallel.plan.md` — все шаги completed.

Константы `c_*` уже скопированы в `pdf_parsing_v2_engine/stamp_fields.py`.
Но два класса из `pdf_parsing/shtamp_extract_classes.py` всё ещё в v1:

1. **`doc_ATTRIBUTES`** — модель документа (создаётся в `utils.path.get_files_single`)
2. **`PageStampAttributes`** — модель страницы (создаётся в `compat.py`)

### Оставшиеся зависимости от v1 (полный список)

```
# Runtime:
pdf_parsing_v2_engine/compat.py:7    → from pdf_parsing.shtamp_extract_classes import PageStampAttributes

# TYPE_CHECKING only (станут runtime после порта):
pdf_parsing_v2/v2_pipeline.py:35     → from pdf_parsing.shtamp_extract_classes import doc_ATTRIBUTES
pdf_parsing_v2_rules/checks.py:54    → from pdf_parsing.shtamp_extract_classes import doc_ATTRIBUTES
pdf_parsing_v2_tags/extract.py:17    → from pdf_parsing.shtamp_extract_classes import doc_ATTRIBUTES

# utils.path (создаёт doc_ATTRIBUTES):
utils/path.py:12                     → from pdf_parsing.shtamp_extract_classes import doc_ATTRIBUTES
utils/path.py:164                    → curr_proj.append(doc_ATTRIBUTES(file_path))
```

### Зависимости `doc_ATTRIBUTES.__init__` от `utils.string_parsing`

Конструктор `doc_ATTRIBUTES(file_path)` вызывает:
- `string_parsing.getDocTypeFromFile(file_name)` → `doc_Type`
- `string_parsing.getDocTitle(file_name)` → `doc_Title_4d`
- `string_parsing.getMarkaFromFileName(file_name)` → `doc_Marka`
- `string_parsing.getDocNumber(file_name)` → `doc_Number`
- `string_parsing.getRevisionFromFileName(file_name)` → `doc_Revision`
- `string_parsing.getOdStyleFileName(file_name)` → `doc_OD_style_file_name`
- `ProjectFileName.scan_title_system(file_name)` → `doc_Short_Title`

Все эти функции — чистые парсеры строк (stateless), живут в `utils/string_parsing.py`
и `utils/file_name_converts.py`. Пакет `utils/` остаётся shared.

### Задача

**Спроектируй и реализуй** замену `doc_ATTRIBUTES` и `PageStampAttributes`
в v2-пакетах. Принципы:

1. **Не разбрасывать сущности.** Новые модели должны жить в ОДНОМ месте,
   а не размазаны по 5 пакетам. Рекомендуемое место — `pdf_parsing_v2_engine/models.py`
   (там уже живут `V2PageResult`, `FieldDef`, `StampTemplate` и др.)
   или отдельный `pdf_parsing_v2_engine/document.py` если models.py станет слишком большим.

2. **`V2Document`** (замена `doc_ATTRIBUTES`):
   - `@dataclass` (не class с ручным `__init__`)
   - Поля: `file_full_path`, `file_name`, `doc_Type`, `doc_Marka`, `doc_Number`,
     `doc_Revision`, `doc_Title_4d`, `doc_Short_Title`, `doc_Short_File_Name`,
     `doc_OD_style_file_name`, `doc_Number_for_sort`
   - `pages: list[V2PageData]` (а не `list[PageStampAttributes]`)
   - `_v2_results: list[V2PageResult]` (как сейчас, но типизировано)
   - **Factory method** `V2Document.from_file_path(file_path: str) -> V2Document`
     — вызывает те же `string_parsing` функции для парсинга имени файла
   - pickle-safe (для будущего использования в PPE)

3. **`V2PageData`** (замена `PageStampAttributes`):
   - `@dataclass`
   - `page_num: int`
   - `page_type: str`
   - `page_marka: str`
   - `page_title: str`
   - `dict_attributes: dict[str, list]` — ТОТ ЖЕ формат что у PageStampAttributes
     (ключи = c_* строки, значения = list). Это нужно для совместимости с `checks.py`
     (20 проверок нормоконтроля работают с `page.dict_attributes[c_18_1][0]` и т.д.)
   - Factory: `V2PageData.from_v2_result(v2r: V2PageResult, doc: V2Document) -> V2PageData`
     — содержит ту же логику что сейчас в `compat.to_page_stamp_attributes`

4. **Обновить `utils/path.py`:**
   - `get_files_single` → возвращает `list[V2Document]`
   - Импорт из v2_engine вместо `pdf_parsing.shtamp_extract_classes`

5. **Обновить `compat.py`:**
   - Функция `to_page_stamp_attributes` → `to_v2_page_data`
   - Убрать импорт `PageStampAttributes` из v1
   - Вся конвертация metadata → dict_attributes остаётся (логика из `_psa_list_from_metadata`)

6. **Обновить все TYPE_CHECKING:**
   - `v2_pipeline.py`, `checks.py`, `extract.py` → `from pdf_parsing_v2_engine.document import V2Document`
   - Заменить TYPE_CHECKING на обычный import (теперь это v2-native класс)

7. **`checks.py`:** заменить `doc_ATTRIBUTES` → `V2Document`, `page.dict_attributes` → тот же API.
   Проверки нормоконтроля работают с `page.dict_attributes[key][0]` — API `V2PageData`
   должен быть полностью совместим.

8. **Не удалять `pdf_parsing/shtamp_extract_classes.py`** — пользователь удалит v1 сам.
   Просто убрать все импорты из v2-пакетов.

### Verification после ЧАСТИ 1

```bash
# Нуль импортов из pdf_parsing в v2 пакетах:
rg "from pdf_parsing\." --type py -g "pdf_parsing_v2*/**" -g "pdf_template_editor/**"
# Ожидание: только compare_v1_v2.py (утилита сравнения, не production)

# Smoke tests:
python -c "from pdf_parsing_v2_engine.document import V2Document, V2PageData"
python -c "from pdf_parsing_v2_rules.checks import run_all_checks"
python -c "from pdf_parsing_v2_engine.compat import to_v2_page_data"
python -c "from pdf_parsing_v2.v2_pipeline import run_v2_pipeline"
python -c "from utils.path import get_files_single"
```

### ВАЖНО: `pdf_template_editor/`

Прочитай `pdf_template_editor/main_window.py` — там используется
`from utils.string_parsing import getDocTypeFromFile`. Это зависимость от `utils/`,
а НЕ от `pdf_parsing/` — не трогай. Но проверь нет ли скрытых импортов из `pdf_parsing.*`.

---

## ЧАСТЬ 2: GUI мониторинг параллели (фаза 5)

### Контекст

Прочитай эти файлы:
- `.cursor/rules/AI_v2_parallel.mdc` — текущая спека (фазы 4A+4B реализованы)
- `pdf_parsing_v2/SPEC_parallel.md` — спека parallel.py (DTOs, callback, worker)
- `pdf_parsing_v2/parallel.py` — реализация (ParallelCallback protocol)
- `pdf_parsing_v2/v2_timing.py` — TimingCollector

### Цель

Создать **компактный** GUI-модуль мониторинга `pdf_v2_monitor/` (PySide6)
для наблюдения за процессом v2 pipeline в реальном времени.

### Требования к GUI

**Дизайн: минимальный, информативный, приятный.**

1. **Одно окно** (QDialog или QMainWindow), можно модальное.
   Открывается из `main.py` кнопкой «Запуск v2 с мониторингом».

2. **Верхняя панель — общий прогресс:**
   - Прогресс-бар: `N / M файлов` (обновляется per-file)
   - Текущая фаза: `Extraction` / `Tags` / `Postprocess` (три этапа pipeline)
   - Общее время: обновляется каждую секунду

3. **Центральная панель — таблица файлов:**
   - Столбцы: `#`, `Файл`, `Статус`, `Страниц`, `Время (с)`, `Ошибка`
   - Статус: иконка/цвет: ⏳ в очереди, 🔄 обработка, ✅ готово, ❌ ошибка
   - Сортировка по порядку исходного индекса
   - При ошибке — красная строка, текст ошибки в столбце (tooltip = полный traceback)
   - Автоскролл к последнему обновлённому файлу

4. **Нижняя панель — сводка:**
   - Общее время / Extraction / Tags / Postprocess
   - Количество ошибок (красным если > 0)
   - Кнопка «Сохранить лог» → `v2_timing_log.xlsx` (через `TimingCollector.to_excel`)
   - Кнопка «Закрыть» (активна только после завершения)

5. **Архитектура связки с pipeline:**
   - Реализовать `MonitorCallback` (класс, реализующий `ParallelCallback` protocol)
   - Callback-методы вызываются из **main process** (не из workers) →
     безопасно для Qt signals
   - `MonitorCallback` эмитит Qt signals → GUI обновляется в GUI thread
   - Pipeline запускается в **QThread** (не в main thread, чтобы GUI не замерзал)
   - `run_v2_pipeline` уже поддерживает `cfg["parallel"]` — нужно добавить
     возможность передать callback

6. **Не усложнять:**
   - Никаких графиков/чартов — только таблица + текст + прогрессбары где уместно
   - Никаких настроек внутри монитора — они в pdf_v2_settings_gui
   - Пакет `pdf_v2_monitor/`: `__init__.py`, `monitor_window.py`, `monitor_callback.py`
   - Интеграция в `main.py`: одна кнопка, один импорт

### Доработка pipeline для callback-пробрасывания

Сейчас `run_v2_pipeline` не принимает callback. Нужно:

```python
def run_v2_pipeline(
    pdf_path: str,
    cfg: dict[str, Any],
    project: str | None = None,
    *,
    extraction_callback=None,  # NEW
) -> str:
```

И пробросить в `run_v2_extraction` / `run_parallel_extraction`.

### Примерная структура пакета

```
pdf_v2_monitor/
├── __init__.py              # реэкспорт MonitorWindow
├── monitor_window.py        # QDialog/QMainWindow с таблицей
├── monitor_callback.py      # MonitorCallback (ParallelCallback + Qt signals)
└── pipeline_thread.py       # QThread обёртка над run_v2_pipeline
```

### Verification

- Запуск `python -m pdf_v2_monitor` (standalone demo с mock-данными)
- Интеграция в `main.py` — кнопка работает, GUI обновляется per-file

---

## Порядок работы

| Этап | Описание | Модель |
|------|----------|--------|
| 1 | [PREMIUM] Спроектировать `V2Document`, `V2PageData` dataclasses | Ты сам |
| 2 | [C2] Создать dataclasses + factory methods | Subagent |
| 3 | [C2] Обновить `compat.py` → `to_v2_page_data` | Subagent |
| 4 | [C2] Обновить `utils/path.py` → `V2Document` | Subagent |
| 5 | [C2] Обновить импорты: v2_pipeline, checks, extract, parallel | Subagent |
| 6 | [V] Верификация: rg + smoke tests | Ты сам |
| 7 | [PREMIUM] Спроектировать GUI монитор | Ты сам |
| 8 | [C2] Создать `pdf_v2_monitor/` по спеке | Subagent |
| 9 | [C2] Доработать pipeline (callback param) | Subagent |
| 10 | [C2] Интеграция в main.py | Subagent |
| 11 | [V] Финальная верификация | Ты сам |

Этапы 2–5 можно параллелить (после этапа 1).
Этапы 8–10 можно параллелить (после этапа 7).

---

## Ключевые файлы для чтения

**Обязательно прочитай перед началом работы:**

- `pdf_parsing/shtamp_extract_classes.py` — текущие doc_ATTRIBUTES + PageStampAttributes
- `pdf_parsing_v2_engine/compat.py` — конвертер V2PageResult → PageStampAttributes
- `pdf_parsing_v2_engine/stamp_fields.py` — c_* константы (уже в v2)
- `pdf_parsing_v2_engine/models.py` — V2PageResult, FieldDef и др.
- `utils/path.py` — get_files_single (создаёт doc_ATTRIBUTES)
- `utils/string_parsing.py` — парсеры имён файлов
- `pdf_parsing_v2/v2_pipeline.py` — run_v2_pipeline, run_v2_extraction
- `pdf_parsing_v2/parallel.py` — run_parallel_extraction, ParallelCallback
- `pdf_parsing_v2/v2_timing.py` — TimingCollector
- `pdf_parsing_v2_rules/checks.py` — 20 проверок (page.dict_attributes[key][0])
- `pdf_parsing_v2_tags/extract.py` — parse_tags
- `.cursor/rules/AI_v2_parallel.mdc` — спека параллели
- `.cursor/plans/v2_decoupling_parallel.plan.md` — завершённый план
- `pdf_parsing_v2/SPEC_parallel.md` — спека parallel.py
- `main.py` — GUI точка входа (для интеграции кнопки монитора)
