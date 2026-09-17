# Спека: `pdf_parsing_v2_rules` — нормоконтроль v2 с профилями проектов

> **Шаг 4A [PREMIUM]** из плана `v2_packages_split`. Для реализации шагом 4B [C2].

## Текущее состояние

`pdf_parsing/rules_check.py` — монолитная функция `rules_check(curr_proj, proj_od_list)`
(~990 строк, 20 проверок c_code 100–1019).

**Project-specific захардкожено:**
- `project_name="AGCC"` в проверке 1000 (имя файла).
- `rev_literal`, `rev_numeral` — допустимые коды ревизий.
- AN-логика (5.3.x процедуры AGCC) в проверках 1005, 1013, 1017.
- IFC/IFR/SUP/CAN текстовые шаблоны в проверке 1017.

---

## 1. `RulesProfile`

```python
@dataclass(frozen=True)
class RulesProfile:
    """Project-specific settings for normcontrol checks."""

    id: str                          # "AGCC.287", "default", …
    project_name: str                # для ProjectFileName.find_full_matches
    list_of_bbb: tuple[str, ...]     # типы документов BBB (BOM, BOE, BOQ, MTO)

    # Ревизии
    rev_literal: tuple[str, ...]     # ("A","B","C","D","E","F","G")
    rev_numeral: tuple[str, ...]     # ("0","01","02",…,"AN01","AN02",…)

    # IFC/IFR шаблоны (1017)
    text_ifc: str = "IFC - Выпущен для строительства"
    text_ifr: str = "IFR - Выпущено для рассмотрения"
    text_sup: str = "SUP - Заменен"
    text_can: str = "CAN - Аннулирован"

    # Какие проверки включены (None = все)
    enabled_checks: frozenset[int] | None = None
```

**frozen=True** — dataclass неизменяемый; безопасен для PPE (pickle/share).

---

## 2. Registry и `get_profile`

```python
_AGCC_PROFILE = RulesProfile(
    id="AGCC.287",
    project_name="AGCC",
    list_of_bbb=("BOM", "BOE", "BOQ", "MTO"),
    rev_literal=("A", "B", "C", "D", "E", "F", "G"),
    rev_numeral=("0", "01", "02", "03", "04", "05", "06", "07", "08", "09",
                 "AN01", "AN02", "AN03", "AN04", "AN05"),
)

_DEFAULT_PROFILE = _AGCC_PROFILE  # первый проект = единственный

_REGISTRY: dict[str, RulesProfile] = {
    "AGCC.287": _AGCC_PROFILE,
    "AGCC": _AGCC_PROFILE,         # alias по project_name
}


def get_profile(project: str | None) -> RulesProfile:
    """Lookup profile by project id or name.

    Args:
        project: project string from ``pdf_v2_config.json`` / CLI.

    Returns:
        Matching profile, or ``_DEFAULT_PROFILE`` with a warning
        if *project* is unknown or None.
    """
    if not project:
        return _DEFAULT_PROFILE
    key = project.strip()
    if key in _REGISTRY:
        return _REGISTRY[key]
    # case-insensitive fallback
    for k, v in _REGISTRY.items():
        if k.lower() == key.lower():
            return v
    print(
        f"[rules] Unknown project {project!r}, "
        f"using default profile {_DEFAULT_PROFILE.id!r}"
    )
    return _DEFAULT_PROFILE
```

**Fallback policy:** unknown project → default profile + warning (не `ValueError`).
Обоснование: pipeline не должен падать из-за нового/неизвестного проекта;
неполные проверки лучше, чем их отсутствие.

---

## 3. Публичный вход

```python
def rules_check_start_v2(
    curr_proj: list[doc_ATTRIBUTES],
    proj_od_list: list | int,
    pdf_path: str,
    path_out_dir: str,
    *,
    profile: RulesProfile | None = None,
    effective_project: str | None = None,
) -> None:
    """Run normcontrol checks and write Excel report.

    Args:
        curr_proj: documents with extracted stamps.
        proj_od_list: OD entries (or 0 if OD not found).
        pdf_path: root PDF folder.
        path_out_dir: output directory for report.
        profile: explicit profile; if None — resolved from *effective_project*.
        effective_project: project id for profile lookup (from config/CLI).
    """
    if profile is None:
        profile = get_profile(effective_project)
    check_list = run_all_checks(curr_proj, proj_od_list, profile)
    flagged = print_check_list(check_list, pdf_path)
    excel_check_list_out(flagged, curr_proj, path_out_dir)
```

---

## 4. Структура пакета

```
pdf_parsing_v2_rules/
├── __init__.py          # re-export: rules_check_start_v2, get_profile, RulesProfile
├── profile.py           # RulesProfile dataclass, _REGISTRY, get_profile
├── checks.py            # run_all_checks → list[CheckRow]; отдельные функции по группам
├── output.py            # check_row, print_check_list, excel_check_list_out, formats_convolution
└── _an_revision.py      # _an_revision_long_from_18_1_rows (выделен, т.к. нетривиальный)
```

### `checks.py` — группировка проверок

Монолитная `rules_check` разбивается на именованные функции:

| Функция | c_code | Описание |
|---------|--------|----------|
| `check_file_names` | 1000 | Имя файла через `ProjectFileName` + `profile.project_name` |
| `prepare_revision_selection` | 100 | Выбор рабочей ревизии из 18_1_* |
| `check_cyrillic` | 1001 | Кириллица в полях |
| `check_layers` | 1002 | Многослойность |
| `check_annotations` | 1003 | Аннотации/закладки |
| `check_format_vs_real` | 1004 | Формат штампа vs реальный |
| `check_revisions_internal` | 1005 | Ревизии внутри документа (вкл. AN) |
| `check_revisions_vs_od` | 1006 | Ревизии doc vs ОД |
| `check_formats_vs_od` | 1007 | Форматы doc vs ОД |
| `check_page_count_stamp` | 1008 | Кол-во листов штамп vs реальное |
| `check_doc_numbers` | 1009 | Номера документа |
| `check_page_numbers` | 1010 | Номера страниц |
| `check_doc_name_vs_od` | 1011 | Наименование doc vs ОД |
| `check_facility_title` | 1012 | Титул/объект vs ОД |
| `check_dates` | 1013 | Даты (10 vs 18.2), AN-логика |
| `check_doc_code` | 1014 | Шифр (1_DOC_TITLE) |
| `check_od_page_count` | 1015 | Кол-во листов ОД (c_7) |
| `check_pdf_extension` | 1016 | Расширение .pdf |
| `check_ifc_ifr` | 1017 | IFC/IFR/SUP/CAN текст |
| `check_mark` | 1018 | Марка consistency |
| `check_duplicate_ext` | 1019 | Задвоение расширений |

Каждая функция:

```python
def check_file_names(
    curr_proj: list[doc_ATTRIBUTES],
    profile: RulesProfile,
) -> list[CheckRow]:
```

**`run_all_checks`** — compose:

```python
def run_all_checks(
    curr_proj: list[doc_ATTRIBUTES],
    proj_od_list: list | int,
    profile: RulesProfile,
) -> list[CheckRow]:
    """Run all enabled checks; return flat list of results."""
    enabled = profile.enabled_checks
    results: list[CheckRow] = []
    brake = False

    if _is_enabled(1000, enabled):
        rows = check_file_names(curr_proj, profile)
        results.extend(rows)
        if any(not r.result for r in rows):
            brake = True

    if not brake and _is_enabled(100, enabled):
        results.extend(prepare_revision_selection(curr_proj, profile))

    # … остальные check_ по порядку …
    return results
```

`brake_flag` из v1 сохранён: если имена файлов некорректны —
остальные проверки пропускаются (зависимость от корректного парсинга имён).

---

## 5. `CheckRow` (замена list → dataclass)

```python
@dataclass
class CheckRow:
    """One normcontrol check result."""
    result: bool
    c_code: int
    c_description: str
    doc_name: str
    page_num: int | str      # int или "нет"
    text: str
```

В v1 — `[result, c_code[0], c_code[1], doc, page_num, text]` (list).
Замена на dataclass → читаемость, type safety, pickle-safe.
Excel-вывод адаптируется в `output.py`.

---

## 6. Зависимости от v1

| Что | Откуда | Комментарий |
|-----|--------|-------------|
| `c_*` константы | `pdf_parsing.shtamp_extract_classes` | ~30 констант; перенос — отдельная задача |
| `doc_ATTRIBUTES`, `PageStampAttributes` | `pdf_parsing.shtamp_extract_classes` | TYPE_CHECKING |
| `list_of_BBB` | ~~`elements_coordinates`~~ → `pdf_parsing_v2_engine.doc_types.LIST_OF_BBB` | **Отвязано** |
| `doc_Osnovnoi`, `doc_Prilagaemie` | ~~`OD_tab_parsing`~~ → `pdf_parsing_v2_od.DOC_OSNOVNOI/DOC_PRILAGAEMIE` | **Отвязано** |
| `string_parsing.*` | `utils.string_parsing` | Utility, оставляем |
| `ProjectFileName` | `utils.file_name_converts` | Utility, оставляем |
| Excel шаблон | `templates/out_template_check_list.xlsx` | Путь: относительно project root |

---

## 7. Подключение в pipeline

В `pdf_parsing_v2/v2_pipeline.py` → `run_v2_postprocess`:

```python
# Заменить:
#     from pdf_parsing import rules_check
#     rules_check.rules_check_start(curr_proj, proj_od_list, pdf_path, out_result_dir)
# На:
    from pdf_parsing_v2_rules import rules_check_start_v2
    rules_check_start_v2(
        curr_proj, proj_od_list, pdf_path, out_result_dir,
        effective_project=effective_project,
    )
```

`effective_project` уже доступен как параметр `run_v2_postprocess`.

---

## 8. Что НЕ делать в этой волне

| Что | Когда |
|-----|-------|
| Переносить `c_*` из `shtamp_extract_classes` в общий domain | Отдельная задача |
| Перерабатывать AN-логику | При запросе от пользователя |
| Добавлять новые проверки | По мере необходимости |
| Разносить checks по файлам `checks/*.py` | Если `checks.py` > 800 строк |

---

## 9. Verification (для C2)

1. `python -c "from pdf_parsing_v2_rules import rules_check_start_v2, get_profile, RulesProfile"` — OK.
2. `get_profile("AGCC.287").project_name == "AGCC"` — OK.
3. `get_profile(None) is get_profile("unknown")` — оба `_DEFAULT_PROFILE`, warning для "unknown".
4. Smoke `v2_pipeline.py` на папке PDF → сравнить `_результат_проверки.xlsx`:
   тот же набор строк (c_code, doc, page_num), те же result-значения.
5. Отдельный кейс: `project=None` → pipeline не падает, rules отрабатывают с default profile.

---

## 10. Инструкция для C2 (шаг 4B)

1. Создать `pdf_parsing_v2_rules/` с файлами: `__init__.py`, `profile.py`, `checks.py`, `output.py`, `_an_revision.py`.
2. В `profile.py`: `RulesProfile` (frozen dataclass), `_REGISTRY`, `get_profile` по спеке выше.
3. В `_an_revision.py`: перенести `_an_revision_long_from_18_1_rows` и RE-компиляцию.
4. В `output.py`: `CheckRow` dataclass, `check_row` (конструктор), `print_check_list`, `excel_check_list_out`, `formats_convolution`.
5. В `checks.py`: `run_all_checks` + 20 функций `check_*` по таблице §4. Каждая принимает `curr_proj` (и/или `proj_od_list`) + `profile`. Импорты `c_*` из `pdf_parsing.shtamp_extract_classes`, `LIST_OF_BBB` из `pdf_parsing_v2_engine.doc_types`, `DOC_OSNOVNOI`/`DOC_PRILAGAEMIE` из `pdf_parsing_v2_od`.
6. В `__init__.py`: re-export `rules_check_start_v2`, `get_profile`, `RulesProfile`.
7. Переключить `v2_pipeline.py` (`run_v2_postprocess`): `from pdf_parsing_v2_rules import rules_check_start_v2`.
   Добавить `effective_project=effective_project` в вызов.
8. **Verification** по §9.

**Файлы для чтения:** `pdf_parsing/rules_check.py` (оригинал), `pdf_parsing/shtamp_extract_classes.py` (константы `c_*`), `pdf_parsing_v2_engine/doc_types.py`, `pdf_parsing_v2_od/__init__.py`, `pdf_parsing_v2/v2_pipeline.py` (текущий).
