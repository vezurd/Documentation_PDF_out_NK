"""Native PySide6 editor for MTO/BBB analysis JSON (``bbb_analysis_config.json``)."""

from __future__ import annotations

import os
from dataclasses import dataclass
from typing import Any, Callable

from PySide6.QtWidgets import (
    QCheckBox,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QMessageBox,
    QPushButton,
    QScrollArea,
    QVBoxLayout,
    QWidget,
)

from base.bbb_config import load_config, save_config
from base.bbb_section_types_excel import get_excel_path


@dataclass(frozen=True)
class _FieldSpec:
    key: str
    label: str
    default: bool = True


_SECTIONS: tuple[tuple[str, tuple[_FieldSpec, ...]], ...] = (
    (
        "Ограничения выбора папки",
        (
            _FieldSpec(
                "folder_rules.search_only_in_dwg",
                "Искать только в папке DWG",
            ),
        ),
    ),
    (
        "Проверка MTO по code_base (mto_vs_code_base)",
        (
            _FieldSpec("mto_vs_code_base.enabled", "Включить mto_vs_code_base"),
            _FieldSpec(
                "mto_vs_code_base.check_by_code",
                "check_by_code — сравнение NAME/TYPE_MARK/UNITS/VENDOR/CODE",
            ),
            _FieldSpec(
                "mto_vs_code_base.check_equipment_codes",
                "check_equipment_codes — проверка кодов оборудования",
            ),
            _FieldSpec(
                "mto_vs_code_base.check_tags_value",
                "check_tags_value — проверка тегов и значений",
            ),
            _FieldSpec(
                "mto_vs_code_base.check_value",
                "check_value — проверка количеств (VALUES)",
            ),
            _FieldSpec(
                "mto_vs_code_base.check_duplicate_tags",
                "check_duplicate_tags — проверка дублей тегов",
            ),
            _FieldSpec(
                "mto_vs_code_base.check_mass",
                "check_mass — проверка массы",
            ),
            _FieldSpec(
                "mto_vs_code_base.check_prohibition",
                "check_prohibition — проверка запрещённых кодов",
                default=False,
            ),
            _FieldSpec(
                "mto_vs_code_base.check_tags_4_2_4_4",
                "check_tags_4_2_4_4 — проверка 94S 4.2 / 4.4",
            ),
            _FieldSpec(
                "mto_vs_code_base.check_position_numeration",
                "check_position_numeration — проверка нумерации позиций",
            ),
        ),
    ),
    (
        "Коррекция MTO (mto_vs_code_base_correction)",
        (
            _FieldSpec(
                "mto_vs_code_base_correction.enabled",
                "Включить mto_vs_code_base_correction",
            ),
            _FieldSpec(
                "mto_vs_code_base_correction.check_by_code",
                "check_by_code — коррекция по code_base",
            ),
            _FieldSpec(
                "mto_vs_code_base_correction.check_tags_value",
                "check_tags_value — проверка тегов и значений",
            ),
            _FieldSpec(
                "mto_vs_code_base_correction.check_value",
                "check_value — проверка количеств (VALUES)",
            ),
            _FieldSpec(
                "mto_vs_code_base_correction.check_duplicate_tags",
                "check_duplicate_tags — проверка дублей тегов",
            ),
            _FieldSpec(
                "mto_vs_code_base_correction.check_tags_4_2_4_4",
                "check_tags_4_2_4_4 — проверка 94S 4.2 / 4.4",
            ),
            _FieldSpec(
                "mto_vs_code_base_correction.check_position_numeration",
                "check_position_numeration — проверка нумерации позиций",
            ),
        ),
    ),
    (
        "Проверка по code_base (bbb_vs_code_base)",
        (
            _FieldSpec("bbb_vs_code_base.enabled", "Включить bbb_vs_code_base"),
            _FieldSpec(
                "bbb_vs_code_base.check_by_code",
                "check_by_code — сравнение NAME/TYPE_MARK/UNITS/VENDOR/CODE",
            ),
            _FieldSpec(
                "bbb_vs_code_base.check_equipment_codes",
                "check_equipment_codes — проверка кодов оборудования",
            ),
            _FieldSpec(
                "bbb_vs_code_base.check_tags_value",
                "check_tags_value — проверка тегов и значений",
            ),
            _FieldSpec(
                "bbb_vs_code_base.check_value",
                "check_value — проверка количеств (VALUES)",
            ),
            _FieldSpec(
                "bbb_vs_code_base.check_duplicate_tags",
                "check_duplicate_tags — проверка дублей тегов",
            ),
            _FieldSpec(
                "bbb_vs_code_base.check_mass",
                "check_mass — проверка массы",
            ),
            _FieldSpec(
                "bbb_vs_code_base.check_prohibition",
                "check_prohibition — проверка запрещённых кодов",
                default=False,
            ),
            _FieldSpec(
                "bbb_vs_code_base.check_work_code_and_mtr_group",
                "check_work_code_and_mtr_group — проверка BBB_WORK_CODE и "
                "BBB_MTR_GROUP по GoogleBase",
            ),
        ),
    ),
    (
        "Сравнение с MTO (bbb_vs_mto)",
        (
            _FieldSpec("bbb_vs_mto.enabled", "Включить bbb_vs_mto"),
            _FieldSpec(
                "bbb_vs_mto.compare_fields",
                "compare_fields — сравнение NAME/TYPE_MARK/VENDOR/UNITS с MTO",
            ),
            _FieldSpec(
                "bbb_vs_mto.compare_values",
                "compare_values — сравнение сумм VALUES BBB↔MTO",
            ),
            _FieldSpec(
                "bbb_vs_mto.compare_tags",
                "compare_tags — сравнение тегов BOE↔MTO",
            ),
            _FieldSpec(
                "bbb_vs_mto.compare_title_marka_revision",
                "compare_title_marka_revision — сравнение TITLE/BBB_MARKA/"
                "BBB_REVISION с именем файла MTO",
            ),
            _FieldSpec(
                "bbb_vs_mto.compare_bom_annotation_with_mto",
                "compare_bom_annotation_with_mto — проверка BOM.ANNOTATION "
                "по сумме слагаемых MTO",
            ),
            _FieldSpec(
                "bbb_vs_mto.check_missing_mto_codes",
                "check_missing_mto_codes — добавлять строки для кодов MTO, "
                "отсутствующих в BOE и BOM",
            ),
        ),
    ),
)

_BBB_VS_MTO_TITLE = "Сравнение с MTO (bbb_vs_mto)"


def _nested_get(config: dict[str, Any], key: str, default: Any) -> Any:
    value: Any = config
    for part in key.split("."):
        if not isinstance(value, dict) or part not in value:
            return default
        value = value[part]
    return value


def _nested_set(config: dict[str, Any], key: str, value: Any) -> None:
    target = config
    parts = key.split(".")
    for part in parts[:-1]:
        child = target.get(part)
        if not isinstance(child, dict):
            child = {}
            target[part] = child
        target = child
    target[parts[-1]] = value


class BbbSettingsPanel(QWidget):
    """Native editor for MTO/BBB flags; merge-on-save keeps unknown JSON keys."""

    def __init__(
        self,
        parent: QWidget | None = None,
        *,
        on_saved: Callable[[], None] | None = None,
    ) -> None:
        super().__init__(parent)
        self._on_saved = on_saved
        self._widgets: dict[str, QCheckBox] = {}
        self._specs: dict[str, _FieldSpec] = {}

        scroll = QScrollArea(self)
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QScrollArea.Shape.NoFrame)

        inner = QWidget()
        layout = QVBoxLayout(inner)
        hint = QLabel(
            "Те же флаги, что окно CTk «Настройки: Проверка MTO + BBB». "
            "Сохранение патчит только известные поля поверх текущего JSON.",
            inner,
        )
        hint.setWordWrap(True)
        layout.addWidget(hint)
        for title, specs in _SECTIONS:
            layout.addWidget(self._build_section(title, specs))
        layout.addStretch(1)
        scroll.setWidget(inner)

        btn_row = QHBoxLayout()
        btn_row.addStretch(1)
        btn_reload = QPushButton("Перезагрузить", self)
        btn_reload.clicked.connect(self.reload_from_disk)
        btn_save = QPushButton("Сохранить", self)
        btn_save.clicked.connect(self.save_to_disk)
        btn_row.addWidget(btn_reload)
        btn_row.addWidget(btn_save)

        root = QVBoxLayout(self)
        root.addWidget(scroll, stretch=1)
        root.addLayout(btn_row)

        self.reload_from_disk()

    def _build_section(self, title: str, specs: tuple[_FieldSpec, ...]) -> QGroupBox:
        box = QGroupBox(title, self)
        v = QVBoxLayout(box)
        if title == _BBB_VS_MTO_TITLE:
            btn = QPushButton("Открыть типы секций MTO", box)
            btn.clicked.connect(self._open_section_types_excel)
            v.addWidget(btn, stretch=0)
        for spec in specs:
            widget = QCheckBox(spec.label, box)
            widget.setToolTip(f"JSON: {spec.key}")
            v.addWidget(widget)
            self._widgets[spec.key] = widget
            self._specs[spec.key] = spec
        return box

    def _open_section_types_excel(self) -> None:
        excel_path = get_excel_path()
        try:
            os.startfile(excel_path)  # type: ignore[attr-defined]
        except FileNotFoundError:
            QMessageBox.warning(
                self,
                "Типы секций MTO",
                f"Файл не найден:\n{excel_path}",
            )
        except OSError as exc:
            QMessageBox.warning(
                self,
                "Типы секций MTO",
                f"Не удалось открыть файл:\n{exc}",
            )

    def collect_config(self) -> dict[str, Any]:
        """Patch checkbox values onto a freshly loaded config (unknown keys kept)."""
        config = load_config()
        for key, widget in self._widgets.items():
            _nested_set(config, key, bool(widget.isChecked()))
        return config

    def reload_from_disk(self) -> None:
        """Reload JSON and populate every checkbox."""
        config = load_config()
        for key, widget in self._widgets.items():
            spec = self._specs[key]
            raw = _nested_get(config, key, spec.default)
            widget.setChecked(bool(raw))

    def save_to_disk(self) -> bool:
        """Save patched config and notify the owner after success."""
        try:
            config = self.collect_config()
            saved = save_config(config)
        except Exception as exc:
            QMessageBox.critical(
                self,
                "Настройки MTO / BBB",
                f"Не удалось сохранить конфиг:\n{exc}",
            )
            return False
        if not saved:
            QMessageBox.critical(
                self,
                "Настройки MTO / BBB",
                "Функция сохранения сообщила об ошибке. Конфиг не сохранён.",
            )
            return False
        if self._on_saved is not None:
            self._on_saved()
        self.reload_from_disk()
        return True
