"""Native PySide6 editor for the RFP and as-build JSON configurations."""

from __future__ import annotations

from dataclasses import dataclass
from typing import Any, Callable, Literal

from PySide6.QtCore import QObject, Qt, QThread, Signal, Slot
from PySide6.QtWidgets import (
    QCheckBox,
    QComboBox,
    QFormLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QScrollArea,
    QSpinBox,
    QVBoxLayout,
    QWidget,
)

from RFQ.packing_list_provider import load_packing_dataset
from RFQ.rfp_parts.ds_registry import DEFAULT_REGISTRY_PATH
from RFQ.tags_rfp_compare.rfp_tags_utils import (
    DEFAULT_UNITS_SPLIT_BAN,
    RFP_PIPELINE_MODE_MTO_VO_ONLY,
    RFP_PIPELINE_MODE_STANDARD,
    RFP_PIPELINE_MODE_WITH_ORPHAN_MTO_VO,
    load_asbuild_config,
    load_config,
    resolve_rfp_pipeline_mode,
    save_asbuild_config,
    save_config,
)


FieldKind = Literal["bool", "text", "int", "list", "choice"]


class _PackingStatusLoader(QObject):
    """Load packing cache off the UI thread (UNC pickle can take tens of seconds)."""

    finished = Signal(object)

    @Slot()
    def run(self) -> None:
        try:
            # Resolve at call time so unit tests can mock ``load_packing_dataset``.
            self.finished.emit(load_packing_dataset())
        except BaseException as exc:
            self.finished.emit(exc)


@dataclass(frozen=True)
class _FieldSpec:
    key: str
    label: str
    kind: FieldKind = "bool"
    default: Any = False
    minimum: int = 0
    maximum: int = 100
    help_text: str = ""
    choices: tuple[tuple[str, str], ...] = ()


_PIPELINE_ITEMS = (
    (
        "Стандарт: RFP из файла",
        RFP_PIPELINE_MODE_STANDARD,
    ),
    (
        "RFP + MTO/VO без секции RFP",
        RFP_PIPELINE_MODE_WITH_ORPHAN_MTO_VO,
    ),
    (
        "Только MTO/VO (RFP не читается)",
        RFP_PIPELINE_MODE_MTO_VO_ONLY,
    ),
)

_ASBUILD_FIXED: dict[str, bool] = {
    "step2.flat_mto_structure": True,
    "step4.filter_to_mto_titles": True,
    "step4.include_packing_lists": False,
    "step4.assign_rfp_mto_code_compare_colors": False,
    "rfp_parts.auto_update_checklist": False,
    "rfp_parts.use_latest_net": False,
}

_ASBUILD_SOURCE_DISABLED: tuple[str, ...] = (
    "rfp_parts.ds_source_dir",
    "rfp_parts.ds_registry_path",
)

_RFP_SOURCE_SPECS: tuple[_FieldSpec, ...] = (
    _FieldSpec(
        "rfp_parts.use_latest_net",
        "Брать последний свод частей",
        default=True,
        help_text=(
            "Вкл.: Step1 читает свод по режиму с вкладки «RFP · Запуск» "
            "(части / ДС / hybrid). "
            "Для частей — самый новый rfp_parts_net.xlsx из "
            "«RFP сводный файл\\YYYY.MM.DD_HH.MM». "
            "При выключенных тегах — соседний rfp_parts_net_no_tags.xlsx. "
            "Если в RFP_Зиновьев появился новый файл или файла свода нет — "
            "свод пересобирается сам (как кнопка «Сбор частей»). "
            "Выкл.: используется поле «Файл RFP» ниже; режимы ДС не трогают."
        ),
    ),
    _FieldSpec(
        "rfp_parts.ds_source_dir",
        "Папка закупочных ДС",
        "text",
        "",
        help_text=(
            "Корень рекурсивного аудита ДС для режимов ds_only / hybrid. "
            "JSON: rfp_parts.ds_source_dir. Пусто — берётся "
            "last_ds_trusted_folder с вкладки RFP · Сбор частей."
        ),
    ),
    _FieldSpec(
        "rfp_parts.ds_registry_path",
        "Реестр ДС / УЛ",
        "text",
        str(DEFAULT_REGISTRY_PATH),
        help_text=(
            "Канон «Реестр_ДС_УЛ.xlsx». Робот обычным запуском файл не пишет. "
            "JSON: rfp_parts.ds_registry_path."
        ),
    ),
    _FieldSpec(
        "paths.rfp_path",
        "Файл RFP, если галка снята",
        "text",
        "",
        help_text=(
            "Запасной путь. Нужен, когда галка снята, для as-build, "
            "или если последнего свода частей ещё нет."
        ),
    ),
)

_SECTIONS: tuple[tuple[str, tuple[_FieldSpec, ...]], ...] = (
    (
        "Пути к файлам и папкам",
        (
            _FieldSpec(
                "paths.rfp_registr_lot_path", "Реестр лотов RFP", "text", ""
            ),
            _FieldSpec(
                "paths.ds_manager_matrix", "Список ДС — фамилии МП", "text", ""
            ),
            _FieldSpec(
                "paths.gem_supply_codes",
                "Коды поставки ГЭМ (8950)",
                "text",
                "",
            ),
            _FieldSpec(
                "paths.units_convert_matrix",
                "Матрица преобразования единиц",
                "text",
                "",
            ),
            _FieldSpec("paths.mto_path", "Папка MTO", "text", ""),
            _FieldSpec("paths.vo_path", "Папка VO", "text", ""),
            _FieldSpec("paths.code_ban_file", "Файл запрещённых кодов", "text", ""),
            _FieldSpec(
                "paths.replacement_table_file", "Таблица замен кодов", "text", ""
            ),
            _FieldSpec(
                "paths.result_dir_base", "Базовая папка результатов", "text", ""
            ),
        ),
    ),
    (
        "Общие настройки и память",
        (
            _FieldSpec(
                "memory_log",
                "Писать memory_log.txt на всех этапах",
                default=True,
            ),
            _FieldSpec(
                "step4.memory_top_stats",
                "Топ аллокаций tracemalloc (медленно)",
            ),
            _FieldSpec(
                "step4.memory_top_n",
                "Количество аллокаций в топе",
                "int",
                15,
                1,
                50,
            ),
        ),
    ),
    (
        "Step1 · Загрузка RFP",
        (
            _FieldSpec("step1.debug", "Отладка Step1"),
            _FieldSpec(
                "load_tags",
                "Читать теги из RFP / MTO / VO / УЛ",
                default=True,
                help_text=(
                    "Вкл.: теги в раскладке и сопоставлении; Step1 читает "
                    "rfp_parts_net.xlsx. Выкл.: Запуск берёт "
                    "rfp_parts_net_no_tags.xlsx (лот как в ДС, без раскрытия "
                    "на теги; если файла нет — пересборка свода в новый штамп, "
                    "боевой net не затирается). Теги MTO/VO/УЛ обнуляются в "
                    "памяти; отдельный кэш УЛ не пишется."
                ),
            ),
            _FieldSpec(
                "step1.skip_split",
                "Не раскладывать RFP в Step1 (split выполнит worker)",
            ),
            _FieldSpec(
                "step1.use_rfp_code_ban",
                "Читать и обновлять файл code_ban",
                default=True,
            ),
        ),
    ),
    (
        "Раскладка позиций по количеству",
        (
            _FieldSpec(
                "units_split_ban.enabled",
                "Не дробить позиции с запрещёнными единицами",
                default=True,
            ),
            _FieldSpec(
                "units_split_ban.units",
                "Запрещённые единицы (через запятую)",
                "list",
                DEFAULT_UNITS_SPLIT_BAN,
            ),
        ),
    ),
    (
        "Step2 · Загрузка MTO",
        (
            _FieldSpec("step2.debug", "Отладка Step2"),
            _FieldSpec(
                "step2.export_load_results_excel",
                "Excel с результатами загрузки MTO",
                default=True,
            ),
            _FieldSpec(
                "step2.export_positions_database_excel",
                "Excel с базой position_row MTO",
            ),
            _FieldSpec(
                "step2.flat_mto_structure",
                "Файлы MTO лежат прямо в корне папки",
            ),
        ),
    ),
    (
        "Step3 · Загрузка VO",
        (
            _FieldSpec("step3.debug", "Отладка Step3"),
            _FieldSpec("step3.export_to_excel", "Выгружать отчёты VO в Excel"),
        ),
    ),
    (
        "Step4 · Отладка и выполнение",
        (
            _FieldSpec("step4.debug", "Общая отладка Step4", default=True),
            _FieldSpec("step4.debug_step4_1", "Отладка Step4.1", default=True),
            _FieldSpec("step4.debug_step4_2", "Отладка Step4.2", default=True),
            _FieldSpec("step4.debug_step4_3", "Отладка Step4.3", default=True),
            _FieldSpec("step4.debug_step4_4", "Отладка Step4.4", default=True),
            _FieldSpec(
                "step4.verbose_progress_messages",
                "Подробные служебные сообщения 4.3/4.4",
                default=True,
            ),
            _FieldSpec("step4.collapse_debug", "Лог схлопывания строк"),
            _FieldSpec(
                "step4.unified_debug", "Единый debug-лог Step4", default=True
            ),
            _FieldSpec("step4.timing_log", "Тайминги в timing_log.xlsx", default=True),
            _FieldSpec(
                "step4.use_multiprocessing",
                "Дополнительный multiprocessing внутри match",
            ),
            _FieldSpec(
                "step4.parallel_by_title_mark",
                "Обрабатывать title/mark в нескольких процессах",
            ),
            _FieldSpec(
                "step4.max_workers",
                "Число workers (0 = автоматически)",
                "int",
                4,
                0,
                64,
            ),
        ),
    ),
    (
        "Step4 · Диагностические списки",
        (
            _FieldSpec(
                "step4.debug_tag", "Теги для отладки (через запятую)", "list", []
            ),
            _FieldSpec(
                "step4.debug_code", "Коды для отладки (через запятую)", "list", []
            ),
            _FieldSpec(
                "step4.debug_title_system",
                "Title/system для отладки (через запятую)",
                "list",
                [],
            ),
        ),
    ),
    (
        "Step4 · Отчёты, фильтры и упаковочные листы",
        (
            _FieldSpec(
                "step4.print_title_systems_table",
                "Печатать таблицу RFP/MTO/VO в консоль",
            ),
            _FieldSpec(
                "step4.export_title_systems_comparison_excel",
                "Excel сравнения title/system",
                default=True,
            ),
            _FieldSpec(
                "step4.assign_rfp_mto_code_compare_colors",
                "Раскрашивать сравнение кодов RFP и MTO",
                default=True,
            ),
            _FieldSpec(
                "step4.filter_to_mto_titles",
                "Оставлять только title/system из папки MTO",
            ),
            _FieldSpec(
                "step4.include_mto_vo_without_rfp_anchor",
                "Добавлять MTO/VO без секции RFP",
                help_text=(
                    "В не-стандартном режиме RFP поле принудительно включено."
                ),
            ),
            _FieldSpec(
                "step4.include_packing_lists",
                "Сравнивать результат RFP с упаковочными листами",
                default=True,
                help_text=(
                    "Если тегов в строке УЛ больше количества, экспорт RFP "
                    "останавливается без создания итогового Excel."
                ),
            ),
            _FieldSpec(
                "step4.export_bcc_accum_matrix",
                "Писать накопительную матрицу BCC",
                default=True,
                help_text=(
                    "Отдельный xlsx рядом с Шаг4: свод по закупочному коду "
                    "из МТО/ДС/УЛ без замен и тегов."
                ),
            ),
            _FieldSpec(
                "step4.ul_match_use_mto_tags",
                "Учитывать теги МТО при посадке УЛ",
                default=True,
                help_text=(
                    "После тегов RFP и безтеговых слотов УЛ сначала искать "
                    "совпадение с тегом МТО той же строки; только потом брать "
                    "УЛ с другим тегом. Выключено — сразу добор по коду без тегов."
                ),
            ),
        ),
    ),
    (
        "rfp_tags_utils",
        (
            _FieldSpec(
                "rfp_tags_utils.finalize_timing_top_n",
                "Количество финальных таймингов",
                "int",
                5,
                1,
                50,
            ),
            _FieldSpec(
                "rfp_tags_utils.save_input_fingerprints",
                "Сохранять input_fingerprints.json",
                default=True,
            ),
        ),
    ),
    (
        "Оптимизация колонок (текущий набор)",
        (
            _FieldSpec(
                "column_optimization.rfp.name.truncate",
                "RFP · обрезать NAME до символов",
                "int",
                200,
                0,
                1000,
            ),
            _FieldSpec(
                "column_optimization.rfp.type_mark.truncate",
                "RFP · обрезать TYPE_MARK до символов",
                "int",
                100,
                0,
                500,
            ),
            _FieldSpec(
                "column_optimization.rfp.DS_SPECIFICATION.load",
                "Загружать RFP DS_SPECIFICATION",
            ),
            _FieldSpec(
                "column_optimization.mto.name.truncate",
                "MTO · обрезать NAME до символов",
                "int",
                200,
                0,
                1000,
            ),
            _FieldSpec(
                "column_optimization.mto.type_mark.truncate",
                "MTO · обрезать TYPE_MARK до символов",
                "int",
                100,
                0,
                500,
            ),
        ),
    ),
)


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


def _parse_comma_list(text: str) -> list[str]:
    return [item.strip() for item in text.split(",") if item.strip()]


class RfpSettingsPanel(QWidget):
    """Edit the main RFP and as-build profiles without launching CTk."""

    def __init__(
        self,
        parent: QWidget | None = None,
        *,
        on_saved: Callable[[], None] | None = None,
        on_goto_packing: Callable[[], None] | None = None,
    ) -> None:
        super().__init__(parent)
        self._on_saved = on_saved
        self._on_goto_packing = on_goto_packing
        self._widgets: dict[str, QWidget] = {}
        self._specs = {spec.key: spec for spec in _RFP_SOURCE_SPECS}
        self._specs.update(
            {spec.key: spec for _, specs in _SECTIONS for spec in specs}
        )
        self._packing_thread: QThread | None = None

        root = QVBoxLayout(self)
        root.setContentsMargins(8, 8, 8, 8)
        root.setSpacing(8)
        root.addLayout(self._build_profile_row())

        scroll = QScrollArea(self)
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QScrollArea.Shape.NoFrame)
        inner = QWidget(scroll)
        form_root = QVBoxLayout(inner)
        form_root.setContentsMargins(2, 2, 8, 2)
        form_root.addWidget(self._build_packing_group())
        form_root.addWidget(self._build_pipeline_group())
        form_root.addWidget(self._build_rfp_source_group())
        for title, specs in _SECTIONS:
            form_root.addWidget(self._build_section(title, specs))
        form_root.addStretch(1)
        scroll.setWidget(inner)
        root.addWidget(scroll, stretch=1)
        root.addLayout(self._build_action_row())

        self.reload_from_disk()
        # Async: full UNC pickle load must not block CenterWindow.show().
        self._refresh_packing_status()

    def closeEvent(self, event) -> None:  # noqa: N802 — Qt API
        thread = self._packing_thread
        if thread is not None and thread.isRunning():
            thread.quit()
            thread.wait(3000)
        super().closeEvent(event)

    @property
    def current_profile(self) -> str:
        """Return ``main`` or ``asbuild`` for the selected profile."""
        return str(self._profile.currentData())

    def _build_profile_row(self) -> QHBoxLayout:
        row = QHBoxLayout()
        row.addWidget(QLabel("Профиль:", self))
        self._profile = QComboBox(self)
        self._profile.addItem("Основной RFP", "main")
        self._profile.addItem("As-build", "asbuild")
        self._profile.currentIndexChanged.connect(self.reload_from_disk)
        row.addWidget(self._profile, stretch=1)
        return row

    def _build_action_row(self) -> QHBoxLayout:
        row = QHBoxLayout()
        row.addStretch(1)
        reload_button = QPushButton("Перезагрузить", self)
        reload_button.clicked.connect(self.reload_from_disk)
        save_button = QPushButton("Сохранить", self)
        save_button.clicked.connect(self.save_to_disk)
        row.addWidget(reload_button)
        row.addWidget(save_button)
        return row

    def _build_packing_group(self) -> QGroupBox:
        box = QGroupBox("Общий кэш упаковочных листов", self)
        layout = QVBoxLayout(box)
        self._packing_status = QLabel("УЛ: проверка…", box)
        self._packing_status.setWordWrap(True)
        layout.addWidget(self._packing_status)
        warning = QLabel(
            "Критично для RFP: если тегов в строке УЛ больше количества, "
            "проверка останавливает экспорт и итоговый Excel не создаётся.",
            box,
        )
        warning.setWordWrap(True)
        warning.setStyleSheet("color: #a30; font-weight: bold;")
        warning.setToolTip("JSON: step4.include_packing_lists")
        layout.addWidget(warning)
        buttons = QHBoxLayout()
        refresh = QPushButton("Обновить статус УЛ", box)
        refresh.clicked.connect(self._refresh_packing_status)
        goto = QPushButton("Перейти к упаковочным листам", box)
        goto.setEnabled(self._on_goto_packing is not None)
        if self._on_goto_packing is not None:
            goto.clicked.connect(self._on_goto_packing)
        buttons.addWidget(refresh)
        buttons.addWidget(goto)
        buttons.addStretch(1)
        layout.addLayout(buttons)
        return box

    def _build_pipeline_group(self) -> QGroupBox:
        box = QGroupBox("Step1 · Режим участия RFP", self)
        form = QFormLayout(box)
        self._pipeline = QComboBox(box)
        for label, value in _PIPELINE_ITEMS:
            self._pipeline.addItem(label, value)
        self._pipeline.setToolTip("JSON: step1.rfp_pipeline_mode")
        self._pipeline.currentIndexChanged.connect(self._apply_dependency_rules)
        form.addRow("Режим пайплайна:", self._pipeline)
        return box

    def _build_rfp_source_group(self) -> QGroupBox:
        box = QGroupBox("Источник файла RFP для запуска", self)
        layout = QVBoxLayout(box)
        comment = QLabel(
            "Режим свода (части / ДС / ДС-RFP) выбирается на вкладке "
            "«RFP · Запуск», в свёрнутом блоке вверху.\n"
            "Галка включена — запуск читает свод выбранного режима. "
            "Галка снята — читается путь «Файл RFP» ниже; режимы ДС не "
            "используются. Если галка включена, а свода ещё нет, запуск "
            "остановится с ошибкой, а не подставит старый файл. "
            "В профиле As-build галка всегда выключена: там свой файл as-build.",
            box,
        )
        comment.setWordWrap(True)
        comment.setStyleSheet("color: #444;")
        layout.addWidget(comment)
        form = QFormLayout()
        form.setFieldGrowthPolicy(QFormLayout.FieldGrowthPolicy.AllNonFixedFieldsGrow)
        for spec in _RFP_SOURCE_SPECS:
            widget = self._make_widget(spec, box)
            tooltip = f"JSON: {spec.key}"
            if spec.help_text:
                tooltip += f"\n{spec.help_text}"
            widget.setToolTip(tooltip)
            label = QLabel(f"{spec.label}:", box)
            label.setToolTip(tooltip)
            label.setWordWrap(True)
            form.addRow(label, widget)
            self._widgets[spec.key] = widget
        layout.addLayout(form)
        return box

    def _build_section(
        self, title: str, specs: tuple[_FieldSpec, ...]
    ) -> QGroupBox:
        box = QGroupBox(title, self)
        form = QFormLayout(box)
        form.setFieldGrowthPolicy(QFormLayout.FieldGrowthPolicy.AllNonFixedFieldsGrow)
        for spec in specs:
            widget = self._make_widget(spec, box)
            tooltip = f"JSON: {spec.key}"
            if spec.help_text:
                tooltip += f"\n{spec.help_text}"
            widget.setToolTip(tooltip)
            label = QLabel(f"{spec.label}:", box)
            label.setToolTip(tooltip)
            label.setWordWrap(True)
            form.addRow(label, widget)
            self._widgets[spec.key] = widget
        return box

    @staticmethod
    def _make_widget(spec: _FieldSpec, parent: QWidget) -> QWidget:
        if spec.kind == "bool":
            return QCheckBox(parent)
        if spec.kind == "int":
            widget = QSpinBox(parent)
            widget.setRange(spec.minimum, spec.maximum)
            return widget
        if spec.kind == "choice":
            widget = QComboBox(parent)
            for label, value in spec.choices:
                widget.addItem(label, value)
            return widget
        return QLineEdit(parent)

    def _load_current_config(self) -> dict[str, Any]:
        return load_asbuild_config() if self.current_profile == "asbuild" else load_config()

    def reload_from_disk(self, *_args: object) -> None:
        """Reload the selected profile and repopulate every field."""
        try:
            config = self._load_current_config()
        except Exception as exc:
            QMessageBox.warning(self, "Настройки RFP", f"Ошибка чтения конфига:\n{exc}")
            return

        mode = resolve_rfp_pipeline_mode(config.get("step1"))
        index = self._pipeline.findData(mode)
        self._pipeline.setCurrentIndex(max(0, index))

        for key, widget in self._widgets.items():
            spec = self._specs[key]
            value = _nested_get(config, key, spec.default)
            if isinstance(widget, QCheckBox):
                widget.setChecked(bool(value))
            elif isinstance(widget, QSpinBox):
                try:
                    widget.setValue(int(value))
                except (TypeError, ValueError):
                    widget.setValue(int(spec.default))
            elif isinstance(widget, QComboBox):
                index = widget.findData(str(value or spec.default))
                widget.setCurrentIndex(max(0, index))
            elif isinstance(widget, QLineEdit):
                if spec.kind == "list":
                    if isinstance(value, (list, tuple)):
                        text = ", ".join(str(item) for item in value)
                    else:
                        text = str(value or "")
                else:
                    text = str(value or "")
                widget.setText(text)

        self._apply_profile_rules()

    def _widget_value(self, spec: _FieldSpec, widget: QWidget) -> Any:
        if isinstance(widget, QCheckBox):
            return widget.isChecked()
        if isinstance(widget, QSpinBox):
            return widget.value()
        if isinstance(widget, QComboBox):
            data = widget.currentData()
            return str(data if data is not None else spec.default)
        if isinstance(widget, QLineEdit):
            return (
                _parse_comma_list(widget.text())
                if spec.kind == "list"
                else widget.text().strip()
            )
        raise TypeError(f"Unsupported widget for {spec.key}")

    def collect_config(self) -> dict[str, Any]:
        """Patch current widget values onto a freshly loaded profile config."""
        config = self._load_current_config()
        for key, widget in self._widgets.items():
            _nested_set(config, key, self._widget_value(self._specs[key], widget))

        _nested_set(
            config,
            "step1.rfp_pipeline_mode",
            str(self._pipeline.currentData()),
        )
        step1 = config.get("step1")
        if isinstance(step1, dict):
            step1.pop("skip_rfp_load", None)

        if self._pipeline.currentData() != RFP_PIPELINE_MODE_STANDARD:
            _nested_set(config, "step4.include_mto_vo_without_rfp_anchor", True)

        # The truncation controls represent loaded columns, matching the legacy form.
        _nested_set(config, "column_optimization.rfp.name.load", True)
        _nested_set(config, "column_optimization.rfp.type_mark.load", True)
        _nested_set(config, "column_optimization.mto.name.load", True)
        _nested_set(config, "column_optimization.mto.type_mark.load", True)

        if self.current_profile == "asbuild":
            for key, value in _ASBUILD_FIXED.items():
                _nested_set(config, key, value)
        return config

    def save_to_disk(self) -> bool:
        """Save the selected profile and notify the owner after success."""
        try:
            config = self.collect_config()
            saved = (
                save_asbuild_config(config)
                if self.current_profile == "asbuild"
                else save_config(config)
            )
        except Exception as exc:
            QMessageBox.critical(
                self, "Настройки RFP", f"Не удалось сохранить конфиг:\n{exc}"
            )
            return False
        if not saved:
            QMessageBox.critical(
                self,
                "Настройки RFP",
                "Функция сохранения сообщила об ошибке. Конфиг не сохранён.",
            )
            return False
        if self._on_saved is not None:
            self._on_saved()
        self.reload_from_disk()
        return True

    def _apply_profile_rules(self) -> None:
        asbuild = self.current_profile == "asbuild"
        for key, forced_value in _ASBUILD_FIXED.items():
            widget = self._widgets.get(key)
            if widget is None:
                continue
            if asbuild and isinstance(widget, QCheckBox):
                widget.setChecked(forced_value)
            widget.setEnabled(not asbuild)
            if asbuild:
                widget.setToolTip(
                    f"JSON: {key}\nПрофиль As-build принудительно сохраняет: "
                    f"{str(forced_value).lower()}."
                )
            else:
                spec = self._specs[key]
                tooltip = f"JSON: {key}"
                if spec.help_text:
                    tooltip += f"\n{spec.help_text}"
                widget.setToolTip(tooltip)
        for key in _ASBUILD_SOURCE_DISABLED:
            widget = self._widgets.get(key)
            if widget is None:
                continue
            widget.setEnabled(not asbuild)
            spec = self._specs.get(key)
            if asbuild:
                widget.setToolTip(
                    f"JSON: {key}\nПрофиль As-build всегда читает «Файл RFP»; "
                    "режимы ДС не используются (use_latest_net=false)."
                )
            elif spec is not None:
                tooltip = f"JSON: {key}"
                if spec.help_text:
                    tooltip += f"\n{spec.help_text}"
                widget.setToolTip(tooltip)
        self._apply_dependency_rules()

    def _apply_dependency_rules(self, *_args: object) -> None:
        widget = self._widgets.get("step4.include_mto_vo_without_rfp_anchor")
        if not isinstance(widget, QCheckBox):
            return
        nonstandard = self._pipeline.currentData() != RFP_PIPELINE_MODE_STANDARD
        if nonstandard:
            widget.setChecked(True)
        widget.setEnabled(not nonstandard)
        widget.setToolTip(
            "JSON: step4.include_mto_vo_without_rfp_anchor\n"
            + (
                "Не-стандартный режим RFP принудительно включает это поле."
                if nonstandard
                else "В стандартном режиме значение задаётся пользователем."
            )
        )

    def _refresh_packing_status(self) -> None:
        """Start a non-blocking packing-cache status check (UNC I/O)."""
        if self._packing_thread is not None and self._packing_thread.isRunning():
            return
        self._packing_status.setText("УЛ: проверка кэша…")
        self._packing_status.setStyleSheet("color: #333;")
        self._packing_status.setToolTip(
            "Чтение кэша УЛ с сетевого диска; окно Центра не блокируется."
        )

        thread = QThread(self)
        loader = _PackingStatusLoader()
        loader.moveToThread(thread)
        thread.started.connect(loader.run)
        loader.finished.connect(self._apply_packing_status)
        loader.finished.connect(thread.quit)
        loader.finished.connect(loader.deleteLater)
        thread.finished.connect(self._on_packing_thread_finished)
        thread.finished.connect(thread.deleteLater)
        self._packing_thread = thread
        thread.start()

    @Slot(object)
    def _apply_packing_status(self, result: object) -> None:
        if isinstance(result, BaseException):
            self._packing_status.setText(f"УЛ: ошибка проверки кэша ({result})")
            self._packing_status.setStyleSheet("color: #a30; font-weight: bold;")
            self._packing_status.setToolTip(str(result))
            return
        # Duck-typed: PackingDataset in production; SimpleNamespace in smoke tests.
        dataset = result
        self._packing_status.setText(dataset.format_short())
        if getattr(dataset.quality, "value", None) == "ok":
            self._packing_status.setStyleSheet("color: #075; font-weight: bold;")
        else:
            self._packing_status.setStyleSheet("color: #a30; font-weight: bold;")
        issues = getattr(dataset, "issues", None) or []
        details = [
            issue.format_line()
            for issue in issues[:10]
            if hasattr(issue, "format_line")
        ]
        self._packing_status.setToolTip(
            "\n".join(details) if details else "Кэш УЛ проверен."
        )

    @Slot()
    def _on_packing_thread_finished(self) -> None:
        self._packing_thread = None
