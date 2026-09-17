"""Settings tab for ``ds_compare_center`` (``ds_compare_config.json``)."""

from __future__ import annotations

from typing import Any, Callable

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
    QVBoxLayout,
    QWidget,
)

from RFQ.ds_compare.ds_compare_config import (
    EXCEL_EXPORT_MODE_BOTH,
    EXCEL_EXPORT_MODE_WITHOUT_COMMENTS,
    EXCEL_EXPORT_MODE_WITH_COMMENTS,
    load_ds_compare_config,
    normalize_ds_vs_mto_output,
    normalize_grouped_compare,
    normalize_in_cabinet_debug,
    save_ds_compare_config,
)


class SettingsPanel(QWidget):
    """Editor for DS vs MTO JSON config (except MTO path presets)."""

    def __init__(
        self,
        parent: QWidget | None = None,
        *,
        on_saved: Callable[[], None] | None = None,
    ) -> None:
        super().__init__(parent)
        self._on_saved = on_saved

        scroll = QScrollArea(self)
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QScrollArea.Shape.NoFrame)

        inner = QWidget()
        layout = QVBoxLayout(inner)
        layout.addWidget(self._build_general_group())
        layout.addWidget(self._build_in_cabinet_group())
        layout.addWidget(self._build_grouped_group())
        layout.addWidget(self._build_output_group())
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

    def _build_general_group(self) -> QGroupBox:
        box = QGroupBox("Загрузка МТО", self)
        form = QFormLayout(box)
        self._chk_flat = QCheckBox("Плоская структура MTO (flat_mto_structure)", box)
        self._chk_cache = QCheckBox("Использовать кэш pickle", box)
        self._chk_force = QCheckBox("Принудительно обновить кэш MTO", box)
        self._chk_cache.setToolTip(
            "Кэш pickle при загрузке xlsx МТО из папки пресета (файлы в "
            "RFQ/ds_compare/cache/, отдельно от кэша результата сравнения ДС↔MTO)."
        )
        self._chk_force.setToolTip(
            "Перечитать xlsx МТО с диска (игнорировать pickle МТО) и на этот запуск "
            "не читать grouped_cmp_*.cache — полный DsMtoComparator."
        )
        form.addRow(self._chk_flat)
        form.addRow(self._chk_cache)
        form.addRow(self._chk_force)
        return box

    def _build_in_cabinet_group(self) -> QGroupBox:
        box = QGroupBox("Отладка IN_CABINET", self)
        v = QVBoxLayout(box)
        self._chk_ic_debug = QCheckBox(
            "Включить трассировку (ds_in_cabinet_trace.txt, ds_mto_compare_debug.log)",
            box,
        )
        v.addWidget(self._chk_ic_debug)
        hint = QLabel(
            "Фильтры (пустое поле = все). Несколько значений — через запятую.",
            box,
        )
        hint.setWordWrap(True)
        v.addWidget(hint)
        form = QFormLayout()
        self._ic_title = QLineEdit(box)
        self._ic_mark = QLineEdit(box)
        self._ic_code = QLineEdit(box)
        self._ic_cabinet = QLineEdit(box)
        form.addRow("Титул:", self._ic_title)
        form.addRow("Марка / раздел:", self._ic_mark)
        form.addRow("Код:", self._ic_code)
        form.addRow("Шкаф (IN_CABINET):", self._ic_cabinet)
        v.addLayout(form)
        return box

    def _build_grouped_group(self) -> QGroupBox:
        box = QGroupBox("Grouped DS vs MTO — ключи группировки", self)
        v = QVBoxLayout(box)
        hint = QLabel(
            "Если снять все галочки, будет дефолт: Титул + Марка/раздел + Код.",
            box,
        )
        hint.setWordWrap(True)
        v.addWidget(hint)
        self._grp_code = QCheckBox("Код (CODE)", box)
        self._grp_title = QCheckBox("Титул (DS_TITLE)", box)
        self._grp_system = QCheckBox("Марка / раздел (DS_SYSTEM)", box)
        self._grp_ds_name = QCheckBox("ДС (DS_NAME)", box)
        self._grp_mto_new = QCheckBox(
            "Группировать новые MTO-позиции теми же ключами (правая часть)",
            box,
        )
        self._grp_cache = QCheckBox(
            "Кэшировать результат сравнения ДС vs MTO (если ДС и MTO не менялись)",
            box,
        )
        self._grp_cache.setToolTip(
            "Сохраняет результат DsMtoComparator (до RFQ) в grouped_cmp_*.cache. "
            "RFQ и Excel пересчитываются каждый раз. Не путать с «кэш pickle» МТО. "
            "После смены логики сравнения в коде — удалите cache или включите "
            "принудительное обновление МТО на один прогон."
        )
        for w in (
            self._grp_code,
            self._grp_title,
            self._grp_system,
            self._grp_ds_name,
            self._grp_mto_new,
            self._grp_cache,
        ):
            v.addWidget(w)
        return box

    def _build_output_group(self) -> QGroupBox:
        box = QGroupBox("Excel-выгрузка DS vs MTO", self)
        v = QVBoxLayout(box)
        self._chk_internal_cols = QCheckBox(
            "Показывать технические имена колонок во второй строке заголовка",
            box,
        )
        self._chk_expand_agg_codes = QCheckBox(
            "Показывать коды агрегированных замен в колонке «Код продукции»",
            box,
        )
        self._chk_expand_agg_codes.setToolTip(
            "Только для Excel-выгрузки: CODE_2 остаётся машинным кодом в расчётах, "
            "а в файле выводится как «код / код_из_агрегации» для фильтра Excel."
        )
        v.addWidget(self._chk_internal_cols)
        v.addWidget(self._chk_expand_agg_codes)
        form = QFormLayout()
        self._cmb_excel_comments = QComboBox(box)
        self._cmb_excel_comments.addItem(
            "Оба варианта (с комментариями и без)", EXCEL_EXPORT_MODE_BOTH
        )
        self._cmb_excel_comments.addItem(
            "С комментариями", EXCEL_EXPORT_MODE_WITH_COMMENTS
        )
        self._cmb_excel_comments.addItem(
            "Без комментариев", EXCEL_EXPORT_MODE_WITHOUT_COMMENTS
        )
        form.addRow("Сохранение xlsx:", self._cmb_excel_comments)
        v.addLayout(form)
        hint = QLabel(
            "«Без комментариев» — второй файл с суффиксом _без_комм "
            "(при «Оба варианта» сохраняются два файла).",
            box,
        )
        hint.setWordWrap(True)
        v.addWidget(hint)
        return box

    def reload_from_disk(self) -> None:
        cfg = load_ds_compare_config()

        self._chk_flat.setChecked(bool(cfg.get("flat_mto_structure", False)))
        self._chk_cache.setChecked(bool(cfg.get("mto_use_cache", True)))
        self._chk_force.setChecked(bool(cfg.get("mto_force_update", False)))

        ic = normalize_in_cabinet_debug(cfg.get("in_cabinet_debug"))
        self._chk_ic_debug.setChecked(bool(ic.get("enabled", False)))
        self._ic_title.setText(ic.get("watch_title", ""))
        self._ic_mark.setText(ic.get("watch_mark", ""))
        self._ic_code.setText(ic.get("watch_code", ""))
        self._ic_cabinet.setText(ic.get("watch_cabinet", ""))

        grp = normalize_grouped_compare(cfg.get("grouped_compare"))
        self._grp_code.setChecked(bool(grp.get("group_by_code", True)))
        self._grp_title.setChecked(bool(grp.get("group_by_title", True)))
        self._grp_system.setChecked(bool(grp.get("group_by_system", True)))
        self._grp_ds_name.setChecked(bool(grp.get("group_by_ds_name", False)))
        self._grp_mto_new.setChecked(bool(grp.get("group_mto_new_positions", True)))
        self._grp_cache.setChecked(bool(grp.get("cache_ds_mto_result", False)))

        out = normalize_ds_vs_mto_output(cfg.get("ds_vs_mto_output"))
        self._chk_internal_cols.setChecked(
            bool(out.get("show_internal_column_names", False))
        )
        self._chk_expand_agg_codes.setChecked(
            bool(out.get("expand_aggregated_replacement_codes", True))
        )
        mode = str(out.get("excel_export_mode", EXCEL_EXPORT_MODE_BOTH))
        idx = self._cmb_excel_comments.findData(mode)
        self._cmb_excel_comments.setCurrentIndex(idx if idx >= 0 else 0)

    def _collect_config(self) -> dict[str, Any]:
        cfg = load_ds_compare_config()
        prev_out = normalize_ds_vs_mto_output(cfg.get("ds_vs_mto_output"))
        cfg["flat_mto_structure"] = self._chk_flat.isChecked()
        cfg["mto_use_cache"] = self._chk_cache.isChecked()
        cfg["mto_force_update"] = self._chk_force.isChecked()
        cfg["in_cabinet_debug"] = {
            "enabled": self._chk_ic_debug.isChecked(),
            "watch_title": self._ic_title.text().strip(),
            "watch_mark": self._ic_mark.text().strip(),
            "watch_code": self._ic_code.text().strip(),
            "watch_cabinet": self._ic_cabinet.text().strip(),
        }
        cfg["grouped_compare"] = {
            "group_by_code": self._grp_code.isChecked(),
            "group_by_title": self._grp_title.isChecked(),
            "group_by_system": self._grp_system.isChecked(),
            "group_by_ds_name": self._grp_ds_name.isChecked(),
            "group_mto_new_positions": self._grp_mto_new.isChecked(),
            "cache_ds_mto_result": self._grp_cache.isChecked(),
        }
        cfg["ds_vs_mto_output"] = {
            "show_internal_column_names": self._chk_internal_cols.isChecked(),
            "expand_aggregated_replacement_codes": self._chk_expand_agg_codes.isChecked(),
            "excel_export_mode": self._cmb_excel_comments.currentData()
            or EXCEL_EXPORT_MODE_BOTH,
            "columns": prev_out.get("columns", []),
        }
        return cfg

    def save_to_disk(self) -> bool:
        cfg = self._collect_config()
        if save_ds_compare_config(cfg):
            QMessageBox.information(
                self,
                "Сохранено",
                "Настройки записаны в ds_compare_config.json.",
            )
            if self._on_saved:
                self._on_saved()
            return True
        QMessageBox.critical(self, "Ошибка", "Не удалось сохранить файл конфигурации.")
        return False
