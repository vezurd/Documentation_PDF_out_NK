"""DS / RFP / leftover main.py control center main window."""

from __future__ import annotations

import json
import os
import re
import subprocess
import sys
from collections.abc import Callable
from pathlib import Path

from PySide6.QtCore import Qt, QProcess, QTimer, Slot
from PySide6.QtGui import QCloseEvent, QResizeEvent, QShowEvent
from PySide6.QtWidgets import (
    QComboBox,
    QFileDialog,
    QGroupBox,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QScrollArea,
    QSizePolicy,
    QSplitter,
    QTabWidget,
    QVBoxLayout,
    QWidget,
)

from RFQ.ds_compare.ds_compare_config import (
    index_from_menu_line,
    load_ds_compare_config,
    menu_lines_and_index,
    normalize_gui_paths,
    normalize_gui_window,
    save_ds_compare_config,
)
from RFQ.ds_compare.ds_grouped_compare import (
    get_last_packing_compare_audit,
    get_last_rfq_quantity_audit,
)
from RFQ.rfp_parts.ds_id_coverage import run_ds_id_coverage_job
from RFQ.rfp_parts.ds_jobs import (
    run_ds_baseline_job,
    run_ds_coverage_job,
    run_ds_hybrid_job,
    run_ds_registry_check_job,
)
from RFQ.tags_rfp_compare.ds_manager_roster import (
    get_last_ds_roster_compare,
    resolve_ds_manager_matrix_path,
    run_ds_roster_compare_job,
)
from RFQ.tags_rfp_compare.rfp_tags_utils import (
    find_latest_bcc_matrix_file,
    find_latest_step4_result_file,
    get_asbuild_config_path,
    load_asbuild_config,
    load_config,
    save_asbuild_config,
)
from ds_compare_center.bbb_settings_panel import BbbSettingsPanel
from ds_compare_center.columns_panel import ColumnsPanel
from ds_compare_center.layout_persistence import DebouncedLayoutSaver
from ds_compare_center.misc_run_panel import MiscRunPanel
from ds_compare_center.mto_paths_panel import MtoPathsPanel
from ds_compare_center.rfp_ds_id_panel import RfpDsIdPanel
from ds_compare_center.rfp_ds_mp_panel import RfpDsMpPanel
from ds_compare_center.rfp_parts_panel import RfpPartsPanel, rfp_parts_reports_dir
from ds_compare_center.rfp_progress_parser import RfpProgressParser
from ds_compare_center.rfp_run_panel import RfpRunPanel
from ds_compare_center.rfp_settings_panel import RfpSettingsPanel
from ds_compare_center.settings_panel import SettingsPanel
from ds_compare_center.tsd_packing_help_panel import TsdPackingHelpPanel
from ds_compare_center.tsd_packing_panel import TsdPackingPanel
from ds_compare_center.upd_panel import UpdPanel
from ds_compare_center.vpn_panel import VpnPanel
from main_v2.function_runner import FunctionJobRunner
from main_v2.help_dialog import show_help_dialog
from main_v2.job_monitor import JobMonitorPanel
from main_v2.job_runner import (
    ProcessJobRunner,
    aggregate_tags_env_overlay,
    build_aggregate_tags_argv,
)
from main_v2.legacy_actions import (
    run_ds_mto_file,
    run_grouped_ds_mto_file,
    run_grouped_ds_mto_rfq_only_file,
    run_merge_ds_dir,
    run_mto_ds_file,
    run_release_zip,
    run_tsd_packing_load,
    run_tsd_zinoviev_compare,
    run_upd_load,
)
from utils.path import open_dir

_TAB_INDEX_RUN = 0
_TAB_INDEX_MTO_PATHS = 1
_TAB_INDEX_SETTINGS = 2
_TAB_INDEX_COLUMNS = 3
_TAB_INDEX_TSD_PACKING = 4
_TAB_INDEX_TSD_HELP = 5
_TAB_INDEX_UPD = 6
_TAB_INDEX_SEP_RFP = 7
_TAB_INDEX_RFP_RUN = 8
_TAB_INDEX_RFP_SETTINGS = 9
_TAB_INDEX_RFP_PARTS = 10
_TAB_INDEX_RFP_DS_ID = 11
_TAB_INDEX_RFP_DS_MP = 12
_TAB_INDEX_SEP_MISC = 13
_TAB_INDEX_MISC_RUN = 14
_TAB_INDEX_BBB_SETTINGS = 15
_TAB_INDEX_VPN = 16
_WINDOW_TITLE = "Центр ДС, RFP и MTO"
_EXCEL_FILTER = "Excel files (*.xlsx *.xlsm *.xls);;All files (*.*)"
# Vertical gaps (~15–20% tighter than original) so block 3 actions fit without scroll on open.
_OPERATION_BLOCK_SPACING = 10
_OPERATION_FIELD_BLOCK_SPACING = 7
_OPERATION_FIELD_LABEL_GAP = 4
_OPERATION_FORM_TO_BUTTON_SPACING = 8
# Run tab: left = steps, right = full-height console (both stretch with window).
_RUN_TAB_LEFT_WIDTH = 520
_RUN_TAB_CONSOLE_WIDTH = 770
_RUN_TAB_CONSOLE_MIN_WIDTH = 420
_RUN_TAB_LEFT_MIN_WIDTH = 240
_JOB_TITLE_TSD_PACKING = "Упаковочные листы (ТСД)"
_JOB_TITLE_TSD_ZINOVIEV = "УЛ: робот vs Зиновьев"
_JOB_TITLE_UPD = "УПД (файлы закачки)"
_JOB_TITLE_RFP = "Сопоставление RFP / MTO / РКД / УЛ"
_JOB_TITLE_ASBUILD = "Сопоставление as-build RFP / MTO / РКД"
_JOB_TITLE_RFP_PARTS = "Сбор RFP из частей — только отчёты"
_JOB_TITLE_DS_BASELINE = "ДС: собрать вход Только ДС"
_JOB_TITLE_DS_HYBRID = "ДС: проверить RFP и наложить"
_JOB_TITLE_DS_COVERAGE = "ДС: только покрытие"
_JOB_TITLE_DS_REGISTRY = "ДС: проверить реестр"
_JOB_TITLES_DS_COCKPIT = frozenset(
    {
        _JOB_TITLE_DS_BASELINE,
        _JOB_TITLE_DS_HYBRID,
        _JOB_TITLE_DS_COVERAGE,
        _JOB_TITLE_DS_REGISTRY,
    }
)
_JOB_TITLE_RFP_DS_ID = "Соответствие ДС: RFP ↔ УЛ"
_JOB_TITLE_RFP_DS_MP = "Соответствие ДС: RFP ↔ фамилии МП"


def _shrink_h(widget: QWidget) -> QWidget:
    """Allow widget to follow a narrow splitter pane (ignore content-based width)."""
    widget.setMinimumWidth(0)
    policy = widget.sizePolicy()
    policy.setHorizontalPolicy(QSizePolicy.Policy.Expanding)
    widget.setSizePolicy(policy)
    return widget


def _project_root() -> Path:
    return Path(__file__).resolve().parents[1]


_TAB_BY_NAME = {
    "run": _TAB_INDEX_RUN,
    "mto_paths": _TAB_INDEX_MTO_PATHS,
    "settings": _TAB_INDEX_SETTINGS,
    "columns": _TAB_INDEX_COLUMNS,
    "packing": _TAB_INDEX_TSD_PACKING,
    "tsd": _TAB_INDEX_TSD_PACKING,
    "tsd_help": _TAB_INDEX_TSD_HELP,
    "packing_help": _TAB_INDEX_TSD_HELP,
    "upd": _TAB_INDEX_UPD,
    "upd_load": _TAB_INDEX_UPD,
    "rfp": _TAB_INDEX_RFP_RUN,
    "rfp_run": _TAB_INDEX_RFP_RUN,
    "rfp_settings": _TAB_INDEX_RFP_SETTINGS,
    "rfp_parts": _TAB_INDEX_RFP_PARTS,
    "parts": _TAB_INDEX_RFP_PARTS,
    "rfp_ds_id": _TAB_INDEX_RFP_DS_ID,
    "ds_id": _TAB_INDEX_RFP_DS_ID,
    "rfp_ul": _TAB_INDEX_RFP_DS_ID,
    "rfp_ds_mp": _TAB_INDEX_RFP_DS_MP,
    "ds_mp": _TAB_INDEX_RFP_DS_MP,
    "rfp_mp": _TAB_INDEX_RFP_DS_MP,
    "rfp_managers": _TAB_INDEX_RFP_DS_MP,
    "misc": _TAB_INDEX_MISC_RUN,
    "mto_run": _TAB_INDEX_MISC_RUN,
    "bbb": _TAB_INDEX_BBB_SETTINGS,
    "mto_settings": _TAB_INDEX_BBB_SETTINGS,
    "vpn": _TAB_INDEX_VPN,
    "cursor_vpn": _TAB_INDEX_VPN,
}


class CenterWindow(QWidget):
    """Standalone GUI for DS compare, RFP, and leftover main.py workflows."""

    def __init__(
        self,
        parent: QWidget | None = None,
        *,
        initial_tab: str = "misc",
    ) -> None:
        super().__init__(parent)
        self.setWindowFlags(Qt.WindowType.Window)
        self.setWindowTitle(_WINDOW_TITLE)
        # Side-by-side Run layout: steps left + full-height console right; fits Full HD.
        self.setMinimumSize(1200, 860)
        self.resize(1410, 980)

        self._last_job_title = ""
        self._function_job_tab = "run"
        self._rfp_progress_parser = RfpProgressParser()
        self._layout_saving_suspended = True
        self._layout_restore_done = False
        self._pending_maximized = False
        self._run_splitter: QSplitter | None = None

        self._runner = FunctionJobRunner(self)
        self._runner.log_received.connect(self._on_log)
        self._runner.finished.connect(self._on_finished)
        self._runner.error.connect(self._on_runner_error)

        self._proc_runner = ProcessJobRunner(self)
        self._proc_runner.log_received.connect(self._on_proc_log)
        self._proc_runner.finished.connect(self._on_proc_finished)
        self._proc_runner.error.connect(self._on_proc_error)

        self._tabs = QTabWidget(self)
        self._run_tab = self._build_run_tab()
        self._mto_paths_panel = MtoPathsPanel(on_saved=self._on_mto_paths_saved)
        self._settings_panel = SettingsPanel()
        self._columns_panel = ColumnsPanel()
        self._tsd_packing_panel = TsdPackingPanel(
            on_run=self._run_tsd_packing,
            on_compare_zinoviev=self._run_tsd_zinoviev_compare,
        )
        self._tsd_packing_help_panel = TsdPackingHelpPanel()
        self._upd_panel = UpdPanel(on_run=self._run_upd_load)
        self._rfp_run_panel = RfpRunPanel(
            on_run_rfp=self._run_rfp_check,
            on_run_asbuild=self._run_asbuild_check,
            on_open_last=self._open_latest_step4,
            on_open_matrix=self._open_latest_bcc_matrix,
        )
        self._rfp_settings_panel = RfpSettingsPanel(
            on_saved=self._on_rfp_settings_saved,
            on_goto_packing=self._goto_tsd_packing,
        )
        self._rfp_parts_panel = RfpPartsPanel(
            on_run=self._run_rfp_parts,
            on_ds_baseline=self._run_ds_baseline,
            on_ds_hybrid=self._run_ds_hybrid,
            on_ds_coverage=self._run_ds_coverage,
            on_ds_registry=self._run_ds_registry,
        )
        self._rfp_ds_id_panel = RfpDsIdPanel(
            on_run=self._run_rfp_ds_id,
        )
        self._rfp_ds_mp_panel = RfpDsMpPanel(
            on_run=self._run_rfp_ds_mp,
            on_open=self._open_ds_manager_book,
        )
        self._misc_run_panel = MiscRunPanel(
            on_launch_pdf=self._launch_pdf_v2_monitor,
            on_launch_rd_catalog=self._launch_rd_catalog,
            on_folder_job=self._run_misc_folder_job,
            on_file_job=self._run_misc_file_job,
            on_zip=self._run_misc_zip,
            on_help=self._show_misc_help,
        )
        self._bbb_settings_panel = BbbSettingsPanel()
        self._vpn_panel = VpnPanel()

        self._rfp_run_panel.monitor.stop_requested.connect(self._proc_runner.request_stop)
        self._rfp_parts_panel.monitor.stop_requested.connect(self._proc_runner.request_stop)

        self._tabs.addTab(self._run_tab, "ДС · Запуск")
        self._tabs.addTab(self._mto_paths_panel, "ДС · Папки МТО")
        self._tabs.addTab(self._settings_panel, "ДС · Настройки")
        self._tabs.addTab(self._columns_panel, "ДС · Столбцы Excel")
        self._tabs.addTab(self._tsd_packing_panel, "ДС · Упаковочные листы")
        self._tabs.addTab(self._tsd_packing_help_panel, "ДС · Формат ТСД")
        self._tabs.addTab(self._upd_panel, "ДС · УПД")
        sep_idx = self._tabs.addTab(QWidget(self), "│ RFP")
        self._tabs.setTabEnabled(sep_idx, False)
        self._tabs.addTab(self._rfp_run_panel, "RFP · Запуск")
        self._tabs.addTab(self._rfp_settings_panel, "RFP · Настройки")
        self._tabs.addTab(self._rfp_parts_panel, "RFP · Сбор частей")
        self._tabs.addTab(self._rfp_ds_id_panel, "RFP · ДС ↔ УЛ")
        self._tabs.addTab(self._rfp_ds_mp_panel, "RFP · ДС ↔ МП")
        sep_misc = self._tabs.addTab(QWidget(self), "│ Прочее")
        self._tabs.setTabEnabled(sep_misc, False)
        self._tabs.addTab(self._misc_run_panel, "Прочее · Запуск")
        self._tabs.addTab(self._bbb_settings_panel, "MTO · Настройки")
        self._tabs.addTab(self._vpn_panel, "Прочее · VPN")

        layout = QVBoxLayout(self)
        layout.setContentsMargins(8, 8, 8, 8)
        layout.addWidget(self._tabs)

        tab_idx = _TAB_BY_NAME.get(initial_tab, _TAB_INDEX_MISC_RUN)
        self._tabs.setCurrentIndex(tab_idx)

        self._reload_paths_from_config()
        self._refresh_mto_combo()
        self._init_layout_persistence()

    def _splitters_by_key(self) -> dict[str, QSplitter]:
        """Named horizontal splitters that share ``gui_window.splitters``."""
        out: dict[str, QSplitter] = {}
        if self._run_splitter is not None:
            out["run"] = self._run_splitter
        packing = getattr(self._tsd_packing_panel, "splitter", None)
        if isinstance(packing, QSplitter):
            out["packing"] = packing
        upd = getattr(self._upd_panel, "splitter", None)
        if isinstance(upd, QSplitter):
            out["upd"] = upd
        rfp_run = getattr(self._rfp_run_panel, "splitter", None)
        if isinstance(rfp_run, QSplitter):
            out["rfp_run"] = rfp_run
        rfp_parts = getattr(self._rfp_parts_panel, "splitter", None)
        if isinstance(rfp_parts, QSplitter):
            out["rfp_parts"] = rfp_parts
        rfp_ds_id = getattr(self._rfp_ds_id_panel, "splitter", None)
        if isinstance(rfp_ds_id, QSplitter):
            out["rfp_ds_id"] = rfp_ds_id
        rfp_ds_mp = getattr(self._rfp_ds_mp_panel, "splitter", None)
        if isinstance(rfp_ds_mp, QSplitter):
            out["rfp_ds_mp"] = rfp_ds_mp
        misc_run = getattr(self._misc_run_panel, "splitter", None)
        if isinstance(misc_run, QSplitter):
            out["misc_run"] = misc_run
        return out

    def _init_layout_persistence(self) -> None:
        """Restore size from config; apply splitters after first show."""
        cfg = load_ds_compare_config()
        win = normalize_gui_window(cfg.get("gui_window"))
        self.resize(win["width"], win["height"])
        self._pending_maximized = bool(win["maximized"])
        self._layout_saver = DebouncedLayoutSaver(
            self,
            snapshot=self._snapshot_gui_window,
            on_flush=self._flush_gui_window_to_disk,
        )
        for splitter in self._splitters_by_key().values():
            splitter.splitterMoved.connect(self._schedule_layout_save)

    def _snapshot_gui_window(self) -> dict:
        geo = self.normalGeometry() if self.isMaximized() else self.geometry()
        splitters: dict[str, list[int]] = {}
        for key, splitter in self._splitters_by_key().items():
            sizes = splitter.sizes()
            if len(sizes) >= 2:
                splitters[key] = [max(1, int(sizes[0])), max(1, int(sizes[1]))]
        return normalize_gui_window(
            {
                "width": int(geo.width()),
                "height": int(geo.height()),
                "maximized": bool(self.isMaximized()),
                "splitters": splitters,
            }
        )

    def _flush_gui_window_to_disk(self, snap: dict) -> None:
        cfg = load_ds_compare_config()
        cfg["gui_window"] = normalize_gui_window(snap)
        save_ds_compare_config(cfg)

    def _schedule_layout_save(self, *_args: object) -> None:
        if self._layout_saving_suspended:
            return
        if hasattr(self, "_layout_saver"):
            self._layout_saver.schedule()

    def _apply_splitters_from_config(self) -> None:
        cfg = load_ds_compare_config()
        win = normalize_gui_window(cfg.get("gui_window"))
        raw_splitters = win.get("splitters") or {}
        self._layout_saving_suspended = True
        try:
            for key, splitter in self._splitters_by_key().items():
                pair = raw_splitters.get(key)
                if isinstance(pair, list) and len(pair) >= 2:
                    splitter.setSizes([int(pair[0]), int(pair[1])])
        finally:
            self._layout_saving_suspended = False

    def _finish_layout_restore(self) -> None:
        """Apply splitter sizes (and maximize) after the first layout pass."""
        if self._layout_restore_done:
            return
        self._layout_restore_done = True
        self._apply_splitters_from_config()
        if self._pending_maximized:
            self.showMaximized()
        if hasattr(self, "_layout_saver"):
            self._layout_saver.sync_last_written_from_snapshot(
                self._snapshot_gui_window()
            )
        self._layout_saving_suspended = False

    def showEvent(self, event: QShowEvent) -> None:
        super().showEvent(event)
        if not self._layout_restore_done:
            QTimer.singleShot(0, self._finish_layout_restore)

    def resizeEvent(self, event: QResizeEvent) -> None:
        super().resizeEvent(event)
        self._schedule_layout_save()

    def closeEvent(self, event: QCloseEvent) -> None:
        if hasattr(self, "_layout_saver"):
            self._layout_saver.flush_now()
        super().closeEvent(event)

    def _build_run_tab(self) -> QWidget:
        tab = QWidget(self)
        splitter = QSplitter(Qt.Orientation.Horizontal, tab)

        scroll = QScrollArea(splitter)
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QScrollArea.Shape.NoFrame)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        scroll.setMinimumWidth(_RUN_TAB_LEFT_MIN_WIDTH)
        scroll.setSizePolicy(
            QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding
        )

        inner = QWidget(scroll)
        # Ignored H: always match scroll viewport (rubber left column).
        inner.setSizePolicy(
            QSizePolicy.Policy.Ignored, QSizePolicy.Policy.Preferred
        )
        left_layout = QVBoxLayout(inner)
        left_layout.setContentsMargins(0, 0, 4, 0)
        left_layout.setSpacing(_OPERATION_BLOCK_SPACING)
        left_layout.setAlignment(Qt.AlignmentFlag.AlignTop)

        self._edit_merge_dir = self._compact_line_edit(QLineEdit(inner))
        left_layout.addWidget(
            self._build_operation_block(
                inner,
                title="1. Объединить ДС из папки",
                description=(
                    "Сводный xlsx из всех файлов ДС в указанной папке (DS_summary). "
                    "Результат подставится в поле «Файл ДС» на шаге 2."
                ),
                form_rows=[
                    (
                        "Папка для merge ДС:",
                        self._path_row(
                            inner,
                            self._edit_merge_dir,
                            browse=lambda: self._browse_directory(
                                self._edit_merge_dir, "last_merge_ds_folder"
                            ),
                        ),
                    ),
                ],
                run_label="Объединить ДС из папки",
                run_handler=self._run_merge_ds,
            )
        )

        self._edit_ds = self._compact_line_edit(QLineEdit(inner))
        self._mto_combo = QComboBox(inner)
        self._compact_form_control(self._mto_combo)
        self._mto_combo.currentTextChanged.connect(self._on_mto_combo_changed)
        left_layout.addWidget(
            self._build_operation_block(
                inner,
                title="2. Список ДС vs MTO",
                description=(
                    "Построчное сравнение сводного или отдельного файла ДС с MTO; "
                    "результат — Excel РОБОТ_СРАВНЕНИЕ_*."
                ),
                form_rows=[
                    (
                        "Файл ДС (xlsx):",
                        self._path_row(
                            inner,
                            self._edit_ds,
                            browse=lambda: self._browse_file(self._edit_ds, "last_ds_file"),
                        ),
                    ),
                    ("Папка МТО (пресет):", self._mto_combo),
                ],
                run_label="Список ДС vs MTO",
                run_handler=self._run_ds_mto,
            )
        )

        self._edit_rfq = self._compact_line_edit(QLineEdit(inner))
        self._rfq_audit_label = QLabel("", inner)
        self._rfq_audit_label.setWordWrap(True)
        self._rfq_audit_label.setStyleSheet("color: #333; font-size: 11px;")
        _shrink_h(self._rfq_audit_label)
        self._packing_audit_label = QLabel("", inner)
        self._packing_audit_label.setWordWrap(True)
        self._packing_audit_label.setStyleSheet("color: #333; font-size: 11px;")
        _shrink_h(self._packing_audit_label)
        rfq_box = self._build_operation_block(
            inner,
            title="3. ДС vs MTO vs RFQ",
            description=(
                "ДС агрегируется по ключам из настроек, затем сравнение с MTO и RFQ. "
                "После этого добавляется накопительная поставка из последнего кэша УЛ. "
                "Используются файл ДС и пресет МТО из шага 2; RFQ необязателен. "
                "При включённом кэше ДС↔MTO повторный прогон с новым RFQ можно ускорить "
                "кнопкой «Только RFQ → Excel»."
            ),
            form_rows=[
                (
                    "Файл RFQ (TPK xlsx):",
                    self._path_row(
                        inner,
                        self._edit_rfq,
                        browse=lambda: self._browse_file(self._edit_rfq, "last_rfq_file"),
                    ),
                ),
            ],
            run_label="ДС vs MTO vs RFQ",
            run_handler=self._run_grouped_ds_mto,
        )
        rfq_only_btn = QPushButton("Только RFQ → Excel (кэш ДС↔MTO)", rfq_box)
        rfq_only_btn.setToolTip(
            "Не запускает DsMtoComparator: нужен сохранённый grouped_cmp_*.cache "
            "и тот же файл ДС/пресет МТО. Обязателен файл RFQ."
        )
        rfq_only_btn.clicked.connect(self._run_grouped_rfq_only)
        rfq_layout = rfq_box.layout()
        if isinstance(rfq_layout, QVBoxLayout):
            rfq_layout.addSpacing(5)
            rfq_layout.addWidget(rfq_only_btn)
            rfq_layout.addWidget(self._rfq_audit_label)
            rfq_layout.addWidget(self._packing_audit_label)
        left_layout.addWidget(rfq_box)
        left_layout.addStretch(1)
        scroll.setWidget(inner)

        self._monitor = JobMonitorPanel(splitter, show_stop=False)
        self._monitor.setMinimumWidth(_RUN_TAB_CONSOLE_MIN_WIDTH)
        self._monitor.setSizePolicy(
            QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding
        )
        splitter.addWidget(scroll)
        splitter.addWidget(self._monitor)
        splitter.setStretchFactor(0, 1)
        splitter.setStretchFactor(1, 1)
        splitter.setSizes([_RUN_TAB_LEFT_WIDTH, _RUN_TAB_CONSOLE_WIDTH])
        splitter.setChildrenCollapsible(False)
        self._run_splitter = splitter

        tab_layout = QVBoxLayout(tab)
        tab_layout.setContentsMargins(0, 0, 0, 0)
        tab_layout.addWidget(splitter)
        return tab

    def _path_row(
        self,
        parent: QWidget,
        edit: QLineEdit,
        *,
        browse: Callable[[], None],
    ) -> QWidget:
        """Path field on top, «Обзор…» below; edit width follows left column."""
        # QLineEdit sizeHint grows with text (UNC paths) — force shrink to parent width.
        edit.setMinimumWidth(0)
        self._compact_form_control(edit)
        row_widget = QWidget(parent)
        row_widget.setSizePolicy(
            QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Maximum
        )
        col = QVBoxLayout(row_widget)
        col.setContentsMargins(0, 0, 0, 0)
        col.setSpacing(_OPERATION_FIELD_LABEL_GAP)
        btn = QPushButton("Обзор…", row_widget)
        btn.setSizePolicy(QSizePolicy.Policy.Fixed, QSizePolicy.Policy.Fixed)
        btn.clicked.connect(browse)
        col.addWidget(edit)
        col.addWidget(btn, alignment=Qt.AlignmentFlag.AlignLeft)
        return row_widget

    def _build_operation_block(
        self,
        parent: QWidget,
        *,
        title: str,
        description: str,
        form_rows: list[tuple[str, QWidget]],
        run_label: str,
        run_handler: Callable[[], None],
    ) -> QGroupBox:
        """One workflow step: description, path fields, run button."""
        box = _shrink_h(QGroupBox(title, parent))
        box.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Maximum)
        layout = QVBoxLayout(box)
        layout.setContentsMargins(10, 13, 10, 10)
        layout.setSpacing(_OPERATION_BLOCK_SPACING)

        desc = QLabel(description, box)
        desc.setWordWrap(True)
        desc.setAlignment(Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignLeft)
        _shrink_h(desc)
        layout.addWidget(desc)

        fields = QWidget(box)
        fields.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Maximum)
        fields.setMinimumWidth(0)
        fields_layout = QVBoxLayout(fields)
        fields_layout.setContentsMargins(0, 0, 0, 0)
        fields_layout.setSpacing(_OPERATION_FIELD_BLOCK_SPACING)
        for label_text, widget in form_rows:
            fields_layout.addWidget(self._labeled_field(fields, label_text, widget))
        layout.addWidget(fields)

        layout.addSpacing(_OPERATION_FORM_TO_BUTTON_SPACING)
        btn = QPushButton(run_label, box)
        btn.setMinimumWidth(0)
        btn.clicked.connect(run_handler)
        layout.addWidget(btn)
        return box

    def _labeled_field(
        self,
        parent: QWidget,
        label_text: str,
        widget: QWidget,
    ) -> QWidget:
        """Field block: caption on top, control below (avoids QFormLayout clipping)."""
        block = QWidget(parent)
        block.setMinimumWidth(0)
        block.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Maximum)
        block_layout = QVBoxLayout(block)
        block_layout.setContentsMargins(0, 0, 0, 0)
        block_layout.setSpacing(_OPERATION_FIELD_LABEL_GAP)
        label = QLabel(label_text, block)
        self._compact_form_control(widget)
        if isinstance(widget, QLineEdit):
            widget.setMinimumWidth(0)
        block_layout.addWidget(label)
        block_layout.addWidget(widget, alignment=Qt.AlignmentFlag.AlignTop)
        return block

    @staticmethod
    def _compact_form_control(widget: QWidget) -> QWidget:
        """Keep single-line inputs at natural height; allow horizontal shrink."""
        widget.setMinimumWidth(0)
        widget.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        return widget

    def _compact_line_edit(self, edit: QLineEdit) -> QLineEdit:
        """Single-line path field — do not stretch vertically."""
        self._compact_form_control(edit)
        return edit

    def _reload_paths_from_config(self) -> None:
        cfg = load_ds_compare_config()
        paths = normalize_gui_paths(cfg.get("gui_paths"))
        self._edit_ds.setText(paths.get("last_ds_file", ""))
        self._edit_rfq.setText(paths.get("last_rfq_file", ""))
        self._edit_merge_dir.setText(paths.get("last_merge_ds_folder", ""))
        if hasattr(self, "_tsd_packing_panel"):
            self._tsd_packing_panel.reload_paths_from_config()
        if hasattr(self, "_upd_panel"):
            self._upd_panel.reload_paths_from_config()
        if hasattr(self, "_rfp_parts_panel"):
            self._rfp_parts_panel.reload_paths_from_config()

    def _refresh_mto_combo(self) -> None:
        cfg = load_ds_compare_config()
        lines, idx = menu_lines_and_index(cfg)
        self._mto_combo.blockSignals(True)
        self._mto_combo.clear()
        self._mto_combo.addItems(lines)
        if lines:
            self._mto_combo.setCurrentIndex(idx)
        self._mto_combo.blockSignals(False)

    def _on_mto_paths_saved(self) -> None:
        self._refresh_mto_combo()

    def _on_rfp_settings_saved(self) -> None:
        self._rfp_run_panel.refresh_from_config()

    def _goto_tsd_packing(self) -> None:
        self._tabs.setCurrentIndex(_TAB_INDEX_TSD_PACKING)

    def _persist_gui_path(self, key: str, value: str) -> None:
        cfg = load_ds_compare_config()
        paths = normalize_gui_paths(cfg.get("gui_paths"))
        paths[key] = value.strip()  # type: ignore[literal-required]
        cfg["gui_paths"] = paths
        save_ds_compare_config(cfg)

    def _browse_file(self, edit: QLineEdit, config_key: str) -> None:
        start = edit.text().strip()
        start_dir = str(Path(start).parent) if start else ""
        path, _ = QFileDialog.getOpenFileName(
            self,
            "Выберите файл",
            start_dir,
            "Excel (*.xlsx);;All files (*.*)",
        )
        if path:
            edit.setText(path)
            self._persist_gui_path(config_key, path)

    def _browse_directory(self, edit: QLineEdit, config_key: str) -> None:
        start = edit.text().strip()
        path = QFileDialog.getExistingDirectory(self, "Выберите папку", start or "")
        if path:
            edit.setText(path)
            self._persist_gui_path(config_key, path)

    def _on_mto_combo_changed(self, choice: str) -> None:
        if not choice or choice == "…":
            return
        cfg = load_ds_compare_config()
        entries = cfg.get("mto_paths") or []
        n = len(entries)
        if n == 0:
            return
        parsed = index_from_menu_line(choice, n)
        if parsed is None:
            lines, _ = menu_lines_and_index(cfg)
            try:
                parsed = lines.index(choice)
            except ValueError:
                return
        cfg["mto_path_selected_index"] = parsed
        save_ds_compare_config(cfg)

    def _save_rfq_path_from_field(self) -> str | None:
        path = self._edit_rfq.text().strip()
        if not path:
            return None
        if not os.path.isfile(path):
            QMessageBox.warning(self, "Файл RFQ", f"Файл не найден:\n{path}")
            return None
        self._persist_gui_path("last_rfq_file", path)
        return path

    def _save_ds_path_from_field(self) -> str | None:
        path = self._edit_ds.text().strip()
        if not path:
            QMessageBox.warning(self, "Файл ДС", "Укажите файл ДС или выберите через «Обзор…».")
            return None
        if not os.path.isfile(path):
            QMessageBox.warning(self, "Файл ДС", f"Файл не найден:\n{path}")
            return None
        self._persist_gui_path("last_ds_file", path)
        return path

    def _save_merge_dir_from_field(self) -> str | None:
        path = self._edit_merge_dir.text().strip()
        if not path:
            QMessageBox.warning(
                self,
                "Папка merge",
                "Укажите папку с файлами ДС или выберите через «Обзор…».",
            )
            return None
        if not os.path.isdir(path):
            QMessageBox.warning(self, "Папка merge", f"Папка не найдена:\n{path}")
            return None
        self._persist_gui_path("last_merge_ds_folder", path)
        return path

    def _start_job(self, title: str, fn, *args, job_tab: str | None = None) -> None:
        if self._runner.is_running() or self._proc_busy():
            QMessageBox.information(self, "Занято", "Дождитесь завершения текущей операции.")
            return
        self._last_job_title = title
        if job_tab == "misc":
            self._function_job_tab = "misc"
            self._misc_run_panel.monitor.start_job(title)
            self._tabs.setCurrentIndex(_TAB_INDEX_MISC_RUN)
        elif title in _JOB_TITLES_DS_COCKPIT or job_tab == "rfp_parts":
            self._function_job_tab = "rfp_parts"
            self._rfp_parts_panel.monitor.start_job(title)
            self._tabs.setCurrentIndex(_TAB_INDEX_RFP_PARTS)
        elif title == _JOB_TITLE_RFP_DS_ID:
            self._function_job_tab = "ds_id"
            self._rfp_ds_id_panel.monitor.start_job(title)
            self._tabs.setCurrentIndex(_TAB_INDEX_RFP_DS_ID)
        elif title == _JOB_TITLE_RFP_DS_MP:
            self._function_job_tab = "ds_mp"
            self._rfp_ds_mp_panel.monitor.start_job(title)
            self._tabs.setCurrentIndex(_TAB_INDEX_RFP_DS_MP)
        elif title in (_JOB_TITLE_TSD_PACKING, _JOB_TITLE_TSD_ZINOVIEV):
            self._function_job_tab = "packing"
            self._tsd_packing_panel.monitor.start_job(title)
            self._tabs.setCurrentIndex(_TAB_INDEX_TSD_PACKING)
        elif title == _JOB_TITLE_UPD:
            self._function_job_tab = "upd"
            self._upd_panel.monitor.start_job(title)
            self._tabs.setCurrentIndex(_TAB_INDEX_UPD)
        else:
            self._function_job_tab = "run"
            self._monitor.start_job(title)
            self._tabs.setCurrentIndex(_TAB_INDEX_RUN)
        self._runner.start(title, fn, *args)

    def _proc_busy(self) -> bool:
        return self._proc_runner.process().state() != QProcess.ProcessState.NotRunning

    def _monitor_for_proc_job(self) -> JobMonitorPanel:
        if self._last_job_title == _JOB_TITLE_RFP_PARTS:
            return self._rfp_parts_panel.monitor
        return self._rfp_run_panel.monitor

    def _is_rfp_pipeline_process_job(self) -> bool:
        return self._last_job_title in (_JOB_TITLE_RFP, _JOB_TITLE_ASBUILD)

    def _start_process_job(self, title: str, argv: list[str], *, tab_index: int) -> None:
        if self._runner.is_running() or self._proc_busy():
            QMessageBox.information(self, "Занято", "Дождитесь завершения текущей операции.")
            return
        self._last_job_title = title
        if self._is_rfp_pipeline_process_job():
            self._rfp_progress_parser = RfpProgressParser()
            self._rfp_run_panel.reset_milestones()
        self._tabs.setCurrentIndex(tab_index)
        monitor = self._monitor_for_proc_job()
        monitor.start_job(title)
        root = _project_root()
        started = self._proc_runner.start(
            argv,
            cwd=root,
            env=aggregate_tags_env_overlay(root),
        )
        if not started:
            if self._is_rfp_pipeline_process_job():
                self._rfp_run_panel.apply_milestone(
                    "prepare", "Error", "Процесс не запущен"
                )
            monitor.finish_job(False, f"{title}: процесс не запущен", None)

    @Slot(str)
    def _on_proc_log(self, chunk: str) -> None:
        self._monitor_for_proc_job().append_log(chunk)
        if self._is_rfp_pipeline_process_job() and not chunk.startswith("[stderr]"):
            for event in self._rfp_progress_parser.feed(chunk):
                self._rfp_run_panel.apply_milestone(event)

    @Slot(str)
    def _on_proc_error(self, message: str) -> None:
        self._monitor_for_proc_job().append_log(f"[error] {message}\n")
        if self._is_rfp_pipeline_process_job():
            self._rfp_run_panel.mark_active_error(f"Ошибка процесса: {message}")

    @Slot(int, int)
    def _on_proc_finished(self, exit_code: int, exit_status: int) -> None:
        success = exit_status == 0 and exit_code == 0
        title = self._last_job_title or "Задача"
        message = f"{title}: готово" if success else f"{title}: код выхода {exit_code}"
        result_path: str | None = None
        if self._is_rfp_pipeline_process_job():
            for event in self._rfp_progress_parser.flush():
                self._rfp_run_panel.apply_milestone(event)
            if not success:
                detail = (
                    f"Процесс аварийно завершён, код {exit_code}"
                    if exit_status != 0
                    else f"Процесс завершён с кодом {exit_code}"
                )
                self._rfp_run_panel.mark_active_error(detail)
            if success:
                result_path = find_latest_step4_result_file()
            else:
                self._maybe_show_quantity_balance_fatal()
        elif self._last_job_title == _JOB_TITLE_RFP_PARTS:
            result_path = str(
                self._rfp_parts_panel.last_reports_dir or rfp_parts_reports_dir()
            )
            self._rfp_parts_panel.on_job_finished(success, message)
            if success and result_path:
                try:
                    open_dir(result_path)
                except Exception:
                    pass
        self._monitor_for_proc_job().finish_job(success, message, result_path)

    def _maybe_show_quantity_balance_fatal(self) -> None:
        """Qt dialog if Step4 left a fatal quantity-balance marker (like main.py)."""
        try:
            from RFQ.tags_rfp_compare.step4.step4_quantity_balance import (
                find_latest_fatal_result_dir,
            )

            cfg = load_config()
            base = (cfg.get("paths") or {}).get("result_dir_base", "")
            if not base:
                return
            result_dir = find_latest_fatal_result_dir(base)
            if not result_dir:
                return
            marker = Path(result_dir) / "quantity_balance_fatal.json"
            if not marker.is_file():
                return
            data = json.loads(marker.read_text(encoding="utf-8"))
            msg = data.get("message", "Ошибка баланса количеств Step4.")
            QMessageBox.critical(self, "Баланс количеств Step4", str(msg))
        except Exception as exc:
            self._rfp_run_panel.monitor.append_log(
                f"[warn] Не удалось показать диалог баланса: {exc}\n"
            )

    def _run_rfp_check(self) -> None:
        root = _project_root()
        argv = build_aggregate_tags_argv(root)
        self._start_process_job(_JOB_TITLE_RFP, argv, tab_index=_TAB_INDEX_RFP_RUN)

    def _run_asbuild_check(self) -> None:
        root = _project_root()
        config_path = get_asbuild_config_path()
        if not Path(config_path).is_file():
            save_asbuild_config(load_asbuild_config())
        argv = build_aggregate_tags_argv(root, asbuild_config_path=config_path)
        self._start_process_job(_JOB_TITLE_ASBUILD, argv, tab_index=_TAB_INDEX_RFP_RUN)

    def _run_rfp_parts(self) -> None:
        from RFQ.rfp_parts.analyze_rfp_parts import make_reports_out_dir

        out_dir = make_reports_out_dir()
        out_dir.mkdir(parents=True, exist_ok=True)
        self._rfp_parts_panel.set_last_reports_dir(out_dir)

        # Backend: RFQ/rfp_parts (python -m RFQ.rfp_parts)
        # Checklist vs folders runs inside analyze by default.
        argv: list[str] = [
            sys.executable,
            "-X",
            "utf8",
            "-m",
            "RFQ.rfp_parts",
            "--out-dir",
            str(out_dir),
        ]
        self._start_process_job(
            _JOB_TITLE_RFP_PARTS, argv, tab_index=_TAB_INDEX_RFP_PARTS
        )

    def _ds_job_paths(self) -> dict[str, str]:
        return self._rfp_parts_panel.ds_job_paths()

    def _run_ds_baseline(self) -> None:
        from RFQ.rfp_parts.ds_hybrid_preflight import ds_baseline_output_dir

        paths = self._ds_job_paths()
        out_dir = ds_baseline_output_dir()
        self._rfp_parts_panel.set_last_reports_dir(out_dir)
        self._start_job(
            _JOB_TITLE_DS_BASELINE,
            run_ds_baseline_job,
            paths["source_root"],
            paths["registry_path"],
            out_dir,
            paths["ul_root"],
            job_tab="rfp_parts",
        )

    def _run_ds_hybrid(self) -> None:
        from RFQ.rfp_parts.ds_hybrid_preflight import ds_hybrid_output_dir

        paths = self._ds_job_paths()
        out_dir = ds_hybrid_output_dir()
        self._rfp_parts_panel.set_last_reports_dir(out_dir)
        self._start_job(
            _JOB_TITLE_DS_HYBRID,
            run_ds_hybrid_job,
            paths["source_root"],
            paths["registry_path"],
            out_dir,
            paths["ul_root"],
            paths["rfp_root"],
            job_tab="rfp_parts",
        )

    def _run_ds_coverage(self) -> None:
        paths = self._ds_job_paths()
        self._start_job(
            _JOB_TITLE_DS_COVERAGE,
            run_ds_coverage_job,
            paths["source_root"],
            paths["registry_path"],
            None,
            paths["ul_root"],
            paths["rfp_root"],
            job_tab="rfp_parts",
        )

    def _run_ds_registry(self) -> None:
        from RFQ.rfp_parts.ds_registry import DEFAULT_RFP_BASE

        paths = self._ds_job_paths()
        self._start_job(
            _JOB_TITLE_DS_REGISTRY,
            run_ds_registry_check_job,
            paths["registry_path"],
            DEFAULT_RFP_BASE,
            paths["ul_root"],
            job_tab="rfp_parts",
        )

    def _run_rfp_ds_id(self) -> None:
        self._start_job(_JOB_TITLE_RFP_DS_ID, run_ds_id_coverage_job)

    def _run_rfp_ds_mp(self) -> None:
        self._start_job(_JOB_TITLE_RFP_DS_MP, run_ds_roster_compare_job)

    def _open_ds_manager_book(self) -> None:
        result = get_last_ds_roster_compare()
        path = ""
        if result is not None:
            if result.sync is not None and result.sync.saved_path:
                path = result.sync.saved_path
            else:
                path = str(result.matrix_path)
        if not path:
            path = str(resolve_ds_manager_matrix_path())
        if not path or not Path(path).is_file():
            QMessageBox.information(
                self,
                "Список ДС — фамилии МП",
                f"Файл не найден:\n{path or '—'}",
            )
            return
        try:
            os.startfile(path)  # type: ignore[attr-defined]
        except Exception as exc:
            QMessageBox.warning(
                self,
                "Список ДС — фамилии МП",
                f"Не удалось открыть:\n{exc}",
            )

    def _open_latest_step4(self) -> None:
        path = find_latest_step4_result_file()
        if not path:
            QMessageBox.information(
                self,
                "Результат Step4",
                "Файл Шаг4_Сопоставление_RFP_MTO_*.xlsx не найден.",
            )
            return
        try:
            os.startfile(path)  # type: ignore[attr-defined]
        except Exception as exc:
            QMessageBox.warning(self, "Результат Step4", f"Не удалось открыть:\n{exc}")

    def _open_latest_bcc_matrix(self) -> None:
        path = find_latest_bcc_matrix_file()
        if not path:
            QMessageBox.information(
                self,
                "Накопительная матрица BCC",
                "Файл Шаг4_Матрица_BCC_*.xlsx не найден.",
            )
            return
        try:
            os.startfile(path)  # type: ignore[attr-defined]
        except Exception as exc:
            QMessageBox.warning(
                self,
                "Накопительная матрица BCC",
                f"Не удалось открыть:\n{exc}",
            )


    def _function_job_monitor(self) -> JobMonitorPanel:
        if self._function_job_tab == "misc":
            return self._misc_run_panel.monitor
        if self._function_job_tab == "rfp_parts":
            return self._rfp_parts_panel.monitor
        if self._function_job_tab == "ds_id":
            return self._rfp_ds_id_panel.monitor
        if self._function_job_tab == "ds_mp":
            return self._rfp_ds_mp_panel.monitor
        if self._last_job_title in (_JOB_TITLE_TSD_PACKING, _JOB_TITLE_TSD_ZINOVIEV):
            return self._tsd_packing_panel.monitor
        if self._last_job_title == _JOB_TITLE_UPD:
            return self._upd_panel.monitor
        return self._monitor

    @Slot(str)
    def _on_log(self, chunk: str) -> None:
        monitor = self._function_job_monitor()
        monitor.append_log(chunk)
        if self._function_job_tab != "rfp_parts":
            return
        for line in chunk.splitlines():
            match = re.search(r"\[ds progress\] FRACTION: (\d+)/(\d+)", line)
            if match:
                monitor.set_progress_fraction(int(match.group(1)), int(match.group(2)))

    @Slot(str)
    def _on_runner_error(self, message: str) -> None:
        self._function_job_monitor().append_log(f"[error] {message}\n")

    @Slot(bool, str, object)
    def _on_finished(self, success: bool, message: str, result_path: object) -> None:
        rp = str(result_path) if result_path else None
        if self._last_job_title in _JOB_TITLES_DS_COCKPIT:
            self._rfp_parts_panel.monitor.finish_job(success, message, rp)
            self._rfp_parts_panel.on_ds_job_finished(success, message, rp)
            return
        if self._last_job_title == _JOB_TITLE_RFP_DS_ID:
            self._rfp_ds_id_panel.monitor.finish_job(success, message, None)
            self._rfp_ds_id_panel.on_job_finished(success, message)
            return
        if self._last_job_title == _JOB_TITLE_RFP_DS_MP:
            self._rfp_ds_mp_panel.monitor.finish_job(success, message, rp)
            self._rfp_ds_mp_panel.on_job_finished(success, message)
            return
        if self._last_job_title in (_JOB_TITLE_TSD_PACKING, _JOB_TITLE_TSD_ZINOVIEV):
            self._tsd_packing_panel.monitor.finish_job(success, message, rp)
            self._tsd_packing_panel.on_job_finished(success, message, rp)
        elif self._last_job_title == _JOB_TITLE_UPD:
            self._upd_panel.monitor.finish_job(success, message, rp)
            self._upd_panel.on_job_finished(success, message, rp)
            if success and rp:
                self._persist_gui_path("last_upd_summary_file", rp)
        else:
            self._function_job_monitor().finish_job(success, message, rp)
        if self._last_job_title in (
            "Grouped ДС vs MTO vs RFQ",
            "Только RFQ → Excel (кэш ДС↔MTO)",
        ):
            audit = get_last_rfq_quantity_audit() if success else None
            if audit is not None:
                style = (
                    "color: #0a5; font-size: 11px;"
                    if audit.quantities_consistent
                    else "color: #a30; font-size: 11px; font-weight: bold;"
                )
                self._rfq_audit_label.setStyleSheet(style)
                self._rfq_audit_label.setText(audit.format_short())
            elif not success:
                self._rfq_audit_label.clear()
            packing_audit = get_last_packing_compare_audit() if success else None
            if packing_audit is not None:
                packing_ok = packing_audit.quality.value == "ok"
                style = (
                    "color: #0a5; font-size: 11px;"
                    if packing_ok
                    else "color: #a30; font-size: 11px; font-weight: bold;"
                )
                self._packing_audit_label.setStyleSheet(style)
                self._packing_audit_label.setText(packing_audit.format_short())
            elif not success:
                self._packing_audit_label.clear()
        if success and rp:
            if (
                self._last_job_title == "Объединить ДС из папки"
                and rp.lower().endswith(".xlsx")
                and os.path.isfile(rp)
            ):
                self._edit_ds.setText(rp)
                self._persist_gui_path("last_ds_file", rp)
            elif self._last_job_title in (
                "Grouped ДС vs MTO vs RFQ",
                "Только RFQ → Excel (кэш ДС↔MTO)",
            ):
                paths = normalize_gui_paths(
                    load_ds_compare_config().get("gui_paths")
                )
                refreshed_ds_path = paths.get("last_ds_file", "")
                if refreshed_ds_path:
                    self._edit_ds.setText(refreshed_ds_path)
            try:
                open_dir(rp)
            except Exception:
                pass
        elif (
            self._last_job_title
            in (_JOB_TITLE_TSD_PACKING, _JOB_TITLE_TSD_ZINOVIEV, _JOB_TITLE_UPD)
            and rp
            and os.path.exists(rp)
        ):
            try:
                open_dir(rp)
            except Exception:
                pass

    def _run_merge_ds(self) -> None:
        path = self._save_merge_dir_from_field()
        if not path:
            return
        self._start_job("Объединить ДС из папки", run_merge_ds_dir, path)

    def _run_tsd_packing(self, path: str) -> None:
        self._start_job(_JOB_TITLE_TSD_PACKING, run_tsd_packing_load, path)

    def _run_upd_load(self, path: str) -> None:
        self._start_job(_JOB_TITLE_UPD, run_upd_load, path)

    def _run_tsd_zinoviev_compare(self) -> None:
        self._start_job(_JOB_TITLE_TSD_ZINOVIEV, run_tsd_zinoviev_compare)

    def _run_ds_mto(self) -> None:
        path = self._save_ds_path_from_field()
        if not path:
            return
        self._start_job("Список ДС vs MTO", run_ds_mto_file, path)

    def _run_grouped_ds_mto(self) -> None:
        path = self._save_ds_path_from_field()
        if not path:
            return
        rfq_path = self._save_rfq_path_from_field()
        self._start_job(
            "Grouped ДС vs MTO vs RFQ",
            run_grouped_ds_mto_file,
            path,
            rfq_path,
        )

    def _run_grouped_rfq_only(self) -> None:
        path = self._save_ds_path_from_field()
        if not path:
            return
        rfq_path = self._save_rfq_path_from_field()
        if not rfq_path:
            return
        self._start_job(
            "Только RFQ → Excel (кэш ДС↔MTO)",
            run_grouped_ds_mto_rfq_only_file,
            path,
            rfq_path,
        )

    def _run_mto_ds(self) -> None:
        path = self._save_ds_path_from_field()
        if not path:
            return
        self._start_job("MTO vs Список ДС", run_mto_ds_file, path)

    def _run_misc_folder_job(self, title: str, fn: Callable) -> None:
        path = QFileDialog.getExistingDirectory(self, title, "")
        if not path:
            return
        self._start_job(title, fn, path, job_tab="misc")

    def _run_misc_file_job(
        self,
        title: str,
        fn: Callable,
        extra_args: tuple[object, ...] = (),
        file_filter: str = _EXCEL_FILTER,
    ) -> None:
        path, _ = QFileDialog.getOpenFileName(self, title, "", file_filter)
        if not path:
            return
        self._start_job(title, fn, path, *extra_args, job_tab="misc")

    def _run_misc_zip(self) -> None:
        self._start_job(
            "Собрать ZIP для коллег",
            run_release_zip,
            _project_root(),
            job_tab="misc",
        )

    def _show_misc_help(self, help_key: str) -> None:
        show_help_dialog(self, initial_key=help_key)

    def _launch_pdf_v2_monitor(self) -> None:
        """Start ``python -m pdf_v2_monitor`` in a separate process."""
        cmd = [sys.executable, "-m", "pdf_v2_monitor"]
        root = _project_root()
        try:
            subprocess.Popen(cmd, cwd=str(root))
        except OSError as exc:
            QMessageBox.warning(
                self,
                "Запуск монитора",
                f"Не удалось запустить процесс:\n{exc}",
            )
            return
        self._tabs.setCurrentIndex(_TAB_INDEX_MISC_RUN)
        self._misc_run_panel.monitor.append_log(
            f"Started: {' '.join(cmd)} cwd={root}\n"
        )

    def _launch_rd_catalog(self) -> None:
        """Start ``python -m rd_catalog`` in a separate process."""

        cmd = [sys.executable, "-m", "rd_catalog"]
        root = _project_root()
        try:
            subprocess.Popen(cmd, cwd=str(root))
        except OSError as exc:
            QMessageBox.warning(
                self,
                "Запуск каталога РД",
                f"Не удалось запустить процесс:\n{exc}",
            )
            return
        self._tabs.setCurrentIndex(_TAB_INDEX_MISC_RUN)
        self._misc_run_panel.monitor.append_log(
            f"Started: {' '.join(cmd)} cwd={root}\n"
        )
