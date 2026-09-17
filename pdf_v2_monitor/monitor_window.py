"""PDF v2 control center: run, settings, live monitor, and reports."""

from __future__ import annotations

import copy
import html
import json
from collections.abc import Iterable
from datetime import datetime
import os
import shutil
import subprocess
import sys
import time
from typing import Any

from PySide6.QtCore import QEvent, QSettings, Qt, QTimer
from PySide6.QtGui import (
    QAction,
    QBrush,
    QColor,
    QFont,
    QGuiApplication,
    QIcon,
    QResizeEvent,
)
from PySide6.QtWidgets import (
    QAbstractItemView,
    QApplication,
    QCheckBox,
    QComboBox,
    QDialog,
    QDialogButtonBox,
    QDoubleSpinBox,
    QFileDialog,
    QFrame,
    QFormLayout,
    QGridLayout,
    QGroupBox,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QMenu,
    QMenuBar,
    QMessageBox,
    QPlainTextEdit,
    QProgressBar,
    QPushButton,
    QScrollArea,
    QSizePolicy,
    QStyledItemDelegate,
    QSpinBox,
    QSplitter,
    QTabWidget,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from main_v2.appearance import apply_theme as apply_main_v2_theme
from pdf_parsing_v2.v2_config import (
    GRID_DETECTED_PREFILTER_KEYS,
    MONITOR_UI_THEME_LABELS,
    PER_PAGE_DOC_TYPES,
    excel_sanitize_illegal_chars_enabled,
    get_default_monitor_ui_config,
    get_default_v2_config,
    load_v2_config,
    normalize_monitor_ui_theme,
    normalize_tag_analysis_ui_config,
    normalize_v2_config,
    resolve_templates_dir,
    save_v2_config,
)
from pdf_v2_monitor.layout_persistence import DebouncedLayoutSaver
from pdf_parsing_v2_rules.output import (
    NORMCONTROL_SUMMARY_PLACEHOLDER_FIELD,
    load_normcontrol_report_from_workbook,
    load_review_state_from_workbook,
    save_review_state_to_workbook,
)
from utils import string_parsing
from pdf_v2_monitor.monitor_callback import MonitorCallback
from pdf_v2_monitor.pipeline_thread import PipelineThread

_COL_IDX = 0
_COL_FILE = 1
_COL_EXTRACT = 2
_COL_GRID_MISMATCH = 3
_COL_PAGES = 4
_COL_EXTRACT_TIME = 5
_COL_TAGS = 6
_COL_TAGS_TIME = 7
_COL_ERROR = 8
_CLR_AFFINE_WARN = QColor(255, 243, 205)
_CLR_AFFINE_BAD = QColor(248, 215, 218)

_STAGE_COL_ITEM = 2
_STAGE_COL_DETAIL = 5

_STATUS_QUEUED = "В очереди"
_STATUS_RUNNING = "Обработка..."
_STATUS_DONE = "Готово"
_STATUS_ERROR = "Ошибка"

_CLR_QUEUED = QColor(180, 180, 180)
_CLR_RUNNING = QColor(70, 130, 230)
_CLR_DONE = QColor(60, 180, 75)
_CLR_ERROR = QColor(220, 50, 50)
_EXTRACTION_DETAIL_THRESHOLD_SEC = 0.2
_RECENT_PDF_PATHS_KEY = "ui/recent_pdf_paths"
_RECENT_PDF_PATHS_LIMIT = 5
_MONITOR_WINDOW_TITLE_BASE = "PDF v2 Control Center"
_TAB_INDEX_RUN = 0
_TAB_INDEX_NORMCONTROL = 1
_TAB_INDEX_OD = 2
_TAB_INDEX_TAG_ANALYSIS = 3
_TAB_INDEX_REPORTS = 4
_DEFAULT_OD_HEADERS = [
    "Doc_Title",
    "Page_Format",
    "Page_Count:0",
    "Document_name",
    "Document_Revision",
    "Комплект",
]
_OD_TABLE_DEFAULT_ROW_HEIGHT = 44


def _monitor_table_selection_hover_stylesheet() -> str:
    """Selection/hover QSS from current Qt color scheme (call after ``setColorScheme`` / palette change)."""
    scheme = QGuiApplication.styleHints().colorScheme()
    if scheme == Qt.ColorScheme.Dark:
        return (
            "QTableWidget::item:selected { background-color: #1e3a5f; color: #e2e8f0; } "
            "QTableWidget::item:hover { background-color: #334155; color: #e2e8f0; }"
        )
    return (
        "QTableWidget::item:selected { background-color: #dbeafe; color: #000000; } "
        "QTableWidget::item:hover { background-color: #f5f5f5; color: #000000; }"
    )
_NON_NATIVE_FILE_DIALOG_OPTION = QFileDialog.Option.DontUseNativeDialog
_MONITOR_WIN_MIN_W = 1100
_MONITOR_WIN_MIN_H = 760
_MONITOR_WIN_MAX_DIM = 20000
_MONITOR_WIN_DEFAULT_W = 1350
_MONITOR_WIN_DEFAULT_H = 1060
# Диалог настроек JSON: оценка высоты области с двумя колонками (Fusion/Windows, ~96 DPI).
# Верх: строка кнопок «Сохранить/Перечитать» + кнопка «Обновить проекты» ≈ 80 px (margins + виджеты).
# Правая колонка обычно выше: grid_m 5 + frame 6 + prefilter 2 + debug 8 + excel 1 = 22 ряда формы;
# левая: monitor 3 + appearance 1 + pipeline 5 + per-page 6 (в т.ч. сетка типов 3 строки) + tags 2 ≈ 17 рядов.
# Ряд QFormLayout ≈ 26–30 px; заголовок QGroupBox + отступы ≈ +32–40 px на группу → правая колонка
# ≈ 22×27 + 5×38 + 4×10 (spacing) ≈ 800–850 px; с запасом DPI/стиля контент scroll ≈ 880–920 px.
# Итого клиент ≈ 80 + 900 ≈ 980 px по вертикали; default height ≈ 1000 даёт viewport scroll ~900+ и почти
# убирает прокрутку на 1080p; при меньшем экране остаётся QScrollArea.
_SETTINGS_DIALOG_DEFAULT_W = 1040
_SETTINGS_DIALOG_DEFAULT_H = 1000
_SETTINGS_DIALOG_MIN_H = 780
_PER_PAGE_DOC_LABELS: dict[str, str] = {
    "DW": "DW",
    "WIR": "WIR",
    "LAY": "LAY",
    "CAE": "CAE",
    "GA": "GA",
    "PL": "PL",
    "NI": "NI",
    "MTO": "MTO",
    "BOE": "BOE",
    "BOM": "BOM",
    "BOQ": "BOQ",
    "OD": "OD",
    "CJ": "CJ",
    "VO": "VO",
}


def _monitor_size_from_cfg(cfg: dict[str, Any]) -> tuple[int, int]:
    """Return clamped (width, height) for the control-center window from v2 config."""
    try:
        w = int(cfg.get("monitor_window_width", _MONITOR_WIN_DEFAULT_W) or _MONITOR_WIN_DEFAULT_W)
    except (TypeError, ValueError):
        w = _MONITOR_WIN_DEFAULT_W
    try:
        h = int(cfg.get("monitor_window_height", _MONITOR_WIN_DEFAULT_H) or _MONITOR_WIN_DEFAULT_H)
    except (TypeError, ValueError):
        h = _MONITOR_WIN_DEFAULT_H
    w = max(_MONITOR_WIN_MIN_W, min(w, _MONITOR_WIN_MAX_DIM))
    h = max(_MONITOR_WIN_MIN_H, min(h, _MONITOR_WIN_MAX_DIM))
    return w, h
_ROLE_ROW_KEY = int(Qt.ItemDataRole.UserRole)
_ROLE_SORT_VALUE = int(Qt.ItemDataRole.UserRole) + 1
# Tag analysis detail / invalid tables: row dict + column key for restoring text after edit.
_TAG_ANALYSIS_ROW_DICT_ROLE = int(Qt.ItemDataRole.UserRole)
_TAG_ANALYSIS_COLKEY_ROLE = int(Qt.ItemDataRole.UserRole) + 1
# Tag analysis filter combos: userData (stable, not tied to Russian labels).
_TAG_FILT_REASON_ALL = "all"
_TAG_FILT_REASON_NOT_FOUND = "not_found"
_TAG_FILT_REASON_HIDDEN = "hidden_by_rule"
_TAG_FILT_TITUL_ALL = "all"
_TAG_FILT_TITUL_EQ_MTO = "eq_mto"
_TAG_FILT_TITUL_NEQ_MTO = "neq_mto"
_TAG_FILTER_BAR_MIN_H = 44
# Список проверок «Анализ тегов»: без находок (severity info) — мягкий зелёный фон.
_CLR_TAG_CHECK_ROW_OK = QColor(220, 242, 228)
_NK_COL_RESOLVED = 0
_NK_COL_STATUS = 1
_NK_COL_CODE = 2
_NK_COL_DESCRIPTION = 3
_NK_COL_DOC = 4
_NK_COL_PAGE = 5
_NK_COL_TEXT = 6
_NK_COL_RESOLVED_AT = 7
_NK_COL_COMMENT = 8
_CLR_NK_RESOLVED = QColor(226, 240, 217)
_NK_COLUMN_WIDTHS_KEY = "ui/nk_column_widths"
_NK_TEXT_COLUMNS = {_NK_COL_DESCRIPTION, _NK_COL_TEXT, _NK_COL_COMMENT}
_NK_ROW_EXTRA_HEIGHT = 14
_NK_TABLE_DEFAULT_ROW_HEIGHT = 120
_NK_FILTER_ALL = "Все"
# Filter control widths (NK xlsx column widths × ~8–9 px + margin)
_NK_FILTER_W_CODE = 108
_NK_FILTER_W_DOC = 288
_NK_FILTER_W_PAGE = 96
_NK_FILTER_W_TEXT = _NK_FILTER_W_DOC


def _nk_row_code_str(row: dict[str, Any]) -> str:
    """Normcontrol row check code as string (matches table column)."""
    return str(row.get("c_code", "") or "")


def _nk_sorted_unique_codes(rows: Iterable[dict[str, Any]]) -> list[str]:
    codes = {_nk_row_code_str(r) for r in rows if _nk_row_code_str(r)}
    return sorted(codes, key=lambda s: (int(s) if s.isdigit() else s, s))


def _nk_sorted_unique_doc_names(rows: Iterable[dict[str, Any]]) -> list[str]:
    names = {str(r.get("doc_name", "") or "").strip() for r in rows}
    names.discard("")
    return sorted(names, key=str.lower)


def _nk_row_page_str(row: dict[str, Any]) -> str:
    """Canonical page label for filters (matches table display)."""
    return str(row.get("page_num", "") or "").strip()


def _nk_sorted_unique_pages(rows: Iterable[dict[str, Any]]) -> list[str]:
    pages = {_nk_row_page_str(r) for r in rows if _nk_row_page_str(r)}
    return sorted(pages, key=lambda s: (int(s), s) if s.isdigit() else (10**12, s))


def _nk_is_summary_placeholder(row: dict[str, Any]) -> bool:
    return bool(row.get(NORMCONTROL_SUMMARY_PLACEHOLDER_FIELD))


def apply_monitor_ui_theme(theme: Any) -> None:
    """Apply control-center UI theme (Qt ColorScheme or Fusion palettes from main_v2)."""
    app = QApplication.instance()
    if app is None:
        return
    key = normalize_monitor_ui_theme(theme)
    hints = app.styleHints()
    if key in ("light_high_contrast", "dark_high_contrast"):
        apply_main_v2_theme(app, key)
        hints.setColorScheme(
            Qt.ColorScheme.Light if key == "light_high_contrast" else Qt.ColorScheme.Dark
        )
        return
    apply_main_v2_theme(app, "system")
    scheme = Qt.ColorScheme.Unknown
    if key == "dark":
        scheme = Qt.ColorScheme.Dark
    elif key == "light":
        scheme = Qt.ColorScheme.Light
    hints.setColorScheme(scheme)


def load_monitor_app_icon() -> QIcon | None:
    """Return packaged ``ico/monitor_favicon.ico`` if present."""
    ico = os.path.join(os.path.dirname(os.path.abspath(__file__)), "ico", "monitor_favicon.ico")
    if os.path.isfile(ico):
        return QIcon(ico)
    return None


class _SortableTableWidgetItem(QTableWidgetItem):
    """Table item with custom sort role support."""

    def __lt__(self, other: object) -> bool:
        if not isinstance(other, QTableWidgetItem):
            return super().__lt__(other)
        left = self.data(_ROLE_SORT_VALUE)
        right = other.data(_ROLE_SORT_VALUE)
        if left is None or right is None:
            return super().__lt__(other)
        try:
            return left < right
        except TypeError:
            return str(left) < str(right)


class _NormcontrolItemDelegate(QStyledItemDelegate):
    """Read-only text editors for copy/select on double-click."""

    def createEditor(self, parent, option, index):
        column = index.column()
        if column in _NK_TEXT_COLUMNS:
            editor = QPlainTextEdit(parent)
            editor.setTabChangesFocus(True)
            editor.setReadOnly(column != _NK_COL_COMMENT)
            editor.document().setDocumentMargin(4)
            editor.setStyleSheet("QPlainTextEdit { padding: 2px 4px 2px 4px; }")
            return editor
        editor = QLineEdit(parent)
        editor.setReadOnly(column != _NK_COL_COMMENT)
        editor.setStyleSheet("QLineEdit { padding: 1px 4px 1px 4px; }")
        return editor

    def setEditorData(self, editor, index) -> None:
        value = str(index.data(Qt.ItemDataRole.DisplayRole) or "")
        if isinstance(editor, QPlainTextEdit):
            editor.setPlainText(value)
            return
        if isinstance(editor, QLineEdit):
            editor.setText(value)
            return
        super().setEditorData(editor, index)

    def setModelData(self, editor, model, index) -> None:
        if index.column() != _NK_COL_COMMENT:
            return
        if isinstance(editor, QPlainTextEdit):
            model.setData(index, editor.toPlainText(), Qt.ItemDataRole.EditRole)
            return
        if isinstance(editor, QLineEdit):
            model.setData(index, editor.text(), Qt.ItemDataRole.EditRole)
            return
        super().setModelData(editor, model, index)

    def sizeHint(self, option, index):
        hint = super().sizeHint(option, index)
        if index.column() in _NK_TEXT_COLUMNS:
            hint.setHeight(hint.height() + _NK_ROW_EXTRA_HEIGHT)
        return hint


class _NormcontrolTableWidget(QTableWidget):
    """NK results table: do not auto-pan horizontally when selecting off-screen cells."""

    def scrollTo(self, index, hint=QAbstractItemView.ScrollHint.EnsureVisible) -> None:
        hbar = self.horizontalScrollBar()
        hx = hbar.value() if hbar is not None else 0
        super().scrollTo(index, hint)
        if hbar is not None:
            hbar.setValue(hx)


class _OdTableReadOnlyDelegate(QStyledItemDelegate):
    """Read-only multiline cell editors for copy/select on double-click."""

    def createEditor(self, parent, option, index):
        editor = QPlainTextEdit(parent)
        editor.setReadOnly(True)
        editor.setTabChangesFocus(True)
        editor.document().setDocumentMargin(4)
        editor.setStyleSheet("QPlainTextEdit { padding: 2px 4px; }")
        return editor

    def setEditorData(self, editor, index) -> None:
        if isinstance(editor, QPlainTextEdit):
            editor.setPlainText(str(index.data(Qt.ItemDataRole.DisplayRole) or ""))
            return
        super().setEditorData(editor, index)

    def setModelData(self, editor, model, index) -> None:
        return

    def sizeHint(self, option, index):
        hint = super().sizeHint(option, index)
        hint.setHeight(hint.height() + 6)
        return hint


class _NormcontrolMultiPickDialog(QDialog):
    """Multi-select checklist for normcontrol filters (codes / documents / pages)."""

    def __init__(
        self,
        parent: QWidget,
        *,
        window_title: str,
        hint: str,
        items: list[str],
        current: frozenset[str] | None,
    ) -> None:
        super().__init__(parent)
        self.setWindowTitle(window_title)
        self.resize(380, 440)
        layout = QVBoxLayout(self)
        hint_lbl = QLabel(hint)
        hint_lbl.setWordWrap(True)
        layout.addWidget(hint_lbl)
        btn_row = QHBoxLayout()
        btn_all = QPushButton("Выбрать все")
        btn_clear = QPushButton("Сбросить все")
        btn_row.addWidget(btn_all)
        btn_row.addWidget(btn_clear)
        btn_row.addStretch(1)
        layout.addLayout(btn_row)
        self._list = QListWidget(self)
        self._list.setSelectionMode(QAbstractItemView.SelectionMode.NoSelection)
        for text in items:
            item = QListWidgetItem(text)
            item.setFlags(
                item.flags()
                | Qt.ItemFlag.ItemIsUserCheckable
                | Qt.ItemFlag.ItemIsEnabled
            )
            if current is None or text in current:
                item.setCheckState(Qt.CheckState.Checked)
            else:
                item.setCheckState(Qt.CheckState.Unchecked)
            self._list.addItem(item)
        layout.addWidget(self._list, 1)
        btn_all.clicked.connect(self._select_all_items)
        btn_clear.clicked.connect(self._clear_all_items)
        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
        )
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def _select_all_items(self) -> None:
        for i in range(self._list.count()):
            item = self._list.item(i)
            if item is not None:
                item.setCheckState(Qt.CheckState.Checked)

    def _clear_all_items(self) -> None:
        for i in range(self._list.count()):
            item = self._list.item(i)
            if item is not None:
                item.setCheckState(Qt.CheckState.Unchecked)

    def selected_or_none_if_all(self) -> frozenset[str] | None:
        """Return None when all or none checked (no filter); else frozenset of checked."""
        checked: list[str] = []
        for i in range(self._list.count()):
            item = self._list.item(i)
            if item is None:
                continue
            if item.checkState() == Qt.CheckState.Checked:
                checked.append(item.text())
        n = self._list.count()
        if n == 0 or len(checked) == n or not checked:
            return None
        return frozenset(checked)


_CLR_STAGE_ERROR_BG = QColor(255, 220, 220)
_CLR_STAGE_WARN_BG = QColor(255, 243, 205)
_CLR_WARN_TEXT = QColor(180, 83, 9)


def _stage_error_user_message(category: str, stage: str, error_msg: str) -> tuple[str, str]:
    """Return (title, body) for a short QMessageBox when a pipeline stage fails."""
    title = f"Ошибка этапа: {category} / {stage}"
    em = error_msg or ""
    if "IllegalCharacterError" in em or "cannot be used in worksheets" in em.lower():
        body = (
            "Не удалось записать текст в Excel: в данных есть символы, запрещённые "
            "форматом xlsx (часто невидимые управляющие символы из полей штампа или PDF).\n\n"
            "На вкладке Run в таблице этапов найдите эту строку (Category / Stage), "
            "полный текст ошибки в колонке Detail с переносами; двойной щелчок по Detail "
            "открывает ячейку для копирования."
        )
        return title, body
    if "PermissionError" in em or "Permission denied" in em:
        body = (
            "Не удалось сохранить файл: нет доступа или файл занят другой программой "
            "(закройте Excel и повторите)."
        )
        return title, body
    body = em.strip() or "Неизвестная ошибка. См. таблицу этапов на вкладке Run (колонка Detail)."
    # Full text stays in the stage table Detail column; avoid huge modal dialogs.
    if len(body) > 6000:
        body = body[:6000] + "\n\n… (продолжение в колонке Detail на вкладке Run)"
    return title, body


def _normalize_local_fs_path(path: str | None) -> str:
    """Normalize paths for Win32 APIs (UNC ``//srv/share``, mixed ``/`` and ``\\``)."""
    s = str(path or "").strip()
    if not s:
        return ""
    return os.path.normpath(s)


class MonitorWindow(QDialog):
    """Qt control center for the PDF v2 pipeline."""

    def __init__(
        self,
        pdf_path: str,
        cfg: dict[str, Any],
        project: str | None = None,
        *,
        initial_tab: str = "run",
        parent=None,
    ) -> None:
        super().__init__(parent)
        self._cfg = normalize_v2_config(cfg)
        self._pdf_path = pdf_path
        self._project = project
        self._initial_tab = initial_tab
        self._result_dir: str | None = None
        self._last_summary: dict[str, Any] = {}
        self._stage_rows_by_key: dict[tuple[str, str, str], int] = {}
        self._file_path_by_row: dict[int, str] = {}
        self._od_pdf_path: str | None = None
        self._monitor_rows: dict[int, dict[str, Any]] = {}
        self._uncovered_doc_type_by_file: dict[str, str] = {}
        self._t_start = 0.0
        self._n_errors = 0
        self._total_files = 0
        self._total_tasks = 0
        self._completed = 0
        self._current_phase = ""
        self._phase_start = 0.0
        self._phase_times: dict[str, float] = {}
        self._pipeline_started = False
        self._settings_store = QSettings("Documentation_PDF_out_NK", "pdf_v2_monitor")
        self._previewed_pdf_path = ""
        self._settings_dialog: QDialog | None = None
        self._normcontrol_all_rows: list[dict[str, Any]] = []
        self._normcontrol_summary_placeholder: dict[str, Any] | None = None
        self._normcontrol_rows_by_key: dict[str, dict[str, Any]] = {}
        self._normcontrol_xlsx_path: str = ""
        self._nk_code_filter: frozenset[str] | None = None
        self._nk_doc_filter: frozenset[str] | None = None
        self._nk_page_filter: frozenset[str] | None = None
        self._nk_known_doc_names: frozenset[str] = frozenset()
        self._nk_reloading_filter_widgets = False

        _icon = load_monitor_app_icon()
        if _icon is not None:
            self.setWindowIcon(_icon)
        self.setMinimumSize(_MONITOR_WIN_MIN_W, _MONITOR_WIN_MIN_H)

        self._build_ui()
        self._setup_timer()
        self._load_cfg_into_widgets(self._cfg)
        self._load_recent_pdf_paths()
        if pdf_path:
            self._set_pdf_path_text(pdf_path)
        self._refresh_project_combo()
        self._refresh_start_availability()
        self._select_initial_tab()
        self._update_window_title_with_size()

    def _update_window_title_with_size(self) -> None:
        """Title bar: base name plus window size as ширина×высота (ш×в)."""
        w, h = self.width(), self.height()
        self.setWindowTitle(f"{_MONITOR_WINDOW_TITLE_BASE} ({w}×{h})")

    def resizeEvent(self, event: QResizeEvent) -> None:
        super().resizeEvent(event)
        self._update_window_title_with_size()

    def changeEvent(self, event: QEvent) -> None:
        super().changeEvent(event)
        if event.type() == QEvent.Type.PaletteChange:
            self.refresh_ui_after_theme_change(self._cfg.get("monitor_ui_theme"))

    def _apply_monitor_table_selection_styles(self) -> None:
        """Re-apply table selection/hover QSS for the current color scheme."""
        qss = _monitor_table_selection_hover_stylesheet()
        self._table.setStyleSheet(qss)
        self._nk_table.setStyleSheet(qss)
        self._od_result_table.setStyleSheet(qss)
        self._tag_detail_table.setStyleSheet(qss)
        self._tag_invalid_table.setStyleSheet(qss)

    def refresh_ui_after_theme_change(self, theme: Any) -> None:
        """Re-apply palette-dependent UI: table chrome, repainted summary tables, monitor/stage rows."""
        _ = normalize_monitor_ui_theme(theme)
        self._apply_monitor_table_selection_styles()
        if self._last_summary:
            self._populate_reports(self._last_summary)
        for row in range(self._table.rowCount()):
            self._render_file_row(row)
        self._refresh_stage_table_error_backgrounds()

    def _refresh_stage_table_error_backgrounds(self) -> None:
        brush = QBrush(_CLR_STAGE_ERROR_BG)
        for row in range(self._stage_table.rowCount()):
            st = self._stage_table.item(row, 3)
            if st is None or st.text() != "error":
                continue
            for col in range(self._stage_table.columnCount()):
                item = self._stage_table.item(row, col)
                if item is not None:
                    item.setBackground(brush)

    def _build_ui(self) -> None:
        root = QVBoxLayout(self)
        root.setContentsMargins(12, 12, 12, 12)
        root.setSpacing(8)

        menu_bar = QMenuBar(self)
        menu_main = menu_bar.addMenu("Главная")
        menu_main.addAction("Настройки", self._open_settings_dialog)
        root.setMenuBar(menu_bar)

        self._tabs = QTabWidget(self)
        self._tabs.addTab(self._build_run_tab(), "Run")
        self._tabs.addTab(self._build_normcontrol_tab(), "Результат Нормоконтроля")
        self._tabs.addTab(self._build_od_table_tab(), "ОД")
        self._tabs.addTab(self._build_tag_analysis_tab(), "Анализ тегов")
        self._tabs.addTab(self._build_reports_tab(), "Reports")
        root.addWidget(self._tabs)
        self._settings_dialog = self._build_settings_dialog()

    def _build_run_tab(self) -> QWidget:
        tab = QWidget(self)
        layout = QVBoxLayout(tab)
        layout.setSpacing(12)

        folder_group = QGroupBox("Папка с PDF")
        folder_layout = QGridLayout(folder_group)
        folder_layout.setColumnStretch(1, 1)
        folder_layout.setHorizontalSpacing(8)
        self._combo_pdf_path = QComboBox()
        self._combo_pdf_path.setEditable(True)
        self._combo_pdf_path.setInsertPolicy(QComboBox.InsertPolicy.NoInsert)
        self._combo_pdf_path.lineEdit().textChanged.connect(self._on_pdf_path_changed)
        self._combo_pdf_path.currentTextChanged.connect(self._on_pdf_path_changed)
        self._btn_browse_pdf = QPushButton("Обзор…")
        self._btn_browse_pdf.clicked.connect(self._browse_pdf_folder)
        lbl_pdf_path = QLabel("PDF папка:")
        folder_layout.addWidget(lbl_pdf_path, 0, 0, alignment=Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
        folder_layout.addWidget(self._combo_pdf_path, 0, 1)
        folder_layout.addWidget(self._btn_browse_pdf, 0, 2)
        layout.addWidget(folder_group)

        buttons_group = QGroupBox("Действия")
        buttons_layout = QGridLayout(buttons_group)
        self._btn_start = QPushButton("Запустить pipeline")
        self._btn_start.clicked.connect(self._on_start_clicked)
        self._btn_open_editor = QPushButton("Открыть редактор шаблонов")
        self._btn_open_editor.clicked.connect(self._launch_template_editor)
        self._combo_project = QComboBox()
        self._combo_project.currentIndexChanged.connect(self._on_project_changed)
        buttons_layout.addWidget(self._btn_start, 0, 0)
        buttons_layout.addWidget(QLabel("Project:"), 0, 1)
        buttons_layout.addWidget(self._combo_project, 0, 2)
        buttons_layout.addWidget(self._btn_open_editor, 0, 3)
        buttons_layout.setColumnStretch(2, 1)
        layout.addWidget(buttons_group)

        self._lbl_effective_cfg = QLabel("")
        self._lbl_effective_cfg.setWordWrap(True)
        self._lbl_effective_cfg.setStyleSheet("font-size: 12px; color: #666;")
        layout.addWidget(self._lbl_effective_cfg)

        monitor_group = QGroupBox("Мониторинг выполнения")
        monitor_layout = QVBoxLayout(monitor_group)
        monitor_layout.setContentsMargins(8, 8, 8, 8)
        monitor_layout.addWidget(self._build_monitor_panel())
        layout.addWidget(monitor_group, 1)
        return tab

    def _build_settings_dialog(self) -> QDialog:
        dialog = QDialog(self)
        dialog.setWindowTitle("Настройки PDF v2")
        dialog.setMinimumWidth(920)
        dialog.setMinimumHeight(_SETTINGS_DIALOG_MIN_H)
        dialog.resize(_SETTINGS_DIALOG_DEFAULT_W, _SETTINGS_DIALOG_DEFAULT_H)
        layout = QVBoxLayout(dialog)

        actions = QHBoxLayout()
        self._btn_save_settings = QPushButton("Сохранить настройки")
        self._btn_save_settings.clicked.connect(self._save_settings)
        self._btn_reload_settings = QPushButton("Перечитать настройки")
        self._btn_reload_settings.clicked.connect(self._reload_settings)
        self._btn_refresh_projects = QPushButton("Обновить список проектов")
        self._btn_refresh_projects.clicked.connect(self._refresh_project_combo)
        self._btn_refresh_projects.setToolTip(
            "Перечитать список проектов из текущей папки шаблонов. "
            "Полезно после смены templates_dir или добавления новых шаблонов."
        )
        actions.addWidget(self._btn_save_settings)
        actions.addWidget(self._btn_reload_settings)
        actions.addWidget(self._btn_refresh_projects)
        actions.addStretch(1)
        layout.addLayout(actions)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.Shape.NoFrame)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)

        two_col = QWidget()
        h_main = QHBoxLayout(two_col)
        h_main.setSpacing(18)
        left_col = QVBoxLayout()
        right_col = QVBoxLayout()
        left_col.setSpacing(10)
        right_col.setSpacing(10)
        h_main.addLayout(left_col, 1)
        h_main.addLayout(right_col, 1)

        monitor_group = QGroupBox("Окно центра управления")
        monitor_form = QFormLayout(monitor_group)
        monitor_form.setLabelAlignment(Qt.AlignmentFlag.AlignRight)
        self._spin_monitor_width = QSpinBox()
        self._spin_monitor_width.setRange(_MONITOR_WIN_MIN_W, _MONITOR_WIN_MAX_DIM)
        self._spin_monitor_width.setSuffix(" px")
        self._spin_monitor_height = QSpinBox()
        self._spin_monitor_height.setRange(_MONITOR_WIN_MIN_H, _MONITOR_WIN_MAX_DIM)
        self._spin_monitor_height.setSuffix(" px")
        monitor_form.addRow("Ширина:", self._spin_monitor_width)
        monitor_form.addRow("Высота:", self._spin_monitor_height)
        self._btn_monitor_capture_size = QPushButton("Сохранить текущий размер окна")
        self._btn_monitor_capture_size.setToolTip(
            "Подставить текущие ширину и высоту окна в поля выше. "
            "Чтобы записать их в pdf_v2_config.json, нажмите «Сохранить настройки»."
        )
        self._btn_monitor_capture_size.clicked.connect(self._capture_monitor_window_size_to_widgets)
        monitor_form.addRow(self._btn_monitor_capture_size)
        left_col.addWidget(monitor_group)

        appearance_group = QGroupBox("Оформление интерфейса")
        appearance_form = QFormLayout(appearance_group)
        appearance_form.setLabelAlignment(Qt.AlignmentFlag.AlignRight)
        self._combo_monitor_theme = QComboBox()
        self._combo_monitor_theme.addItem("Как в системе", "system")
        self._combo_monitor_theme.addItem("Светлая", "light")
        self._combo_monitor_theme.addItem("Светлая контрастная", "light_high_contrast")
        self._combo_monitor_theme.addItem("Тёмная", "dark")
        self._combo_monitor_theme.addItem("Тёмная контрастная", "dark_high_contrast")
        self._combo_monitor_theme.currentIndexChanged.connect(
            lambda _i: apply_monitor_ui_theme_full(
                self._combo_monitor_theme.currentData(), window=self
            )
        )
        appearance_form.addRow("Тема:", self._combo_monitor_theme)
        left_col.addWidget(appearance_group)

        pipeline_group = QGroupBox("Конфигурация pipeline")
        pipeline_form = QFormLayout(pipeline_group)
        pipeline_form.setLabelAlignment(Qt.AlignmentFlag.AlignRight)
        self._chk_enabled = QCheckBox("Показывать v2 в main.py")
        pipeline_form.addRow("Включено:", self._chk_enabled)
        td_wrap = QWidget()
        td_layout = QHBoxLayout(td_wrap)
        td_layout.setContentsMargins(0, 0, 0, 0)
        self._edit_templates_dir = QLineEdit()
        self._btn_browse_templates = QPushButton("Обзор…")
        self._btn_browse_templates.clicked.connect(self._browse_templates_dir)
        td_layout.addWidget(self._edit_templates_dir, 1)
        td_layout.addWidget(self._btn_browse_templates)
        pipeline_form.addRow("Папка шаблонов:", td_wrap)
        self._chk_parallel = QCheckBox("Включить parallel extraction")
        self._chk_parallel.stateChanged.connect(lambda _v: self._update_effective_cfg_label())
        pipeline_form.addRow("Parallel extraction:", self._chk_parallel)
        self._spin_max_workers = QSpinBox()
        self._spin_max_workers.setRange(0, 64)
        self._spin_max_workers.setSpecialValueText("auto")
        self._spin_max_workers.valueChanged.connect(lambda _v: self._update_effective_cfg_label())
        pipeline_form.addRow("Потоки extract (max workers):", self._spin_max_workers)
        left_col.addWidget(pipeline_group)

        self._grp_per_page = QGroupBox("Постраничное извлечение (per-page)")
        per_page_form = QFormLayout(self._grp_per_page)
        per_page_form.setLabelAlignment(Qt.AlignmentFlag.AlignRight)
        self._chk_per_page_enabled = QCheckBox("Включить selective per-page parallel")
        self._chk_per_page_enabled.stateChanged.connect(lambda _v: self._update_effective_cfg_label())
        per_page_form.addRow("Режим:", self._chk_per_page_enabled)
        self._spin_per_page_min_pages = QSpinBox()
        self._spin_per_page_min_pages.setRange(1, 999)
        self._spin_per_page_min_pages.valueChanged.connect(lambda _v: self._update_effective_cfg_label())
        per_page_form.addRow("Мин. число страниц:", self._spin_per_page_min_pages)
        self._spin_per_page_workers = QSpinBox()
        self._spin_per_page_workers.setRange(0, 64)
        self._spin_per_page_workers.setSpecialValueText("auto")
        self._spin_per_page_workers.valueChanged.connect(lambda _v: self._update_effective_cfg_label())
        per_page_form.addRow("Потоки:", self._spin_per_page_workers)
        self._spin_per_page_chunk_size = QSpinBox()
        self._spin_per_page_chunk_size.setRange(0, 999)
        self._spin_per_page_chunk_size.setSpecialValueText("auto")
        self._spin_per_page_chunk_size.valueChanged.connect(lambda _v: self._update_effective_cfg_label())
        per_page_form.addRow("Размер чанка (стр.):", self._spin_per_page_chunk_size)
        doc_types_wrap = QWidget()
        doc_types_grid = QGridLayout(doc_types_wrap)
        doc_types_grid.setContentsMargins(0, 0, 0, 0)
        doc_types_grid.setHorizontalSpacing(10)
        doc_types_grid.setVerticalSpacing(4)
        self._chk_per_page_doc_types: dict[str, QCheckBox] = {}
        for idx, doc_type in enumerate(PER_PAGE_DOC_TYPES):
            checkbox = QCheckBox(_PER_PAGE_DOC_LABELS.get(doc_type, doc_type))
            checkbox.stateChanged.connect(lambda _v, _doc=doc_type: self._update_effective_cfg_label())
            self._chk_per_page_doc_types[doc_type] = checkbox
            row = idx // 4
            col = idx % 4
            doc_types_grid.addWidget(checkbox, row, col)
        per_page_form.addRow("Типы документов:", doc_types_wrap)
        left_col.addWidget(self._grp_per_page)

        tags_group = QGroupBox("Разбор тегов (после извлечения)")
        tags_form = QFormLayout(tags_group)
        tags_form.setLabelAlignment(Qt.AlignmentFlag.AlignRight)
        self._spin_tags_max_workers = QSpinBox()
        self._spin_tags_max_workers.setRange(0, 64)
        self._spin_tags_max_workers.setSpecialValueText("auto")
        self._spin_tags_max_workers.valueChanged.connect(lambda _v: self._update_effective_cfg_label())
        tags_form.addRow("Потоки (tags max workers):", self._spin_tags_max_workers)
        self._combo_tags_backend = QComboBox()
        self._combo_tags_backend.addItem("fitz", "fitz")
        self._combo_tags_backend.addItem("pdfminer", "pdfminer")
        self._combo_tags_backend.currentIndexChanged.connect(lambda _v: self._update_effective_cfg_label())
        tags_form.addRow("Движок текста для тегов:", self._combo_tags_backend)
        left_col.addWidget(tags_group)
        left_col.addStretch(1)

        grid_m_group = QGroupBox("Колонка «Разъезд» (пороги и веса)")
        grid_m_group.setToolTip(
            "Оценка несоответствия сетки штампа и PDF на вкладке Run (целое, больше = хуже). "
            "Настраиваются пороги подсветки и веса членов оценки."
        )
        grid_m_form = QFormLayout(grid_m_group)
        grid_m_form.setLabelAlignment(Qt.AlignmentFlag.AlignRight)
        self._spin_grid_ok = QSpinBox()
        self._spin_grid_ok.setRange(0, 99_999_999)
        self._spin_grid_ok.setToolTip(
            "Порог «нормально»: при оценке ≤ этого показывается «OK» без жёлтой/красной заливки. "
            "Оценка считается в движке из глобального |sx−sy|, суммы сдвигов линий (pt) и числа строк без ребра PDF."
        )
        self._spin_grid_warn = QSpinBox()
        self._spin_grid_warn.setRange(0, 99_999_999)
        self._spin_grid_warn.setToolTip(
            "Верхняя граница жёлтой подсветки числа оценки. Выше — красная. "
            "При сохранении не будет меньше порога «нормально»."
        )
        grid_m_form.addRow("Норма (≤ зелёный «OK»):", self._spin_grid_ok)
        grid_m_form.addRow("Жёлтый до (≤):", self._spin_grid_warn)

        self._spin_weight_global = QDoubleSpinBox()
        self._spin_weight_global.setRange(0.0, 10_000_000.0)
        self._spin_weight_global.setDecimals(1)
        self._spin_weight_global.setSingleStep(1000.0)
        self._spin_weight_global.setToolTip("Вес члена «глобаль»: умножается на |sx−sy| из подгонки bbox штампа.")
        self._spin_weight_walk = QDoubleSpinBox()
        self._spin_weight_walk.setRange(0.0, 5000.0)
        self._spin_weight_walk.setDecimals(2)
        self._spin_weight_walk.setSingleStep(1.0)
        self._spin_weight_walk.setToolTip(
            "Вес суммы |фактическая линия − интерполяция| по всем линиям каркаса (в пунктах PDF)."
        )
        self._spin_weight_no_match = QDoubleSpinBox()
        self._spin_weight_no_match.setRange(0.0, 500_000.0)
        self._spin_weight_no_match.setDecimals(1)
        self._spin_weight_no_match.setSingleStep(500.0)
        self._spin_weight_no_match.setToolTip(
            "Штраф за каждую линию шаблона без сопоставленного ребра PDF (интерполяция)."
        )
        grid_m_form.addRow("Вес |sx−sy| (глобаль):", self._spin_weight_global)
        grid_m_form.addRow("Вес суммы pt по линиям:", self._spin_weight_walk)
        grid_m_form.addRow("Вес строки без ребра PDF:", self._spin_weight_no_match)

        self._spin_grid_ok.valueChanged.connect(lambda _v: self._update_effective_cfg_label())
        self._spin_grid_warn.valueChanged.connect(lambda _v: self._update_effective_cfg_label())
        self._spin_weight_global.valueChanged.connect(lambda _v: self._update_effective_cfg_label())
        self._spin_weight_walk.valueChanged.connect(lambda _v: self._update_effective_cfg_label())
        self._spin_weight_no_match.valueChanged.connect(lambda _v: self._update_effective_cfg_label())
        right_col.addWidget(grid_m_group)

        frame_group = QGroupBox("Рамка листа и привязка таблиц (find_tables)")
        frame_form = QFormLayout(frame_group)
        frame_form.setLabelAlignment(Qt.AlignmentFlag.AlignRight)
        self._spin_frame_left_mm = QDoubleSpinBox()
        self._spin_frame_left_mm.setRange(0.0, 100.0)
        self._spin_frame_left_mm.setDecimals(1)
        self._spin_frame_left_mm.setSingleStep(1.0)
        frame_form.addRow("Полоса слева, мм:", self._spin_frame_left_mm)
        self._spin_frame_right_mm = QDoubleSpinBox()
        self._spin_frame_right_mm.setRange(0.0, 100.0)
        self._spin_frame_right_mm.setDecimals(1)
        self._spin_frame_right_mm.setSingleStep(1.0)
        frame_form.addRow("Полоса справа, мм:", self._spin_frame_right_mm)
        self._spin_frame_top_mm = QDoubleSpinBox()
        self._spin_frame_top_mm.setRange(0.0, 100.0)
        self._spin_frame_top_mm.setDecimals(1)
        self._spin_frame_top_mm.setSingleStep(1.0)
        frame_form.addRow("Полоса сверху, мм:", self._spin_frame_top_mm)
        self._spin_frame_bottom_mm = QDoubleSpinBox()
        self._spin_frame_bottom_mm.setRange(0.0, 100.0)
        self._spin_frame_bottom_mm.setDecimals(1)
        self._spin_frame_bottom_mm.setSingleStep(1.0)
        frame_form.addRow("Полоса снизу, мм:", self._spin_frame_bottom_mm)
        self._spin_snap_x = QDoubleSpinBox()
        self._spin_snap_x.setRange(0.0, 50.0)
        self._spin_snap_x.setDecimals(2)
        self._spin_snap_x.setSingleStep(0.1)
        frame_form.addRow("snap_x (допуск, pt):", self._spin_snap_x)
        self._spin_snap_y = QDoubleSpinBox()
        self._spin_snap_y.setRange(0.0, 50.0)
        self._spin_snap_y.setDecimals(2)
        self._spin_snap_y.setSingleStep(0.1)
        frame_form.addRow("snap_y (допуск, pt):", self._spin_snap_y)
        right_col.addWidget(frame_group)

        self._grp_grid_prefilter = QGroupBox("Префильтр «островов» в сетке")
        self._grp_grid_prefilter.setToolTip(
            "Ключи grid_detected_prefilter_*: отсекает лишние «острова» ячеек find_tables внутри search bbox."
        )
        pf_form = QFormLayout(self._grp_grid_prefilter)
        pf_form.setLabelAlignment(Qt.AlignmentFlag.AlignRight)
        self._chk_grid_pf_enabled = QCheckBox("Включить префильтр перед extent / alignment")
        self._chk_grid_pf_enabled.setToolTip(
            "Отсекает лишние «острова» ячеек find_tables внутри search bbox. "
            "При ложном разрезе одного штампа на два острова увеличьте пол ниже."
        )
        self._chk_grid_pf_enabled.stateChanged.connect(lambda _v: self._update_effective_cfg_label())
        pf_form.addRow("Режим:", self._chk_grid_pf_enabled)
        self._spin_grid_pf_adjacency_floor_pt = QDoubleSpinBox()
        self._spin_grid_pf_adjacency_floor_pt.setRange(0.25, 30.0)
        self._spin_grid_pf_adjacency_floor_pt.setDecimals(2)
        self._spin_grid_pf_adjacency_floor_pt.setSingleStep(0.25)
        self._spin_grid_pf_adjacency_floor_pt.setToolTip(
            "Нижняя граница допуска смежности ячеек в pt (ещё сверху берётся max с "
            "grid_tolerance_detected_mm×SCALE). "
            "~2.834 pt ≈ 1 mm при типографском pt/mm из coord_transform.SCALE. "
            "Значение по умолчанию 2.5 pt ≈ 0.88 mm."
        )
        self._spin_grid_pf_adjacency_floor_pt.valueChanged.connect(
            lambda _v: self._update_effective_cfg_label()
        )
        pf_form.addRow("Смежность: пол pt (adjacency_tol_floor):", self._spin_grid_pf_adjacency_floor_pt)
        right_col.addWidget(self._grp_grid_prefilter)

        debug_group = QGroupBox("Отладка и экспорт артефактов")
        debug_form = QFormLayout(debug_group)
        debug_form.setLabelAlignment(Qt.AlignmentFlag.AlignRight)
        self._chk_export_debug_excel = QCheckBox("Экспортировать debug Excel")
        debug_form.addRow("Debug Excel:", self._chk_export_debug_excel)
        self._chk_debug_visual = QCheckBox("Сохранять debug PNG")
        debug_form.addRow("Debug PNG:", self._chk_debug_visual)
        self._chk_debug_visual_overlay = QCheckBox("Overlay bbox поверх страницы")
        debug_form.addRow("Overlay на странице:", self._chk_debug_visual_overlay)
        self._chk_debug_verbose_log = QCheckBox("Подробный console log")
        debug_form.addRow("Подробный лог:", self._chk_debug_verbose_log)
        self._chk_timing_log = QCheckBox("Экспортировать timing Excel")
        self._chk_timing_log.stateChanged.connect(lambda _v: self._update_effective_cfg_label())
        debug_form.addRow("Timing log (xlsx):", self._chk_timing_log)
        self._combo_extraction_mode = QComboBox()
        self._combo_extraction_mode.addItem("char_center", "char_center")
        self._combo_extraction_mode.addItem("get_textbox", "get_textbox")
        debug_form.addRow("Режим извлечения текста:", self._combo_extraction_mode)
        dv_wrap = QWidget()
        dv_layout = QHBoxLayout(dv_wrap)
        dv_layout.setContentsMargins(0, 0, 0, 0)
        self._edit_debug_visual_dir = QLineEdit()
        self._btn_browse_visual_dir = QPushButton("Обзор…")
        self._btn_browse_visual_dir.clicked.connect(self._browse_debug_visual_dir)
        dv_layout.addWidget(self._edit_debug_visual_dir, 1)
        dv_layout.addWidget(self._btn_browse_visual_dir)
        debug_form.addRow("Папка debug PNG:", dv_wrap)
        right_col.addWidget(debug_group)

        excel_group = QGroupBox("Excel (openpyxl)")
        excel_form = QFormLayout(excel_group)
        excel_form.setLabelAlignment(Qt.AlignmentFlag.AlignRight)
        self._chk_excel_sanitize = QCheckBox(
            "Удалять из строк символы, запрещённые в xlsx (как при вставке в Excel)"
        )
        self._chk_excel_sanitize.setToolTip(
            "Касается файлов «результат проверки НК» и «v2_timing_log» при записи через openpyxl. "
            "Снимите галочку только для отладки: при выключении редкие «мусорные» символы из PDF "
            "могут снова приводить к ошибке сохранения."
        )
        self._chk_excel_sanitize.stateChanged.connect(lambda _v: self._update_effective_cfg_label())
        excel_form.addRow("Строки в xlsx:", self._chk_excel_sanitize)
        right_col.addWidget(excel_group)
        right_col.addStretch(1)

        scroll.setWidget(two_col)
        layout.addWidget(scroll, 1)
        return dialog

    def _build_monitor_panel(self) -> QWidget:
        panel = QWidget(self)
        layout = QVBoxLayout(panel)
        layout.setSpacing(8)

        top = QHBoxLayout()
        self._lbl_phase = QLabel("Готов к запуску")
        self._lbl_phase.setStyleSheet("font-weight: bold; font-size: 13px;")
        self._progress = QProgressBar()
        self._progress.setMinimum(0)
        self._progress.setMaximum(1)
        self._progress.setValue(0)
        self._progress.setFormat("0 / 0 задач")
        self._lbl_elapsed = QLabel("00:00")
        self._lbl_elapsed.setStyleSheet("font-family: monospace; font-size: 13px;")
        top.addWidget(self._lbl_phase)
        top.addWidget(self._progress, 1)
        top.addWidget(self._lbl_elapsed)
        layout.addLayout(top)

        self._lbl_summary = QLabel("После старта здесь появится сводка выполнения.")
        self._lbl_summary.setWordWrap(True)
        layout.addWidget(self._lbl_summary)

        splitter = QSplitter(Qt.Orientation.Vertical)
        self._table = self._create_table(
            ["#", "Файл", "Extract", "Разъезд", "Стр.", "Extract, c", "Tags", "Tags, c", "Ошибка"]
        )
        self._table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self._table.cellDoubleClicked.connect(self._open_file_from_monitor_row)
        self._table.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self._table.customContextMenuRequested.connect(self._show_monitor_context_menu)
        hdr = self._table.horizontalHeader()
        hdr.setSectionResizeMode(_COL_IDX, QHeaderView.ResizeMode.Fixed)
        hdr.setSectionResizeMode(_COL_FILE, QHeaderView.ResizeMode.Stretch)
        hdr.setSectionResizeMode(_COL_EXTRACT, QHeaderView.ResizeMode.Fixed)
        hdr.setSectionResizeMode(_COL_GRID_MISMATCH, QHeaderView.ResizeMode.Fixed)
        hdr.setSectionResizeMode(_COL_PAGES, QHeaderView.ResizeMode.Fixed)
        hdr.setSectionResizeMode(_COL_EXTRACT_TIME, QHeaderView.ResizeMode.Fixed)
        hdr.setSectionResizeMode(_COL_TAGS, QHeaderView.ResizeMode.Fixed)
        hdr.setSectionResizeMode(_COL_TAGS_TIME, QHeaderView.ResizeMode.Fixed)
        hdr.setSectionResizeMode(_COL_ERROR, QHeaderView.ResizeMode.Stretch)
        self._table.setColumnWidth(_COL_IDX, 45)
        self._table.setColumnWidth(_COL_EXTRACT, 100)
        self._table.setColumnWidth(_COL_GRID_MISMATCH, 96)
        self._table.setColumnWidth(_COL_PAGES, 55)
        self._table.setColumnWidth(_COL_EXTRACT_TIME, 90)
        self._table.setColumnWidth(_COL_TAGS, 100)
        self._table.setColumnWidth(_COL_TAGS_TIME, 90)
        gh_grid = self._table.horizontalHeaderItem(_COL_GRID_MISMATCH)
        if gh_grid is not None:
            gh_grid.setToolTip(
                "Разъезд сетки — насколько каркас штампа на странице PDF не совпал с шаблоном: "
                "целое число, чем выше — тем хуже. Совпадение без подсветки если «OK» "
                "(см. пороги в Настройки)."
            )
        splitter.addWidget(self._table)

        self._stage_table = self._create_table(
            ["Category", "Stage", "Item", "Status", "Time (s)", "Detail"]
        )
        self._stage_table.setWordWrap(True)
        self._stage_table.setTextElideMode(Qt.TextElideMode.ElideNone)
        self._stage_table.setEditTriggers(
            QAbstractItemView.EditTrigger.DoubleClicked | QAbstractItemView.EditTrigger.EditKeyPressed
        )
        self._stage_table.cellDoubleClicked.connect(self._open_file_from_stage_row)
        self._stage_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        self._stage_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        self._stage_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.Stretch)
        self._stage_table.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeMode.ResizeToContents)
        self._stage_table.horizontalHeader().setSectionResizeMode(4, QHeaderView.ResizeMode.ResizeToContents)
        self._stage_table.horizontalHeader().setSectionResizeMode(5, QHeaderView.ResizeMode.Stretch)
        splitter.addWidget(self._stage_table)
        splitter.setSizes([420, 260])
        layout.addWidget(splitter, 1)
        return panel

    def _build_reports_tab(self) -> QWidget:
        tab = QWidget(self)
        layout = QVBoxLayout(tab)

        buttons = QHBoxLayout()
        self._btn_open_dir = QPushButton("Открыть папку результатов")
        self._btn_open_dir.clicked.connect(self._open_result_dir)
        self._btn_open_timing = QPushButton("Открыть timing xlsx")
        self._btn_open_timing.clicked.connect(
            lambda: self._open_path(self._artifact_path("timing_path"))
        )
        self._btn_open_debug_report = QPushButton("Открыть debug xlsx")
        self._btn_open_debug_report.clicked.connect(
            lambda: self._open_path(self._artifact_path("debug_report_path"))
        )
        buttons.addWidget(self._btn_open_dir)
        buttons.addWidget(self._btn_open_timing)
        buttons.addWidget(self._btn_open_debug_report)
        buttons.addStretch(1)
        layout.addLayout(buttons)

        self._report_tabs = QTabWidget()
        self._report_summary = QPlainTextEdit()
        self._report_summary.setReadOnly(True)
        self._report_extraction = self._create_table([])
        self._report_tags = self._create_table([])
        self._report_od = self._create_table([])
        self._report_rules = self._create_table([])
        self._report_tabs.addTab(self._report_summary, "Summary")
        self._report_tabs.addTab(self._report_extraction, "Extraction")
        self._report_tabs.addTab(self._report_tags, "Tags")
        self._report_tabs.addTab(self._report_od, "OD")
        self._report_tabs.addTab(self._report_rules, "Rules")
        layout.addWidget(self._report_tabs, 1)
        return tab

    def _create_table(self, headers: list[str]) -> QTableWidget:
        table = QTableWidget(0, len(headers))
        if headers:
            table.setHorizontalHeaderLabels(headers)
        table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        table.verticalHeader().setVisible(False)
        table.setAlternatingRowColors(True)
        return table

    def _build_normcontrol_tab(self) -> QWidget:
        tab = QWidget(self)
        layout = QVBoxLayout(tab)
        layout.setSpacing(8)

        self._chk_nk_show_positive = QCheckBox("Показывать положительные")
        self._chk_nk_show_positive.stateChanged.connect(lambda _v: self._refresh_normcontrol_table())
        self._chk_nk_hide_resolved = QCheckBox("Скрыть исправленные")
        self._chk_nk_hide_resolved.stateChanged.connect(lambda _v: self._refresh_normcontrol_table())

        top = QHBoxLayout()
        self._btn_nk_save = QPushButton("Сохранить отметки")
        self._btn_nk_save.clicked.connect(self._save_normcontrol_review_state)
        self._btn_nk_import = QPushButton("Импортировать отметки из НК xlsx")
        self._btn_nk_import.clicked.connect(self._import_normcontrol_review_state)
        self._btn_open_normcontrol = QPushButton("Открыть результат проверки НК")
        self._btn_open_normcontrol.clicked.connect(self._open_normcontrol_xlsx)
        self._btn_nk_copy = QPushButton("Сохранить как…")
        self._btn_nk_copy.clicked.connect(self._save_normcontrol_copy_as)
        self._btn_nk_open_results_dir = QPushButton("Открыть папку результатов")
        self._btn_nk_open_results_dir.clicked.connect(self._open_result_dir)
        self._lbl_nk_stats = QLabel("НК: данных пока нет.")
        self._lbl_nk_stats.setWordWrap(True)
        top.addWidget(self._btn_nk_save)
        top.addWidget(self._btn_nk_import)
        top.addWidget(self._btn_open_normcontrol)
        top.addWidget(self._btn_nk_copy)
        top.addWidget(self._btn_nk_open_results_dir)
        top.addStretch(1)
        layout.addLayout(top)
        layout.addWidget(self._lbl_nk_stats)

        nk_filters = QGroupBox("Фильтры таблицы", self)
        nk_grid = QGridLayout(nk_filters)
        nk_grid.setHorizontalSpacing(6)
        nk_grid.setVerticalSpacing(4)
        nk_grid.setContentsMargins(6, 6, 6, 6)
        self._nk_cmb_code = QComboBox(nk_filters)
        self._nk_cmb_code.setFixedWidth(_NK_FILTER_W_CODE)
        self._nk_cmb_code.setMaxVisibleItems(24)
        self._nk_cmb_code.currentIndexChanged.connect(self._on_nk_code_combo_index_changed)
        self._btn_nk_codes = QPushButton("Коды…", nk_filters)
        self._btn_nk_codes.setToolTip("Несколько кодов проверки")
        self._btn_nk_codes.clicked.connect(self._open_nk_codes_filter_dialog)
        self._btn_nk_codes.setFixedWidth(self._btn_nk_codes.sizeHint().width())
        self._nk_cmb_doc = QComboBox(nk_filters)
        self._nk_cmb_doc.setFixedWidth(_NK_FILTER_W_DOC)
        self._nk_cmb_doc.setMaxVisibleItems(24)
        self._nk_cmb_doc.setToolTip(
            "Точное имя документа. Несколько — «Документы…». Подстрока по имени — «Текст»."
        )
        self._nk_cmb_doc.currentIndexChanged.connect(self._on_nk_doc_combo_index_changed)
        self._btn_nk_docs = QPushButton("Документы…", nk_filters)
        self._btn_nk_docs.setToolTip("Выбрать несколько документов")
        self._btn_nk_docs.clicked.connect(self._open_nk_docs_filter_dialog)
        self._btn_nk_docs.setFixedWidth(self._btn_nk_docs.sizeHint().width())
        self._nk_cmb_page = QComboBox(nk_filters)
        self._nk_cmb_page.setFixedWidth(_NK_FILTER_W_PAGE)
        self._nk_cmb_page.setMaxVisibleItems(24)
        self._nk_cmb_page.currentIndexChanged.connect(self._on_nk_page_combo_index_changed)
        self._btn_nk_pages = QPushButton("Страницы…", nk_filters)
        self._btn_nk_pages.setToolTip("Выбрать несколько страниц")
        self._btn_nk_pages.clicked.connect(self._open_nk_pages_filter_dialog)
        self._btn_nk_pages.setFixedWidth(self._btn_nk_pages.sizeHint().width())
        self._btn_nk_reset_filters = QPushButton("Сбросить фильтры", nk_filters)
        self._btn_nk_reset_filters.clicked.connect(self._reset_nk_column_filters)
        self._btn_nk_reset_filters.setFixedWidth(self._btn_nk_reset_filters.sizeHint().width())
        r0 = 0
        lbl_code = QLabel("Код:", nk_filters)
        lbl_code.setSizePolicy(QSizePolicy.Policy.Fixed, QSizePolicy.Policy.Fixed)
        lbl_doc = QLabel("Документ:", nk_filters)
        lbl_doc.setSizePolicy(QSizePolicy.Policy.Fixed, QSizePolicy.Policy.Fixed)
        lbl_page = QLabel("Стр.:", nk_filters)
        lbl_page.setSizePolicy(QSizePolicy.Policy.Fixed, QSizePolicy.Policy.Fixed)
        nk_grid.addWidget(lbl_code, r0, 0)
        nk_grid.addWidget(self._nk_cmb_code, r0, 1)
        nk_grid.addWidget(self._btn_nk_codes, r0, 2)
        nk_grid.addWidget(lbl_doc, r0, 3)
        nk_grid.addWidget(self._nk_cmb_doc, r0, 4)
        nk_grid.addWidget(self._btn_nk_docs, r0, 5)
        nk_grid.addWidget(lbl_page, r0, 6)
        nk_grid.addWidget(self._nk_cmb_page, r0, 7)
        nk_grid.addWidget(self._btn_nk_pages, r0, 8)
        nk_grid.addWidget(self._btn_nk_reset_filters, r0, 9)
        self._nk_filter_debounce = QTimer(self)
        self._nk_filter_debounce.setSingleShot(True)
        self._nk_filter_debounce.setInterval(180)
        self._nk_filter_debounce.timeout.connect(self._refresh_normcontrol_table)
        self._nk_ed_text = QLineEdit(nk_filters)
        self._nk_ed_text.setFixedWidth(_NK_FILTER_W_TEXT)
        self._nk_ed_text.setPlaceholderText("Текст: описание, документ, замечание, комментарий, код")
        self._nk_ed_text.textChanged.connect(self._schedule_nk_filter_debounced_refresh)
        self._nk_ed_text.editingFinished.connect(self._nk_flush_filter_refresh)
        self._nk_ed_text.returnPressed.connect(self._nk_flush_filter_refresh)
        lbl_text = QLabel("Текст:", nk_filters)
        lbl_text.setSizePolicy(QSizePolicy.Policy.Fixed, QSizePolicy.Policy.Fixed)
        nk_text_row = QHBoxLayout()
        nk_text_row.setSpacing(6)
        nk_text_row.setContentsMargins(0, 0, 0, 0)
        nk_text_row.addWidget(lbl_text)
        nk_text_row.addWidget(self._nk_ed_text)
        nk_text_row.addWidget(self._chk_nk_show_positive)
        nk_text_row.addWidget(self._chk_nk_hide_resolved)
        nk_text_row.addStretch(1)
        nk_grid.addLayout(nk_text_row, 1, 0, 1, 10)
        nk_filters.setSizePolicy(QSizePolicy.Policy.Maximum, QSizePolicy.Policy.Preferred)
        nk_filters_wrap = QHBoxLayout()
        nk_filters_wrap.setContentsMargins(0, 0, 0, 0)
        nk_filters_wrap.addWidget(nk_filters, 0, Qt.AlignmentFlag.AlignLeft)
        nk_filters_wrap.addStretch(1)
        layout.addLayout(nk_filters_wrap)

        self._nk_table = _NormcontrolTableWidget(0, 9, tab)
        self._nk_table.setHorizontalHeaderLabels(
            [
                "Исправлено",
                "Статус",
                "Код",
                "Описание",
                "Документ",
                "Стр.",
                "Замечание",
                "Дата отметки",
                "Комментарий",
            ]
        )
        self._nk_table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectItems)
        self._nk_table.setAlternatingRowColors(True)
        self._nk_table.setWordWrap(True)
        self._nk_table.setTextElideMode(Qt.TextElideMode.ElideNone)
        self._nk_table.setSortingEnabled(True)
        self._nk_table.setEditTriggers(
            QAbstractItemView.EditTrigger.DoubleClicked
            | QAbstractItemView.EditTrigger.EditKeyPressed
            | QAbstractItemView.EditTrigger.SelectedClicked
        )
        self._nk_table.setItemDelegate(_NormcontrolItemDelegate(self._nk_table))
        self._nk_table.itemChanged.connect(self._on_nk_item_changed)
        self._nk_table.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self._nk_table.customContextMenuRequested.connect(self._show_normcontrol_context_menu)
        header = self._nk_table.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        nk_vh = self._nk_table.verticalHeader()
        nk_vh.setSectionResizeMode(QHeaderView.ResizeMode.Fixed)
        nk_vh.setDefaultSectionSize(_NK_TABLE_DEFAULT_ROW_HEIGHT)
        self._nk_table.setColumnWidth(_NK_COL_RESOLVED, 90)
        self._nk_table.setColumnWidth(_NK_COL_STATUS, 90)
        self._nk_table.setColumnWidth(_NK_COL_CODE, 90)
        self._nk_table.setColumnWidth(_NK_COL_DESCRIPTION, 260)
        self._nk_table.setColumnWidth(_NK_COL_DOC, 260)
        self._nk_table.setColumnWidth(_NK_COL_PAGE, 70)
        self._nk_table.setColumnWidth(_NK_COL_TEXT, 520)
        self._nk_table.setColumnWidth(_NK_COL_RESOLVED_AT, 150)
        self._nk_table.setColumnWidth(_NK_COL_COMMENT, 220)
        self._restore_normcontrol_column_widths()
        header.sectionResized.connect(self._on_normcontrol_section_resized)
        layout.addWidget(self._nk_table, 1)
        self._set_normcontrol_buttons_enabled(False)
        self._reload_nk_filter_options()
        return tab

    def _build_od_table_tab(self) -> QWidget:
        """OD general-data table (same columns as console / PrettyTable)."""
        tab = QWidget(self)
        layout = QVBoxLayout(tab)
        layout.setSpacing(8)
        self._lbl_od_info = QLabel("")
        self._lbl_od_info.setWordWrap(True)
        self._lbl_od_info.setTextInteractionFlags(
            Qt.TextInteractionFlag.TextSelectableByMouse
        )
        layout.addWidget(self._lbl_od_info)
        self._od_result_table = QTableWidget(0, len(_DEFAULT_OD_HEADERS))
        self._od_result_table.setHorizontalHeaderLabels(list(_DEFAULT_OD_HEADERS))
        self._od_result_table.setWordWrap(True)
        self._od_result_table.setTextElideMode(Qt.TextElideMode.ElideNone)
        self._od_result_table.setAlternatingRowColors(True)
        self._od_result_table.setSelectionBehavior(
            QAbstractItemView.SelectionBehavior.SelectItems
        )
        self._od_result_table.setEditTriggers(
            QAbstractItemView.EditTrigger.DoubleClicked
            | QAbstractItemView.EditTrigger.EditKeyPressed
        )
        self._od_result_table.setItemDelegate(_OdTableReadOnlyDelegate(self._od_result_table))
        od_hdr = self._od_result_table.horizontalHeader()
        od_hdr.setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        od_hdr.setStretchLastSection(True)
        od_vh = self._od_result_table.verticalHeader()
        od_vh.setSectionResizeMode(QHeaderView.ResizeMode.Fixed)
        od_vh.setDefaultSectionSize(_OD_TABLE_DEFAULT_ROW_HEIGHT)
        self._od_result_table.setColumnWidth(0, 300)
        self._od_result_table.setColumnWidth(1, 88)
        self._od_result_table.setColumnWidth(2, 120)
        self._od_result_table.setColumnWidth(3, 280)
        self._od_result_table.setColumnWidth(4, 120)
        self._od_result_table.setColumnWidth(5, 140)
        self._od_result_table.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self._od_result_table.customContextMenuRequested.connect(self._show_od_context_menu)
        layout.addWidget(self._od_result_table, 1)
        self._clear_od_table_tab()
        return tab

    def _clear_od_table_tab(self) -> None:
        """Reset OD tab before a new run or on startup."""
        self._od_pdf_path = None
        self._lbl_od_info.setText("ОД: данных пока нет (запустите pipeline).")
        self._lbl_od_info.setStyleSheet("font-size: 12px; color: #666666;")
        self._od_result_table.clearContents()
        self._od_result_table.setRowCount(0)
        self._od_result_table.setColumnCount(len(_DEFAULT_OD_HEADERS))
        self._od_result_table.setHorizontalHeaderLabels(list(_DEFAULT_OD_HEADERS))

    def _populate_od_table_from_summary(self, od_table: dict[str, Any] | None) -> None:
        """Fill OD tab from ``summary['od_table']``."""
        od_table = od_table or {}
        if not od_table:
            self._clear_od_table_tab()
            return
        status = str(od_table.get("status") or "")
        err = (od_table.get("error") or "").strip()
        path = _normalize_local_fs_path(od_table.get("od_pdf_path") or "")
        self._od_pdf_path = path or None
        columns = list(od_table.get("columns") or [])
        rows_raw = od_table.get("rows") or []
        warnings_raw = od_table.get("warnings") or []
        warnings = [str(w).strip() for w in warnings_raw if str(w).strip()]

        lines: list[str] = []
        if status == "ok":
            lines.append("Разбор ведомости ОД выполнен без ошибок.")
            if path:
                lines.append(f"Файл: {path}")
            lines.append(f"Строк в таблице: {len(rows_raw)}.")
            if warnings:
                lines.append("Замечания:")
                lines.extend(f"• {w}" for w in warnings)
            self._lbl_od_info.setStyleSheet("font-size: 12px; color: #14532d;")
        elif status == "no_rows":
            lines.append(
                "Файл ОД прочитан, строк ведомости не извлечено "
                "(таблица пуста или не распознаны заголовки/строки)."
            )
            if path:
                lines.append(f"Файл: {path}")
            if warnings:
                lines.append("Диагностика:")
                lines.extend(f"• {w}" for w in warnings)
            self._lbl_od_info.setStyleSheet("font-size: 12px; color: #92400e;")
        elif status == "no_od_file":
            lines.append(
                "В папке PDF не найден файл с типом документа OD — таблица ведомости недоступна."
            )
            self._lbl_od_info.setStyleSheet("font-size: 12px; color: #92400e;")
        elif status == "error":
            lines.append("При разборе ведомости ОД произошла ошибка:")
            lines.append(err or "(текст ошибки отсутствует)")
            if path:
                lines.append(f"Файл: {path}")
            if warnings:
                lines.append("Дополнительно:")
                lines.extend(f"• {w}" for w in warnings)
            self._lbl_od_info.setStyleSheet("font-size: 12px; color: #b91c1c;")
        else:
            lines.append(f"Неизвестный статус ОД: {status!r}.")
            self._lbl_od_info.setStyleSheet("font-size: 12px; color: #666666;")
        self._lbl_od_info.setText("\n".join(lines))

        if not columns:
            columns = list(_DEFAULT_OD_HEADERS)
        self._od_result_table.clear()
        self._od_result_table.setColumnCount(len(columns))
        self._od_result_table.setHorizontalHeaderLabels(columns)
        n_rows = len(rows_raw) if isinstance(rows_raw, list) else 0
        self._od_result_table.setRowCount(n_rows)
        for r in range(n_rows):
            row = rows_raw[r]
            if not isinstance(row, (list, tuple)):
                row = []
            for c in range(len(columns)):
                val = row[c] if c < len(row) else ""
                item = QTableWidgetItem(str(val))
                item.setFlags(
                    Qt.ItemFlag.ItemIsSelectable
                    | Qt.ItemFlag.ItemIsEnabled
                    | Qt.ItemFlag.ItemIsEditable
                )
                item.setTextAlignment(
                    int(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignTop)
                )
                self._od_result_table.setItem(r, c, item)
        od_hdr = self._od_result_table.horizontalHeader()
        od_hdr.setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        od_hdr.setStretchLastSection(True)

    def _build_tag_analysis_tab(self) -> QWidget:
        """Structured tag analysis from ``summary['tag_analysis']``."""
        tab = QWidget(self)
        layout = QVBoxLayout(tab)
        layout.setSpacing(8)
        self._lbl_tag_analysis = QLabel("")
        self._lbl_tag_analysis.setWordWrap(True)
        self._lbl_tag_analysis.setTextInteractionFlags(
            Qt.TextInteractionFlag.TextSelectableByMouse
        )
        layout.addWidget(self._lbl_tag_analysis)
        btn_row = QHBoxLayout()
        self._btn_tag_open_txt = QPushButton("Открыть txt отчёт")
        self._btn_tag_open_txt.clicked.connect(self._on_tag_analysis_open_txt)
        self._btn_tag_open_dir = QPushButton("Открыть папку результатов")
        self._btn_tag_open_dir.clicked.connect(self._on_tag_analysis_open_result_dir)
        btn_row.addWidget(self._btn_tag_open_txt)
        btn_row.addWidget(self._btn_tag_open_dir)
        btn_row.addStretch(1)
        layout.addLayout(btn_row)

        self._tag_outer_split = QSplitter(Qt.Orientation.Vertical, tab)
        self._tag_inner_split = QSplitter(Qt.Orientation.Horizontal, tab)

        self._lw_tag_checks = QListWidget()
        self._lw_tag_checks.setMinimumWidth(160)
        self._lw_tag_checks.currentRowChanged.connect(self._on_tag_analysis_check_changed)
        self._tag_inner_split.addWidget(self._lw_tag_checks)

        right_panel = QWidget(tab)
        rp_layout = QVBoxLayout(right_panel)
        rp_layout.setContentsMargins(0, 0, 0, 0)
        rp_layout.setSpacing(6)

        self._tag_filter_bar = QWidget(right_panel)
        fl = QHBoxLayout(self._tag_filter_bar)
        fl.setContentsMargins(0, 0, 0, 0)
        self._tag_filter_bar.setMinimumHeight(_TAG_FILTER_BAR_MIN_H)
        self._lbl_tag_filter_reason = QLabel("Причина:")
        self._combo_tag_extra_reason = QComboBox()
        for label, fid in (
            ("Все", _TAG_FILT_REASON_ALL),
            ("not_found", _TAG_FILT_REASON_NOT_FOUND),
            ("hidden_by_rule", _TAG_FILT_REASON_HIDDEN),
        ):
            self._combo_tag_extra_reason.addItem(label, fid)
        self._combo_tag_extra_reason.currentIndexChanged.connect(self._on_tag_analysis_filter_changed)
        self._lbl_tag_filter_wbs = QLabel("Титул:")
        self._combo_tag_extra_wbs = QComboBox()
        for label, fid in (
            ("Все", _TAG_FILT_TITUL_ALL),
            ("= Титул как в МТО", _TAG_FILT_TITUL_EQ_MTO),
            ("≠ Титул НЕ как в МТО", _TAG_FILT_TITUL_NEQ_MTO),
        ):
            self._combo_tag_extra_wbs.addItem(label, fid)
        self._combo_tag_extra_wbs.currentIndexChanged.connect(self._on_tag_analysis_filter_changed)
        self._lbl_tag_hf = QLabel("Порог кол-ва:")
        self._spin_tag_hf = QSpinBox()
        self._spin_tag_hf.setMinimum(2)
        self._spin_tag_hf.setMaximum(9999)
        self._spin_tag_hf.valueChanged.connect(self._on_tag_analysis_hf_spin_changed)
        fl.addWidget(self._lbl_tag_filter_reason)
        fl.addWidget(self._combo_tag_extra_reason)
        fl.addWidget(self._lbl_tag_filter_wbs)
        fl.addWidget(self._combo_tag_extra_wbs)
        fl.addWidget(self._lbl_tag_hf)
        fl.addWidget(self._spin_tag_hf)
        fl.addStretch(1)
        rp_layout.addWidget(self._tag_filter_bar)

        self._tag_detail_table = QTableWidget(0, 0)
        self._tag_detail_table.setAlternatingRowColors(True)
        self._tag_detail_table.setSelectionBehavior(
            QAbstractItemView.SelectionBehavior.SelectRows
        )
        self._tag_detail_table.setSortingEnabled(True)
        self._tag_detail_table.setEditTriggers(
            QAbstractItemView.EditTrigger.DoubleClicked
            | QAbstractItemView.EditTrigger.EditKeyPressed
        )
        self._tag_detail_table.verticalHeader().setDefaultSectionSize(28)
        self._tag_detail_table.itemChanged.connect(self._on_tag_detail_table_item_changed)
        self._tag_detail_table.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self._tag_detail_table.customContextMenuRequested.connect(
            self._on_tag_analysis_table_context_menu
        )
        hdr = self._tag_detail_table.horizontalHeader()
        hdr.setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        hdr.setStretchLastSection(True)
        rp_layout.addWidget(self._tag_detail_table, 1)
        self._tag_inner_split.addWidget(right_panel)

        self._tag_outer_split.addWidget(self._tag_inner_split)

        inv_lab = QLabel("Некорректные теги (разбор)")
        inv_w = QWidget(tab)
        inv_l = QVBoxLayout(inv_w)
        inv_l.setContentsMargins(0, 0, 0, 0)
        inv_l.addWidget(inv_lab)
        self._tag_invalid_table = QTableWidget(0, 4)
        self._tag_invalid_table.setHorizontalHeaderLabels(["Лист", "Тег", "Ошибка", "PDF"])
        self._tag_invalid_table.setAlternatingRowColors(True)
        self._tag_invalid_table.setEditTriggers(
            QAbstractItemView.EditTrigger.DoubleClicked
            | QAbstractItemView.EditTrigger.EditKeyPressed
        )
        self._tag_invalid_table.verticalHeader().setDefaultSectionSize(28)
        self._tag_invalid_table.itemChanged.connect(self._on_tag_invalid_table_item_changed)
        self._tag_invalid_table.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self._tag_invalid_table.customContextMenuRequested.connect(
            self._on_tag_invalid_table_context_menu
        )
        inv_l.addWidget(self._tag_invalid_table)
        self._tag_outer_split.addWidget(inv_w)
        self._tag_outer_split.setStretchFactor(0, 3)
        self._tag_outer_split.setStretchFactor(1, 1)
        layout.addWidget(self._tag_outer_split, 1)

        self._tag_layout_saving_suspended = False
        self._tag_layout_saver = DebouncedLayoutSaver(
            self,
            snapshot=self._snapshot_tag_analysis_ui_dict,
            on_flush=self._flush_tag_analysis_layout_to_disk,
        )
        self._tag_outer_split.splitterMoved.connect(self._schedule_tag_analysis_layout_save)
        self._tag_inner_split.splitterMoved.connect(self._schedule_tag_analysis_layout_save)
        self._tag_detail_table.horizontalHeader().sectionResized.connect(
            self._on_tag_analysis_detail_section_resized
        )
        self._tag_invalid_table.horizontalHeader().sectionResized.connect(
            self._on_tag_analysis_invalid_section_resized
        )

        self._tag_analysis_payload: dict[str, Any] = {}
        self._tag_analysis_check_index: int = -1
        self._clear_tag_analysis_tab()
        return tab

    def _schedule_tag_analysis_layout_save(self, *_args: object) -> None:
        if getattr(self, "_tag_layout_saving_suspended", False):
            return
        if hasattr(self, "_tag_layout_saver"):
            self._tag_layout_saver.schedule()

    def _on_tag_analysis_detail_section_resized(self, *_args: object) -> None:
        self._schedule_tag_analysis_layout_save()

    def _on_tag_analysis_invalid_section_resized(self, *_args: object) -> None:
        self._schedule_tag_analysis_layout_save()

    def _snapshot_tag_analysis_ui_dict(self) -> dict[str, Any]:
        ta_prev: dict[str, Any] = {}
        mu = self._cfg.get("monitor_ui")
        if isinstance(mu, dict):
            raw_ta = mu.get("tag_analysis")
            if isinstance(raw_ta, dict):
                ta_prev = raw_ta
        base = normalize_tag_analysis_ui_config(ta_prev)
        dcw = dict(base.get("detail_column_widths") or {})
        ch = self._current_tag_analysis_check()
        if isinstance(ch, dict) and self._tag_detail_table.columnCount():
            cid = str(ch.get("id") or "").strip() or "_default"
            dcw[cid] = [
                self._tag_detail_table.columnWidth(i)
                for i in range(self._tag_detail_table.columnCount())
            ]
        base["detail_column_widths"] = dcw
        try:
            vs = self._tag_outer_split.sizes()
            if len(vs) >= 2:
                base["split_vertical"] = [max(40, int(vs[0])), max(80, int(vs[1]))]
        except (TypeError, ValueError, AttributeError):
            pass
        try:
            hs = self._tag_inner_split.sizes()
            if len(hs) >= 2:
                base["split_horizontal"] = [max(120, int(hs[0])), max(200, int(hs[1]))]
        except (TypeError, ValueError, AttributeError):
            pass
        if self._tag_invalid_table.columnCount() >= 4:
            base["invalid_column_widths"] = [
                self._tag_invalid_table.columnWidth(i) for i in range(4)
            ]
        return base

    def _flush_tag_analysis_layout_to_disk(self, snap: dict[str, Any]) -> None:
        cfg = load_v2_config()
        merged_mu = dict(cfg.get("monitor_ui") or get_default_monitor_ui_config())
        merged_mu["tag_analysis"] = normalize_tag_analysis_ui_config(snap)
        cfg["monitor_ui"] = merged_mu
        cfg = normalize_v2_config(cfg)
        save_v2_config(cfg)
        self._cfg["monitor_ui"] = cfg["monitor_ui"]

    def _apply_tag_analysis_layout_from_cfg(self, cfg: dict[str, Any] | None) -> None:
        if not hasattr(self, "_tag_outer_split"):
            return
        mu = cfg.get("monitor_ui") if isinstance(cfg, dict) else None
        mu = mu if isinstance(mu, dict) else {}
        ta = normalize_tag_analysis_ui_config(mu.get("tag_analysis"))
        self._tag_layout_saving_suspended = True
        try:
            sv = ta.get("split_vertical")
            if isinstance(sv, list) and len(sv) >= 2:
                self._tag_outer_split.setSizes([int(sv[0]), int(sv[1])])
            sh = ta.get("split_horizontal")
            if isinstance(sh, list) and len(sh) >= 2:
                self._tag_inner_split.setSizes([int(sh[0]), int(sh[1])])
            inv_w = ta.get("invalid_column_widths")
            if isinstance(inv_w, list) and len(inv_w) >= 4:
                for i in range(4):
                    try:
                        self._tag_invalid_table.setColumnWidth(i, int(inv_w[i]))
                    except (TypeError, ValueError):
                        continue
        finally:
            self._tag_layout_saving_suspended = False
        if hasattr(self, "_tag_layout_saver"):
            self._tag_layout_saver.sync_last_written_from_snapshot(ta)

    def _restore_tag_analysis_detail_column_widths(self, check_id: str) -> None:
        mu = self._cfg.get("monitor_ui")
        if not isinstance(mu, dict):
            return
        ta = mu.get("tag_analysis")
        if not isinstance(ta, dict):
            return
        dcw = ta.get("detail_column_widths")
        if not isinstance(dcw, dict):
            return
        cid = check_id.strip() or "_default"
        wlist = dcw.get(cid)
        if not isinstance(wlist, list):
            return
        self._tag_layout_saving_suspended = True
        try:
            for i, w in enumerate(wlist):
                if i < self._tag_detail_table.columnCount():
                    try:
                        self._tag_detail_table.setColumnWidth(i, int(w))
                    except (TypeError, ValueError):
                        continue
        finally:
            self._tag_layout_saving_suspended = False

    def _clear_tag_analysis_tab(self) -> None:
        self._tag_analysis_payload = {}
        self._tag_analysis_check_index = -1
        if hasattr(self, "_lbl_tag_analysis"):
            self._lbl_tag_analysis.setTextFormat(Qt.TextFormat.PlainText)
            self._lbl_tag_analysis.setText("Анализ тегов: данных пока нет (запустите pipeline).")
            self._lbl_tag_analysis.setStyleSheet("font-size: 12px; color: #666666;")
        if hasattr(self, "_lw_tag_checks"):
            self._lw_tag_checks.blockSignals(True)
            self._lw_tag_checks.clear()
            self._lw_tag_checks.blockSignals(False)
        if hasattr(self, "_tag_detail_table"):
            self._tag_detail_table.clear()
            self._tag_detail_table.setRowCount(0)
            self._tag_detail_table.setColumnCount(0)
        if hasattr(self, "_tag_invalid_table"):
            self._tag_invalid_table.setRowCount(0)
        if hasattr(self, "_lbl_tag_filter_reason"):
            self._lbl_tag_filter_reason.setVisible(False)
            self._combo_tag_extra_reason.setVisible(False)
            self._lbl_tag_filter_wbs.setVisible(False)
            self._combo_tag_extra_wbs.setVisible(False)
            self._lbl_tag_hf.setVisible(False)
            self._spin_tag_hf.setVisible(False)
        if hasattr(self, "_btn_tag_open_txt"):
            self._btn_tag_open_txt.setEnabled(False)
            self._btn_tag_open_dir.setEnabled(False)

    def _html_tag_analysis_summary(self, payload: dict[str, Any]) -> str:
        """Rich text: Run-like gray (#666); status value red when not ok."""
        status = str(payload.get("status") or "")
        err = str(payload.get("error") or "").strip()
        path = _normalize_local_fs_path(payload.get("text_report_path") or "")
        mto_wbs = str(payload.get("mto_wbs") or "").strip()
        esc = html.escape
        ok_color = "#666666"
        bad_red = "#dc3232"
        parts: list[str] = []
        if status == "ok":
            parts.append(f'<span style="color:{ok_color};">Статус: {esc(status)}</span>')
        else:
            parts.append(f'<span style="color:{ok_color};">Статус: </span>')
            parts.append(f'<span style="color:{bad_red};">{esc(status)}</span>')
        if mto_wbs:
            parts.append(f'<span style="color:{ok_color};">  |  Титул МТО: {esc(mto_wbs)}</span>')
        if path:
            parts.append(f'<span style="color:{ok_color};">  |  Отчёт: {esc(path)}</span>')
        if err:
            parts.append(f'<span style="color:{ok_color};">  |  Ошибка: {esc(err)}</span>')
        for w in payload.get("warnings") or []:
            ws = str(w).strip()
            if ws:
                parts.append(f'<span style="color:{ok_color};">  |  {esc(ws)}</span>')
        return "".join(parts)

    def _populate_tag_analysis_from_summary(self, payload: dict[str, Any] | None) -> None:
        payload = payload or {}
        self._tag_analysis_payload = dict(payload)
        if not self._tag_analysis_payload:
            self._clear_tag_analysis_tab()
            return
        self._lbl_tag_analysis.setTextFormat(Qt.TextFormat.RichText)
        self._lbl_tag_analysis.setText(self._html_tag_analysis_summary(self._tag_analysis_payload))
        self._lbl_tag_analysis.setStyleSheet("font-size: 12px;")
        rep_path = _normalize_local_fs_path(self._tag_analysis_payload.get("text_report_path") or "")
        self._btn_tag_open_txt.setEnabled(bool(rep_path and os.path.isfile(rep_path)))
        self._btn_tag_open_dir.setEnabled(bool(self._result_dir and os.path.isdir(self._result_dir)))

        self._lw_tag_checks.blockSignals(True)
        self._lw_tag_checks.clear()
        checks = self._tag_analysis_payload.get("checks") or []
        if isinstance(checks, list):
            for idx, ch in enumerate(checks):
                if not isinstance(ch, dict):
                    continue
                title = str(ch.get("title") or ch.get("id") or "?")
                sev = str(ch.get("severity") or "info")
                it = QListWidgetItem(title)
                it.setData(int(Qt.ItemDataRole.UserRole), idx)
                if sev == "info":
                    it.setBackground(QBrush(_CLR_TAG_CHECK_ROW_OK))
                self._lw_tag_checks.addItem(it)
        self._lw_tag_checks.blockSignals(False)
        if self._lw_tag_checks.count():
            self._lw_tag_checks.setCurrentRow(0)
        else:
            self._tag_analysis_check_index = -1
            self._tag_detail_table.setRowCount(0)

        inv = self._tag_analysis_payload.get("invalid_tags") or []
        self._tag_invalid_table.setRowCount(0)
        if isinstance(inv, list) and inv:
            inv_keys = ("sheet", "tag", "error", "pdf_path")
            self._tag_invalid_table.blockSignals(True)
            self._tag_invalid_table.setRowCount(len(inv))
            for r, row in enumerate(inv):
                if not isinstance(row, dict):
                    continue
                vals = [
                    str(row.get("sheet", "")),
                    str(row.get("tag", "")),
                    str(row.get("error", "")),
                    str(row.get("pdf_path", "")),
                ]
                for c, val in enumerate(vals):
                    cell = QTableWidgetItem(val)
                    cell.setFlags(
                        Qt.ItemFlag.ItemIsSelectable
                        | Qt.ItemFlag.ItemIsEnabled
                        | Qt.ItemFlag.ItemIsEditable
                    )
                    cell.setData(_TAG_ANALYSIS_ROW_DICT_ROLE, dict(row))
                    cell.setData(_TAG_ANALYSIS_COLKEY_ROLE, inv_keys[c])
                    self._tag_invalid_table.setItem(r, c, cell)
            self._tag_invalid_table.blockSignals(False)
            self._apply_tag_analysis_invalid_widths_from_cfg()

    def _apply_tag_analysis_invalid_widths_from_cfg(self) -> None:
        mu = self._cfg.get("monitor_ui")
        if not isinstance(mu, dict):
            return
        ta = mu.get("tag_analysis")
        if not isinstance(ta, dict):
            return
        inv_w = ta.get("invalid_column_widths")
        if not isinstance(inv_w, list) or len(inv_w) < 4:
            return
        self._tag_layout_saving_suspended = True
        try:
            for i in range(4):
                try:
                    self._tag_invalid_table.setColumnWidth(i, int(inv_w[i]))
                except (TypeError, ValueError):
                    continue
        finally:
            self._tag_layout_saving_suspended = False

    def _current_tag_analysis_check(self) -> dict[str, Any] | None:
        checks = self._tag_analysis_payload.get("checks")
        if not isinstance(checks, list):
            return None
        i = self._tag_analysis_check_index
        if i < 0 or i >= len(checks):
            return None
        ch = checks[i]
        return ch if isinstance(ch, dict) else None

    def _on_tag_analysis_check_changed(self, row: int) -> None:
        if row < 0:
            self._tag_analysis_check_index = -1
            self._lbl_tag_filter_reason.setVisible(False)
            self._combo_tag_extra_reason.setVisible(False)
            self._lbl_tag_filter_wbs.setVisible(False)
            self._combo_tag_extra_wbs.setVisible(False)
            self._lbl_tag_hf.setVisible(False)
            self._spin_tag_hf.setVisible(False)
            self._tag_detail_table.setRowCount(0)
            return
        self._tag_analysis_check_index = row
        ch = self._current_tag_analysis_check()
        if not ch:
            self._lbl_tag_filter_reason.setVisible(False)
            self._combo_tag_extra_reason.setVisible(False)
            self._lbl_tag_filter_wbs.setVisible(False)
            self._combo_tag_extra_wbs.setVisible(False)
            self._lbl_tag_hf.setVisible(False)
            self._spin_tag_hf.setVisible(False)
            self._tag_detail_table.setRowCount(0)
            return
        cid = str(ch.get("id") or "")
        show_filters = cid in ("extra_in_sheets", "high_frequency")
        self._lbl_tag_filter_reason.setVisible(show_filters and cid == "extra_in_sheets")
        self._combo_tag_extra_reason.setVisible(show_filters and cid == "extra_in_sheets")
        self._lbl_tag_filter_wbs.setVisible(show_filters and cid == "extra_in_sheets")
        self._combo_tag_extra_wbs.setVisible(show_filters and cid == "extra_in_sheets")
        self._lbl_tag_hf.setVisible(show_filters and cid == "high_frequency")
        self._spin_tag_hf.setVisible(show_filters and cid == "high_frequency")
        if cid == "high_frequency":
            meta = ch.get("meta") if isinstance(ch.get("meta"), dict) else {}
            thr = int(meta.get("threshold") or self._tag_analysis_payload.get("high_frequency_threshold") or 3)
            self._spin_tag_hf.blockSignals(True)
            self._spin_tag_hf.setValue(max(2, thr))
            self._spin_tag_hf.blockSignals(False)
        self._tag_analysis_apply_detail_filter()

    def _tag_analysis_filtered_rows_for_ui(self, ch: dict[str, Any]) -> list[dict[str, Any]]:
        rows = ch.get("rows")
        if not isinstance(rows, list):
            return []
        out = [r for r in rows if isinstance(r, dict)]
        cid = str(ch.get("id") or "")
        if cid == "extra_in_sheets":
            reason_mode = self._combo_tag_extra_reason.currentData()
            if reason_mode not in (None, "", _TAG_FILT_REASON_ALL):
                out = [r for r in out if str(r.get("reason", "")) == str(reason_mode)]
            titul_mode = self._combo_tag_extra_wbs.currentData()
            mto_wbs = str(self._tag_analysis_payload.get("mto_wbs") or "").strip()
            if titul_mode == _TAG_FILT_TITUL_EQ_MTO and mto_wbs:
                out = [r for r in out if str(r.get("wbs", "")).strip() == mto_wbs]
            elif titul_mode == _TAG_FILT_TITUL_NEQ_MTO and mto_wbs:
                out = [r for r in out if str(r.get("wbs", "")).strip() != mto_wbs]
        elif cid == "high_frequency":
            thr = int(self._spin_tag_hf.value())
            out = [r for r in out if int(r.get("count") or 0) >= thr]
        return out

    def _tag_analysis_apply_detail_filter(self) -> None:
        ch = self._current_tag_analysis_check()
        if not ch:
            return
        columns = ch.get("columns")
        keys = ch.get("column_keys")
        if not isinstance(columns, list):
            columns = []
        if not isinstance(keys, list) or len(keys) != len(columns):
            keys = [f"c{i}" for i in range(len(columns))]
        rows_vis = self._tag_analysis_filtered_rows_for_ui(ch)
        cid_key = str(ch.get("id") or "").strip() or "_default"
        self._tag_detail_table.blockSignals(True)
        self._tag_detail_table.setSortingEnabled(False)
        self._tag_detail_table.clear()
        ncols = len(columns)
        if len(keys) != ncols:
            keys = [f"c{i}" for i in range(ncols)]
        self._tag_detail_table.setColumnCount(ncols)
        self._tag_detail_table.setHorizontalHeaderLabels([str(c) for c in columns])
        self._tag_detail_table.setRowCount(len(rows_vis))
        for r, rowd in enumerate(rows_vis):
            for c, k in enumerate(keys):
                val = rowd.get(k, "")
                cell = QTableWidgetItem("" if val is None else str(val))
                cell.setFlags(
                    Qt.ItemFlag.ItemIsSelectable
                    | Qt.ItemFlag.ItemIsEnabled
                    | Qt.ItemFlag.ItemIsEditable
                )
                cell.setData(_TAG_ANALYSIS_ROW_DICT_ROLE, dict(rowd))
                cell.setData(_TAG_ANALYSIS_COLKEY_ROLE, k)
                self._tag_detail_table.setItem(r, c, cell)
        self._restore_tag_analysis_detail_column_widths(cid_key)
        self._tag_detail_table.setSortingEnabled(True)
        try:
            self._tag_detail_table.sortByColumn(-1, Qt.SortOrder.AscendingOrder)
        except TypeError:
            pass
        self._tag_detail_table.blockSignals(False)

    def _on_tag_analysis_filter_changed(self, *_args: object) -> None:
        self._tag_analysis_apply_detail_filter()

    def _on_tag_analysis_hf_spin_changed(self, *_args: object) -> None:
        self._tag_analysis_apply_detail_filter()

    def _on_tag_analysis_open_txt(self) -> None:
        p = _normalize_local_fs_path(self._tag_analysis_payload.get("text_report_path") or "")
        self._open_path(p)

    def _on_tag_analysis_open_result_dir(self) -> None:
        self._open_containing_folder(self._result_dir)

    def _tag_detail_row_dict(self, table: QTableWidget, table_row: int) -> dict[str, Any] | None:
        for c in range(table.columnCount()):
            it = table.item(table_row, c)
            if it is None:
                continue
            d = it.data(_TAG_ANALYSIS_ROW_DICT_ROLE)
            if isinstance(d, dict):
                return d
        return None

    def _on_tag_detail_table_item_changed(self, item: QTableWidgetItem) -> None:
        if item.tableWidget() is not self._tag_detail_table:
            return
        d = item.data(_TAG_ANALYSIS_ROW_DICT_ROLE)
        k = item.data(_TAG_ANALYSIS_COLKEY_ROLE)
        if not isinstance(d, dict) or not isinstance(k, str):
            return
        expected = "" if d.get(k) is None else str(d.get(k))
        if item.text() == expected:
            return
        self._tag_detail_table.blockSignals(True)
        item.setText(expected)
        self._tag_detail_table.blockSignals(False)

    def _on_tag_invalid_table_item_changed(self, item: QTableWidgetItem) -> None:
        if item.tableWidget() is not self._tag_invalid_table:
            return
        d = item.data(_TAG_ANALYSIS_ROW_DICT_ROLE)
        k = item.data(_TAG_ANALYSIS_COLKEY_ROLE)
        if not isinstance(d, dict) or not isinstance(k, str):
            return
        expected = "" if d.get(k) is None else str(d.get(k))
        if item.text() == expected:
            return
        self._tag_invalid_table.blockSignals(True)
        item.setText(expected)
        self._tag_invalid_table.blockSignals(False)

    def _on_tag_analysis_table_context_menu(self, pos) -> None:
        idx = self._tag_detail_table.indexAt(pos)
        if not idx.isValid():
            return
        row = idx.row()
        data_row = self._tag_detail_row_dict(self._tag_detail_table, row)
        if data_row is None:
            return
        pdf_path = str(data_row.get("pdf_path", "") or "").strip()
        page_num = data_row.get("page_num")
        tag = str(data_row.get("tag", "") or "").strip()
        menu = QMenu(self)
        act_open_pdf = menu.addAction("Открыть PDF (на страницу, если возможно)")
        act_open_plain = menu.addAction("Открыть PDF")
        act_open_editor = menu.addAction("Открыть в редакторе шаблонов")
        act_open_folder = menu.addAction("Открыть папку")
        act_props = menu.addAction("Свойства тега…")
        act_open_pdf.setEnabled(bool(pdf_path))
        act_open_plain.setEnabled(bool(pdf_path))
        act_open_editor.setEnabled(bool(pdf_path))
        act_open_folder.setEnabled(bool(pdf_path))
        act_props.setEnabled(bool(tag))
        chosen = menu.exec(self._tag_detail_table.viewport().mapToGlobal(pos))
        if chosen == act_open_pdf:
            self._open_pdf_best_effort(pdf_path, page_num)
        elif chosen == act_open_plain:
            self._open_path(pdf_path)
        elif chosen == act_open_editor:
            self._launch_template_editor(pdf_path)
        elif chosen == act_open_folder:
            self._open_containing_folder(pdf_path)
        elif chosen == act_props:
            self._tag_analysis_show_properties_for_table_row(self._tag_detail_table, row)

    def _on_tag_invalid_table_context_menu(self, pos) -> None:
        idx = self._tag_invalid_table.indexAt(pos)
        if not idx.isValid():
            return
        row = idx.row()
        data_row = self._tag_detail_row_dict(self._tag_invalid_table, row)
        if data_row is None:
            return
        pdf_path = str(data_row.get("pdf_path", "") or "").strip()
        page_num = data_row.get("page_num")
        menu = QMenu(self)
        act_open_pdf = menu.addAction("Открыть PDF (на страницу, если возможно)")
        act_open_plain = menu.addAction("Открыть PDF")
        act_open_editor = menu.addAction("Открыть в редакторе шаблонов")
        act_open_folder = menu.addAction("Открыть папку")
        act_props = menu.addAction("Свойства тега…")
        tag = str(data_row.get("tag", "") or "").strip()
        act_open_pdf.setEnabled(bool(pdf_path))
        act_open_plain.setEnabled(bool(pdf_path))
        act_open_editor.setEnabled(bool(pdf_path))
        act_open_folder.setEnabled(bool(pdf_path))
        act_props.setEnabled(bool(tag))
        chosen = menu.exec(self._tag_invalid_table.viewport().mapToGlobal(pos))
        if chosen == act_open_pdf:
            self._open_pdf_best_effort(pdf_path, page_num)
        elif chosen == act_open_plain:
            self._open_path(pdf_path)
        elif chosen == act_open_editor:
            self._launch_template_editor(pdf_path)
        elif chosen == act_open_folder:
            self._open_containing_folder(pdf_path)
        elif chosen == act_props:
            self._tag_analysis_show_properties_for_table_row(self._tag_invalid_table, row)

    def _tag_analysis_show_properties_for_table_row(self, table: QTableWidget, row: int) -> None:
        d = self._tag_detail_row_dict(table, row)
        if not isinstance(d, dict):
            return
        tag = str(d.get("tag") or "").strip()
        if not tag:
            QMessageBox.information(self, "Тег", "В строке нет поля «тег».")
            return
        from tags.tag_classes import TagClass

        tc = TagClass(tag, strict=False)
        lines = [
            f"tag: {tc.tag}",
            f"is_valid: {tc.is_valid}",
            f"parse_error: {tc.parse_error or '—'}",
            f"wbs: {tc.wbs}",
            f"system_code: {tc.system_code}",
            f"equipment: {tc.equipment}",
            f"sequence_number: {tc.sequence_number}",
            f"additional_code: {tc.additional_code}",
            f"floor: {tc.floor}",
            f"sum_system_number: {tc.sum_system_number}",
            f"sub_system: {tc.sub_system}",
            f"out_row: {tc.out_row}",
            f"category_name: {tc.category_name}",
            f"is_hidden: {tc.is_hidden()}",
        ]
        QMessageBox.information(self, "Свойства тега", "\n".join(lines))

    def _setup_timer(self) -> None:
        self._timer = QTimer(self)
        self._timer.setInterval(500)
        self._timer.timeout.connect(self._tick_elapsed)

    def _select_initial_tab(self) -> None:
        index_by_name = {
            "run": _TAB_INDEX_RUN,
            "monitor": _TAB_INDEX_RUN,
            "settings": _TAB_INDEX_RUN,
            "normcontrol": _TAB_INDEX_NORMCONTROL,
            "nk": _TAB_INDEX_NORMCONTROL,
            "od": _TAB_INDEX_OD,
            "tags": _TAB_INDEX_TAG_ANALYSIS,
            "tag_analysis": _TAB_INDEX_TAG_ANALYSIS,
            "reports": _TAB_INDEX_REPORTS,
        }
        self._tabs.setCurrentIndex(index_by_name.get(self._initial_tab, 0))

    def _open_settings_dialog(self) -> None:
        if self._settings_dialog is None:
            self._settings_dialog = self._build_settings_dialog()
        self._settings_dialog.show()
        self._settings_dialog.raise_()
        self._settings_dialog.activateWindow()

    def set_pdf_folder(self, folder: str) -> None:
        self._set_pdf_path_text(_normalize_local_fs_path(folder))
        self._refresh_start_availability()

    def _current_pdf_path(self) -> str:
        return (self._combo_pdf_path.currentText() or "").strip()

    def _normcontrol_file_dialog_start_dir(self) -> str:
        """Prefer Run-tab PDF folder, then result dir, then folder of current НК xlsx."""
        pdf = _normalize_local_fs_path(self._current_pdf_path())
        if pdf and os.path.isdir(pdf):
            return pdf
        rd = _normalize_local_fs_path(self._result_dir or "")
        if rd and os.path.isdir(rd):
            return rd
        xlsx = _normalize_local_fs_path(self._normcontrol_xlsx_path or "")
        if xlsx and os.path.isfile(xlsx):
            d = os.path.dirname(xlsx)
            if d and os.path.isdir(d):
                return d
        return ""

    def _set_pdf_path_text(self, folder: str) -> None:
        text = str(folder or "").strip()
        self._combo_pdf_path.setEditText(text)

    def _on_pdf_path_changed(self, _text: str) -> None:
        self._refresh_start_availability()
        if self._pipeline_started:
            return
        self._preview_pdf_folder(self._current_pdf_path())

    def _load_recent_pdf_paths(self) -> None:
        raw = self._settings_store.value(_RECENT_PDF_PATHS_KEY, [])
        if isinstance(raw, str):
            items = [raw] if raw else []
        else:
            items = [str(x).strip() for x in (raw or []) if str(x).strip()]
        unique_items: list[str] = []
        for item in items:
            if item not in unique_items:
                unique_items.append(item)
        self._combo_pdf_path.clear()
        self._combo_pdf_path.addItems(unique_items[:_RECENT_PDF_PATHS_LIMIT])

    def _remember_pdf_path(self, folder: str) -> None:
        folder = str(folder or "").strip()
        if not folder:
            return
        raw = self._settings_store.value(_RECENT_PDF_PATHS_KEY, [])
        if isinstance(raw, str):
            items = [raw] if raw else []
        else:
            items = [str(x).strip() for x in (raw or []) if str(x).strip()]
        updated = [folder]
        for item in items:
            if item and item != folder and item not in updated:
                updated.append(item)
        updated = updated[:_RECENT_PDF_PATHS_LIMIT]
        self._settings_store.setValue(_RECENT_PDF_PATHS_KEY, updated)
        self._load_recent_pdf_paths()
        self._set_pdf_path_text(folder)

    def _preview_pdf_folder(self, folder: str) -> None:
        folder = _normalize_local_fs_path(folder)
        if folder == self._previewed_pdf_path:
            return
        self._previewed_pdf_path = folder
        self._monitor_rows.clear()
        self._file_path_by_row.clear()
        self._total_files = 0
        self._total_tasks = 0
        self._completed = 0
        self._progress.setMaximum(1)
        self._progress.setValue(0)
        self._progress.setFormat("0 / 0 задач")
        self._table.setRowCount(0)
        self._stage_table.setRowCount(0)
        if not folder or not os.path.isdir(folder):
            self._lbl_summary.setText("Укажите существующую папку с PDF.")
            return

        pdf_paths = self._list_pdf_files(folder)
        if not pdf_paths:
            self._lbl_summary.setText("В выбранной папке PDF не найдены.")
            return

        self._ensure_monitor_rows(len(pdf_paths))
        for row, file_path in enumerate(pdf_paths):
            self._file_path_by_row[row] = file_path
            state = self._monitor_rows[row]
            state["file_name"] = os.path.basename(file_path)
            self._render_file_row(row)
        self._lbl_summary.setText(f"Найдено PDF для запуска: {len(pdf_paths)}")

    def _list_pdf_files(self, folder: str) -> list[str]:
        try:
            entries = [
                entry.path
                for entry in os.scandir(folder)
                if entry.is_file() and entry.name.lower().endswith(".pdf")
            ]
        except OSError:
            return []
        return sorted(entries, key=lambda path: os.path.basename(path).lower())

    def _on_project_changed(self, _index: int) -> None:
        self._update_effective_cfg_label()
        if not hasattr(self, "_btn_save_settings"):
            return
        try:
            cfg = self._collect_cfg_from_widgets()
            save_v2_config(cfg)
        except Exception:
            pass

    def _load_cfg_into_widgets(self, cfg: dict[str, Any]) -> None:
        self._cfg = normalize_v2_config(cfg)
        self._chk_enabled.setChecked(bool(self._cfg.get("enabled", True)))
        self._edit_templates_dir.setText(str(self._cfg.get("templates_dir", "")).strip())
        self._chk_parallel.setChecked(bool(self._cfg.get("parallel", True)))
        self._spin_max_workers.setValue(int(self._cfg.get("max_workers", 0) or 0))
        per_page_cfg = dict(self._cfg.get("per_page") or {})
        self._chk_per_page_enabled.setChecked(bool(per_page_cfg.get("enabled", True)))
        self._spin_per_page_min_pages.setValue(int(per_page_cfg.get("min_pages", 4) or 4))
        self._spin_per_page_workers.setValue(int(per_page_cfg.get("max_workers", 0) or 0))
        self._spin_per_page_chunk_size.setValue(int(per_page_cfg.get("chunk_size_pages", 0) or 0))
        doc_type_cfg = dict(per_page_cfg.get("doc_types") or {})
        for doc_type, checkbox in self._chk_per_page_doc_types.items():
            checkbox.setChecked(bool(doc_type_cfg.get(doc_type, False)))
        self._spin_tags_max_workers.setValue(int(self._cfg.get("tags_max_workers", 0) or 0))
        tags_backend = str(self._cfg.get("tags_text_backend", "fitz") or "fitz")
        idx_backend = self._combo_tags_backend.findData(tags_backend)
        self._combo_tags_backend.setCurrentIndex(max(idx_backend, 0))
        self._spin_frame_left_mm.setValue(float(self._cfg.get("find_frame_border_band_left_mm", 20.0)))
        self._spin_frame_right_mm.setValue(float(self._cfg.get("find_frame_border_band_right_mm", 50.0)))
        self._spin_frame_top_mm.setValue(float(self._cfg.get("find_frame_border_band_top_mm", 20.0)))
        self._spin_frame_bottom_mm.setValue(float(self._cfg.get("find_frame_border_band_bottom_mm", 20.0)))
        self._spin_snap_x.setValue(float(self._cfg.get("find_tables_snap_x_tolerance", 2.2)))
        self._spin_snap_y.setValue(float(self._cfg.get("find_tables_snap_y_tolerance", 2.0)))
        self._chk_grid_pf_enabled.setChecked(
            bool(self._cfg.get("grid_detected_prefilter_enabled", True))
        )
        self._spin_grid_pf_adjacency_floor_pt.setValue(
            float(self._cfg.get("grid_detected_prefilter_adjacency_tol_floor_pt", 2.5))
        )
        self._chk_export_debug_excel.setChecked(bool(self._cfg.get("export_debug_excel", True)))
        self._chk_debug_visual.setChecked(bool(self._cfg.get("debug_visual", False)))
        self._chk_debug_visual_overlay.setChecked(bool(self._cfg.get("debug_visual_overlay", False)))
        self._chk_debug_verbose_log.setChecked(bool(self._cfg.get("debug_verbose_log", False)))
        self._chk_timing_log.setChecked(bool(self._cfg.get("timing_log", False)))
        excel_cfg = dict(self._cfg.get("excel") or {})
        self._chk_excel_sanitize.setChecked(bool(excel_cfg.get("sanitize_illegal_chars", True)))
        self._edit_debug_visual_dir.setText(str(self._cfg.get("debug_visual_dir", "")).strip())
        mode = str(self._cfg.get("text_extraction_mode", "char_center"))
        idx = self._combo_extraction_mode.findData(mode)
        self._combo_extraction_mode.setCurrentIndex(max(idx, 0))
        mw, mh = _monitor_size_from_cfg(self._cfg)
        self._spin_monitor_width.setValue(mw)
        self._spin_monitor_height.setValue(mh)
        self._spin_grid_ok.setValue(int(self._cfg.get("monitor_grid_mismatch_ok_max", 3000)))
        self._spin_grid_warn.setValue(int(self._cfg.get("monitor_grid_mismatch_warn_max", 45000)))
        self._spin_weight_global.setValue(float(self._cfg.get("monitor_grid_mismatch_weight_global", 100_000.0)))
        self._spin_weight_walk.setValue(float(self._cfg.get("monitor_grid_mismatch_weight_walk_pt", 50.0)))
        self._spin_weight_no_match.setValue(float(self._cfg.get("monitor_grid_mismatch_weight_no_match", 25_000.0)))
        theme = normalize_monitor_ui_theme(self._cfg.get("monitor_ui_theme"))
        idx_theme = self._combo_monitor_theme.findData(theme)
        self._combo_monitor_theme.blockSignals(True)
        self._combo_monitor_theme.setCurrentIndex(max(idx_theme, 0))
        self._combo_monitor_theme.blockSignals(False)
        self._update_effective_cfg_label()
        self._apply_monitor_window_size_from_cfg()
        self._apply_tag_analysis_layout_from_cfg(self._cfg)
        apply_monitor_ui_theme_full(theme, window=self)

    def _apply_monitor_window_size_from_cfg(self) -> None:
        w, h = _monitor_size_from_cfg(self._cfg)
        self.resize(w, h)

    def _grid_detected_prefilter_cfg_slice(self) -> dict[str, Any]:
        """Flat cfg keys for `grid_detected_prefilter` (GUI edits + preserved JSON-only keys)."""
        base_def = get_default_v2_config()
        prev = getattr(self, "_cfg", {}) or {}
        out: dict[str, Any] = {}
        for key in GRID_DETECTED_PREFILTER_KEYS:
            if key == "grid_detected_prefilter_enabled":
                out[key] = self._chk_grid_pf_enabled.isChecked()
            elif key == "grid_detected_prefilter_adjacency_tol_floor_pt":
                out[key] = float(self._spin_grid_pf_adjacency_floor_pt.value())
            elif key in (
                "grid_detected_prefilter_min_cells",
                "grid_detected_prefilter_min_component_cells",
            ):
                raw = prev.get(key, base_def[key])
                try:
                    out[key] = max(1, int(raw))
                except (TypeError, ValueError):
                    out[key] = int(base_def[key])
            else:
                raw = prev.get(key, base_def[key])
                try:
                    out[key] = float(raw)
                except (TypeError, ValueError):
                    out[key] = float(base_def[key])
        return out

    def _capture_monitor_window_size_to_widgets(self) -> None:
        self._spin_monitor_width.setValue(self.width())
        self._spin_monitor_height.setValue(self.height())

    def _collect_cfg_from_widgets(self) -> dict[str, Any]:
        project_filter = ""
        if self._combo_project.currentIndex() >= 0:
            project_filter = str(self._combo_project.currentData() or "")
        mu_src = self._cfg.get("monitor_ui")
        monitor_ui = (
            copy.deepcopy(mu_src)
            if isinstance(mu_src, dict)
            else get_default_monitor_ui_config()
        )
        cfg = normalize_v2_config(
            {
                "enabled": self._chk_enabled.isChecked(),
                "templates_dir": self._edit_templates_dir.text().strip(),
                "project": project_filter,
                "parallel": self._chk_parallel.isChecked(),
                "max_workers": int(self._spin_max_workers.value()),
                "per_page": {
                    "enabled": self._chk_per_page_enabled.isChecked(),
                    "min_pages": int(self._spin_per_page_min_pages.value()),
                    "max_workers": int(self._spin_per_page_workers.value()),
                    "chunk_size_pages": int(self._spin_per_page_chunk_size.value()),
                    "doc_types": {
                        doc_type: checkbox.isChecked()
                        for doc_type, checkbox in self._chk_per_page_doc_types.items()
                    },
                },
                "tags_max_workers": int(self._spin_tags_max_workers.value()),
                "tags_text_backend": self._combo_tags_backend.currentData(),
                "find_frame_border_band_left_mm": float(self._spin_frame_left_mm.value()),
                "find_frame_border_band_right_mm": float(self._spin_frame_right_mm.value()),
                "find_frame_border_band_top_mm": float(self._spin_frame_top_mm.value()),
                "find_frame_border_band_bottom_mm": float(self._spin_frame_bottom_mm.value()),
                "find_tables_snap_x_tolerance": float(self._spin_snap_x.value()),
                "find_tables_snap_y_tolerance": float(self._spin_snap_y.value()),
                **self._grid_detected_prefilter_cfg_slice(),
                "export_debug_excel": self._chk_export_debug_excel.isChecked(),
                "debug_visual": self._chk_debug_visual.isChecked(),
                "debug_visual_overlay": self._chk_debug_visual_overlay.isChecked(),
                "debug_verbose_log": self._chk_debug_verbose_log.isChecked(),
                "timing_log": self._chk_timing_log.isChecked(),
                "excel": {
                    "sanitize_illegal_chars": self._chk_excel_sanitize.isChecked(),
                },
                "debug_visual_dir": self._edit_debug_visual_dir.text().strip(),
                "text_extraction_mode": self._combo_extraction_mode.currentData(),
                "monitor_window_width": int(self._spin_monitor_width.value()),
                "monitor_window_height": int(self._spin_monitor_height.value()),
                "monitor_grid_mismatch_ok_max": int(self._spin_grid_ok.value()),
                "monitor_grid_mismatch_warn_max": max(
                    int(self._spin_grid_ok.value()),
                    int(self._spin_grid_warn.value()),
                ),
                "monitor_grid_mismatch_weight_global": float(self._spin_weight_global.value()),
                "monitor_grid_mismatch_weight_walk_pt": float(self._spin_weight_walk.value()),
                "monitor_grid_mismatch_weight_no_match": float(self._spin_weight_no_match.value()),
                "monitor_ui_theme": normalize_monitor_ui_theme(self._combo_monitor_theme.currentData()),
                "monitor_ui": monitor_ui,
            }
        )
        self._cfg = cfg
        self._project = str(cfg.get("project") or "") or None
        self._update_effective_cfg_label()
        return cfg

    def _update_effective_cfg_label(self) -> None:
        cfg = self._collect_preview_cfg()
        workers = cfg.get("max_workers")
        workers_text = "auto" if not workers else str(workers)
        tags_workers = cfg.get("tags_max_workers")
        tags_workers_text = "auto" if not tags_workers else str(tags_workers)
        tags_backend = cfg.get("tags_text_backend") or "fitz"
        per_page_cfg = dict(cfg.get("per_page") or {})
        per_page_workers = per_page_cfg.get("max_workers")
        per_page_workers_text = "auto" if not per_page_workers else str(per_page_workers)
        per_page_chunk = per_page_cfg.get("chunk_size_pages")
        per_page_chunk_text = "auto" if not per_page_chunk else str(per_page_chunk)
        enabled_doc_types = [
            doc_type
            for doc_type, enabled in dict(per_page_cfg.get("doc_types") or {}).items()
            if enabled
        ]
        enabled_doc_types_text = ",".join(enabled_doc_types) if enabled_doc_types else "-"
        frame_band = (
            f"L{cfg.get('find_frame_border_band_left_mm')}/"
            f"R{cfg.get('find_frame_border_band_right_mm')}/"
            f"T{cfg.get('find_frame_border_band_top_mm')}/"
            f"B{cfg.get('find_frame_border_band_bottom_mm')}"
        )
        project = cfg.get("project") or "(all projects)"
        theme_key = normalize_monitor_ui_theme(cfg.get("monitor_ui_theme"))
        theme_label = MONITOR_UI_THEME_LABELS.get(theme_key, theme_key)
        text = (
            f"Templates: {cfg.get('templates_dir')}  |  "
            f"Project: {project}  |  "
            f"Тема: {theme_label}  |  "
            f"Parallel: {'on' if cfg.get('parallel') else 'off'}  |  "
            f"Workers: {workers_text}  |  "
            f"Per-page: {'on' if per_page_cfg.get('enabled') else 'off'}  |  "
            f"Per-page docs: {enabled_doc_types_text}  |  "
            f"Per-page min pages: {per_page_cfg.get('min_pages')}  |  "
            f"Per-page workers: {per_page_workers_text}  |  "
            f"Per-page chunk: {per_page_chunk_text}  |  "
            f"Tags workers: {tags_workers_text}  |  "
            f"Tags backend: {tags_backend}  |  "
            f"Frame band mm: {frame_band}  |  "
            f"Timing export: {'on' if cfg.get('timing_log') else 'off'}  |  "
            f"Excel sanitize: {'on' if excel_sanitize_illegal_chars_enabled(cfg) else 'off'}  |  "
            f"Cell prefilter: {'on' if cfg.get('grid_detected_prefilter_enabled', True) else 'off'} "
            f"(adj floor {float(cfg.get('grid_detected_prefilter_adjacency_tol_floor_pt', 2.5)):.2f} pt)  |  "
            f"Разъезд Run: OK≤{int(cfg.get('monitor_grid_mismatch_ok_max', 3000))} "
            f"жёлт.≤{int(cfg.get('monitor_grid_mismatch_warn_max', 45000))}"
        )
        self._lbl_effective_cfg.setText(text)

    def _collect_preview_cfg(self) -> dict[str, Any]:
        project_filter = ""
        if hasattr(self, "_combo_project") and self._combo_project.currentIndex() >= 0:
            project_filter = str(self._combo_project.currentData() or "")
        preview: dict[str, Any] = {
            "templates_dir": self._edit_templates_dir.text().strip(),
            "project": project_filter,
            "parallel": self._chk_parallel.isChecked(),
            "max_workers": int(self._spin_max_workers.value()),
            "per_page": {
                "enabled": self._chk_per_page_enabled.isChecked(),
                "min_pages": int(self._spin_per_page_min_pages.value()),
                "max_workers": int(self._spin_per_page_workers.value()),
                "chunk_size_pages": int(self._spin_per_page_chunk_size.value()),
                "doc_types": {
                    doc_type: checkbox.isChecked()
                    for doc_type, checkbox in self._chk_per_page_doc_types.items()
                },
            },
            "tags_max_workers": int(self._spin_tags_max_workers.value()),
            "tags_text_backend": self._combo_tags_backend.currentData(),
            "find_frame_border_band_left_mm": float(self._spin_frame_left_mm.value()),
            "find_frame_border_band_right_mm": float(self._spin_frame_right_mm.value()),
            "find_frame_border_band_top_mm": float(self._spin_frame_top_mm.value()),
            "find_frame_border_band_bottom_mm": float(self._spin_frame_bottom_mm.value()),
            "timing_log": self._chk_timing_log.isChecked(),
            **self._grid_detected_prefilter_cfg_slice(),
            "excel": {
                "sanitize_illegal_chars": self._chk_excel_sanitize.isChecked(),
            },
        }
        if hasattr(self, "_spin_grid_ok"):
            preview["monitor_grid_mismatch_ok_max"] = int(self._spin_grid_ok.value())
            preview["monitor_grid_mismatch_warn_max"] = max(
                int(self._spin_grid_ok.value()),
                int(self._spin_grid_warn.value()),
            )
            preview["monitor_grid_mismatch_weight_global"] = float(self._spin_weight_global.value())
            preview["monitor_grid_mismatch_weight_walk_pt"] = float(self._spin_weight_walk.value())
            preview["monitor_grid_mismatch_weight_no_match"] = float(self._spin_weight_no_match.value())
        if hasattr(self, "_combo_monitor_theme"):
            preview["monitor_ui_theme"] = normalize_monitor_ui_theme(self._combo_monitor_theme.currentData())
        return normalize_v2_config(preview)

    def _save_settings(self) -> None:
        if hasattr(self, "_tag_layout_saver"):
            self._tag_layout_saver.flush_now()
        snap_ta = self._snapshot_tag_analysis_ui_dict()
        cfg = self._collect_cfg_from_widgets()
        merged_mu = dict(cfg.get("monitor_ui") or get_default_monitor_ui_config())
        merged_mu["tag_analysis"] = normalize_tag_analysis_ui_config(snap_ta)
        cfg["monitor_ui"] = merged_mu
        cfg = normalize_v2_config(cfg)
        save_v2_config(cfg)
        self._cfg = cfg
        if hasattr(self, "_tag_layout_saver"):
            self._tag_layout_saver.sync_last_written_from_snapshot(
                cfg["monitor_ui"]["tag_analysis"]
            )
        self._apply_monitor_window_size_from_cfg()
        self._lbl_summary.setText("Настройки PDF v2 сохранены.")

    def _reload_settings(self) -> None:
        from pdf_parsing_v2.v2_config import load_v2_config

        self._load_cfg_into_widgets(load_v2_config())
        self._refresh_project_combo()
        self._lbl_summary.setText("Настройки PDF v2 перечитаны из файла.")

    def _browse_pdf_folder(self) -> None:
        start = _normalize_local_fs_path(self._current_pdf_path())
        if not start or not os.path.isdir(start):
            start = ""
        folder = QFileDialog.getExistingDirectory(self, "Выберите папку с PDF", start)
        if folder:
            self.set_pdf_folder(folder)

    def _browse_templates_dir(self) -> None:
        folder = QFileDialog.getExistingDirectory(self, "Выберите папку шаблонов")
        if folder:
            try:
                root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
                rel = os.path.relpath(folder, root)
            except ValueError:
                rel = folder
            self._edit_templates_dir.setText(rel.replace("\\", "/"))
            self._refresh_project_combo()

    def _browse_debug_visual_dir(self) -> None:
        folder = QFileDialog.getExistingDirectory(self, "Папка debug visual")
        if folder:
            try:
                root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
                rel = os.path.relpath(folder, root)
            except ValueError:
                rel = folder
            self._edit_debug_visual_dir.setText(rel.replace("\\", "/"))

    def _refresh_project_combo(self) -> None:
        try:
            from pdf_parsing_v2_engine.template_loader import list_project_combo_entries

            templates_dir = resolve_templates_dir(
                {"templates_dir": self._edit_templates_dir.text().strip()}
            )
            entries = (
                list_project_combo_entries(templates_dir)
                if os.path.isdir(templates_dir)
                else []
            )
        except Exception:
            entries = []

        current = str(self._cfg.get("project") or self._project or "")
        self._combo_project.blockSignals(True)
        self._combo_project.clear()
        self._combo_project.addItem("(all projects)", "")
        selected_index = 0
        for idx, (display_name, folder_name) in enumerate(entries, start=1):
            self._combo_project.addItem(display_name, folder_name)
            if folder_name == current:
                selected_index = idx
        self._combo_project.setCurrentIndex(selected_index)
        self._combo_project.blockSignals(False)
        self._update_effective_cfg_label()

    def _refresh_start_availability(self) -> None:
        path = _normalize_local_fs_path(self._current_pdf_path())
        self._btn_start.setEnabled(bool(path and os.path.isdir(path) and not self._pipeline_started))

    def _on_start_clicked(self) -> None:
        if self._pipeline_started:
            return
        path = _normalize_local_fs_path(self._current_pdf_path())
        if not path or not os.path.isdir(path):
            QMessageBox.warning(self, "Запуск", "Укажите существующую папку с PDF.")
            return
        self.start_pipeline()

    def start_pipeline(self) -> None:
        self._pipeline_started = True
        self._refresh_start_availability()
        self._reset_monitor_state()

        run_cfg = self._collect_cfg_from_widgets()
        self._pdf_path = _normalize_local_fs_path(self._current_pdf_path())
        self._remember_pdf_path(self._pdf_path)
        self._t_start = time.perf_counter()
        self._timer.start()
        self._tabs.setCurrentIndex(0)

        self._callback = MonitorCallback(self)
        self._callback.sig_file_start.connect(self._on_file_start)
        self._callback.sig_file_done.connect(self._on_file_done)
        self._callback.sig_file_error.connect(self._on_file_error)
        self._callback.sig_tag_file_start.connect(self._on_tag_file_start)
        self._callback.sig_tag_file_done.connect(self._on_tag_file_done)
        self._callback.sig_tag_file_error.connect(self._on_tag_file_error)
        self._callback.sig_batch_progress.connect(self._on_batch_progress)
        self._callback.sig_phase_changed.connect(self._on_phase_changed)
        self._callback.sig_stage_start.connect(self._on_stage_start)
        self._callback.sig_stage_done.connect(self._on_stage_done)
        self._callback.sig_stage_error.connect(self._on_stage_error)
        self._callback.sig_pipeline_finished.connect(self._on_pipeline_finished)

        self._thread = PipelineThread(
            self._pdf_path,
            run_cfg,
            self._project,
            self._callback,
            parent=self,
        )
        self._thread.sig_error.connect(self._on_pipeline_error)
        self._thread.finished.connect(self._on_thread_finished)
        self._thread.start()

    def _reset_monitor_state(self) -> None:
        self._result_dir = None
        self._last_summary = {}
        self._normcontrol_all_rows = []
        self._normcontrol_summary_placeholder = None
        self._normcontrol_rows_by_key = {}
        self._normcontrol_xlsx_path = ""
        self._stage_rows_by_key.clear()
        self._file_path_by_row.clear()
        self._monitor_rows.clear()
        self._uncovered_doc_type_by_file.clear()
        self._n_errors = 0
        self._total_files = 0
        self._total_tasks = 0
        self._completed = 0
        self._phase_times.clear()
        self._current_phase = ""
        self._phase_start = 0.0
        self._progress.setMaximum(1)
        self._progress.setValue(0)
        self._progress.setFormat("0 / 0 задач")
        self._lbl_phase.setText("Запуск pipeline...")
        self._lbl_summary.setStyleSheet("font-size: 12px;")
        self._lbl_summary.setText("Pipeline запущен.")
        self._table.setRowCount(0)
        self._stage_table.setRowCount(0)
        self._report_summary.clear()
        self._report_extraction.setRowCount(0)
        self._report_tags.setRowCount(0)
        self._report_od.setRowCount(0)
        self._report_rules.setRowCount(0)
        self._btn_open_dir.setEnabled(False)
        self._btn_nk_open_results_dir.setEnabled(False)
        self._btn_open_timing.setEnabled(False)
        self._btn_open_debug_report.setEnabled(False)
        self._btn_open_normcontrol.setEnabled(False)
        self._set_normcontrol_buttons_enabled(False)
        self._nk_table.blockSignals(True)
        self._nk_table.setRowCount(0)
        self._nk_table.blockSignals(False)
        self._lbl_nk_stats.setText("НК: данных пока нет.")
        self._nk_code_filter = None
        self._nk_doc_filter = None
        self._nk_page_filter = None
        self._nk_known_doc_names = frozenset()
        self._nk_filter_debounce.stop()
        self._nk_ed_text.blockSignals(True)
        self._nk_ed_text.clear()
        self._nk_ed_text.blockSignals(False)
        self._reload_nk_filter_options()
        self._clear_od_table_tab()
        self._clear_tag_analysis_tab()

    def _on_thread_finished(self) -> None:
        self._pipeline_started = False
        self._refresh_start_availability()

    def _on_file_start(self, file_path: str, index: int, total: int) -> None:
        self._ensure_monitor_rows(total)
        self._file_path_by_row[index] = file_path
        state = self._monitor_rows[index]
        state["file_name"] = os.path.basename(file_path)
        state["extract_status"] = _STATUS_RUNNING
        self._apply_uncovered_doc_type_warning_to_state(state, file_path)
        self._render_file_row(index)

    def _grid_mismatch_ui_from_detail(self, file_detail: object) -> tuple[str, str, QColor | None]:
        cfg = normalize_v2_config(getattr(self, "_cfg", None))
        ok_max = int(cfg.get("monitor_grid_mismatch_ok_max", 3000))
        warn_max = int(cfg.get("monitor_grid_mismatch_warn_max", 45000))
        if warn_max < ok_max:
            warn_max = ok_max
        if not isinstance(file_detail, dict):
            return "—", "", None
        worst = file_detail.get("stamp_grid_mismatch_worst_score")
        by_page = file_detail.get("stamp_grid_mismatch_by_page")
        tip_lines: list[str] = []
        if isinstance(by_page, list):
            for row in sorted(by_page, key=lambda x: int(x.get("page", 0))):
                p = int(row.get("page", 0))
                sc = int(row.get("score", 0))
                bd = row.get("breakdown") if isinstance(row.get("breakdown"), dict) else {}
                parts: list[str] = []
                if bd.get("global") is not None:
                    parts.append(f"глобал {int(bd['global'])}")
                if bd.get("walk_weighted") is not None:
                    parts.append(f"линии {int(bd['walk_weighted'])}")
                if bd.get("no_match_weighted") is not None:
                    parts.append(f"без ребра {int(bd['no_match_weighted'])}")
                extra = " · ".join(parts)
                tip_lines.append(f"Стр. {p} — {sc}" + (f" ({extra})" if extra else "") + ";")
        tooltip = "\n".join(tip_lines)
        if worst is None:
            return "—", tooltip, None
        w = int(worst)
        if w <= ok_max:
            return "OK", tooltip, None
        if w <= warn_max:
            return str(w), tooltip, _CLR_AFFINE_WARN
        return str(w), tooltip, _CLR_AFFINE_BAD

    def _refresh_monitor_grid_score_from_timing(self, summary: dict[str, Any]) -> None:
        timing = summary.get("timing") or {}
        files_info = timing.get("files") or {}
        for row in range(self._table.rowCount()):
            path = self._file_path_by_row.get(row)
            if not path:
                continue
            fname = os.path.basename(path)
            info = files_info.get(fname)
            if not isinstance(info, dict):
                continue
            worst = info.get("stamp_grid_mismatch_worst_score")
            js = info.get("stamp_grid_mismatch_by_page_json")
            detail: dict[str, Any] = {}
            if worst is not None:
                detail["stamp_grid_mismatch_worst_score"] = worst
            if isinstance(js, str) and js.strip():
                try:
                    detail["stamp_grid_mismatch_by_page"] = json.loads(js)
                except json.JSONDecodeError:
                    detail["stamp_grid_mismatch_by_page"] = []
            if not detail:
                continue
            t, tip, bg = self._grid_mismatch_ui_from_detail(detail)
            state = self._monitor_rows.setdefault(row, self._make_empty_row_state())
            state["grid_score_text"] = t
            state["grid_score_tooltip"] = tip
            state["grid_score_bg"] = bg
            self._render_file_row(row)

    def _on_file_done(
        self,
        file_path: str,
        index: int,
        n_pages: int,
        elapsed_sec: float,
        file_detail: object = None,
    ) -> None:
        state = self._monitor_rows[index]
        state["file_name"] = os.path.basename(file_path)
        state["extract_status"] = _STATUS_DONE
        state["pages"] = str(n_pages)
        state["extract_time"] = f"{elapsed_sec:.3f}"
        gs_t, gs_tip, gs_bg = self._grid_mismatch_ui_from_detail(file_detail)
        state["grid_score_text"] = gs_t
        state["grid_score_tooltip"] = gs_tip
        state["grid_score_bg"] = gs_bg
        self._apply_uncovered_doc_type_warning_to_state(state, file_path)
        self._render_file_row(index)

    def _on_file_error(self, file_path: str, index: int, error_msg: str) -> None:
        state = self._monitor_rows[index]
        state["file_name"] = os.path.basename(file_path)
        state["extract_status"] = _STATUS_ERROR
        state["error_is_warning"] = False
        state["error"] = self._merge_error_text("", error_msg)
        self._render_file_row(index)
        self._n_errors += 1

    def _on_tag_file_start(self, file_path: str, index: int, total: int) -> None:
        self._ensure_monitor_rows(total)
        self._file_path_by_row[index] = file_path
        state = self._monitor_rows[index]
        state["file_name"] = os.path.basename(file_path)
        state["tags_status"] = _STATUS_RUNNING
        self._render_file_row(index)

    def _on_tag_file_done(self, file_path: str, index: int, elapsed_sec: float) -> None:
        state = self._monitor_rows[index]
        state["file_name"] = os.path.basename(file_path)
        state["tags_status"] = _STATUS_DONE
        state["tags_time"] = f"{elapsed_sec:.3f}"
        self._render_file_row(index)

    def _on_tag_file_error(self, file_path: str, index: int, error_msg: str) -> None:
        state = self._monitor_rows[index]
        state["file_name"] = os.path.basename(file_path)
        state["tags_status"] = _STATUS_ERROR
        if state.get("error_is_warning"):
            state["error"] = ""
            state["error_is_warning"] = False
        state["error"] = self._merge_error_text(state["error"], error_msg)
        self._render_file_row(index)
        self._n_errors += 1

    def _on_batch_progress(self, completed: int, total: int) -> None:
        self._completed = completed
        if total > 0:
            self._total_tasks = total
            self._progress.setMaximum(total)
        self._progress.setValue(completed)
        self._progress.setFormat(f"{completed} / {total} задач")

    def _on_phase_changed(self, phase: str) -> None:
        now = time.perf_counter()
        if self._current_phase:
            self._phase_times[self._current_phase] = now - self._phase_start
        self._current_phase = phase
        self._phase_start = now
        self._lbl_phase.setText(phase)

    def _on_stage_start(
        self,
        category: str,
        stage: str,
        file_name: object,
        detail: object,
    ) -> None:
        key = (category, stage, str(file_name or ""))
        row = self._stage_table.rowCount()
        self._stage_table.insertRow(row)
        self._stage_rows_by_key[key] = row
        self._set_stage_row(
            row,
            category,
            stage,
            str(file_name or ""),
            "running",
            "",
            self._detail_text(detail),
        )

    def _on_stage_done(
        self,
        category: str,
        stage: str,
        elapsed_sec: float,
        file_name: object,
        detail: object,
    ) -> None:
        key = (category, stage, str(file_name or ""))
        row = self._stage_rows_by_key.get(key)
        if row is None:
            row = self._stage_table.rowCount()
            self._stage_table.insertRow(row)
        self._set_stage_row(
            row,
            category,
            stage,
            str(file_name or ""),
            "done",
            f"{elapsed_sec:.3f}",
            self._detail_text(detail),
        )
        if stage == "load_templates":
            self._handle_uncovered_doc_types_from_detail(detail)

    def _on_stage_error(
        self,
        category: str,
        stage: str,
        error_msg: str,
        file_name: object,
        detail: object,
    ) -> None:
        key = (category, stage, str(file_name or ""))
        row = self._stage_rows_by_key.get(key)
        if row is None:
            row = self._stage_table.rowCount()
            self._stage_table.insertRow(row)
        text = self._detail_text(detail)
        detail_text = f"{text} | {error_msg}" if text else error_msg
        self._set_stage_row(
            row,
            category,
            stage,
            str(file_name or ""),
            "error",
            "",
            detail_text,
        )
        brush = QBrush(_CLR_STAGE_ERROR_BG)
        for col in range(self._stage_table.columnCount()):
            item = self._stage_table.item(row, col)
            if item is not None:
                item.setBackground(brush)
        self._n_errors += 1
        self._tabs.setCurrentIndex(_TAB_INDEX_RUN)
        self._stage_table.selectRow(row)
        first = self._stage_table.item(row, 0)
        if first is not None:
            self._stage_table.scrollToItem(
                first,
                QAbstractItemView.ScrollHint.PositionAtCenter,
            )
        pop_title, pop_body = _stage_error_user_message(category, stage, error_msg)
        QMessageBox.warning(self, pop_title, pop_body)

    def _on_pipeline_finished(
        self,
        result_dir: str,
        total_elapsed: float,
        summary: dict[str, Any],
    ) -> None:
        if self._current_phase:
            self._phase_times[self._current_phase] = time.perf_counter() - self._phase_start
        self._result_dir = result_dir
        self._last_summary = summary or {}
        self._timer.stop()
        self._update_elapsed()
        self._lbl_phase.setText("Готово")
        parts = [f"Общее время: {total_elapsed:.1f}с", f"Ошибок: {self._n_errors}"]
        uncovered_n = int((summary or {}).get("uncovered_n_files") or 0)
        uncovered_types = [
            str(item)
            for item in ((summary or {}).get("uncovered_doc_types") or [])
            if str(item)
        ]
        if uncovered_n or uncovered_types:
            type_text = ",".join(uncovered_types) if uncovered_types else "—"
            parts.append(f"Без шаблона: {uncovered_n} ({type_text})")
        for phase, duration in self._phase_times.items():
            parts.append(f"{phase}: {duration:.1f}с")
        self._lbl_summary.setText("  |  ".join(parts))
        if self._n_errors:
            self._lbl_summary.setStyleSheet("font-size: 12px; color: #dc3232;")
        elif uncovered_n or uncovered_types:
            self._lbl_summary.setStyleSheet("font-size: 12px; color: #b45309;")
        self._populate_reports(summary or {})
        self._refresh_monitor_grid_score_from_timing(summary or {})
        self._btn_open_dir.setEnabled(bool(result_dir))
        self._btn_nk_open_results_dir.setEnabled(bool(result_dir))
        self._btn_open_timing.setEnabled(bool(self._artifact_path("timing_path")))
        self._btn_open_debug_report.setEnabled(bool(self._artifact_path("debug_report_path")))
        self._btn_open_normcontrol.setEnabled(bool(self._normcontrol_xlsx_path))
        self._tabs.setCurrentIndex(_TAB_INDEX_NORMCONTROL)

    def _on_pipeline_error(self, error_msg: str) -> None:
        self._timer.stop()
        self._lbl_phase.setText("Критическая ошибка")
        self._lbl_summary.setText(error_msg.split("\n")[0][:200])
        self._lbl_summary.setToolTip(error_msg)
        self._lbl_summary.setStyleSheet("font-size: 12px; color: #dc3232;")
        self._pipeline_started = False
        self._refresh_start_availability()

    def _make_empty_row_state(self) -> dict[str, Any]:
        return {
            "file_name": "",
            "extract_status": _STATUS_QUEUED,
            "grid_score_text": "—",
            "grid_score_tooltip": "",
            "grid_score_bg": None,
            "pages": "",
            "extract_time": "",
            "tags_status": _STATUS_QUEUED,
            "tags_time": "",
            "error": "",
            "error_is_warning": False,
        }

    def _ensure_monitor_rows(self, total_files: int) -> None:
        if self._total_files != 0:
            return
        self._total_files = total_files
        self._total_tasks = total_files * 2
        self._progress.setMaximum(max(self._total_tasks, 1))
        self._progress.setFormat(f"0 / {self._total_tasks} задач")
        self._table.setRowCount(total_files)
        for row in range(total_files):
            self._monitor_rows[row] = self._make_empty_row_state()
            self._render_file_row(row)

    def _render_file_row(self, row: int) -> None:
        state = self._monitor_rows.get(row, self._make_empty_row_state())
        self._set_file_row(
            row,
            state["file_name"],
            state["extract_status"],
            str(state.get("grid_score_text") or "—"),
            state["pages"],
            state["extract_time"],
            state["tags_status"],
            state["tags_time"],
            state["error"],
        )
        grid_it = self._table.item(row, _COL_GRID_MISMATCH)
        if grid_it:
            grid_it.setToolTip(str(state.get("grid_score_tooltip") or ""))
            bg = state.get("grid_score_bg")
            if bg is None:
                grid_it.setData(int(Qt.ItemDataRole.BackgroundRole), None)
            else:
                grid_it.setBackground(QBrush(bg))
        self._set_file_cell_color(row, _COL_EXTRACT, self._status_color(state["extract_status"]))
        self._set_file_cell_color(row, _COL_TAGS, self._status_color(state["tags_status"]))
        if state["error"]:
            error_item = self._table.item(row, _COL_ERROR)
            if state.get("error_is_warning"):
                self._set_file_cell_color(row, _COL_ERROR, _CLR_WARN_TEXT)
                if error_item:
                    error_item.setBackground(QBrush(_CLR_STAGE_WARN_BG))
                    error_item.setToolTip(state["error"])
            else:
                self._set_file_cell_color(row, _COL_ERROR, _CLR_ERROR)
                if error_item:
                    error_item.setToolTip(state["error"])

    def _status_color(self, status: str) -> QColor:
        if status == _STATUS_RUNNING:
            return _CLR_RUNNING
        if status == _STATUS_DONE:
            return _CLR_DONE
        if status == _STATUS_ERROR:
            return _CLR_ERROR
        return _CLR_QUEUED

    def _merge_error_text(self, existing: str, incoming: str) -> str:
        if existing:
            return existing
        return incoming.split("\n")[0][:160]

    def _set_file_row(
        self,
        row: int,
        file_name: str,
        extract_status: str,
        grid_score_text: str,
        pages: str,
        extract_time_s: str,
        tags_status: str,
        tags_time_s: str,
        error: str,
    ) -> None:
        values = [
            str(row + 1),
            file_name,
            extract_status,
            grid_score_text,
            pages,
            extract_time_s,
            tags_status,
            tags_time_s,
            error,
        ]
        alignments = [
            Qt.AlignmentFlag.AlignCenter,
            Qt.AlignmentFlag.AlignLeft,
            Qt.AlignmentFlag.AlignCenter,
            Qt.AlignmentFlag.AlignCenter,
            Qt.AlignmentFlag.AlignCenter,
            Qt.AlignmentFlag.AlignRight,
            Qt.AlignmentFlag.AlignCenter,
            Qt.AlignmentFlag.AlignRight,
            Qt.AlignmentFlag.AlignLeft,
        ]
        for col, value in enumerate(values):
            item = QTableWidgetItem(value)
            item.setTextAlignment(int(alignments[col] | Qt.AlignmentFlag.AlignVCenter))
            self._table.setItem(row, col, item)

    def _set_stage_row(
        self,
        row: int,
        category: str,
        stage: str,
        item_name: str,
        status: str,
        elapsed_sec: str,
        detail: str,
    ) -> None:
        values = [category, stage, item_name, status, elapsed_sec, detail]
        base_flags = Qt.ItemFlag.ItemIsSelectable | Qt.ItemFlag.ItemIsEnabled
        for col, value in enumerate(values):
            item = QTableWidgetItem(value)
            if col == _STAGE_COL_DETAIL:
                item.setTextAlignment(
                    int(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignTop)
                )
                item.setFlags(base_flags | Qt.ItemFlag.ItemIsEditable)
            else:
                if col == 4:
                    item.setTextAlignment(
                        int(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
                    )
                item.setFlags(base_flags)
            self._stage_table.setItem(row, col, item)
        self._stage_table.resizeRowToContents(row)

    def _set_file_cell_color(self, row: int, col: int, color: QColor) -> None:
        brush = QBrush(color)
        item = self._table.item(row, col)
        if item:
            item.setForeground(brush)

    def _tick_elapsed(self) -> None:
        self._update_elapsed()

    def _update_elapsed(self) -> None:
        elapsed = max(0.0, time.perf_counter() - self._t_start)
        minutes, seconds = divmod(int(elapsed), 60)
        self._lbl_elapsed.setText(f"{minutes:02d}:{seconds:02d}")

    def _detail_text(self, detail: object) -> str:
        if isinstance(detail, dict):
            skip = {"uncovered_files", "uncovered_pairs"}
            parts = [
                f"{key}={value}"
                for key, value in detail.items()
                if key not in skip and value not in ("", None, {}, [])
            ]
            return "; ".join(parts)
        return str(detail or "")

    def _apply_uncovered_doc_type_warning_to_state(
        self,
        state: dict[str, Any],
        file_path: str,
    ) -> None:
        doc_type = self._uncovered_doc_type_by_file.get(os.path.basename(file_path))
        if not doc_type:
            return
        warning = f"Нет шаблона для типа {doc_type}"
        if state.get("error_is_warning") or not state.get("error"):
            state["error"] = warning
            state["error_is_warning"] = True

    def _handle_uncovered_doc_types_from_detail(self, detail: object) -> None:
        if not isinstance(detail, dict):
            return
        types = [str(item) for item in (detail.get("uncovered_doc_types") or []) if str(item)]
        n_files = int(detail.get("uncovered_n_files") or 0)
        if not types and not n_files:
            return
        self._uncovered_doc_type_by_file.clear()
        pairs = detail.get("uncovered_pairs") or []
        if isinstance(pairs, list):
            for item in pairs:
                if not isinstance(item, (list, tuple)) or len(item) < 2:
                    continue
                name = os.path.basename(str(item[0] or ""))
                doc_type = str(item[1] or "").strip() or "?"
                if name:
                    self._uncovered_doc_type_by_file[name] = doc_type
        for row, path in self._file_path_by_row.items():
            state = self._monitor_rows.get(row)
            if not state:
                continue
            self._apply_uncovered_doc_type_warning_to_state(state, path)
            self._render_file_row(row)
        key = ("pipeline", "load_templates", "")
        row = self._stage_rows_by_key.get(key)
        if row is not None:
            status_item = self._stage_table.item(row, 3)
            if status_item is not None:
                status_item.setText("warning")
            brush = QBrush(_CLR_STAGE_WARN_BG)
            for col in range(self._stage_table.columnCount()):
                item = self._stage_table.item(row, col)
                if item is not None:
                    item.setBackground(brush)
        from pdf_parsing_v2.v2_pipeline import format_uncovered_doc_types_warning

        QMessageBox.warning(
            self,
            "Нет шаблона штампа",
            format_uncovered_doc_types_warning(detail),
        )

    def _populate_reports(self, summary: dict[str, Any]) -> None:
        timing = summary.get("timing", {})
        artifacts = summary.get("artifacts", {})
        normcontrol = summary.get("normcontrol", {}) or {}
        lines = [
            f"Result dir: {summary.get('result_dir', '')}",
            f"Project: {summary.get('effective_project') or '(all projects)'}",
            f"Parallel enabled: {summary.get('parallel_enabled')}",
            f"Used parallel: {summary.get('used_parallel')}",
            f"Max workers: {summary.get('max_workers')}",
            f"Files: {summary.get('n_files')}",
            f"Pages: {summary.get('n_pages')}",
            (
                "Uncovered doc types: "
                + (
                    f"{summary.get('uncovered_n_files') or 0} "
                    f"({', '.join(str(t) for t in (summary.get('uncovered_doc_types') or []) if str(t))})"
                    if (summary.get("uncovered_n_files") or summary.get("uncovered_doc_types"))
                    else "0"
                )
            ),
        ]
        for phase in timing.get("phases", []):
            lines.append(
                f"Phase {phase.get('label')}: {round(float(phase.get('elapsed_sec', 0.0)), 3)} s"
            )
        if artifacts:
            lines.append("")
            lines.append("Artifacts:")
            for key, value in artifacts.items():
                lines.append(f"  {key}: {value}")
        self._report_summary.setPlainText("\n".join(lines))

        self._fill_table_from_rows(
            self._report_extraction,
            self._rows_from_file_timings(timing.get("files", {})),
        )
        self._fill_table_from_rows(
            self._report_tags,
            self._rows_from_tag_timings(timing.get("stages", [])),
        )
        self._fill_table_from_rows(
            self._report_od,
            self._normalize_report_rows(
                [row for row in timing.get("stages", []) if row.get("category") == "od"]
            ),
        )
        self._fill_table_from_rows(
            self._report_rules,
            self._normalize_report_rows(
                [row for row in timing.get("stages", []) if row.get("category") == "rules"]
            ),
        )
        self._normcontrol_xlsx_path = str(normcontrol.get("xlsx_path", "") or artifacts.get("normcontrol_path", "") or "")
        self._normcontrol_all_rows = [dict(row) for row in normcontrol.get("all_rows", [])]
        ph = normcontrol.get("summary_placeholder_row")
        self._normcontrol_summary_placeholder = dict(ph) if isinstance(ph, dict) else None
        self._normcontrol_rows_by_key = {
            str(row.get("row_key", "")): row for row in self._normcontrol_all_rows if str(row.get("row_key", ""))
        }
        self._reload_nk_filter_options()
        self._refresh_normcontrol_table()
        self._populate_od_table_from_summary(summary.get("od_table", {}) or {})
        self._populate_tag_analysis_from_summary(summary.get("tag_analysis") or {})

    def _nk_summary_placeholder_matches_filters(self, ph: dict[str, Any]) -> bool:
        """Return whether the synthetic «no errors» row should be visible under current filters."""
        if self._nk_code_filter is not None and self._nk_code_filter:
            if _nk_row_code_str(ph) not in self._nk_code_filter:
                return False
        if self._nk_doc_filter is not None and self._nk_doc_filter:
            if str(ph.get("doc_name", "") or "").strip() not in self._nk_doc_filter:
                return False
        if self._nk_page_filter is not None and self._nk_page_filter:
            if _nk_row_page_str(ph) not in self._nk_page_filter:
                return False
        needle = self._nk_ed_text.text().strip().casefold()
        if needle:
            blob = " ".join(
                (
                    str(ph.get("c_description", "") or ""),
                    str(ph.get("doc_name", "") or ""),
                    str(ph.get("text", "") or ""),
                    str(ph.get("comment", "") or ""),
                    _nk_row_code_str(ph),
                )
            ).casefold()
            if needle not in blob:
                return False
        return True

    def _rows_from_file_timings(self, file_timings: dict[str, dict[str, Any]]) -> list[dict[str, Any]]:
        visible_cols = self._visible_extraction_columns(file_timings)
        rows: list[dict[str, Any]] = []
        for file_name in sorted(file_timings):
            row = {"file_name": file_name}
            for key, value in file_timings[file_name].items():
                if key not in visible_cols:
                    continue
                row[key] = self._normalize_report_value(value)
            rows.append(row)
        return rows

    def _rows_from_tag_timings(self, stage_rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
        rows: list[dict[str, Any]] = []
        for row in stage_rows:
            if row.get("category") != "tags" or row.get("stage") != "parse_tags_file":
                continue
            rows.append(
                {
                    "file_name": self._normalize_report_value(row.get("file_name", "")),
                    "elapsed_sec": self._normalize_report_value(row.get("elapsed_sec", "")),
                    "status": self._normalize_report_value(row.get("status", "")),
                    "error": self._normalize_report_value(row.get("error", "")),
                }
            )
        rows.sort(key=lambda item: str(item.get("file_name", "")))
        return rows

    def _visible_extraction_columns(
        self,
        file_timings: dict[str, dict[str, Any]],
    ) -> set[str]:
        visible = {"elapsed_sec", "n_pages", "error"}
        max_by_key: dict[str, float] = {}
        for info in file_timings.values():
            for key, value in info.items():
                if key in visible:
                    continue
                if isinstance(value, (int, float)):
                    max_by_key[key] = max(max_by_key.get(key, 0.0), float(value))
                elif value not in ("", None, {}):
                    visible.add(key)
        for key, value in max_by_key.items():
            if value >= _EXTRACTION_DETAIL_THRESHOLD_SEC:
                visible.add(key)
        return visible

    def _normalize_report_rows(self, rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
        return [
            {key: self._normalize_report_value(value) for key, value in row.items()}
            for row in rows
        ]

    def _normalize_report_value(self, value: Any) -> Any:
        if isinstance(value, float):
            return round(value, 3)
        return value

    def _set_normcontrol_buttons_enabled(self, enabled: bool) -> None:
        has_rows = bool(self._normcontrol_all_rows or self._normcontrol_summary_placeholder)
        self._btn_nk_save.setEnabled(enabled and bool(self._normcontrol_xlsx_path))
        self._btn_nk_import.setEnabled(True)
        self._btn_nk_copy.setEnabled(enabled and has_rows and bool(self._normcontrol_xlsx_path))

    def _schedule_nk_filter_debounced_refresh(self, *_args: object) -> None:
        """Debounce table rebuild for doc/page/text filters (avoids UI stalls)."""
        self._nk_filter_debounce.stop()
        self._nk_filter_debounce.start()

    def _nk_flush_filter_refresh(self) -> None:
        """Apply filters immediately (focus out / Enter); coalesces with debounce timer."""
        self._nk_filter_debounce.stop()
        self._refresh_normcontrol_table()

    def _update_normcontrol_stats(self, visible_rows: int) -> None:
        total_checks = len(self._normcontrol_all_rows)
        total_errors = sum(1 for row in self._normcontrol_all_rows if not row.get("result", False))
        resolved_errors = sum(
            1 for row in self._normcontrol_all_rows if (not row.get("result", False)) and row.get("resolved", False)
        )
        active_errors = total_errors - resolved_errors
        self._lbl_nk_stats.setText(
            "НК: "
            f"всего проверок {total_checks} | "
            f"ошибок {total_errors} | "
            f"исправлено {resolved_errors} | "
            f"активных {active_errors} | "
            f"видимых строк {visible_rows}"
        )

    def _reload_nk_filter_options(self) -> None:
        """Rebuild code/doc/page combos from ``_normcontrol_all_rows``."""
        rows = self._normcontrol_all_rows
        codes = _nk_sorted_unique_codes(rows)
        docs = _nk_sorted_unique_doc_names(rows)
        pages = _nk_sorted_unique_pages(rows)
        self._nk_known_doc_names = frozenset(docs)
        self._nk_reloading_filter_widgets = True
        try:
            if self._nk_code_filter is not None:
                avail = frozenset(codes)
                inter = frozenset(c for c in self._nk_code_filter if c in avail)
                if not inter:
                    self._nk_code_filter = None
                else:
                    self._nk_code_filter = inter
            if self._nk_doc_filter is not None:
                availd = frozenset(docs)
                interd = frozenset(x for x in self._nk_doc_filter if x in availd)
                if not interd:
                    self._nk_doc_filter = None
                else:
                    self._nk_doc_filter = interd
            if self._nk_page_filter is not None:
                availp = frozenset(pages)
                interp = frozenset(x for x in self._nk_page_filter if x in availp)
                if not interp:
                    self._nk_page_filter = None
                else:
                    self._nk_page_filter = interp
            cmb = self._nk_cmb_code
            cmb.clear()
            cmb.addItem(_NK_FILTER_ALL, None)
            for c in codes:
                cmb.addItem(c, c)
            self._nk_cmb_doc.blockSignals(True)
            self._nk_cmb_doc.clear()
            self._nk_cmb_doc.addItem(_NK_FILTER_ALL)
            for d in docs:
                self._nk_cmb_doc.addItem(d)
            self._nk_cmb_doc.blockSignals(False)
            self._nk_cmb_page.blockSignals(True)
            self._nk_cmb_page.clear()
            self._nk_cmb_page.addItem(_NK_FILTER_ALL)
            for p in pages:
                self._nk_cmb_page.addItem(p)
            self._nk_cmb_page.blockSignals(False)
            self._sync_nk_code_combo_from_state()
            self._sync_nk_doc_combo_from_state()
            self._sync_nk_page_combo_from_state()
        finally:
            self._nk_reloading_filter_widgets = False

    def _sync_nk_code_combo_from_state(self) -> None:
        """Sync quick code combo with ``_nk_code_filter`` (multi-select disables combo)."""
        cmb = self._nk_cmb_code
        cmb.blockSignals(True)
        try:
            if cmb.count() > 0:
                cmb.setItemText(0, _NK_FILTER_ALL)
            sel = self._nk_code_filter
            if sel is None or not sel:
                cmb.setEnabled(True)
                cmb.setCurrentIndex(0)
            elif len(sel) == 1:
                code = next(iter(sel))
                cmb.setEnabled(True)
                idx = cmb.findData(code, Qt.ItemDataRole.UserRole, Qt.MatchFlag.MatchExactly)
                if idx < 0:
                    idx = cmb.findText(code, Qt.MatchFlag.MatchExactly)
                if idx >= 0:
                    cmb.setCurrentIndex(idx)
                else:
                    cmb.setCurrentIndex(0)
                    self._nk_code_filter = None
            else:
                if cmb.count() > 0:
                    cmb.setItemText(0, f"«Коды…»: выбрано {len(sel)}")
                cmb.setCurrentIndex(0)
                cmb.setEnabled(False)
        finally:
            cmb.blockSignals(False)

    def _sync_nk_doc_combo_from_state(self) -> None:
        """Sync document combo with ``_nk_doc_filter`` (multi-select disables combo)."""
        cmb = self._nk_cmb_doc
        cmb.blockSignals(True)
        try:
            if cmb.count() > 0:
                cmb.setItemText(0, _NK_FILTER_ALL)
            sel = self._nk_doc_filter
            if sel is None or not sel:
                cmb.setEnabled(True)
                cmb.setCurrentIndex(0)
            elif len(sel) == 1:
                name = next(iter(sel))
                cmb.setEnabled(True)
                idx = cmb.findText(name, Qt.MatchFlag.MatchExactly)
                if idx >= 0:
                    cmb.setCurrentIndex(idx)
                else:
                    cmb.setCurrentIndex(0)
                    self._nk_doc_filter = None
            else:
                if cmb.count() > 0:
                    cmb.setItemText(0, f"«Документы…»: {len(sel)}")
                cmb.setCurrentIndex(0)
                cmb.setEnabled(False)
        finally:
            cmb.blockSignals(False)

    def _sync_nk_page_combo_from_state(self) -> None:
        """Sync page combo with ``_nk_page_filter`` (multi-select disables combo)."""
        cmb = self._nk_cmb_page
        cmb.blockSignals(True)
        try:
            if cmb.count() > 0:
                cmb.setItemText(0, _NK_FILTER_ALL)
            sel = self._nk_page_filter
            if sel is None or not sel:
                cmb.setEnabled(True)
                cmb.setCurrentIndex(0)
            elif len(sel) == 1:
                page = next(iter(sel))
                cmb.setEnabled(True)
                idx = cmb.findText(page, Qt.MatchFlag.MatchExactly)
                if idx >= 0:
                    cmb.setCurrentIndex(idx)
                else:
                    cmb.setCurrentIndex(0)
                    self._nk_page_filter = None
            else:
                if cmb.count() > 0:
                    cmb.setItemText(0, f"«Страницы…»: {len(sel)}")
                cmb.setCurrentIndex(0)
                cmb.setEnabled(False)
        finally:
            cmb.blockSignals(False)

    def _on_nk_code_combo_index_changed(self, index: int) -> None:
        if self._nk_reloading_filter_widgets or not self._nk_cmb_code.isEnabled():
            return
        if index <= 0:
            self._nk_code_filter = None
        else:
            cmb = self._nk_cmb_code
            data = cmb.itemData(index, Qt.ItemDataRole.UserRole)
            code = str(data) if data is not None else cmb.itemText(index)
            self._nk_code_filter = frozenset({code})
        self._refresh_normcontrol_table()

    def _on_nk_doc_combo_index_changed(self, index: int) -> None:
        if self._nk_reloading_filter_widgets or not self._nk_cmb_doc.isEnabled():
            return
        if index <= 0:
            self._nk_doc_filter = None
        else:
            self._nk_doc_filter = frozenset({self._nk_cmb_doc.currentText().strip()})
        self._refresh_normcontrol_table()

    def _on_nk_page_combo_index_changed(self, index: int) -> None:
        if self._nk_reloading_filter_widgets or not self._nk_cmb_page.isEnabled():
            return
        if index <= 0:
            self._nk_page_filter = None
        else:
            self._nk_page_filter = frozenset({self._nk_cmb_page.currentText().strip()})
        self._refresh_normcontrol_table()

    def _apply_normcontrol_extra_filters(self, rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
        """Column filters (phase 1) after checkbox filters."""
        out = rows
        if self._nk_code_filter is not None:
            sel = self._nk_code_filter
            out = [r for r in out if _nk_row_code_str(r) in sel]
        if self._nk_doc_filter is not None:
            df = self._nk_doc_filter
            out = [r for r in out if str(r.get("doc_name", "") or "").strip() in df]
        if self._nk_page_filter is not None:
            pf = self._nk_page_filter
            out = [r for r in out if _nk_row_page_str(r) in pf]
        needle = self._nk_ed_text.text().strip().casefold()
        if needle:

            def _row_matches_text(r: dict[str, Any]) -> bool:
                blob = " ".join(
                    (
                        str(r.get("c_description", "") or ""),
                        str(r.get("doc_name", "") or ""),
                        str(r.get("text", "") or ""),
                        str(r.get("comment", "") or ""),
                        _nk_row_code_str(r),
                    )
                ).casefold()
                return needle in blob

            out = [r for r in out if _row_matches_text(r)]
        return out

    def _reset_nk_column_filters(self) -> None:
        self._nk_filter_debounce.stop()
        self._nk_code_filter = None
        self._nk_doc_filter = None
        self._nk_page_filter = None
        self._nk_reloading_filter_widgets = True
        try:
            self._sync_nk_code_combo_from_state()
            self._sync_nk_doc_combo_from_state()
            self._sync_nk_page_combo_from_state()
            self._nk_ed_text.blockSignals(True)
            self._nk_ed_text.clear()
            self._nk_ed_text.blockSignals(False)
        finally:
            self._nk_reloading_filter_widgets = False
        self._refresh_normcontrol_table()

    def _open_nk_codes_filter_dialog(self) -> None:
        codes = _nk_sorted_unique_codes(self._normcontrol_all_rows)
        if not codes:
            QMessageBox.information(self, "НК", "Нет кодов для фильтрации.")
            return
        dlg = _NormcontrolMultiPickDialog(
            self,
            window_title="Фильтр по кодам НК",
            hint="Отметьте коды для отображения. Если отмечены все — фильтр по коду отключён.",
            items=codes,
            current=self._nk_code_filter,
        )
        if dlg.exec() != QDialog.DialogCode.Accepted:
            return
        self._nk_code_filter = dlg.selected_or_none_if_all()
        self._nk_reloading_filter_widgets = True
        try:
            self._sync_nk_code_combo_from_state()
        finally:
            self._nk_reloading_filter_widgets = False
        self._refresh_normcontrol_table()

    def _open_nk_docs_filter_dialog(self) -> None:
        docs = _nk_sorted_unique_doc_names(self._normcontrol_all_rows)
        if not docs:
            QMessageBox.information(self, "НК", "Нет документов для фильтрации.")
            return
        dlg = _NormcontrolMultiPickDialog(
            self,
            window_title="Фильтр по документам",
            hint="Отметьте документы для отображения. Если отмечены все — фильтр по документу отключён.",
            items=docs,
            current=self._nk_doc_filter,
        )
        if dlg.exec() != QDialog.DialogCode.Accepted:
            return
        self._nk_doc_filter = dlg.selected_or_none_if_all()
        self._nk_reloading_filter_widgets = True
        try:
            self._sync_nk_doc_combo_from_state()
        finally:
            self._nk_reloading_filter_widgets = False
        self._refresh_normcontrol_table()

    def _open_nk_pages_filter_dialog(self) -> None:
        pages = _nk_sorted_unique_pages(self._normcontrol_all_rows)
        if not pages:
            QMessageBox.information(self, "НК", "Нет страниц для фильтрации.")
            return
        dlg = _NormcontrolMultiPickDialog(
            self,
            window_title="Фильтр по страницам",
            hint="Отметьте страницы для отображения. Если отмечены все — фильтр по странице отключён.",
            items=pages,
            current=self._nk_page_filter,
        )
        if dlg.exec() != QDialog.DialogCode.Accepted:
            return
        self._nk_page_filter = dlg.selected_or_none_if_all()
        self._nk_reloading_filter_widgets = True
        try:
            self._sync_nk_page_combo_from_state()
        finally:
            self._nk_reloading_filter_widgets = False
        self._refresh_normcontrol_table()

    def _refresh_normcontrol_table(self) -> None:
        rows = list(self._normcontrol_all_rows)
        if not self._chk_nk_show_positive.isChecked():
            rows = [row for row in rows if not row.get("result", False)]
        if self._chk_nk_hide_resolved.isChecked():
            rows = [row for row in rows if not row.get("resolved", False)]
        rows = self._apply_normcontrol_extra_filters(rows)

        ph = self._normcontrol_summary_placeholder
        display_rows: list[dict[str, Any]] = []
        if ph is not None and self._nk_summary_placeholder_matches_filters(ph):
            display_rows.append(ph)
        display_rows.extend(rows)

        self._update_normcontrol_stats(len(display_rows))
        self._set_normcontrol_buttons_enabled(
            bool(self._normcontrol_all_rows or self._normcontrol_summary_placeholder)
        )

        self._nk_table.setUpdatesEnabled(False)
        self._nk_table.blockSignals(True)
        try:
            self._nk_table.setSortingEnabled(False)
            self._nk_table.horizontalHeader().setSortIndicatorShown(False)
            self._nk_table.setRowCount(len(display_rows))
            for row_idx, row in enumerate(display_rows):
                self._populate_normcontrol_row(row_idx, row)
            self._nk_table.setSortingEnabled(True)
        finally:
            self._nk_table.blockSignals(False)
            self._nk_table.setUpdatesEnabled(True)

    def _populate_normcontrol_row(self, row_idx: int, row: dict[str, Any]) -> None:
        summary = _nk_is_summary_placeholder(row)
        row_key = str(row.get("row_key", ""))
        resolved = bool(row.get("resolved", False))
        result_ok = bool(row.get("result", False))
        status_text = "OK" if result_ok else "Ошибка"
        page_num = row.get("page_num", "")
        page_sort = row.get("page_sort", (1, str(page_num)))

        resolved_item = _SortableTableWidgetItem("")
        if summary:
            resolved_item.setFlags(Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsSelectable)
            resolved_item.setCheckState(Qt.CheckState.Unchecked)
        else:
            resolved_item.setFlags(
                Qt.ItemFlag.ItemIsEnabled
                | Qt.ItemFlag.ItemIsSelectable
                | Qt.ItemFlag.ItemIsUserCheckable
            )
            resolved_item.setCheckState(Qt.CheckState.Checked if resolved else Qt.CheckState.Unchecked)
        resolved_item.setData(_ROLE_ROW_KEY, row_key)
        resolved_item.setData(_ROLE_SORT_VALUE, 1 if resolved else 0)
        self._nk_table.setItem(row_idx, _NK_COL_RESOLVED, resolved_item)

        status_item = _SortableTableWidgetItem(status_text)
        status_item.setFlags(
            Qt.ItemFlag.ItemIsEnabled
            | Qt.ItemFlag.ItemIsSelectable
            | (Qt.ItemFlag.ItemIsEditable if not summary else Qt.ItemFlag.NoItemFlags)
        )
        status_item.setData(_ROLE_ROW_KEY, row_key)
        status_item.setData(_ROLE_SORT_VALUE, ("\u0000", 1 if result_ok else 0) if summary else (1 if result_ok else 0))
        self._nk_table.setItem(row_idx, _NK_COL_STATUS, status_item)

        code_item = _SortableTableWidgetItem(str(row.get("c_code", "")))
        code_item.setFlags(
            Qt.ItemFlag.ItemIsEnabled
            | Qt.ItemFlag.ItemIsSelectable
            | (Qt.ItemFlag.ItemIsEditable if not summary else Qt.ItemFlag.NoItemFlags)
        )
        code_item.setData(_ROLE_ROW_KEY, row_key)
        code_item.setData(
            _ROLE_SORT_VALUE,
            ("\u0000", int(row.get("c_code", 0) or 0)) if summary else int(row.get("c_code", 0) or 0),
        )
        self._nk_table.setItem(row_idx, _NK_COL_CODE, code_item)

        _nk_editable = (
            Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsSelectable | Qt.ItemFlag.ItemIsEditable
        )
        _nk_readonly = Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsSelectable
        for col_idx, key in (
            (_NK_COL_DESCRIPTION, "c_description"),
            (_NK_COL_DOC, "doc_name"),
            (_NK_COL_TEXT, "text"),
            (_NK_COL_RESOLVED_AT, "resolved_at"),
            (_NK_COL_COMMENT, "comment"),
        ):
            item = _SortableTableWidgetItem(str(row.get(key, "")))
            item.setFlags(_nk_readonly if summary else _nk_editable)
            item.setData(_ROLE_ROW_KEY, row_key)
            sort_val = str(row.get(key, "")).lower()
            item.setData(_ROLE_SORT_VALUE, ("\u0000", sort_val) if summary else sort_val)
            self._nk_table.setItem(row_idx, col_idx, item)

        page_item = _SortableTableWidgetItem(str(page_num))
        page_item.setFlags(
            Qt.ItemFlag.ItemIsEnabled
            | Qt.ItemFlag.ItemIsSelectable
            | (Qt.ItemFlag.ItemIsEditable if not summary else Qt.ItemFlag.NoItemFlags)
        )
        page_item.setData(_ROLE_ROW_KEY, row_key)
        page_item.setData(_ROLE_SORT_VALUE, ("\u0000", page_sort) if summary else page_sort)
        self._nk_table.setItem(row_idx, _NK_COL_PAGE, page_item)
        self._apply_normcontrol_row_style(row_idx, row)

    def _apply_normcontrol_row_style(self, row_idx: int, row: dict[str, Any]) -> None:
        resolved = bool(row.get("resolved", False))
        summary = _nk_is_summary_placeholder(row)
        gray = QBrush(QColor(128, 128, 128))
        bg = _CLR_NK_RESOLVED if resolved and not summary else None
        for col_idx in range(self._nk_table.columnCount()):
            item = self._nk_table.item(row_idx, col_idx)
            if item is None:
                continue
            if summary:
                item.setForeground(gray)
            else:
                item.setForeground(QBrush())
            if bg is not None:
                item.setBackground(QBrush(bg))
            else:
                item.setBackground(QBrush())

    def _find_normcontrol_row(self, row_key: str) -> dict[str, Any] | None:
        return self._normcontrol_rows_by_key.get(row_key)

    def _on_nk_item_changed(self, item: QTableWidgetItem) -> None:
        row_key = str(item.data(_ROLE_ROW_KEY) or "")
        row = self._find_normcontrol_row(row_key)
        if row is None:
            return
        if item.column() == _NK_COL_RESOLVED:
            resolved = item.checkState() == Qt.CheckState.Checked
            row["resolved"] = resolved
            row["resolved_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S") if resolved else ""
            self._nk_table.blockSignals(True)
            resolved_at_item = self._nk_table.item(item.row(), _NK_COL_RESOLVED_AT)
            if resolved_at_item is not None:
                resolved_at_item.setText(str(row["resolved_at"]))
                resolved_at_item.setData(_ROLE_SORT_VALUE, str(row["resolved_at"]).lower())
            self._apply_normcontrol_row_style(item.row(), row)
            self._nk_table.blockSignals(False)
            if self._chk_nk_hide_resolved.isChecked() and resolved:
                self._refresh_normcontrol_table()
            else:
                self._update_normcontrol_stats(self._nk_table.rowCount())
            return
        if item.column() == _NK_COL_COMMENT:
            row["comment"] = item.text()
            item.setData(_ROLE_SORT_VALUE, item.text().lower())
            self._nk_table.resizeRowToContents(item.row())

    def _open_normcontrol_xlsx(self) -> None:
        self._open_path(self._normcontrol_xlsx_path or self._artifact_path("normcontrol_path"))

    def _save_normcontrol_review_state(self) -> None:
        if not self._normcontrol_xlsx_path:
            QMessageBox.warning(self, "НК", "Файл результата проверки НК пока недоступен.")
            return
        try:
            save_review_state_to_workbook(
                self._normcontrol_all_rows,
                self._normcontrol_xlsx_path,
                sanitize_illegal_chars=excel_sanitize_illegal_chars_enabled(self._cfg),
            )
        except PermissionError:
            QMessageBox.warning(
                self,
                "НК",
                "Не удалось сохранить отметки: файл открыт в Excel или недоступен для записи.",
            )
            return
        except Exception as exc:
            QMessageBox.warning(self, "НК", f"Не удалось сохранить отметки:\n{exc}")
            return
        QMessageBox.information(self, "НК", "Отметки сохранены в xlsx.")

    def _import_normcontrol_review_state(self) -> None:
        xlsx_path, _ = QFileDialog.getOpenFileName(
            self,
            "Импортировать отметки из НК xlsx",
            self._normcontrol_file_dialog_start_dir(),
            "Excel files (*.xlsx)",
            options=_NON_NATIVE_FILE_DIALOG_OPTION,
        )
        if not xlsx_path:
            return
        try:
            report = load_normcontrol_report_from_workbook(xlsx_path)
        except Exception as exc:
            QMessageBox.warning(self, "НК", f"Не удалось импортировать отметки:\n{exc}")
            return
        imported_rows = [dict(row) for row in report.get("all_rows", [])]
        ph_imp = report.get("summary_placeholder_row")
        ph_imp = dict(ph_imp) if isinstance(ph_imp, dict) else None
        if not imported_rows and not ph_imp:
            review_state = load_review_state_from_workbook(xlsx_path)
            if not review_state or not self._normcontrol_all_rows:
                QMessageBox.warning(self, "НК", "В выбранном xlsx не найдены данные НК.")
                return
            for row in self._normcontrol_all_rows:
                state = review_state.get(str(row.get("row_key", "")))
                if not state:
                    continue
                row["resolved"] = bool(state.get("resolved", False))
                row["resolved_at"] = str(state.get("resolved_at", "") or "")
                row["comment"] = str(state.get("comment", "") or "")
            self._refresh_normcontrol_table()
            QMessageBox.information(self, "НК", "Отметки импортированы.")
            return
        self._normcontrol_all_rows = imported_rows
        self._normcontrol_summary_placeholder = ph_imp
        self._normcontrol_rows_by_key = {
            str(row.get("row_key", "")): row for row in self._normcontrol_all_rows if str(row.get("row_key", ""))
        }
        self._normcontrol_xlsx_path = xlsx_path
        self._btn_open_normcontrol.setEnabled(True)
        self._nk_code_filter = None
        self._nk_doc_filter = None
        self._nk_page_filter = None
        self._nk_filter_debounce.stop()
        self._nk_ed_text.blockSignals(True)
        self._nk_ed_text.clear()
        self._nk_ed_text.blockSignals(False)
        self._reload_nk_filter_options()
        self._refresh_normcontrol_table()
        self._tabs.setCurrentIndex(_TAB_INDEX_NORMCONTROL)
        QMessageBox.information(self, "НК", "НК xlsx загружен.")

    def _save_normcontrol_copy_as(self) -> None:
        if not self._normcontrol_xlsx_path or not os.path.exists(self._normcontrol_xlsx_path):
            QMessageBox.warning(self, "НК", "Исходный xlsx результата проверки НК не найден.")
            return
        target_path, _ = QFileDialog.getSaveFileName(
            self,
            "Сохранить НК xlsx как",
            self._normcontrol_xlsx_path,
            "Excel files (*.xlsx)",
            options=_NON_NATIVE_FILE_DIALOG_OPTION,
        )
        if not target_path:
            return
        try:
            shutil.copyfile(self._normcontrol_xlsx_path, target_path)
            save_review_state_to_workbook(
                self._normcontrol_all_rows,
                target_path,
                sanitize_illegal_chars=excel_sanitize_illegal_chars_enabled(self._cfg),
            )
        except Exception as exc:
            QMessageBox.warning(self, "НК", f"Не удалось сохранить копию xlsx:\n{exc}")
            return
        QMessageBox.information(self, "НК", "Копия xlsx сохранена.")

    def _normcontrol_row_from_table(self, table_row: int) -> dict[str, Any] | None:
        item = self._nk_table.item(table_row, _NK_COL_DOC)
        row_key = str(item.data(_ROLE_ROW_KEY) if item is not None else "")
        if not row_key:
            item = self._nk_table.item(table_row, _NK_COL_RESOLVED)
            row_key = str(item.data(_ROLE_ROW_KEY) if item is not None else "")
        return self._find_normcontrol_row(row_key)

    def _show_normcontrol_context_menu(self, pos) -> None:
        row = self._nk_table.rowAt(pos.y())
        if row < 0:
            return
        data_row = self._normcontrol_row_from_table(row)
        if data_row is None:
            return
        pdf_path = str(data_row.get("pdf_path", "") or "")
        menu = QMenu(self)
        act_open_pdf = menu.addAction("Открыть PDF (на страницу, если возможно)")
        act_open_plain = menu.addAction("Открыть PDF")
        act_open_editor = menu.addAction("Открыть в редакторе шаблонов")
        act_open_folder = menu.addAction("Открыть папку")
        chosen = menu.exec(self._nk_table.viewport().mapToGlobal(pos))
        if chosen == act_open_pdf:
            self._open_pdf_best_effort(pdf_path, data_row.get("page_num"))
        elif chosen == act_open_plain:
            self._open_path(pdf_path)
        elif chosen == act_open_editor:
            self._launch_template_editor(pdf_path)
        elif chosen == act_open_folder:
            self._open_containing_folder(pdf_path)

    def _open_pdf_best_effort(self, pdf_path: str, page_num: Any) -> None:
        pdf_path = _normalize_local_fs_path(pdf_path)
        if not pdf_path or not os.path.exists(pdf_path):
            return
        if isinstance(page_num, int):
            page_num_int = page_num
        else:
            text = str(page_num or "").strip()
            page_num_int = int(text) if text.isdigit() else None
        if page_num_int is not None and self._try_open_pdf_at_page(pdf_path, page_num_int):
            return
        self._open_path(pdf_path)

    def _try_open_pdf_at_page(self, pdf_path: str, page_num: int) -> bool:
        pdf_path = _normalize_local_fs_path(pdf_path)
        if not pdf_path:
            return False
        candidates = [
            os.path.join(os.environ.get("LOCALAPPDATA", ""), "SumatraPDF", "SumatraPDF.exe"),
            os.path.join(os.environ.get("ProgramFiles", ""), "SumatraPDF", "SumatraPDF.exe"),
            os.path.join(os.environ.get("ProgramFiles(x86)", ""), "SumatraPDF", "SumatraPDF.exe"),
            os.path.join(os.environ.get("ProgramFiles", ""), "Adobe", "Acrobat DC", "Acrobat", "Acrobat.exe"),
            os.path.join(os.environ.get("ProgramFiles(x86)", ""), "Adobe", "Acrobat Reader DC", "Reader", "AcroRd32.exe"),
        ]
        for exe_path in candidates:
            exe_path = str(exe_path or "").strip()
            if not exe_path or not os.path.exists(exe_path):
                continue
            try:
                base = os.path.basename(exe_path).lower()
                if "sumatra" in base:
                    subprocess.Popen([exe_path, "-page", str(page_num), pdf_path])
                    return True
                if "acro" in base:
                    subprocess.Popen([exe_path, "/A", f"page={page_num}", pdf_path])
                    return True
            except Exception:
                continue
        return False

    def _restore_normcontrol_column_widths(self) -> None:
        raw = self._settings_store.value(_NK_COLUMN_WIDTHS_KEY, [])
        if isinstance(raw, (list, tuple)):
            widths = [int(x) for x in raw]
        elif isinstance(raw, str):
            widths = [int(x) for x in raw.split(",") if str(x).strip().isdigit()]
        else:
            widths = []
        if not widths:
            return
        for idx, width in enumerate(widths):
            if 0 <= idx < self._nk_table.columnCount() and width > 20:
                self._nk_table.setColumnWidth(idx, width)

    def _on_normcontrol_section_resized(self, _logical_index: int, _old_size: int, _new_size: int) -> None:
        widths = [self._nk_table.columnWidth(i) for i in range(self._nk_table.columnCount())]
        self._settings_store.setValue(_NK_COLUMN_WIDTHS_KEY, widths)

    def _fill_table_from_rows(self, table: QTableWidget, rows: list[dict[str, Any]]) -> None:
        headers: list[str] = []
        for row in rows:
            for key in row.keys():
                if key not in headers:
                    headers.append(key)
        table.clear()
        if not headers:
            table.setColumnCount(1)
            table.setRowCount(1)
            table.setHorizontalHeaderLabels(["Info"])
            table.setItem(0, 0, QTableWidgetItem("No data"))
            return
        table.setColumnCount(len(headers))
        table.setHorizontalHeaderLabels(headers)
        table.setRowCount(len(rows))
        for row_idx, row in enumerate(rows):
            for col_idx, key in enumerate(headers):
                table.setItem(
                    row_idx,
                    col_idx,
                    QTableWidgetItem(str(self._normalize_report_value(row.get(key, "")))),
                )
        table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.ResizeToContents)
        if headers:
            table.horizontalHeader().setSectionResizeMode(len(headers) - 1, QHeaderView.ResizeMode.Stretch)

    def _artifact_path(self, key: str) -> str:
        return str((self._last_summary.get("artifacts", {}) or {}).get(key, "") or "")

    def _open_result_dir(self) -> None:
        self._open_path(self._result_dir)

    def _open_containing_folder(self, path: str | None) -> None:
        """Open the parent folder of a file (or the folder itself if ``path`` is a directory)."""
        path = _normalize_local_fs_path(path)
        if not path:
            return
        if os.path.isfile(path):
            folder = os.path.dirname(os.path.abspath(path))
        elif os.path.isdir(path):
            folder = os.path.abspath(path)
        else:
            parent = os.path.dirname(os.path.abspath(path))
            folder = parent if os.path.isdir(parent) else ""
        if folder and os.path.isdir(folder):
            os.startfile(folder)

    def _open_path(self, path: str | None) -> None:
        path = _normalize_local_fs_path(path)
        if not path:
            return
        if os.path.exists(path):
            os.startfile(path)

    def _open_file_from_monitor_row(self, row: int, _col: int) -> None:
        self._open_path(self._monitor_row_path(row))

    def _open_file_from_stage_row(self, row: int, col: int) -> None:
        if col == _STAGE_COL_DETAIL:
            return
        item = self._stage_table.item(row, _STAGE_COL_ITEM)
        if item and item.text():
            p = _normalize_local_fs_path(item.text())
            if p and os.path.exists(p):
                os.startfile(p)

    def _monitor_row_path(self, row: int) -> str | None:
        return self._file_path_by_row.get(row)

    def _show_monitor_context_menu(self, pos) -> None:
        row = self._table.rowAt(pos.y())
        if row < 0:
            return
        pdf_path = self._monitor_row_path(row)
        if not pdf_path:
            return
        menu = QMenu(self)
        act_open_pdf = menu.addAction("Открыть PDF")
        act_open_editor = menu.addAction("Открыть в редакторе шаблонов")
        act_open_folder = menu.addAction("Открыть папку")
        act_od_read: QAction | None = None
        act_od_read_debug: QAction | None = None
        if self._pdf_path_is_od_document(pdf_path):
            act_od_read = menu.addAction("ОД чтение ведомости")
            act_od_read_debug = menu.addAction("ОД чтение ведомости (отладка)")
        chosen = menu.exec(self._table.viewport().mapToGlobal(pos))
        if chosen == act_open_pdf:
            self._open_path(pdf_path)
        elif chosen == act_open_editor:
            self._launch_template_editor(pdf_path)
        elif chosen == act_open_folder:
            self._open_containing_folder(pdf_path)
        elif act_od_read is not None and chosen == act_od_read:
            self._run_od_read_manifest(pdf_path, od_debug=False)
        elif act_od_read_debug is not None and chosen == act_od_read_debug:
            self._run_od_read_manifest(pdf_path, od_debug=True)

    def _show_od_context_menu(self, pos) -> None:
        pdf_path = _normalize_local_fs_path((self._od_pdf_path or "").strip())
        if not pdf_path:
            return
        menu = QMenu(self)
        act_open_pdf = menu.addAction("Открыть PDF")
        act_open_folder = menu.addAction("Открыть папку")
        act_open_editor = menu.addAction("Открыть в редакторе шаблонов")
        act_open_pdf.setEnabled(os.path.isfile(pdf_path))
        act_open_folder.setEnabled(os.path.isfile(pdf_path) or os.path.isdir(pdf_path))
        act_open_editor.setEnabled(os.path.isfile(pdf_path))
        chosen = menu.exec(self._od_result_table.viewport().mapToGlobal(pos))
        if chosen == act_open_pdf:
            self._open_path(pdf_path)
        elif chosen == act_open_folder:
            self._open_containing_folder(pdf_path)
        elif chosen == act_open_editor:
            self._launch_template_editor(pdf_path)

    def _pdf_path_is_od_document(self, pdf_path: str) -> bool:
        try:
            return string_parsing.getDocTypeFromFile(os.path.basename(pdf_path)) == "OD"
        except Exception:
            return False

    def _format_od_debug_trace(self, entries: list[dict[str, str]]) -> str:
        """Human-readable multi-block trace for standalone OD debug."""
        lines: list[str] = []
        for i, e in enumerate(entries, start=1):
            lines.append("═" * 72)
            lines.append(f"#{i:03d}  function: {e.get('function', '')}")
            lines.append(f"      stage:    {e.get('stage', '')}")
            lines.append("─" * 72)
            lines.append(e.get("body", "").rstrip())
            lines.append("")
        if not entries:
            lines.append("(трассировка пуста — шаги не записаны)")
        return "\n".join(lines).rstrip() + "\n"

    def _show_od_debug_dialog(self, entries: list[dict[str, str]], footer: str) -> None:
        dlg = QDialog(self)
        dlg.setWindowTitle("ОД — отладочная трассировка")
        dlg.resize(920, 620)
        layout = QVBoxLayout(dlg)
        hint = QLabel(
            "Порядок блоков соответствует этапам: параметры → геометрия обрезки → "
            "pdfplumber (сырые таблицы) → объединённый текст страницы → разбор строк → "
            "_get_page_count (форматы листов)."
        )
        hint.setWordWrap(True)
        layout.addWidget(hint)
        text = QPlainTextEdit()
        text.setReadOnly(True)
        mono = QFont("Consolas", 10)
        mono.setStyleHint(QFont.StyleHint.Monospace)
        text.setFont(mono)
        body = self._format_od_debug_trace(entries)
        if footer.strip():
            body = body + "\n" + "═" * 72 + "\n" + footer.strip() + "\n"
        text.setPlainText(body)
        layout.addWidget(text)
        btn_row = QHBoxLayout()
        copy_btn = QPushButton("Копировать в буфер")
        close_btn = QPushButton("Закрыть")
        btn_row.addStretch()
        btn_row.addWidget(copy_btn)
        btn_row.addWidget(close_btn)
        layout.addLayout(btn_row)

        def _copy() -> None:
            QGuiApplication.clipboard().setText(text.toPlainText())

        copy_btn.clicked.connect(_copy)
        close_btn.clicked.connect(dlg.accept)
        dlg.exec()

    def _run_od_read_manifest(self, od_pdf_full_path: str, *, od_debug: bool = False) -> None:
        """Parse OD table only (no full pipeline) and show the ОД tab."""
        od_pdf_full_path = _normalize_local_fs_path(od_pdf_full_path)
        if not od_pdf_full_path or not os.path.isfile(od_pdf_full_path):
            QMessageBox.warning(self, "ОД", "Файл ОД не найден по пути из таблицы.")
            return
        from pdf_parsing_v2 import load_all_templates, load_project_templates
        from pdf_parsing_v2_od.od_engine_findtables_debug import (
            collect_od_manifest_findtables_debug,
        )
        from pdf_parsing_v2_od.od_parsing import (
            build_od_table_payload,
            error_message_for_invalid_od_parse_result,
            od_table_parsing_f,
        )

        td = resolve_templates_dir(self._cfg)
        eff = (self._project or self._cfg.get("project") or "").strip() or None
        tpl_list = (
            load_project_templates(td, eff) if eff else load_all_templates(td)
        )

        geom_out: list = []
        debug_entries: list[dict[str, str]] = []
        QApplication.setOverrideCursor(Qt.CursorShape.WaitCursor)
        try:
            od_warnings: list[str] = []
            raw = od_table_parsing_f(
                pdf_full_path=od_pdf_full_path,
                debug_print_out_list=0,
                warnings_out=od_warnings,
                cfg=self._cfg,
                templates=tpl_list or [],
                curr_proj=None,
                manifest_geometry_out=geom_out,
                debug_trace_out=debug_entries if od_debug else None,
            )
            if not isinstance(raw, list):
                od_payload = build_od_table_payload(
                    [],
                    od_pdf_full_path,
                    status="error",
                    error=error_message_for_invalid_od_parse_result(od_pdf_full_path),
                    warnings=od_warnings,
                )
            elif len(raw) == 0:
                od_payload = build_od_table_payload(
                    [],
                    od_pdf_full_path,
                    status="no_rows",
                    warnings=od_warnings,
                )
            else:
                od_payload = build_od_table_payload(
                    raw,
                    od_pdf_full_path,
                    status="ok",
                    warnings=od_warnings or None,
                )
        except Exception as e:
            od_payload = build_od_table_payload(
                [],
                od_pdf_full_path,
                status="error",
                error=str(e),
            )
        finally:
            QApplication.restoreOverrideCursor()
        try:
            if tpl_list:
                dbg = collect_od_manifest_findtables_debug(
                    od_pdf_full_path,
                    tpl_list,
                    self._cfg,
                    curr_proj=None,
                    geometry_rows=geom_out if geom_out else None,
                )
                if dbg:
                    od_payload["engine_findtables_debug"] = dbg
        except Exception as exc:
            od_payload["engine_findtables_debug"] = {"error": str(exc), "pages": []}
        self._populate_od_table_from_summary(od_payload)
        self._tabs.setCurrentIndex(_TAB_INDEX_OD)
        if od_debug:
            err = (od_payload.get("error") or "").strip()
            warn_lines = od_payload.get("warnings") or []
            wtxt = ""
            if warn_lines:
                wtxt = "\n".join(str(w) for w in warn_lines)
            footer_parts = [
                f"summary status: {od_payload.get('status')!r}",
                f"payload error: {err}" if err else "payload error: (none)",
            ]
            if wtxt:
                footer_parts.append("warnings:\n" + wtxt)
            self._show_od_debug_dialog(debug_entries, footer="\n".join(footer_parts))

    def _launch_template_editor(self, pdf_path: str | None = None) -> None:
        root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        cmd = [sys.executable, "-m", "pdf_template_editor"]
        if pdf_path:
            cmd.append(_normalize_local_fs_path(pdf_path))
        subprocess.Popen(cmd, cwd=root)


def apply_monitor_ui_theme_full(theme: Any, *, window: MonitorWindow | None = None) -> None:
    """Set Qt ``ColorScheme`` from ``theme`` and refresh open control-center chrome (tables, NK, …).

    Passing ``window=`` is recommended while a ``MonitorWindow`` is not yet a top-level widget
    (e.g. during ``__init__`` / settings reload).
    """
    apply_monitor_ui_theme(theme)
    app = QApplication.instance()
    if app is None:
        return
    app.processEvents()
    if window is not None:
        window.refresh_ui_after_theme_change(theme)
        return
    for w in app.topLevelWidgets():
        if isinstance(w, MonitorWindow):
            w.refresh_ui_after_theme_change(theme)
