"""RFP parts collection tab — DS/hybrid cockpit plus legacy ``RFQ.rfp_parts``."""

from __future__ import annotations

import html
import os
from collections.abc import Callable
from pathlib import Path

from PySide6.QtCore import Qt, QUrl
from PySide6.QtGui import QBrush, QColor
from PySide6.QtWidgets import (
    QAbstractItemView,
    QFileDialog,
    QGroupBox,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QSizePolicy,
    QTabWidget,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from RFQ.ds_compare.ds_compare_config import (
    load_ds_compare_config,
    normalize_gui_paths,
    save_ds_compare_config,
)
from RFQ.rfp_parts.analyze_rfp_parts import (
    COLLISIONS_XLSX_NAME,
    DEFAULT_PARTS_DIR,
    DEFAULT_REPORTS_BASE_DIR,
    NET_XLSX_NAME,
    find_latest_rfp_parts_run_dir,
    resolve_latest_rfp_parts_collisions_xlsx,
    resolve_latest_rfp_parts_net_xlsx,
)
from RFQ.rfp_parts.ds_baseline import BASELINE_XLSX_NAME
from RFQ.rfp_parts.ds_hybrid_preflight import (
    ds_baseline_output_dir,
    ds_hybrid_output_dir,
)
from RFQ.rfp_parts.ds_jobs import CockpitRow, DsCockpitSnapshot, get_last_ds_cockpit
from RFQ.rfp_parts.ds_registry import DEFAULT_REGISTRY_PATH
from RFQ.rfp_parts.ds_rfp_hybrid import (
    HYBRID_REPORT_PREFIX,
    HYBRID_XLSX_NAME,
    find_latest_hybrid_report,
)
from RFQ.rfp_parts.file_status import (
    DUPLICATE_TAGS_XLSX_NAME,
    STATUS_ERROR,
    STATUS_OK,
    load_file_status_payload,
    resolve_duplicate_tags_xlsx,
)
from ds_compare_center.split_layout import (
    WrappingLabel,
    WrappingPlainText,
    build_side_by_side,
    shrink_h,
)
from utils.path import open_dir

_PATH_FIELD_STYLE = (
    "QTextBrowser { color: #333; font-size: 11px; background: transparent; "
    "border: none; outline: none; padding: 0; }"
    "QTextBrowser a { color: #0645ad; text-decoration: underline; }"
)
_FILE_STATUS_COLORS = {
    STATUS_OK: (QColor("#1a7f37"), QColor("#e6f4ea")),
    STATUS_ERROR: (QColor("#b42318"), QColor("#fce8e6")),
}
_TONE_COLORS = {
    "ok": (QColor("#333333"), QColor("#ffffff")),
    "match": (QColor("#1a7f37"), QColor("#e6f4ea")),
    "warn": (QColor("#9a6700"), QColor("#fff8c5")),
    "error": (QColor("#b42318"), QColor("#fce8e6")),
}
_BANNER_STYLES = {
    "ok": "color: #333; font-size: 11px;",
    "match": "color: #1a7f37; font-size: 11px;",
    "warn": "color: #9a6700; font-size: 11px;",
    "error": "color: #b42318; font-size: 11px;",
}
_FILE_STATUS_GROUP_STYLE = (
    "QGroupBox { padding-top: 2px; margin-top: 6px; }"
    "QGroupBox::title { subcontrol-origin: margin; subcontrol-position: top left; "
    "left: 8px; padding: 0 4px; }"
)
_FILE_STATUS_TABLE_STYLE = (
    "QTableWidget { font-size: 11px; }"
    "QTableWidget::item { padding: 1px 4px; }"
)
_FILE_STATUS_ROW_HEIGHT = 22
_REGISTRY_HEADERS = (
    "ID ДС",
    "Статус",
    "Группы",
    "Ключ RFP",
    "Папка УЛ",
    "Режим",
    "Замечание",
)
_GROUP_HEADERS = (
    "Группа",
    "ДС источники",
    "Ключ RFP",
    "УЛ",
    "Блок overlay",
    "Статус",
)
_DS_FILE_HEADERS = ("Контур", "Файл", "ID / ключ", "Статус")
_COVERAGE_HEADERS = ("Группа", "Файлы ДС", "RFP", "УЛ", "Статус")


def rfp_parts_reports_base_dir() -> Path:
    """UNC folder ``…\\_RFP\\RFP сводный файл`` for stamped report runs."""
    return DEFAULT_REPORTS_BASE_DIR


def find_latest_rfp_parts_reports_dir() -> Path | None:
    """Newest stamp folder ``YYYY.MM.DD_HH.MM`` under the reports base."""
    return find_latest_rfp_parts_run_dir()


def rfp_parts_reports_dir() -> Path:
    """Last run reports folder, else the UNC reports base (open-folder fallback)."""
    return find_latest_rfp_parts_reports_dir() or rfp_parts_reports_base_dir()


def _file_href(path: Path) -> str:
    """``file:`` URL for local/UNC path (safe for HTML href)."""
    return html.escape(QUrl.fromLocalFile(str(path)).toString(), quote=True)


def _path_link_html(path: Path) -> str:
    """Clickable path that opens Explorer via ``path_link_activated``."""
    return f'<a href="{_file_href(path)}">{html.escape(str(path))}</a>'


def _path_status_line_html(label: str, path: Path) -> str:
    mark = "OK" if _safe_exists(path) else "NO"
    return f"[{mark}] {html.escape(label)}: {_path_link_html(path)}"


def _safe_exists(path: Path) -> bool:
    try:
        return path.exists()
    except OSError:
        return False


def _safe_is_file(path: Path) -> bool:
    try:
        return path.is_file()
    except OSError:
        return False


def _safe_is_dir(path: Path) -> bool:
    try:
        return path.is_dir()
    except OSError:
        return False


class RfpPartsPanel(QWidget):
    """DS/hybrid cockpit plus legacy RFP parts diagnostics."""

    def __init__(
        self,
        parent: QWidget | None = None,
        *,
        on_run: Callable[[], None] | None = None,
        on_ds_baseline: Callable[[], None] | None = None,
        on_ds_hybrid: Callable[[], None] | None = None,
        on_ds_coverage: Callable[[], None] | None = None,
        on_ds_registry: Callable[[], None] | None = None,
    ) -> None:
        """Create the panel.

        Args:
            parent: Qt parent.
            on_run: Start legacy ``python -m RFQ.rfp_parts``.
            on_ds_baseline: Start ``run_ds_baseline_job``.
            on_ds_hybrid: Start ``run_ds_hybrid_job``.
            on_ds_coverage: Start ``run_ds_coverage_job``.
            on_ds_registry: Start ``run_ds_registry_check_job``.
        """
        super().__init__(parent)
        self._on_run = on_run
        self._on_ds_baseline = on_ds_baseline
        self._on_ds_hybrid = on_ds_hybrid
        self._on_ds_coverage = on_ds_coverage
        self._on_ds_registry = on_ds_registry
        self._last_reports_dir: Path | None = find_latest_rfp_parts_reports_dir()

        splitter, self.monitor = build_side_by_side(
            self, build_left=self._build_left, show_stop=True
        )
        self.splitter = splitter
        root = QVBoxLayout(self)
        root.setContentsMargins(8, 8, 8, 8)
        root.addWidget(splitter)
        self.reload_paths_from_config()
        self._refresh_status()
        self._refresh_paths()
        self._refresh_file_status()
        last = get_last_ds_cockpit()
        if last is not None:
            self.fill_from_snapshot(last)

    def set_last_reports_dir(self, path: Path) -> None:
        """Remember the out-dir used for the current/last run."""
        self._last_reports_dir = path
        self._refresh_status()

    @property
    def last_reports_dir(self) -> Path | None:
        """Out-dir of the last run (if any)."""
        return self._last_reports_dir

    def ds_job_paths(self) -> dict[str, str]:
        """Return folder/file paths currently shown on the cockpit."""
        gui = normalize_gui_paths(load_ds_compare_config().get("gui_paths"))
        registry = self._edit_registry.text().strip() or str(DEFAULT_REGISTRY_PATH)
        return {
            "source_root": self._edit_ds_folder.text().strip(),
            "registry_path": registry,
            "ul_root": str(gui.get("last_tsd_packing_folder", "")).strip(),
            "rfp_root": str(DEFAULT_PARTS_DIR),
        }

    def reload_paths_from_config(self) -> None:
        """Load ``last_ds_trusted_folder`` / ``last_ds_registry_file`` into edits."""
        gui = normalize_gui_paths(load_ds_compare_config().get("gui_paths"))
        if hasattr(self, "_edit_ds_folder"):
            self._edit_ds_folder.setText(gui.get("last_ds_trusted_folder", ""))
        if hasattr(self, "_edit_registry"):
            self._edit_registry.setText(
                gui.get("last_ds_registry_file", "") or str(DEFAULT_REGISTRY_PATH)
            )
        self._refresh_paths()

    def _build_left(self, left: QWidget, left_layout: QVBoxLayout) -> None:
        left_layout.addWidget(self._build_registry_group(left), stretch=0)
        left_layout.addWidget(self._build_source_group(left), stretch=0)
        left_layout.addWidget(self._build_cockpit_group(left), stretch=1)
        left_layout.addWidget(self._build_file_status_group(left), stretch=1)
        left_layout.addWidget(self._build_actions_group(left), stretch=0)

    def _build_registry_group(self, parent: QWidget) -> QGroupBox:
        box = shrink_h(QGroupBox("Реестр ДС", parent))
        box.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Maximum)
        v = QVBoxLayout(box)
        v.setContentsMargins(10, 13, 10, 10)
        v.setSpacing(6)
        self._registry_banner = WrappingLabel("Реестр ещё не проверялся.", box)
        v.addWidget(self._registry_banner)
        row = QHBoxLayout()
        btn_check = QPushButton("Проверить реестр", box)
        btn_check.setMinimumWidth(0)
        btn_check.clicked.connect(self._click_ds_registry)
        row.addWidget(btn_check)
        btn_open = QPushButton("Открыть реестр", box)
        btn_open.setMinimumWidth(0)
        btn_open.clicked.connect(self._open_registry)
        row.addWidget(btn_open)
        row.addStretch(1)
        v.addLayout(row)
        return box

    def _build_source_group(self, parent: QWidget) -> QGroupBox:
        box = shrink_h(QGroupBox("Пути ДС / RFP / УЛ", parent))
        box.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Maximum)
        v = QVBoxLayout(box)
        v.setContentsMargins(10, 13, 10, 10)
        v.setSpacing(8)

        hint = WrappingLabel(
            "Доверенный вход Step1 собирается из закупочных ДС и канонического "
            "реестра. Overlay RFP атомарно на группу поставки. Legacy "
            "«Проверить части RFP» по-прежнему вызывает python -m RFQ.rfp_parts "
            "и не подменяет rfp_parts_net.xlsx этим контуром.",
            box,
        )
        v.addWidget(hint)

        v.addWidget(QLabel("Папка доверенных ДС", box))
        self._edit_ds_folder = QLineEdit(box)
        self._edit_ds_folder.setMinimumWidth(0)
        v.addWidget(self._edit_ds_folder)
        btn_ds = QPushButton("Обзор папки ДС…", box)
        btn_ds.setMinimumWidth(0)
        btn_ds.clicked.connect(self._browse_ds_folder)
        v.addWidget(btn_ds, alignment=Qt.AlignmentFlag.AlignLeft)

        v.addWidget(QLabel("Реестр ДС (канон UNC, робот не заменяет)", box))
        self._edit_registry = QLineEdit(box)
        self._edit_registry.setMinimumWidth(0)
        v.addWidget(self._edit_registry)
        btn_reg = QPushButton("Обзор реестра…", box)
        btn_reg.setMinimumWidth(0)
        btn_reg.clicked.connect(self._browse_registry)
        v.addWidget(btn_reg, alignment=Qt.AlignmentFlag.AlignLeft)

        self._paths_label = WrappingPlainText("", box)
        self._paths_label.setStyleSheet(_PATH_FIELD_STYLE)
        self._paths_label.path_link_activated.connect(self._open_path_link)
        v.addWidget(self._paths_label)

        btn_refresh_paths = QPushButton("Обновить статус путей", box)
        btn_refresh_paths.setMinimumWidth(0)
        btn_refresh_paths.clicked.connect(self._on_refresh_paths_clicked)
        v.addWidget(btn_refresh_paths)

        self._status = WrappingPlainText("", box)
        self._status.setStyleSheet(_PATH_FIELD_STYLE)
        self._status.path_link_activated.connect(self._open_path_link)
        v.addWidget(self._status)
        return box

    def _build_cockpit_group(self, parent: QWidget) -> QGroupBox:
        box = shrink_h(QGroupBox("Сводка ДС / hybrid", parent))
        box.setStyleSheet(_FILE_STATUS_GROUP_STYLE)
        box.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        v = QVBoxLayout(box)
        v.setContentsMargins(8, 4, 8, 6)
        v.setSpacing(4)
        self._cockpit_summary = QLabel("Нет данных cockpit.", box)
        self._cockpit_summary.setWordWrap(True)
        self._cockpit_summary.setTextFormat(Qt.TextFormat.RichText)
        self._cockpit_summary.setAlignment(
            Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter
        )
        v.addWidget(self._cockpit_summary)

        tabs = QTabWidget(box)
        tabs.setDocumentMode(True)
        self._table_registry = self._make_table(box, _REGISTRY_HEADERS)
        self._table_groups = self._make_table(box, _GROUP_HEADERS)
        self._table_files = self._make_table(box, _DS_FILE_HEADERS)
        self._table_coverage = self._make_table(box, _COVERAGE_HEADERS)
        tabs.addTab(self._table_registry, "Реестр по ДС")
        tabs.addTab(self._table_groups, "Итог по группам")
        tabs.addTab(self._table_files, "Файлы")
        tabs.addTab(self._table_coverage, "Покрытие")
        self._cockpit_tabs = tabs
        v.addWidget(tabs, stretch=1)
        return box

    def _make_table(self, parent: QWidget, headers: tuple[str, ...]) -> QTableWidget:
        table = QTableWidget(parent)
        table.setColumnCount(len(headers))
        table.setHorizontalHeaderLabels(list(headers))
        table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        table.setWordWrap(False)
        table.verticalHeader().setVisible(False)
        table.verticalHeader().setDefaultSectionSize(_FILE_STATUS_ROW_HEIGHT)
        table.verticalHeader().setMinimumSectionSize(_FILE_STATUS_ROW_HEIGHT)
        table.verticalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Fixed)
        table.setSortingEnabled(False)
        table.setMinimumHeight(140)
        table.setStyleSheet(_FILE_STATUS_TABLE_STYLE)
        header = table.horizontalHeader()
        for col in range(len(headers) - 1):
            header.setSectionResizeMode(col, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(
            len(headers) - 1, QHeaderView.ResizeMode.Stretch
        )
        header.setStretchLastSection(True)
        return table

    def _build_file_status_group(self, parent: QWidget) -> QGroupBox:
        box = shrink_h(QGroupBox("Файлы последнего прогона частей RFP", parent))
        box.setStyleSheet(_FILE_STATUS_GROUP_STYLE)
        box.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        v = QVBoxLayout(box)
        v.setContentsMargins(8, 4, 8, 6)
        v.setSpacing(3)

        self._file_status_summary = QLabel("Нет данных последнего прогона.", box)
        self._file_status_summary.setWordWrap(False)
        self._file_status_summary.setAlignment(
            Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter
        )
        self._file_status_summary.setSizePolicy(
            QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed
        )
        v.addWidget(self._file_status_summary)

        self._tag_remarks_link = QLabel("", box)
        self._tag_remarks_link.setWordWrap(False)
        self._tag_remarks_link.setTextFormat(Qt.TextFormat.RichText)
        self._tag_remarks_link.setTextInteractionFlags(
            Qt.TextInteractionFlag.TextBrowserInteraction
        )
        self._tag_remarks_link.setOpenExternalLinks(False)
        self._tag_remarks_link.linkActivated.connect(self._on_tag_remarks_link)
        self._tag_remarks_link.setVisible(False)
        self._tag_remarks_link.setSizePolicy(
            QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed
        )
        v.addWidget(self._tag_remarks_link)

        table = QTableWidget(box)
        table.setColumnCount(3)
        table.setHorizontalHeaderLabels(["Файл", "Статус", "Описание"])
        table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        table.setWordWrap(False)
        table.setAlternatingRowColors(False)
        table.verticalHeader().setVisible(False)
        table.verticalHeader().setDefaultSectionSize(_FILE_STATUS_ROW_HEIGHT)
        table.verticalHeader().setMinimumSectionSize(_FILE_STATUS_ROW_HEIGHT)
        table.verticalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Fixed)
        table.setSortingEnabled(False)
        table.setMinimumHeight(120)
        table.setSizePolicy(
            QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding
        )
        header = table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(2, QHeaderView.ResizeMode.Stretch)
        header.setStretchLastSection(True)
        table.setStyleSheet(_FILE_STATUS_TABLE_STYLE)
        self._file_status_table = table
        v.addWidget(table, stretch=1)
        return box

    def _build_actions_group(self, parent: QWidget) -> QGroupBox:
        box = shrink_h(QGroupBox("Запуск и отчёты", parent))
        box.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Maximum)
        v = QVBoxLayout(box)
        v.setContentsMargins(10, 13, 10, 10)
        v.setSpacing(8)

        btn_baseline = QPushButton("Собрать вход Только ДС", box)
        btn_baseline.setMinimumWidth(0)
        btn_baseline.setToolTip("Аудит папки ДС → Свод ДС для запуска.xlsx")
        btn_baseline.clicked.connect(self._click_ds_baseline)
        v.addWidget(btn_baseline)

        btn_hybrid = QPushButton("Проверить RFP и наложить", box)
        btn_hybrid.setMinimumWidth(0)
        btn_hybrid.setToolTip("Baseline + overlay корневого RFP_Зиновьев")
        btn_hybrid.clicked.connect(self._click_ds_hybrid)
        v.addWidget(btn_hybrid)

        btn_cov = QPushButton("Только покрытие", box)
        btn_cov.setMinimumWidth(0)
        btn_cov.setToolTip("Имена файлов и папок, без разбора qty")
        btn_cov.clicked.connect(self._click_ds_coverage)
        v.addWidget(btn_cov)

        btn_run = QPushButton(
            "Проверить части RFP и сформировать отчёт", box
        )
        btn_run.setMinimumWidth(0)
        btn_run.setToolTip("python -X utf8 -m RFQ.rfp_parts --out-dir …")
        btn_run.clicked.connect(self._click_run)
        v.addWidget(btn_run)

        btn_open_report = QPushButton("Открыть rfp_parts_report.txt", box)
        btn_open_report.setMinimumWidth(0)
        btn_open_report.clicked.connect(self._open_report_txt)
        v.addWidget(btn_open_report)

        btn_open_net = QPushButton("Открыть свод частей (xlsx)", box)
        btn_open_net.setMinimumWidth(0)
        btn_open_net.clicked.connect(self._open_net_xlsx)
        v.addWidget(btn_open_net)

        btn_open_ds = QPushButton("Открыть Свод ДС для запуска.xlsx", box)
        btn_open_ds.setMinimumWidth(0)
        btn_open_ds.clicked.connect(self._open_baseline_xlsx)
        v.addWidget(btn_open_ds)

        btn_open_hybrid = QPushButton("Открыть Свод ДС-RFP для запуска.xlsx", box)
        btn_open_hybrid.setMinimumWidth(0)
        btn_open_hybrid.clicked.connect(self._open_hybrid_xlsx)
        v.addWidget(btn_open_hybrid)

        btn_open_hybrid_report = QPushButton("Открыть отчёт сверки ДС-RFP", box)
        btn_open_hybrid_report.setMinimumWidth(0)
        btn_open_hybrid_report.clicked.connect(self._open_hybrid_report_xlsx)
        v.addWidget(btn_open_hybrid_report)

        btn_open_collisions = QPushButton("Открыть коллизии частей (xlsx)", box)
        btn_open_collisions.setMinimumWidth(0)
        btn_open_collisions.clicked.connect(self._open_collisions_xlsx)
        v.addWidget(btn_open_collisions)

        btn_open = QPushButton("Открыть последнюю папку отчётов", box)
        btn_open.setMinimumWidth(0)
        btn_open.clicked.connect(self._open_reports)
        v.addWidget(btn_open)

        artifacts = WrappingLabel(
            f"ДС/hybrid: {BASELINE_XLSX_NAME}, {HYBRID_XLSX_NAME}, "
            f"{HYBRID_REPORT_PREFIX}_<штамп>.xlsx. Части: {NET_XLSX_NAME}, "
            f"{COLLISIONS_XLSX_NAME}, {DUPLICATE_TAGS_XLSX_NAME}. "
            "Канон реестра робот не перезаписывает.",
            box,
        )
        artifacts.setStyleSheet("color: #666; font-size: 10px;")
        v.addWidget(artifacts)
        return box

    def _persist_ds_paths(self) -> None:
        cfg = load_ds_compare_config()
        gui = normalize_gui_paths(cfg.get("gui_paths"))
        gui["last_ds_trusted_folder"] = self._edit_ds_folder.text().strip()
        registry = self._edit_registry.text().strip() or str(DEFAULT_REGISTRY_PATH)
        gui["last_ds_registry_file"] = registry
        cfg["gui_paths"] = gui
        save_ds_compare_config(cfg)

    def _browse_ds_folder(self) -> None:
        start = self._edit_ds_folder.text().strip() or str(Path.home())
        path = QFileDialog.getExistingDirectory(self, "Папка доверенных ДС", start)
        if path:
            self._edit_ds_folder.setText(path)
            self._persist_ds_paths()
            self._refresh_paths()

    def _browse_registry(self) -> None:
        start = self._edit_registry.text().strip()
        start_dir = str(Path(start).parent) if start else ""
        path, _ = QFileDialog.getOpenFileName(
            self,
            "Реестр ДС",
            start_dir,
            "Excel (*.xlsx *.xlsm);;All files (*.*)",
        )
        if path:
            self._edit_registry.setText(path)
            self._persist_ds_paths()
            self._refresh_paths()

    def _on_refresh_paths_clicked(self) -> None:
        self.reload_paths_from_config()
        self._refresh_file_status()

    def _refresh_paths(self) -> None:
        latest = find_latest_rfp_parts_reports_dir()
        gui = normalize_gui_paths(load_ds_compare_config().get("gui_paths"))
        ul_text = str(gui.get("last_tsd_packing_folder", "")).strip()
        registry_text = self._edit_registry.text().strip() or str(DEFAULT_REGISTRY_PATH)
        ds_text = self._edit_ds_folder.text().strip()
        lines = [
            html.escape("Источники ДС/hybrid и частей RFP:"),
        ]
        if ds_text:
            lines.append(_path_status_line_html("доверенные ДС", Path(ds_text)))
        else:
            lines.append(html.escape("[NO] доверенные ДС: —"))
        lines.append(_path_status_line_html("реестр", Path(registry_text)))
        lines.append(_path_status_line_html("RFP_Зиновьев", DEFAULT_PARTS_DIR))
        if ul_text:
            lines.append(_path_status_line_html("УЛ (ТСД)", Path(ul_text)))
        else:
            lines.append(html.escape("[NO] УЛ (ТСД): —"))
        lines.extend(
            [
                html.escape("Своды ДС/hybrid для Запуска:"),
                _path_status_line_html("_ds_baseline", ds_baseline_output_dir()),
                _path_status_line_html("_ds_hybrid", ds_hybrid_output_dir()),
                html.escape("Выход отчётов частей (штамп YYYY.MM.DD_HH.MM):"),
                _path_status_line_html("база", rfp_parts_reports_base_dir()),
            ]
        )
        if latest is not None:
            lines.append(_path_status_line_html("последняя папка", latest))
        self._paths_label.setHtmlText("<br>".join(lines))

    def _effective_reports_dir(self) -> Path:
        if self._last_reports_dir and _safe_is_dir(self._last_reports_dir):
            return self._last_reports_dir
        return rfp_parts_reports_dir()

    def _report_txt_path(self) -> Path | None:
        reports = self._effective_reports_dir()
        path = reports / "rfp_parts_report.txt"
        return path if _safe_is_file(path) else None

    def _refresh_status(self) -> None:
        reports = self._effective_reports_dir()
        report_txt = reports / "rfp_parts_report.txt"
        net = reports / NET_XLSX_NAME
        collisions = reports / COLLISIONS_XLSX_NAME
        baseline = ds_baseline_output_dir() / BASELINE_XLSX_NAME
        if not _safe_is_file(baseline):
            baseline = reports / BASELINE_XLSX_NAME
        hybrid = ds_hybrid_output_dir() / HYBRID_XLSX_NAME
        if not _safe_is_file(hybrid):
            hybrid = reports / HYBRID_XLSX_NAME
        hybrid_report = find_latest_hybrid_report(
            ds_hybrid_output_dir()
        ) or find_latest_hybrid_report(reports)
        bits: list[str] = []
        if _safe_is_file(baseline):
            bits.append(f"Свод ДС: {_path_link_html(baseline)}")
        if _safe_is_file(hybrid):
            bits.append(f"Гибрид: {_path_link_html(hybrid)}")
        if hybrid_report is not None and _safe_is_file(hybrid_report):
            bits.append(f"Сверка: {_path_link_html(hybrid_report)}")
        if _safe_is_file(report_txt):
            bits.append(f"Отчёт частей: {_path_link_html(report_txt)}")
        if _safe_is_file(net):
            bits.append(f"Свод частей: {_path_link_html(net)}")
        if _safe_is_file(collisions):
            bits.append(f"Коллизии: {_path_link_html(collisions)}")
        if bits:
            self._status.setHtmlText("<br>".join(bits))
        elif reports != rfp_parts_reports_base_dir():
            self._status.setHtmlText(
                f"Папка отчётов: {_path_link_html(reports)}"
            )
        else:
            self._status.setHtmlText(
                "Отчётов ещё нет. Будут в "
                f"{_path_link_html(rfp_parts_reports_base_dir())}"
                "\\YYYY.MM.DD_HH.MM\\"
            )

    def _refresh_file_status(self) -> None:
        """Fill the left-pane file table from the latest stamp folder."""
        table = getattr(self, "_file_status_table", None)
        summary = getattr(self, "_file_status_summary", None)
        link = getattr(self, "_tag_remarks_link", None)
        if table is None or summary is None:
            return
        run_dir = self._effective_reports_dir()
        payload = load_file_status_payload(run_dir)
        files = payload.get("files") or []
        ok_n = int(payload.get("ok") or 0)
        error_n = int(payload.get("error") or 0)
        tag_n = int(payload.get("tag_remarks") or 0)
        if not files:
            summary.setText("Нет данных последнего прогона.")
        else:
            summary.setText(
                f"Всего {len(files)}: ОК {ok_n} · ошибки {error_n}"
            )
        self._refresh_tag_remarks_link(run_dir, tag_n, link)
        table.setRowCount(0)
        table.setRowCount(len(files))
        v_center = Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter
        for row_idx, item in enumerate(files):
            status = str(item.get("status") or STATUS_OK)
            status_text = str(item.get("status_label") or "ОК")
            detail = str(item.get("detail") or "")
            if status == STATUS_OK and not detail:
                detail = "ОК"
            fg, bg = _FILE_STATUS_COLORS.get(
                status, _FILE_STATUS_COLORS[STATUS_OK]
            )
            values = (
                str(item.get("file_name") or ""),
                status_text,
                detail,
            )
            for col, text in enumerate(values):
                cell = QTableWidgetItem(text)
                cell.setForeground(QBrush(fg))
                cell.setBackground(QBrush(bg))
                cell.setToolTip(text)
                cell.setFlags(
                    Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsSelectable
                )
                cell.setTextAlignment(v_center)
                table.setItem(row_idx, col, cell)
            table.setRowHeight(row_idx, _FILE_STATUS_ROW_HEIGHT)

    def _refresh_tag_remarks_link(
        self,
        run_dir: Path,
        tag_n: int,
        link: QLabel | None,
    ) -> None:
        if link is None:
            return
        path = resolve_duplicate_tags_xlsx(run_dir)
        if path is None:
            link.clear()
            link.setVisible(False)
            return
        count_txt = f" ({tag_n})" if tag_n else ""
        link.setText(
            f'Дубли тегов{count_txt}: '
            f'<a href="{_file_href(path)}">{html.escape(DUPLICATE_TAGS_XLSX_NAME)}</a>'
        )
        link.setVisible(True)

    def _fill_table(self, table: QTableWidget, rows: list[CockpitRow]) -> None:
        table.setRowCount(0)
        table.setRowCount(len(rows))
        v_center = Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter
        for row_idx, row in enumerate(rows):
            fg, bg = _TONE_COLORS.get(row.tone, _TONE_COLORS["ok"])
            for col, text in enumerate(row.cells):
                cell = QTableWidgetItem(text)
                cell.setForeground(QBrush(fg))
                cell.setBackground(QBrush(bg))
                cell.setToolTip(text)
                cell.setFlags(
                    Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsSelectable
                )
                cell.setTextAlignment(v_center)
                table.setItem(row_idx, col, cell)
            table.setRowHeight(row_idx, _FILE_STATUS_ROW_HEIGHT)

    def fill_from_snapshot(self, snapshot: DsCockpitSnapshot) -> None:
        """Populate banner, counters and cockpit tables from a job snapshot.

        Args:
            snapshot: Last DS/hybrid/coverage/registry job outcome.
        """
        banner = (
            f"{snapshot.summary}\n"
            f"Формат: {snapshot.registry_format} v{snapshot.format_version}; "
            f"активных {snapshot.active_count}; групп {snapshot.group_count}."
        )
        self._registry_banner.setText(banner)
        self._registry_banner.setStyleSheet(
            _BANNER_STYLES.get(snapshot.banner_tone, _BANNER_STYLES["ok"])
        )
        err = f'<span style="color:#b42318">ERROR={snapshot.error_count}</span>'
        warn = f'<span style="color:#9a6700">WARN={snapshot.warn_count}</span>'
        match = f'<span style="color:#1a7f37">MATCH={snapshot.match_count}</span>'
        self._cockpit_summary.setText(
            f"{err} · {warn} · OVERLAY={snapshot.overlay_count} · {match} · "
            f"MISMATCH={snapshot.mismatch_count} · DS_ONLY={snapshot.ds_only_count} · "
            f"RFP_ONLY={snapshot.rfp_only_count} · BLOCKED={snapshot.blocked_count}<br>"
            f"без кода={snapshot.empty_code} · qty={snapshot.qty_errors} · "
            f"вне Google={snapshot.unknown_google} · теги={snapshot.tag_warnings} · "
            f"дубли={snapshot.duplicates} · файлов ДС={snapshot.file_count} · "
            f"RFP={snapshot.rfp_file_count}"
        )
        self._fill_table(self._table_registry, snapshot.registry_rows)
        self._fill_table(self._table_groups, snapshot.group_rows)
        self._fill_table(self._table_files, snapshot.file_rows)
        self._fill_table(self._table_coverage, snapshot.coverage_rows)
        if snapshot.output_dir is not None:
            self._last_reports_dir = snapshot.output_dir
        self._refresh_status()

    def _on_tag_remarks_link(self, href: str) -> None:
        local = QUrl(href).toLocalFile() or href
        self._open_path_link(local)

    def _open_path_link(self, path: str) -> None:
        """Open a file in its app, or a folder in Explorer."""
        try:
            target = Path(path)
            try:
                is_file = target.is_file()
            except OSError:
                is_file = False
            if is_file:
                os.startfile(path)  # type: ignore[attr-defined]
            else:
                open_dir(path)
        except Exception as exc:
            QMessageBox.warning(
                self,
                "Путь",
                f"Не удалось открыть в проводнике:\n{path}\n\n{exc}",
            )

    def _click_run(self) -> None:
        if self._on_run is None:
            return
        self._on_run()

    def _require_ds_folder(self) -> bool:
        path = self._edit_ds_folder.text().strip()
        if not path:
            QMessageBox.warning(
                self,
                "Папка ДС",
                "Укажите папку доверенных ДС или выберите через «Обзор папки ДС…».",
            )
            return False
        if not _safe_is_dir(Path(path)):
            QMessageBox.warning(self, "Папка ДС", f"Папка не найдена:\n{path}")
            return False
        return True

    def _click_ds_baseline(self) -> None:
        if not self._require_ds_folder():
            return
        self._persist_ds_paths()
        if self._on_ds_baseline is not None:
            self._on_ds_baseline()

    def _click_ds_hybrid(self) -> None:
        if not self._require_ds_folder():
            return
        self._persist_ds_paths()
        if self._on_ds_hybrid is not None:
            self._on_ds_hybrid()

    def _click_ds_coverage(self) -> None:
        self._persist_ds_paths()
        if self._on_ds_coverage is not None:
            self._on_ds_coverage()

    def _click_ds_registry(self) -> None:
        self._persist_ds_paths()
        if self._on_ds_registry is not None:
            self._on_ds_registry()

    def _open_registry(self) -> None:
        path = Path(
            self._edit_registry.text().strip() or str(DEFAULT_REGISTRY_PATH)
        )
        if not _safe_is_file(path):
            QMessageBox.information(
                self, "Реестр ДС", f"Файл не найден:\n{path}"
            )
            return
        self._open_path_link(str(path))

    def _open_reports(self) -> None:
        path = self._effective_reports_dir()
        path.mkdir(parents=True, exist_ok=True)
        open_dir(str(path))

    def _open_report_txt(self) -> None:
        path = self._report_txt_path()
        if path is None:
            QMessageBox.information(
                self,
                "Отчёт",
                "Файл rfp_parts_report.txt не найден.\n"
                "Сначала выполните «Проверить части RFP и сформировать отчёт».",
            )
            return
        try:
            os.startfile(path)  # type: ignore[attr-defined]
        except Exception as exc:
            QMessageBox.warning(self, "Отчёт", f"Не удалось открыть:\n{exc}")

    def _open_xlsx_artifact(
        self,
        *,
        title: str,
        preferred: Path | None,
        resolve_latest: Callable[[], Path | None],
        missing_hint: str,
    ) -> None:
        path = preferred if preferred and _safe_is_file(preferred) else resolve_latest()
        if path is None or not _safe_is_file(path):
            QMessageBox.information(self, title, missing_hint)
            return
        try:
            os.startfile(path)  # type: ignore[attr-defined]
        except Exception as exc:
            QMessageBox.warning(self, title, f"Не удалось открыть:\n{exc}")

    def _open_named_in_reports(self, title: str, name: str, hint: str) -> None:
        reports = self._effective_reports_dir()
        self._open_xlsx_artifact(
            title=title,
            preferred=reports / name,
            resolve_latest=lambda: (
                reports / name if _safe_is_file(reports / name) else None
            ),
            missing_hint=hint,
        )

    def _open_net_xlsx(self) -> None:
        reports = self._effective_reports_dir()
        self._open_xlsx_artifact(
            title="Свод RFP",
            preferred=reports / NET_XLSX_NAME,
            resolve_latest=resolve_latest_rfp_parts_net_xlsx,
            missing_hint=(
                f"Файл {NET_XLSX_NAME} не найден.\n"
                "Сначала выполните «Проверить части RFP и сформировать отчёт»."
            ),
        )

    def _open_collisions_xlsx(self) -> None:
        reports = self._effective_reports_dir()
        self._open_xlsx_artifact(
            title="Коллизии",
            preferred=reports / COLLISIONS_XLSX_NAME,
            resolve_latest=resolve_latest_rfp_parts_collisions_xlsx,
            missing_hint=(
                f"Файл {COLLISIONS_XLSX_NAME} не найден.\n"
                "Сначала выполните «Проверить части RFP и сформировать отчёт»."
            ),
        )

    def _open_baseline_xlsx(self) -> None:
        preferred = ds_baseline_output_dir() / BASELINE_XLSX_NAME
        fallback = self._effective_reports_dir() / BASELINE_XLSX_NAME
        self._open_xlsx_artifact(
            title="Свод ДС",
            preferred=preferred,
            resolve_latest=lambda: (
                preferred
                if _safe_is_file(preferred)
                else (fallback if _safe_is_file(fallback) else None)
            ),
            missing_hint=(
                f"Файл {BASELINE_XLSX_NAME} не найден.\n"
                "Сначала выполните «Собрать вход Только ДС»."
            ),
        )

    def _open_hybrid_xlsx(self) -> None:
        preferred = ds_hybrid_output_dir() / HYBRID_XLSX_NAME
        fallback = self._effective_reports_dir() / HYBRID_XLSX_NAME
        self._open_xlsx_artifact(
            title="Гибрид ДС-RFP",
            preferred=preferred,
            resolve_latest=lambda: (
                preferred
                if _safe_is_file(preferred)
                else (fallback if _safe_is_file(fallback) else None)
            ),
            missing_hint=(
                f"Файл {HYBRID_XLSX_NAME} не найден.\n"
                "Сначала выполните «Проверить RFP и наложить»."
            ),
        )

    def _open_hybrid_report_xlsx(self) -> None:
        preferred = find_latest_hybrid_report(ds_hybrid_output_dir())
        fallback = find_latest_hybrid_report(self._effective_reports_dir())
        found = preferred or fallback
        self._open_xlsx_artifact(
            title="Отчёт сверки",
            preferred=found,
            resolve_latest=lambda: preferred or fallback,
            missing_hint=(
                f"Файл {HYBRID_REPORT_PREFIX}_<штамп>.xlsx не найден.\n"
                "Сначала выполните «Проверить RFP и наложить»."
            ),
        )

    def on_job_finished(self, success: bool, message: str) -> None:
        """Refresh status line after legacy parts diagnostics finish."""
        latest = find_latest_rfp_parts_reports_dir()
        if latest is not None:
            self._last_reports_dir = latest
        self._refresh_status()
        self._refresh_paths()
        self._refresh_file_status()
        if message:
            prefix = "OK. " if success else "Ошибка. "
            report = self._report_txt_path()
            bits = [html.escape(prefix + message)]
            if report is not None:
                bits.append(_path_link_html(report))
            net = resolve_latest_rfp_parts_net_xlsx()
            if net is not None:
                bits.append(f"Свод: {_path_link_html(net)}")
            collisions = resolve_latest_rfp_parts_collisions_xlsx()
            if collisions is not None:
                bits.append(f"Коллизии: {_path_link_html(collisions)}")
            self._status.setHtmlText("<br>".join(bits))

    def on_ds_job_finished(self, success: bool, message: str, result_path: str | None) -> None:
        """Refresh cockpit tables after a FunctionJobRunner DS job."""
        snapshot = get_last_ds_cockpit()
        if snapshot is not None:
            self.fill_from_snapshot(snapshot)
        self._refresh_paths()
        prefix = "OK. " if success else "Ошибка. "
        bits = [html.escape(prefix + (message or ""))]
        if result_path:
            bits.append(_path_link_html(Path(result_path)))
        if snapshot is not None:
            if snapshot.baseline_path is not None:
                bits.append(f"Свод ДС: {_path_link_html(snapshot.baseline_path)}")
            if snapshot.hybrid_path is not None:
                bits.append(f"Гибрид: {_path_link_html(snapshot.hybrid_path)}")
            if snapshot.report_path is not None:
                bits.append(f"Сверка: {_path_link_html(snapshot.report_path)}")
        self._status.setHtmlText("<br>".join(bits))
