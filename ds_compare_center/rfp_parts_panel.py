"""RFP parts collection tab — diagnostics from ``RFQ.rfp_parts``."""

from __future__ import annotations

import html
import os
from collections.abc import Callable
from pathlib import Path

from PySide6.QtCore import Qt, QUrl
from PySide6.QtGui import QBrush, QColor
from PySide6.QtWidgets import (
    QAbstractItemView,
    QGroupBox,
    QHeaderView,
    QLabel,
    QMessageBox,
    QPushButton,
    QSizePolicy,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
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
    mark = "OK" if path.exists() else "NO"
    return (
        f"[{mark}] {html.escape(label)}: {_path_link_html(path)}"
    )


class RfpPartsPanel(QWidget):
    """Run RFP parts diagnostics; open reports."""

    def __init__(
        self,
        parent: QWidget | None = None,
        *,
        on_run: Callable[[], None] | None = None,
    ) -> None:
        """Create the panel.

        Args:
            parent: Qt parent.
            on_run: start diagnostics job (no manifest/strict flags).
        """
        super().__init__(parent)
        self._on_run = on_run
        self._last_reports_dir: Path | None = find_latest_rfp_parts_reports_dir()

        splitter, self.monitor = build_side_by_side(
            self, build_left=self._build_left, show_stop=True
        )
        root = QVBoxLayout(self)
        root.setContentsMargins(8, 8, 8, 8)
        root.addWidget(splitter)
        self._refresh_status()
        self._refresh_paths()
        self._refresh_file_status()

    def set_last_reports_dir(self, path: Path) -> None:
        """Remember the out-dir used for the current/last run."""
        self._last_reports_dir = path
        self._refresh_status()

    @property
    def last_reports_dir(self) -> Path | None:
        """Out-dir of the last run (if any)."""
        return self._last_reports_dir

    def _build_left(self, left: QWidget, left_layout: QVBoxLayout) -> None:
        left_layout.addWidget(self._build_source_group(left), stretch=0)
        left_layout.addWidget(self._build_file_status_group(left), stretch=1)
        left_layout.addWidget(self._build_actions_group(left), stretch=0)

    def _build_source_group(self, parent: QWidget) -> QGroupBox:
        box = shrink_h(QGroupBox("Сбор RFP из частей (диагностика)", parent))
        box.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Maximum)
        v = QVBoxLayout(box)
        v.setContentsMargins(10, 13, 10, 10)
        v.setSpacing(8)

        hint = WrappingLabel(
            "Обход папки частей RFP_Зиновьев и суммирование в свод. "
            "Сверки со списком ДС пока нет (реестр не подключён). "
            "Файл без прочитанных позиций — ошибка в таблице слева и в отчёте. "
            "Qty частей — блок «Закупка по Лоту» (Excel №17–18). "
            "Отчёты: …\\RFP сводный файл\\YYYY.MM.DD_HH.MM\\ "
            f"(net={NET_XLSX_NAME} со столбцами-фильтрами коллизий, "
            f"ещё {COLLISIONS_XLSX_NAME}). "
            "Это только проверка и отчёт: боевой step1 не запускается. "
            "Какой файл возьмёт «RFP · Запуск» — галка «Брать последний свод частей» "
            "во вкладке «RFP · Настройки».",
            box,
        )
        v.addWidget(hint)

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

    def _build_file_status_group(self, parent: QWidget) -> QGroupBox:
        box = shrink_h(QGroupBox("Файлы последнего прогона", parent))
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
        table.setMinimumHeight(180)
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

        btn_open_net = QPushButton("Открыть свод (xlsx)", box)
        btn_open_net.setMinimumWidth(0)
        btn_open_net.clicked.connect(self._open_net_xlsx)
        v.addWidget(btn_open_net)

        btn_open_collisions = QPushButton("Открыть коллизии (xlsx)", box)
        btn_open_collisions.setMinimumWidth(0)
        btn_open_collisions.clicked.connect(self._open_collisions_xlsx)
        v.addWidget(btn_open_collisions)

        btn_open = QPushButton("Открыть последнюю папку отчётов", box)
        btn_open.setMinimumWidth(0)
        btn_open.clicked.connect(self._open_reports)
        v.addWidget(btn_open)

        artifacts = WrappingLabel(
            f"В папке прогона: {NET_XLSX_NAME}, {COLLISIONS_XLSX_NAME}, "
            "rfp_parts_report.txt, rfp_parts_records.xlsx, diagnostics.xlsx, "
            f"{DUPLICATE_TAGS_XLSX_NAME}, rfp_parts_file_status.json. "
            "Коллизии: лист «К исправлению» — файл / лист / строка, где ед. изм. ≠ Google.",
            box,
        )
        artifacts.setStyleSheet("color: #666; font-size: 10px;")
        v.addWidget(artifacts)
        return box

    def _on_refresh_paths_clicked(self) -> None:
        self._refresh_paths()
        self._refresh_file_status()

    def _refresh_paths(self) -> None:
        latest = find_latest_rfp_parts_reports_dir()
        lines = [
            html.escape(
                "Источники (DEFAULT в RFQ.rfp_parts; пока только чтение):"
            ),
            _path_status_line_html("части", DEFAULT_PARTS_DIR),
            html.escape("Выход отчётов (штамп YYYY.MM.DD_HH.MM):"),
            _path_status_line_html("база", rfp_parts_reports_base_dir()),
        ]
        if latest is not None:
            lines.append(_path_status_line_html("последняя папка", latest))
        self._paths_label.setHtmlText("<br>".join(lines))

    def _effective_reports_dir(self) -> Path:
        if self._last_reports_dir and self._last_reports_dir.is_dir():
            return self._last_reports_dir
        return rfp_parts_reports_dir()

    def _report_txt_path(self) -> Path | None:
        reports = self._effective_reports_dir()
        path = reports / "rfp_parts_report.txt"
        return path if path.is_file() else None

    def _refresh_status(self) -> None:
        reports = self._effective_reports_dir()
        report_txt = reports / "rfp_parts_report.txt"
        net = reports / NET_XLSX_NAME
        collisions = reports / COLLISIONS_XLSX_NAME
        bits: list[str] = []
        if report_txt.is_file():
            bits.append(f"Отчёт: {_path_link_html(report_txt)}")
        if net.is_file():
            bits.append(f"Свод: {_path_link_html(net)}")
        if collisions.is_file():
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
        path = preferred if preferred and preferred.is_file() else resolve_latest()
        if path is None or not path.is_file():
            QMessageBox.information(self, title, missing_hint)
            return
        try:
            os.startfile(path)  # type: ignore[attr-defined]
        except Exception as exc:
            QMessageBox.warning(self, title, f"Не удалось открыть:\n{exc}")

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

    def on_job_finished(self, success: bool, message: str) -> None:
        """Refresh status line after diagnostics finish."""
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
