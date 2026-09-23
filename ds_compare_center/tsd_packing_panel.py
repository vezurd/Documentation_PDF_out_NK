"""TSD packing-lists tab for ``ds_compare_center``."""

from __future__ import annotations

import html
import os
import re
from collections.abc import Callable
from pathlib import Path

from PySide6.QtCore import Qt, QUrl
from PySide6.QtWidgets import (
    QFileDialog,
    QGroupBox,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QScrollArea,
    QSizePolicy,
    QSplitter,
    QTextBrowser,
    QVBoxLayout,
    QWidget,
)

from RFQ.ds_compare.ds_compare_config import (
    load_ds_compare_config,
    normalize_gui_paths,
    save_ds_compare_config,
)
from RFQ.ds_compare.tsd_packing_load import (
    TSD_PACKING_CRITICAL_REPORT_NAME,
    find_latest_tsd_packing_summary,
    get_last_tsd_critical_report,
    tsd_packing_cache_dir,
)
from RFQ.ds_compare.tsd_zinoviev_compare import DEFAULT_ZINOVIEV_XLSX
from main_v2.job_monitor import JobMonitorPanel
from utils.path import open_dir

# Match Run tab: left controls + full-height console on the right.
_LEFT_WIDTH = 520
_CONSOLE_WIDTH = 770
_CONSOLE_MIN_WIDTH = 420
_LEFT_MIN_WIDTH = 240

_CRITICAL_BROWSER_STYLE = (
    "QTextBrowser { font-size: 11px; color: #222; }"
    "QTextBrowser a { color: #0645ad; text-decoration: underline; }"
)

_XLSX_SUFFIXES = (".xlsx", ".xls", ".xlsm")

# PackingIssue location: [rel\file.xlsx · sheet · строка N] or [file.xlsx]
_BRACKET_XLSX_RE = re.compile(
    r"\[(?P<path>[^\[\]\n]+?\.(?:xlsx|xls|xlsm))(?=\s*·|\s*\])",
    re.IGNORECASE,
)
# Absolute/UNC path embedded in a longer remark line.
_INLINE_ABS_XLSX_RE = re.compile(
    r"(?P<path>(?:\\\\|[A-Za-z]:[\\/])[^\[\]\n]*?\.(?:xlsx|xls|xlsm))",
    re.IGNORECASE,
)
# ``Ошибка чтения файла — folder\file.xlsx`` (relative, no brackets).
_AFTER_DASH_XLSX_RE = re.compile(
    r"[—\-]\s*(?P<path>[^\[\]\n]+?\.(?:xlsx|xls|xlsm))\s*$",
    re.IGNORECASE,
)
_ROOT_LINE_RE = re.compile(r"^root:\s*(.+)\s*$", re.IGNORECASE | re.MULTILINE)


def _shrink_h(widget: QWidget) -> QWidget:
    """Allow widget to follow a narrow splitter pane (ignore content-based width)."""
    widget.setMinimumWidth(0)
    policy = widget.sizePolicy()
    policy.setHorizontalPolicy(QSizePolicy.Policy.Expanding)
    widget.setSizePolicy(policy)
    return widget


def _has_xlsx_suffix(text: str) -> bool:
    lower = text.lower()
    return any(lower.endswith(suf) for suf in _XLSX_SUFFIXES)


def _looks_like_xlsx_path(line: str) -> bool:
    """True for absolute/UNC Windows paths ending with an Excel extension."""
    text = line.strip()
    if not text or not _has_xlsx_suffix(text):
        return False
    if text.startswith("\\\\"):
        return True
    return len(text) >= 3 and text[1] == ":" and text[2] in "\\/"


def _looks_like_rel_xlsx_path(line: str) -> bool:
    """True for relative paths like ``folder\\file.xlsx`` (not absolute/UNC)."""
    text = line.strip()
    if not text or not _has_xlsx_suffix(text):
        return False
    if _looks_like_xlsx_path(text):
        return False
    return ("\\" in text or "/" in text) and "[" not in text and "]" not in text


def _extract_root_from_detail(detail: str) -> str:
    """Return ``root:`` value from a critical report file body, if present."""
    match = _ROOT_LINE_RE.search(detail)
    return match.group(1).strip() if match else ""


def _file_href(path: str) -> str:
    """``file:`` URL safe for an HTML href attribute."""
    return html.escape(QUrl.fromLocalFile(path).toString(), quote=True)


def _resolve_xlsx_path(path: str, root: str) -> str:
    """Turn a relative packing-list path into an absolute path using *root*."""
    text = path.strip().strip('"')
    if not text:
        return text
    if _looks_like_xlsx_path(text):
        return str(Path(text))
    if root:
        return str(Path(root) / text)
    return text


def _find_xlsx_path_spans(line: str) -> list[tuple[int, int, str]]:
    """Return non-overlapping ``(start, end, path)`` spans for xlsx paths in *line*."""
    candidates: list[tuple[int, int, str]] = []
    for regex in (_BRACKET_XLSX_RE, _INLINE_ABS_XLSX_RE, _AFTER_DASH_XLSX_RE):
        for match in regex.finditer(line):
            candidates.append(
                (match.start("path"), match.end("path"), match.group("path"))
            )

    stripped = line.strip()
    if _looks_like_xlsx_path(stripped) or _looks_like_rel_xlsx_path(stripped):
        start = line.find(stripped)
        if start >= 0:
            candidates.append((start, start + len(stripped), stripped))

    candidates.sort(key=lambda item: (item[0], -(item[1] - item[0])))
    spans: list[tuple[int, int, str]] = []
    occupied_end = -1
    for start, end, path in candidates:
        if start < occupied_end:
            continue
        spans.append((start, end, path))
        occupied_end = end
    return spans


def _linkify_critical_line(line: str, root: str) -> str:
    """Escape one remark line and wrap detected xlsx paths in file links."""
    stripped = line.strip()
    if stripped.lower().startswith("root:"):
        folder = stripped.split(":", 1)[1].strip()
        if folder:
            return (
                "root: "
                f'<a href="{_file_href(folder)}">{html.escape(folder)}</a>'
            )
        return html.escape(line)

    spans = _find_xlsx_path_spans(line)
    if not spans:
        return html.escape(line)

    parts: list[str] = []
    last = 0
    for start, end, path in spans:
        parts.append(html.escape(line[last:start]))
        abs_path = _resolve_xlsx_path(path, root)
        parts.append(
            f'<a href="{_file_href(abs_path)}">{html.escape(path)}</a>'
        )
        last = end
    parts.append(html.escape(line[last:]))
    return "".join(parts)


def _critical_detail_to_html(detail: str, *, root: str = "") -> str:
    """Escape remark text; turn xlsx paths (absolute or under *root*) into links."""
    if not detail.strip():
        return ""
    effective_root = _extract_root_from_detail(detail) or root.strip()
    parts = [
        _linkify_critical_line(line, effective_root) for line in detail.splitlines()
    ]
    return "<br>".join(parts)


class TsdPackingPanel(QWidget):
    """Source folder + load TSD packing lists into ``RFQ/ds_compare/cache``."""

    def __init__(
        self,
        parent: QWidget | None = None,
        *,
        on_run: Callable[[str], None] | None = None,
        on_compare_zinoviev: Callable[[], None] | None = None,
        on_check_one: Callable[[str], None] | None = None,
    ) -> None:
        super().__init__(parent)
        self._on_run = on_run
        self._on_compare_zinoviev = on_compare_zinoviev
        self._on_check_one = on_check_one

        splitter = QSplitter(Qt.Orientation.Horizontal, self)

        scroll = QScrollArea(splitter)
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QScrollArea.Shape.NoFrame)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        scroll.setMinimumWidth(_LEFT_MIN_WIDTH)
        scroll.setSizePolicy(
            QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding
        )

        left = QWidget(scroll)
        # Ignored H: always match scroll viewport width (rubber left column).
        # Preferred→Expanding V: fill viewport height so critical block can stretch.
        left.setSizePolicy(
            QSizePolicy.Policy.Ignored, QSizePolicy.Policy.Expanding
        )
        left_layout = QVBoxLayout(left)
        left_layout.setContentsMargins(0, 0, 4, 0)
        left_layout.addWidget(self._build_source_group(), stretch=0)
        left_layout.addWidget(self._build_actions_group(), stretch=0)
        left_layout.addWidget(self._build_critical_group(), stretch=1)
        scroll.setWidget(left)

        self.monitor = JobMonitorPanel(splitter, show_stop=False)
        self.monitor.setMinimumWidth(_CONSOLE_MIN_WIDTH)
        self.monitor.setSizePolicy(
            QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding
        )

        splitter.addWidget(scroll)
        splitter.addWidget(self.monitor)
        splitter.setStretchFactor(0, 1)
        splitter.setStretchFactor(1, 1)
        splitter.setSizes([_LEFT_WIDTH, _CONSOLE_WIDTH])
        splitter.setChildrenCollapsible(False)
        self.splitter = splitter

        root = QVBoxLayout(self)
        root.setContentsMargins(8, 8, 8, 8)
        root.addWidget(splitter)

        self.reload_paths_from_config()
        self._refresh_status()

    def _build_source_group(self) -> QGroupBox:
        box = _shrink_h(QGroupBox("Источник ТСД (упаковочные листы)", self))
        box.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Maximum)
        v = QVBoxLayout(box)
        v.setContentsMargins(10, 13, 10, 10)
        hint = QLabel(
            "Рекурсивный обход xlsx. Master* пропускаются; все остальные вкладки "
            "сканируются на position_row. Свод/кэш и отчёты: "
            r"\\bcc\eng\PrDoc\377_НИПИГАЗ\АГХК\КСБ\RFP_MTO_VO\_RFP\УЛ сводный файл.",
            box,
        )
        hint.setWordWrap(True)
        _shrink_h(hint)
        v.addWidget(hint)

        self._edit_root = QLineEdit(box)
        self._edit_root.setMinimumWidth(0)
        self._edit_root.setSizePolicy(
            QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed
        )
        v.addWidget(self._edit_root)

        btn_browse = QPushButton("Обзор…", box)
        btn_browse.setSizePolicy(QSizePolicy.Policy.Fixed, QSizePolicy.Policy.Fixed)
        btn_browse.clicked.connect(self._browse_root)
        v.addWidget(btn_browse, alignment=Qt.AlignmentFlag.AlignLeft)
        return box

    def _build_actions_group(self) -> QGroupBox:
        box = _shrink_h(QGroupBox("Чтение и свод", self))
        box.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Maximum)
        v = QVBoxLayout(box)
        v.setContentsMargins(10, 13, 10, 10)
        v.setSpacing(8)

        self._btn_run = QPushButton("Прочитать и проанализировать", box)
        self._btn_run.setMinimumWidth(0)
        self._btn_run.clicked.connect(self._run_load)
        self._btn_one = QPushButton("Проверить один файл УЛ", box)
        self._btn_one.setMinimumWidth(0)
        self._btn_one.setToolTip(
            "Тот же разбор вкладок и критичных замечаний, что полный прогон, "
            "только выбранный xlsx. Отчёт в _проверка_одного_файла\\УЛ. "
            "Кэш и свод полного прогона остаются на месте."
        )
        self._btn_one.clicked.connect(self._run_one_file)
        self._btn_compare_zin = QPushButton(
            "Сравнить свод робота с Зиновьевым",
            box,
        )
        self._btn_compare_zin.setMinimumWidth(0)
        self._btn_compare_zin.setToolTip(
            "Сверка кэша робота (tsd_packing_rows.cache) с ручным файлом:\n"
            + str(DEFAULT_ZINOVIEV_XLSX)
            + "\nРезультат → папка _результат_сравнения_ГГГГ_ММ_ДД_ЧЧ_ММ"
        )
        self._btn_compare_zin.clicked.connect(self._run_compare_zinoviev)
        self._btn_open_cache = QPushButton("Открыть папку свода", box)
        self._btn_open_cache.setMinimumWidth(0)
        self._btn_open_cache.clicked.connect(self._open_cache_folder)
        v.addWidget(self._btn_run)
        v.addWidget(self._btn_one)
        one_hint = QLabel(
            "Один файл — тот же читатель, что «Прочитать и проанализировать». "
            "Отчёт в папке _проверка_одного_файла\\УЛ. "
            "Кэш и свод полного прогона остаются на месте.",
            box,
        )
        one_hint.setWordWrap(True)
        one_hint.setStyleSheet("color: #555; font-size: 11px;")
        _shrink_h(one_hint)
        v.addWidget(one_hint)
        v.addWidget(self._btn_compare_zin)
        v.addWidget(self._btn_open_cache)

        self._status = QLabel("", box)
        self._status.setWordWrap(True)
        self._status.setStyleSheet("color: #333; font-size: 11px;")
        _shrink_h(self._status)
        v.addWidget(self._status)
        return box

    def _build_critical_group(self) -> QGroupBox:
        box = _shrink_h(QGroupBox("Критичные замечания по входным файлам", self))
        box.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        v = QVBoxLayout(box)
        v.setContentsMargins(10, 13, 10, 10)
        v.setSpacing(6)

        self._critical_verdict = QLabel(
            "После прогона здесь будет вердикт: OK или «есть замечания».",
            box,
        )
        self._critical_verdict.setWordWrap(True)
        self._critical_verdict.setStyleSheet("color: #333; font-size: 12px;")
        _shrink_h(self._critical_verdict)
        v.addWidget(self._critical_verdict, stretch=0)

        self._critical_text = QTextBrowser(box)
        self._critical_text.setReadOnly(True)
        self._critical_text.setOpenLinks(False)
        self._critical_text.setOpenExternalLinks(False)
        self._critical_text.setPlaceholderText(
            "Список замечаний (сдвиг столбцов, файлы без позиций, ошибки чтения)…"
        )
        self._critical_text.setMinimumHeight(120)
        self._critical_text.setMinimumWidth(0)
        self._critical_text.setStyleSheet(_CRITICAL_BROWSER_STYLE)
        self._critical_text.setSizePolicy(
            QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding
        )
        self._critical_text.anchorClicked.connect(self._on_critical_path_clicked)
        v.addWidget(self._critical_text, stretch=1)

        hint = QLabel(
            "Имена xlsx в замечаниях — кликабельные ссылки (откроется исходный файл). "
            "Если OK — детальные txt-отчёты можно не открывать. "
            "Полный список также пишется в tsd_packing_critical.txt в папке свода.",
            box,
        )
        hint.setWordWrap(True)
        hint.setStyleSheet("color: #666; font-size: 10px;")
        _shrink_h(hint)
        v.addWidget(hint, stretch=0)
        return box

    def reload_paths_from_config(self) -> None:
        cfg = load_ds_compare_config()
        gui = normalize_gui_paths(cfg.get("gui_paths"))
        self._edit_root.setText(gui.get("last_tsd_packing_folder", ""))

    def set_status(self, text: str) -> None:
        self._status.setText(text)

    def _persist_root(self, path: str) -> None:
        cfg = load_ds_compare_config()
        gui = normalize_gui_paths(cfg.get("gui_paths"))
        gui["last_tsd_packing_folder"] = path
        cfg["gui_paths"] = gui
        save_ds_compare_config(cfg)

    def _browse_root(self) -> None:
        start = self._edit_root.text().strip() or str(Path.home())
        path = QFileDialog.getExistingDirectory(self, "Папка ТСД", start)
        if path:
            self._edit_root.setText(path)
            self._persist_root(path)

    def _run_load(self) -> None:
        path = self._edit_root.text().strip()
        if not path:
            QMessageBox.warning(
                self,
                "Упаковочные листы",
                "Укажите папку-источник ТСД или выберите через «Обзор…».",
            )
            return
        if not os.path.isdir(path):
            QMessageBox.warning(
                self,
                "Упаковочные листы",
                f"Папка не найдена:\n{path}",
            )
            return
        self._persist_root(path)
        if self._on_run is None:
            QMessageBox.warning(self, "Упаковочные листы", "Обработчик запуска не задан.")
            return
        self._on_run(path)

    def _run_one_file(self) -> None:
        start = self._edit_root.text().strip() or str(Path.home())
        path, _selected = QFileDialog.getOpenFileName(
            self,
            "Один файл УЛ",
            start,
            "Excel (*.xlsx)",
        )
        if not path:
            return
        if self._on_check_one is None:
            QMessageBox.warning(self, "Упаковочные листы", "Обработчик запуска не задан.")
            return
        self._on_check_one(path)

    def _run_compare_zinoviev(self) -> None:
        if self._on_compare_zinoviev is None:
            QMessageBox.warning(
                self,
                "Сравнение с Зиновьевым",
                "Обработчик сравнения не задан.",
            )
            return
        if not DEFAULT_ZINOVIEV_XLSX.is_file():
            QMessageBox.warning(
                self,
                "Сравнение с Зиновьевым",
                f"Файл Зиновьева не найден:\n{DEFAULT_ZINOVIEV_XLSX}",
            )
            return
        self._on_compare_zinoviev()

    def _cache_dir_path(self) -> Path:
        return tsd_packing_cache_dir()

    def _refresh_status(self) -> None:
        cache_dir = self._cache_dir_path()
        summary = find_latest_tsd_packing_summary(cache_dir)
        if summary is not None:
            self._status.setText(f"Свод: {summary}")
            self._btn_open_cache.setEnabled(True)
        elif cache_dir.is_dir():
            self._status.setText(f"Папка свода: {cache_dir} (файл свода ещё не создан)")
            self._btn_open_cache.setEnabled(True)
        else:
            self._status.setText("Свод ещё не создавался.")
            self._btn_open_cache.setEnabled(False)
        self._load_critical_from_disk_or_memory()

    def _set_critical_ui(self, ok: bool, verdict: str, detail: str) -> None:
        style = (
            "color: #0a5; font-size: 12px; font-weight: bold;"
            if ok
            else "color: #a30; font-size: 12px; font-weight: bold;"
        )
        self._critical_verdict.setStyleSheet(style)
        self._critical_verdict.setText(verdict)
        root = self._edit_root.text().strip()
        html_body = _critical_detail_to_html(detail, root=root)
        if html_body:
            self._critical_text.setHtml(html_body)
        else:
            self._critical_text.clear()

    def _on_critical_path_clicked(self, url: QUrl) -> None:
        """Open linked packing-list xlsx (or folder) from a critical remark."""
        path = url.toLocalFile()
        if not path:
            path = url.toString(QUrl.UrlFormattingOption.PreferLocalFile)
        if not path:
            return
        # QUrl UNC → ``//server/share/...``; Path restores ``\\server\share\...``.
        path = str(Path(path))
        try:
            if os.path.isfile(path):
                os.startfile(path)  # type: ignore[attr-defined]
            elif os.path.isdir(path):
                open_dir(path)
            else:
                QMessageBox.warning(
                    self,
                    "Критичные замечания",
                    f"Файл или папка не найдены:\n{path}",
                )
        except Exception as exc:
            QMessageBox.warning(
                self,
                "Критичные замечания",
                f"Не удалось открыть:\n{path}\n\n{exc}",
            )

    def _load_critical_from_disk_or_memory(self) -> None:
        report = get_last_tsd_critical_report()
        if report is not None:
            self._set_critical_ui(
                report.ok, report.format_short(), report.format_detail()
            )
            return
        path = self._cache_dir_path() / TSD_PACKING_CRITICAL_REPORT_NAME
        if path.is_file():
            text = path.read_text(encoding="utf-8")
            ok = "критичных замечаний нет" in text.lower() or "(замечаний нет)" in text
            # Prefer first non-empty line as verdict.
            lines = [ln.strip() for ln in text.splitlines() if ln.strip()]
            verdict = lines[1] if len(lines) > 1 and lines[0].startswith("root:") else (
                lines[0] if lines else "Есть сохранённый отчёт замечаний."
            )
            if verdict.lower().startswith("root:"):
                verdict = lines[1] if len(lines) > 1 else verdict
            self._set_critical_ui(ok, verdict, text)
            return
        self._set_critical_ui(
            True,
            "Замечаний пока нет — сначала выполните «Прочитать и проанализировать».",
            "",
        )

    def _open_cache_folder(self) -> None:
        cache_dir = self._cache_dir_path()
        if not cache_dir.exists():
            cache_dir.mkdir(parents=True, exist_ok=True)
        open_dir(str(cache_dir))

    def on_job_finished(self, success: bool, message: str, result_path: str | None) -> None:
        """Update status after a packing-list job finishes."""
        self._status.setText(message or ("Готово." if success else "Ошибка."))
        self._refresh_status()
        report = get_last_tsd_critical_report()
        if report is not None:
            self._set_critical_ui(
                report.ok, report.format_short(), report.format_detail()
            )
        elif not success:
            self._set_critical_ui(
                False,
                "Ошибка прогона — смотрите лог справа и отчёты в папке свода.",
                message or "",
            )
        if result_path and message:
            self._status.setText(message)
