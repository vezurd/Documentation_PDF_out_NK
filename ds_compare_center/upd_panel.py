"""UPD upload-files tab for ``ds_compare_center``."""

from __future__ import annotations

import os
from collections.abc import Callable
from pathlib import Path

from PySide6.QtCore import Qt
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
    QVBoxLayout,
    QWidget,
)

from RFQ.ds_compare.ds_compare_config import (
    load_ds_compare_config,
    normalize_gui_paths,
    save_ds_compare_config,
)
from RFQ.ds_compare.upd_load import (
    DEFAULT_UPD_OUTPUT_DIR,
    DEFAULT_UPD_ROOT,
    find_latest_upd_summary,
    get_last_upd_load_result,
    upd_output_dir,
)
from main_v2.job_monitor import JobMonitorPanel
from utils.path import open_dir

_LEFT_WIDTH = 520
_CONSOLE_WIDTH = 770
_CONSOLE_MIN_WIDTH = 420
_LEFT_MIN_WIDTH = 240


def _shrink_h(widget: QWidget) -> QWidget:
    """Allow widget to follow a narrow splitter pane (ignore content-based width)."""
    widget.setMinimumWidth(0)
    policy = widget.sizePolicy()
    policy.setHorizontalPolicy(QSizePolicy.Policy.Expanding)
    widget.setSizePolicy(policy)
    return widget


class UpdPanel(QWidget):
    """Source folder + merge 1C UPD upload xlsx into a summary workbook."""

    def __init__(
        self,
        parent: QWidget | None = None,
        *,
        on_run: Callable[[str], None] | None = None,
    ) -> None:
        super().__init__(parent)
        self._on_run = on_run

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
        left.setSizePolicy(
            QSizePolicy.Policy.Ignored, QSizePolicy.Policy.Expanding
        )
        left_layout = QVBoxLayout(left)
        left_layout.setContentsMargins(0, 0, 4, 0)
        left_layout.addWidget(self._build_source_group(), stretch=0)
        left_layout.addWidget(self._build_actions_group(), stretch=0)
        left_layout.addWidget(self._build_stats_group(), stretch=1)
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
        box = _shrink_h(QGroupBox("Источник УПД (файлы закачки 1С)", self))
        box.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Maximum)
        v = QVBoxLayout(box)
        v.setContentsMargins(10, 13, 10, 10)
        hint = QLabel(
            "Рекурсивный обход xlsx. Читаются вкладки формата закачки "
            "«Реализация 2_0» (compact и wide). Если в книге есть Лист1 / "
            "Лист_1 / общая — остальные вкладки (6600, спец40, …) не "
            "дублируются. Свод: "
            + str(DEFAULT_UPD_OUTPUT_DIR),
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

        self._btn_run = QPushButton("Прочитать и сформировать свод", box)
        self._btn_run.setMinimumWidth(0)
        self._btn_run.clicked.connect(self._run_load)
        self._btn_open_summary = QPushButton("Открыть свод", box)
        self._btn_open_summary.setMinimumWidth(0)
        self._btn_open_summary.clicked.connect(self._open_summary)
        self._btn_open_cache = QPushButton("Открыть папку свода", box)
        self._btn_open_cache.setMinimumWidth(0)
        self._btn_open_cache.clicked.connect(self._open_cache_folder)
        v.addWidget(self._btn_run)
        v.addWidget(self._btn_open_summary)
        v.addWidget(self._btn_open_cache)

        self._status = QLabel("", box)
        self._status.setWordWrap(True)
        self._status.setStyleSheet("color: #333; font-size: 11px;")
        _shrink_h(self._status)
        v.addWidget(self._status)
        return box

    def _build_stats_group(self) -> QGroupBox:
        box = _shrink_h(QGroupBox("Краткая статистика", self))
        box.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        v = QVBoxLayout(box)
        v.setContentsMargins(10, 13, 10, 10)
        self._stats = QLabel(
            "После прогона здесь будут файлы, position_row, compact/wide "
            "и замечания по формату.",
            box,
        )
        self._stats.setWordWrap(True)
        self._stats.setAlignment(
            Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignLeft
        )
        self._stats.setStyleSheet("color: #222; font-size: 12px;")
        _shrink_h(self._stats)
        v.addWidget(self._stats, stretch=1)
        return box

    def reload_paths_from_config(self) -> None:
        cfg = load_ds_compare_config()
        gui = normalize_gui_paths(cfg.get("gui_paths"))
        self._edit_root.setText(
            gui.get("last_upd_folder") or DEFAULT_UPD_ROOT
        )

    def set_status(self, text: str) -> None:
        self._status.setText(text)

    def _persist_root(self, path: str) -> None:
        cfg = load_ds_compare_config()
        gui = normalize_gui_paths(cfg.get("gui_paths"))
        gui["last_upd_folder"] = path
        cfg["gui_paths"] = gui
        save_ds_compare_config(cfg)

    def _browse_root(self) -> None:
        start = self._edit_root.text().strip() or str(Path.home())
        path = QFileDialog.getExistingDirectory(self, "Папка УПД", start)
        if path:
            self._edit_root.setText(path)
            self._persist_root(path)

    def _run_load(self) -> None:
        path = self._edit_root.text().strip()
        if not path:
            QMessageBox.warning(
                self,
                "УПД",
                "Укажите папку с файлами закачки УПД или выберите через «Обзор…».",
            )
            return
        if not os.path.isdir(path):
            QMessageBox.warning(self, "УПД", f"Папка не найдена:\n{path}")
            return
        self._persist_root(path)
        if self._on_run is None:
            QMessageBox.warning(self, "УПД", "Обработчик запуска не задан.")
            return
        self._on_run(path)

    def _cache_dir_path(self) -> Path:
        return upd_output_dir()

    def _refresh_status(self) -> None:
        cache_dir = self._cache_dir_path()
        summary = find_latest_upd_summary(cache_dir)
        if summary is not None:
            self._status.setText(f"Свод: {summary}")
            self._btn_open_cache.setEnabled(True)
            self._btn_open_summary.setEnabled(True)
        elif cache_dir.is_dir():
            self._status.setText(f"Папка свода: {cache_dir} (файл свода ещё не создан)")
            self._btn_open_cache.setEnabled(True)
            self._btn_open_summary.setEnabled(False)
        else:
            self._status.setText("Свод ещё не создавался.")
            self._btn_open_cache.setEnabled(False)
            self._btn_open_summary.setEnabled(False)
        self._load_stats_from_last_run()

    def _load_stats_from_last_run(self) -> None:
        result = get_last_upd_load_result()
        if result is None:
            return
        self._set_stats_from_result(result)

    def _set_stats_from_result(self, result) -> None:
        stats = result.stats
        ok = result.success and not stats.failures
        style = (
            "color: #0a5; font-size: 12px;"
            if ok
            else "color: #a30; font-size: 12px;"
        )
        self._stats.setStyleSheet(style)
        lines = [
            stats.format_short(),
            f"листов загружено={stats.sheets_loaded}, пропущено (дубли титулов)={stats.sheets_skipped}",
            f"без позиций={stats.files_without_position}, empty_row={stats.empty_rows}, other_row={stats.other_rows}",
            f"свод: {result.summary_path}",
        ]
        if stats.failures:
            lines.append("замечания:")
            lines.extend(stats.failures[:12])
            if len(stats.failures) > 12:
                lines.append(f"… ещё {len(stats.failures) - 12}, см. {result.report_path}")
        self._stats.setText("\n".join(lines))

    def _open_summary(self) -> None:
        summary = find_latest_upd_summary(self._cache_dir_path())
        if summary is None or not summary.is_file():
            QMessageBox.information(self, "УПД", "Сводный xlsx ещё не создан.")
            return
        try:
            os.startfile(str(summary))  # type: ignore[attr-defined]
        except Exception as exc:
            QMessageBox.warning(self, "УПД", f"Не удалось открыть:\n{summary}\n\n{exc}")

    def _open_cache_folder(self) -> None:
        cache_dir = self._cache_dir_path()
        cache_dir.mkdir(parents=True, exist_ok=True)
        open_dir(str(cache_dir))

    def on_job_finished(
        self, success: bool, message: str, result_path: str | None
    ) -> None:
        """Update status after an UPD merge job finishes."""
        self._status.setText(message or ("Готово." if success else "Ошибка."))
        self._refresh_status()
        result = get_last_upd_load_result()
        if result is not None:
            self._set_stats_from_result(result)
        elif not success:
            self._stats.setStyleSheet("color: #a30; font-size: 12px;")
            self._stats.setText(message or "Ошибка прогона — смотрите лог справа.")
        if result_path and message:
            self._status.setText(message)
