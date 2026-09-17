"""Main Qt window: tabbed shell, splitter log area, menu/toolbar."""

from __future__ import annotations

import subprocess
import sys
from collections.abc import Callable
from pathlib import Path
from typing import Any

from PySide6.QtCore import QProcess, Qt
from PySide6.QtGui import QAction, QActionGroup, QKeySequence
from PySide6.QtWidgets import (
    QApplication,
    QFileDialog,
    QGridLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QMainWindow,
    QMenuBar,
    QMessageBox,
    QPlainTextEdit,
    QPushButton,
    QSplitter,
    QStatusBar,
    QTabWidget,
    QToolBar,
    QVBoxLayout,
    QWidget,
)

try:
    from main_v2.job_monitor import JobMonitorPanel
except ImportError:  # pragma: no cover - optional panel
    JobMonitorPanel = None  # type: ignore[misc, assignment]

from main_v2.appearance import (
    apply_theme,
    available_themes,
    read_saved_theme,
    theme_label,
    write_saved_theme,
)
from main_v2.cards import WorkflowToolCard
from main_v2.function_runner import FunctionJobRunner
from main_v2.help_dialog import show_help_dialog
from main_v2.job_runner import (
    ProcessJobRunner,
    aggregate_tags_env_overlay,
    build_aggregate_tags_argv,
)
from main_v2 import legacy_actions
from GUI.gui_constants import GuiConst
from RFQ.tags_rfp_compare.rfp_tags_utils import (
    find_latest_step4_result_file,
    get_asbuild_config_path,
    load_asbuild_config,
    save_asbuild_config,
)


def project_root() -> Path:
    """Returns repository root (parent of the ``main_v2`` package)."""
    return Path(__file__).resolve().parent.parent


class MainWindow(QMainWindow):
    """Tabbed shell with a lower process/log pane and PDF monitor launcher."""

    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setWindowTitle("Documentation — main_v2 (Qt6)")
        self.resize(960, 640)
        self._root = project_root()
        self._active_job_title = ""
        self._current_theme_id = apply_theme(self, read_saved_theme())
        self._build_ui()
        self._job_runner = ProcessJobRunner(self)
        self._function_runner = FunctionJobRunner(self)
        self._wire_job_runner()
        self._wire_function_runner()
        self._build_menu_toolbar()
        status = QStatusBar()
        self.setStatusBar(status)
        status.showMessage(f"Схема: {theme_label(self._current_theme_id)}", 5000)

    def _append_log(self, line: str) -> None:
        """Appends one line to the bottom process panel."""
        text = line if line.endswith("\n") else f"{line}\n"
        panel = getattr(self, "_log_panel", None)
        if panel is not None and hasattr(panel, "append_log"):
            panel.append_log(text)
            return
        if hasattr(self, "_log") and self._log is not None:
            self._log.appendPlainText(line.rstrip("\n"))

    def _build_ui(self) -> None:
        central = QWidget()
        self.setCentralWidget(central)
        outer = QVBoxLayout(central)
        splitter = QSplitter(Qt.Orientation.Vertical)
        tabs = QTabWidget()
        tabs.addTab(self._tab_pdf_v2(), "PDF / v2")
        tabs.addTab(self._tab_mto_bbb(), "MTO / BBB")
        tabs.addTab(self._tab_comparisons(), "Сравнения")
        tabs.addTab(self._tab_rfq_asbuild(), "RFQ / as-build")
        tabs.addTab(self._tab_services(), "Сервисы")
        splitter.addWidget(tabs)

        if JobMonitorPanel is not None:
            self._log_panel: QWidget = JobMonitorPanel()
        else:
            self._log = QPlainTextEdit()
            self._log.setReadOnly(True)
            self._log.setPlaceholderText(
                "Панель процесса / лога (placeholder). "
                "При появлении main_v2.job_monitor.JobMonitorPanel она подставится сюда."
            )
            wrap = QWidget()
            wl = QVBoxLayout(wrap)
            wl.addWidget(QLabel("Процесс / мониторинг"))
            wl.addWidget(self._log)
            self._log_panel = wrap

        splitter.addWidget(self._log_panel)
        splitter.setStretchFactor(0, 3)
        splitter.setStretchFactor(1, 1)
        outer.addWidget(splitter)

    def _wire_job_runner(self) -> None:
        """Connect the shared process runner to the monitor panel."""
        panel = getattr(self, "_log_panel", None)
        if panel is not None and hasattr(panel, "stop_requested"):
            panel.stop_requested.connect(self._job_runner.request_stop)
        self._job_runner.started.connect(self._on_job_started)
        self._job_runner.log_received.connect(self._append_log)
        self._job_runner.finished.connect(self._on_job_finished)
        self._job_runner.error.connect(self._on_job_error)

    def _wire_function_runner(self) -> None:
        """Connect blocking legacy-call runner to the monitor panel."""
        self._function_runner.started.connect(self._on_function_job_started)
        self._function_runner.log_received.connect(self._append_log)
        self._function_runner.finished.connect(self._on_function_job_finished)
        self._function_runner.error.connect(self._on_function_job_error)

    def _tab_pdf_v2(self) -> QWidget:
        w = QWidget()
        lay = QVBoxLayout(w)
        lay.addWidget(
            QLabel(
                "Полный конвейер PDF v2 — отдельное окно монитора. "
                "Кнопка также в меню и на панели инструментов."
            )
        )
        row = QHBoxLayout()
        btn = QPushButton("Центр управления PDF v2")
        btn.clicked.connect(self._launch_pdf_v2_monitor)
        row.addWidget(btn)
        row.addStretch(1)
        lay.addLayout(row)
        lay.addStretch(1)
        return w

    def _tab_mto_bbb(self) -> QWidget:
        w = QWidget()
        lay = QVBoxLayout(w)
        check_box = self._button_group(
            "Проверка MTO / BBB",
            (
                (
                    "Открыть папку DWG с MTO",
                    lambda: self._run_folder_action(
                        "MTO DWG (+BBB)",
                        legacy_actions.run_mto_dwg_with_optional_bbb,
                    ),
                ),
                ("Настройки MTO / BBB", lambda: self._launch_settings_dialog("bbb")),
            ),
        )
        files_box = self._button_group(
            "Открыть входные файлы",
            (
                (
                    "Открыть файл MTO",
                    lambda: self._run_file_action(
                        "MTO file",
                        legacy_actions.run_google_file_att,
                        GuiConst.MTO_FILE,
                    ),
                ),
                (
                    "Открыть файл CO.xlsx",
                    lambda: self._run_file_action(
                        "BOOT file",
                        legacy_actions.run_google_file_att,
                        GuiConst.BOOT_FILE,
                    ),
                ),
                (
                    "Открыть файл RFQ",
                    lambda: self._run_file_action("RFQ file", legacy_actions.run_rfq_file),
                ),
                (
                    "Открыть файл output",
                    lambda: self._run_file_action(
                        "OUTPUT file",
                        legacy_actions.run_google_file_att,
                        GuiConst.OUTPUT_FILE,
                    ),
                ),
            ),
        )
        prep_box = self._button_group(
            "Подготовка Excel",
            (
                (
                    "Убрать зачёркивания MTO",
                    lambda: self._run_file_action(
                        "Убрать зачёркивания MTO",
                        legacy_actions.remove_strikethrough_mto,
                    ),
                ),
                (
                    "Убрать зачёркивания BOE/BOM/BOQ",
                    lambda: self._run_folder_action(
                        "Убрать зачёркивания BBB",
                        legacy_actions.remove_strikethrough_bbb_dir,
                    ),
                ),
                (
                    "Для 1C",
                    lambda: self._run_folder_action(
                        "Подготовка BBB для 1C",
                        legacy_actions.prepare_bbb_for_1c,
                    ),
                ),
            ),
        )
        lay.addWidget(check_box)
        lay.addWidget(files_box)
        lay.addWidget(prep_box)
        lay.addStretch(1)
        return w

    def _tab_comparisons(self) -> QWidget:
        w = QWidget()
        lay = QVBoxLayout(w)
        lay.addWidget(
            self._button_group(
                "Кабельные журналы",
                (
                    (
                        "КЖ: выбрать папку",
                        lambda: self._run_folder_action(
                            "Кабельные журналы",
                            legacy_actions.run_cj_dir,
                        ),
                    ),
                ),
            )
        )
        lay.addWidget(
            self._button_group(
                "MTO revisions",
                (
                    (
                        "MTO vs MTO (2 шт.)",
                        lambda: self._run_folder_action(
                            "MTO vs MTO",
                            legacy_actions.run_mto_compare_dir,
                        ),
                    ),
                    (
                        "MTO vs MTO (мульти)",
                        lambda: self._run_folder_action(
                            "MTO multi",
                            legacy_actions.run_mto_multi_compare_dir,
                        ),
                    ),
                    (
                        "MTO: цепочка ревизий",
                        lambda: self._run_folder_action(
                            "MTO цепочка ревизий",
                            legacy_actions.run_mto_chain_compare_dir,
                        ),
                    ),
                ),
            )
        )
        lay.addWidget(
            self._button_group(
                "ДС",
                (
                    (
                        "Объединить ДС из папки",
                        lambda: self._run_folder_action(
                            "Объединить ДС",
                            legacy_actions.run_merge_ds_dir,
                        ),
                    ),
                    (
                        "Список ДС vs MTO (файл)",
                        lambda: self._run_file_action(
                            "ДС vs MTO",
                            legacy_actions.run_ds_mto_file,
                        ),
                    ),
                    (
                        "Grouped ДС vs MTO (файл)",
                        lambda: self._run_file_action(
                            "Grouped ДС vs MTO",
                            legacy_actions.run_grouped_ds_mto_file,
                        ),
                    ),
                    (
                        "Настройки папок MTO",
                        lambda: self._launch_settings_dialog("ds"),
                    ),
                ),
            )
        )
        lay.addWidget(
            self._button_group(
                "MTO vs ДС",
                (
                    (
                        "MTO vs Список DS (файл)",
                        lambda: self._run_file_action(
                            "MTO vs ДС",
                            legacy_actions.run_mto_ds_file,
                        ),
                    ),
                ),
            )
        )
        lay.addStretch(1)
        return w

    def _tab_rfq_asbuild(self) -> QWidget:
        w = QWidget()
        lay = QVBoxLayout(w)
        rfp = WorkflowToolCard(
            "RFP / RFQ",
            on_run=self._placeholder_run_rfp,
            on_settings=self._placeholder_settings_rfp,
            on_open_last=self._placeholder_open_last_rfp,
        )
        asb = WorkflowToolCard(
            "as-build",
            on_run=self._placeholder_run_as_build,
            on_settings=self._placeholder_settings_as_build,
            on_open_last=self._placeholder_open_last_as_build,
        )
        lay.addWidget(rfp)
        lay.addWidget(asb)
        lay.addStretch(1)
        return w

    def _tab_services(self) -> QWidget:
        w = QWidget()
        lay = QVBoxLayout(w)
        lay.addWidget(
            self._button_group(
                "Сервисы",
                (
                    (
                        "Открыть NanoCAD.db",
                        lambda: self._run_file_action(
                            "NanoCAD.db",
                            legacy_actions.run_nanocad_db,
                            file_filter="Database (*.db);;All files (*.*)",
                        ),
                    ),
                    (
                        "Собрать ZIP для коллег",
                        lambda: self._start_function_job(
                            "Собрать ZIP для коллег",
                            legacy_actions.run_release_zip,
                            self._root,
                        ),
                    ),
                ),
            )
        )
        lay.addStretch(1)
        return w

    def _button_group(
        self,
        title: str,
        actions: tuple[tuple[str, Callable[[], None]], ...],
    ) -> QGroupBox:
        """Build a compact group of action buttons."""
        group = QGroupBox(title)
        grid = QGridLayout(group)
        for i, (text, callback) in enumerate(actions):
            btn = QPushButton(text)
            btn.clicked.connect(callback)
            grid.addWidget(btn, i // 2, i % 2)
        return group

    def _build_menu_toolbar(self) -> None:
        menubar: QMenuBar = self.menuBar()
        file_menu = menubar.addMenu("Файл")
        act_quit = QAction("Выход", self)
        act_quit.setShortcut(QKeySequence.StandardKey.Quit)

        def _quit_app() -> None:
            app = QApplication.instance()
            if app is not None:
                app.quit()

        act_quit.triggered.connect(_quit_app)
        file_menu.addAction(act_quit)

        view_menu = menubar.addMenu("Вид")
        theme_sub = view_menu.addMenu("Цветовая схема")
        self._theme_action_group = QActionGroup(self)
        self._theme_action_group.setExclusive(True)
        self._theme_actions: dict[str, QAction] = {}
        for entry in available_themes():
            act = QAction(entry.label, self)
            act.setCheckable(True)
            act.setData(entry.theme_id)
            self._theme_action_group.addAction(act)
            theme_sub.addAction(act)
            self._theme_actions[entry.theme_id] = act
        self._theme_action_group.blockSignals(True)
        current = self._theme_actions.get(self._current_theme_id)
        if current is not None:
            current.setChecked(True)
        self._theme_action_group.blockSignals(False)
        self._theme_action_group.triggered.connect(self._on_theme_selected)

        tools_menu = menubar.addMenu("Инструменты")
        act_monitor = QAction("Центр управления PDF v2", self)
        act_monitor.triggered.connect(self._launch_pdf_v2_monitor)
        tools_menu.addAction(act_monitor)

        help_menu = menubar.addMenu("Справка")
        act_help = QAction("Справка по цепочкам…", self)
        act_help.triggered.connect(lambda: show_help_dialog(self))
        help_menu.addAction(act_help)

        tb = QToolBar("Главная")
        tb.setMovable(True)
        self.addToolBar(tb)
        tb.addAction(act_monitor)

    def _on_theme_selected(self, action: QAction) -> None:
        """Persist and apply the color scheme chosen from the View menu."""
        raw = action.data()
        theme_id = str(raw) if raw is not None else "system"
        applied = apply_theme(self, theme_id)
        write_saved_theme(applied)
        self._current_theme_id = applied
        self._theme_action_group.blockSignals(True)
        for tid, act in self._theme_actions.items():
            act.setChecked(tid == applied)
        self._theme_action_group.blockSignals(False)
        self.statusBar().showMessage(f"Схема: {theme_label(applied)}", 5000)

    def _launch_pdf_v2_monitor(self) -> None:
        """Starts ``python -m pdf_v2_monitor`` in a separate process."""
        cmd = [sys.executable, "-m", "pdf_v2_monitor"]
        try:
            subprocess.Popen(cmd, cwd=str(self._root))
            self.statusBar().showMessage("Запущен pdf_v2_monitor", 5000)
            self._append_log(f"Started: {' '.join(cmd)} cwd={self._root}")
        except OSError as e:
            QMessageBox.warning(
                self,
                "Запуск монитора",
                f"Не удалось запустить процесс:\n{e}",
            )
            self._append_log(f"ERROR launch pdf_v2_monitor: {e}")

    def _placeholder_run_rfp(self) -> None:
        self._run_aggregate_tags_job("RFP/RFQ")

    def _placeholder_settings_rfp(self) -> None:
        self._launch_settings_dialog("rfq")

    def _placeholder_open_last_rfp(self) -> None:
        self._open_latest_step4_result()

    def _placeholder_run_as_build(self) -> None:
        self._run_aggregate_tags_job("as-build", asbuild=True)

    def _placeholder_settings_as_build(self) -> None:
        self._launch_settings_dialog("asbuild")

    def _placeholder_open_last_as_build(self) -> None:
        self._open_latest_step4_result()

    def _run_aggregate_tags_job(self, title: str, *, asbuild: bool = False) -> None:
        """Run RFQ/as-build aggregate tags script with live QProcess logging."""
        if self._function_runner.is_running():
            QMessageBox.information(self, "Задача выполняется", "Дождитесь завершения текущей задачи.")
            return
        config_path: str | None = None
        if asbuild:
            config_path = get_asbuild_config_path()
            if not Path(config_path).is_file():
                save_asbuild_config(load_asbuild_config())
                self._append_log(f"Создан конфиг as-build: {config_path}")

        panel = getattr(self, "_log_panel", None)
        if panel is not None and hasattr(panel, "start_job"):
            panel.start_job(title)
        self._active_job_title = title
        argv = build_aggregate_tags_argv(
            self._root,
            asbuild_config_path=config_path,
        )
        self._append_log(f"Started: {' '.join(argv)}")
        started = self._job_runner.start(
            argv,
            cwd=self._root,
            env=aggregate_tags_env_overlay(self._root),
        )
        if not started:
            self._finish_monitor(False, f"{title}: процесс не запущен")

    def _on_job_started(self) -> None:
        self.statusBar().showMessage(f"{self._active_job_title}: выполняется", 5000)

    def _on_job_finished(self, exit_code: int, exit_status: int) -> None:
        success = exit_status == 0 and exit_code == 0
        title = self._active_job_title or "Задача"
        result_path = find_latest_step4_result_file() if success else None
        message = f"{title}: готово" if success else f"{title}: код выхода {exit_code}"
        self._append_log(f"\n{message}")
        self._finish_monitor(success, message, result_path=result_path)
        self.statusBar().showMessage(message, 8000)

    def _on_job_error(self, message: str) -> None:
        title = self._active_job_title or "Задача"
        self._append_log(f"[error] {message}")
        self.statusBar().showMessage(f"{title}: ошибка запуска", 8000)

    def _run_file_action(
        self,
        title: str,
        fn: Callable[..., Any],
        *args: Any,
        file_filter: str = "Excel files (*.xlsx *.xlsm *.xls);;All files (*.*)",
    ) -> None:
        path, _ = QFileDialog.getOpenFileName(self, title, "", file_filter)
        if not path:
            return
        self._start_function_job(title, fn, path, *args)

    def _run_folder_action(
        self,
        title: str,
        fn: Callable[..., Any],
        *args: Any,
    ) -> None:
        path = QFileDialog.getExistingDirectory(self, title, "")
        if not path:
            return
        self._start_function_job(title, fn, path, *args)

    def _start_function_job(self, title: str, fn: Callable[..., Any], *args: Any) -> None:
        """Start a blocking legacy callable in the background worker."""
        if self._job_runner.process().state() != QProcess.ProcessState.NotRunning:
            QMessageBox.information(self, "Задача выполняется", "Дождитесь завершения текущего процесса.")
            return
        panel = getattr(self, "_log_panel", None)
        if panel is not None and hasattr(panel, "start_job"):
            panel.start_job(title)
        self._active_job_title = title
        if not self._function_runner.start(title, fn, *args):
            self._finish_monitor(False, f"{title}: задача не запущена")

    def _on_function_job_started(self) -> None:
        self.statusBar().showMessage(f"{self._active_job_title}: выполняется", 5000)

    def _on_function_job_finished(
        self,
        success: bool,
        message: str,
        result_path: object,
    ) -> None:
        path = str(result_path) if result_path is not None else None
        self._finish_monitor(success, message, result_path=path)
        self.statusBar().showMessage(message, 8000)

    def _on_function_job_error(self, message: str) -> None:
        self._append_log(f"[error] {message}")
        self.statusBar().showMessage(f"{self._active_job_title}: ошибка", 8000)

    def _launch_settings_dialog(self, dialog: str) -> None:
        """Open a legacy CTk settings dialog in a separate process."""
        cmd = [sys.executable, "-m", "main_v2.settings_launchers", dialog]
        try:
            subprocess.Popen(cmd, cwd=str(self._root))
            self._append_log(f"Started settings: {' '.join(cmd)}")
            self.statusBar().showMessage("Окно настроек запущено", 5000)
        except OSError as e:
            QMessageBox.warning(
                self,
                "Настройки",
                f"Не удалось запустить окно настроек:\n{e}",
            )

    def _finish_monitor(
        self,
        success: bool,
        message: str,
        *,
        result_path: str | None = None,
    ) -> None:
        panel = getattr(self, "_log_panel", None)
        if panel is not None and hasattr(panel, "finish_job"):
            panel.finish_job(success, message, result_path=result_path)

    def _open_latest_step4_result(self) -> None:
        """Open the latest Step4 RFQ/MTO result workbook if it exists."""
        path = find_latest_step4_result_file()
        if not path:
            QMessageBox.information(
                self,
                "Последний результат",
                "Файл Шаг4_Сопоставление_RFP_MTO_*.xlsx не найден.",
            )
            return
        try:
            if sys.platform.startswith("win"):
                # os.startfile is available only on Windows.
                import os

                os.startfile(path)  # type: ignore[attr-defined]
            else:
                subprocess.Popen([sys.executable, "-m", "webbrowser", path])
            self._append_log(f"Открыт: {path}")
        except OSError as e:
            QMessageBox.warning(
                self,
                "Последний результат",
                f"Не удалось открыть файл:\n{e}",
            )
