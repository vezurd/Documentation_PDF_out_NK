"""Remaining main.py workflows (MTO/BBB, comparisons, services) for the control center."""

from __future__ import annotations

from collections.abc import Callable
from typing import Any

from PySide6.QtWidgets import (
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QSizePolicy,
    QVBoxLayout,
    QWidget,
)

from GUI.gui_constants import GuiConst
from ds_compare_center.split_layout import build_side_by_side, shrink_h
from main_v2 import legacy_actions

_EXCEL_FILTER = "Excel files (*.xlsx *.xlsm *.xls);;All files (*.*)"
_DB_FILTER = "Database (*.db);;All files (*.*)"

FolderJobFn = Callable[[str, Callable[..., Any]], None]
FileJobFn = Callable[[str, Callable[..., Any], tuple[Any, ...], str], None]


class MiscRunPanel(QWidget):
    """Run leftover ``main.py`` actions (not DS/RFP) with a live Job monitor."""

    def __init__(
        self,
        parent: QWidget | None = None,
        *,
        on_launch_pdf: Callable[[], None] | None = None,
        on_launch_rd_catalog: Callable[[], None] | None = None,
        on_folder_job: FolderJobFn | None = None,
        on_file_job: FileJobFn | None = None,
        on_zip: Callable[[], None] | None = None,
        on_help: Callable[[str], None] | None = None,
    ) -> None:
        super().__init__(parent)
        self._on_launch_pdf = on_launch_pdf
        self._on_launch_rd_catalog = on_launch_rd_catalog
        self._on_folder_job = on_folder_job
        self._on_file_job = on_file_job
        self._on_zip = on_zip
        self._on_help = on_help

        splitter, self.monitor = build_side_by_side(
            self, build_left=self._build_left, show_stop=False
        )
        self.splitter = splitter
        root = QVBoxLayout(self)
        root.setContentsMargins(8, 8, 8, 8)
        root.addWidget(splitter)

    def _build_left(self, left: QWidget, left_layout: QVBoxLayout) -> None:
        left_layout.addWidget(self._build_pdf_group(left), stretch=0)
        left_layout.addWidget(self._build_rd_catalog_group(left), stretch=0)
        left_layout.addWidget(self._build_mto_bbb_group(left), stretch=0)
        left_layout.addWidget(self._build_files_group(left), stretch=0)
        left_layout.addWidget(self._build_excel_group(left), stretch=0)
        left_layout.addWidget(self._build_cj_group(left), stretch=0)
        left_layout.addWidget(self._build_revisions_group(left), stretch=0)
        left_layout.addWidget(self._build_services_group(left), stretch=0)
        left_layout.addStretch(1)

    def _build_pdf_group(self, parent: QWidget) -> QGroupBox:
        box = shrink_h(QGroupBox("PDF v2", parent))
        box.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Maximum)
        v = QVBoxLayout(box)
        v.setContentsMargins(10, 13, 10, 10)
        hint = QLabel(
            "Полный прогон папки PDF: шаблоны, извлечение, теги, ОД, НК — "
            "отдельное окно монитора.",
            box,
        )
        hint.setWordWrap(True)
        shrink_h(hint)
        v.addWidget(hint)
        v.addLayout(
            self._action_row(
                box,
                "Центр управления PDF v2",
                self._click_pdf,
                help_key="open_pdf_folder",
            )
        )
        return box

    def _build_rd_catalog_group(self, parent: QWidget) -> QGroupBox:
        box = shrink_h(QGroupBox("Каталог РД", parent))
        box.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Maximum)
        layout = QVBoxLayout(box)
        layout.setContentsMargins(10, 13, 10, 10)
        hint = QLabel(
            "Фактические версии РД и готовность актуальных MTO для робота.",
            box,
        )
        hint.setWordWrap(True)
        shrink_h(hint)
        layout.addWidget(hint)
        layout.addLayout(
            self._action_row(
                box,
                "Открыть каталог РД",
                self._click_rd_catalog,
            )
        )
        return box

    def _build_mto_bbb_group(self, parent: QWidget) -> QGroupBox:
        box = shrink_h(QGroupBox("Проверка MTO / BBB", parent))
        box.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Maximum)
        v = QVBoxLayout(box)
        v.setContentsMargins(10, 13, 10, 10)
        v.addLayout(
            self._action_row(
                box,
                GuiConst.dict[GuiConst.MTO_DIR][1],
                lambda: self._folder(
                    "MTO DWG (+BBB)",
                    legacy_actions.run_mto_dwg_with_optional_bbb,
                ),
                help_key="mto_file",
            )
        )
        return box

    def _build_files_group(self, parent: QWidget) -> QGroupBox:
        box = shrink_h(QGroupBox("Открыть входные файлы", parent))
        box.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Maximum)
        v = QVBoxLayout(box)
        v.setContentsMargins(10, 13, 10, 10)
        v.addLayout(
            self._action_row(
                box,
                GuiConst.dict[GuiConst.MTO_FILE][1],
                lambda: self._file(
                    "MTO file",
                    legacy_actions.run_google_file_att,
                    GuiConst.MTO_FILE,
                ),
                help_key="mto_file",
            )
        )
        v.addLayout(
            self._action_row(
                box,
                GuiConst.dict[GuiConst.BOOT_FILE][1],
                lambda: self._file(
                    "BOOT file",
                    legacy_actions.run_google_file_att,
                    GuiConst.BOOT_FILE,
                ),
                help_key="boot_file",
            )
        )
        v.addLayout(
            self._action_row(
                box,
                "Открыть файл RFQ",
                lambda: self._file("RFQ file", legacy_actions.run_rfq_file),
                help_key="open_rfq_folder",
            )
        )
        v.addLayout(
            self._action_row(
                box,
                GuiConst.dict[GuiConst.OUTPUT_FILE][1],
                lambda: self._file(
                    "OUTPUT file",
                    legacy_actions.run_google_file_att,
                    GuiConst.OUTPUT_FILE,
                ),
                help_key="output_file",
            )
        )
        return box

    def _build_excel_group(self, parent: QWidget) -> QGroupBox:
        box = shrink_h(QGroupBox("Подготовка Excel", parent))
        box.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Maximum)
        v = QVBoxLayout(box)
        v.setContentsMargins(10, 13, 10, 10)
        v.addLayout(
            self._action_row(
                box,
                "Убрать зачёркивания MTO",
                lambda: self._file(
                    "Убрать зачёркивания MTO",
                    legacy_actions.remove_strikethrough_mto,
                ),
            )
        )
        v.addLayout(
            self._action_row(
                box,
                "Убрать зачёркивания BOE/BOM/BOQ",
                lambda: self._folder(
                    "Убрать зачёркивания BBB",
                    legacy_actions.remove_strikethrough_bbb_dir,
                ),
            )
        )
        v.addLayout(
            self._action_row(
                box,
                "Для 1C",
                lambda: self._folder(
                    "Подготовка BBB для 1C",
                    legacy_actions.prepare_bbb_for_1c,
                ),
            )
        )
        return box

    def _build_cj_group(self, parent: QWidget) -> QGroupBox:
        box = shrink_h(QGroupBox("Кабельные журналы", parent))
        box.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Maximum)
        v = QVBoxLayout(box)
        v.setContentsMargins(10, 13, 10, 10)
        v.addLayout(
            self._action_row(
                box,
                "КЖ: выбрать папку",
                lambda: self._folder("Кабельные журналы", legacy_actions.run_cj_dir),
                help_key="button_open_cj_dir",
            )
        )
        return box

    def _build_revisions_group(self, parent: QWidget) -> QGroupBox:
        box = shrink_h(QGroupBox("MTO revisions", parent))
        box.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Maximum)
        v = QVBoxLayout(box)
        v.setContentsMargins(10, 13, 10, 10)
        v.addLayout(
            self._action_row(
                box,
                "MTO vs MTO (2 шт.)",
                lambda: self._folder(
                    "MTO vs MTO",
                    legacy_actions.run_mto_compare_dir,
                ),
                help_key="open_mto_mto_folder",
            )
        )
        v.addLayout(
            self._action_row(
                box,
                "MTO vs MTO (мульти)",
                lambda: self._folder(
                    "MTO multi",
                    legacy_actions.run_mto_multi_compare_dir,
                ),
                help_key="open_mto_mto_multi_folder",
            )
        )
        v.addLayout(
            self._action_row(
                box,
                "MTO: цепочка ревизий",
                lambda: self._folder(
                    "MTO цепочка ревизий",
                    legacy_actions.run_mto_chain_compare_dir,
                ),
                help_key="open_mto_mto_chain_folder",
            )
        )
        return box

    def _build_services_group(self, parent: QWidget) -> QGroupBox:
        box = shrink_h(QGroupBox("Сервисы", parent))
        box.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Maximum)
        v = QVBoxLayout(box)
        v.setContentsMargins(10, 13, 10, 10)
        v.addLayout(
            self._action_row(
                box,
                "Открыть NanoCAD.db",
                lambda: self._file(
                    "NanoCAD.db",
                    legacy_actions.run_nanocad_db,
                    file_filter=_DB_FILTER,
                ),
                help_key="open_nanocad_db",
            )
        )
        v.addLayout(
            self._action_row(
                box,
                "Собрать ZIP для коллег",
                self._click_zip,
                help_key="open_project_release_zip",
            )
        )
        return box

    def _action_row(
        self,
        parent: QWidget,
        label: str,
        on_click: Callable[[], None],
        *,
        help_key: str | None = None,
    ) -> QHBoxLayout:
        row = QHBoxLayout()
        btn = QPushButton(label, parent)
        btn.setMinimumWidth(0)
        btn.clicked.connect(on_click)
        row.addWidget(btn, stretch=1)
        if help_key is not None:
            help_btn = QPushButton("i", parent)
            help_btn.setFixedSize(30, 30)
            help_btn.setToolTip("Справка")
            help_btn.clicked.connect(
                lambda _checked=False, k=help_key: self._click_help(k)
            )
            row.addWidget(help_btn)
        return row

    def _folder(self, title: str, fn: Callable[..., Any]) -> None:
        if self._on_folder_job is not None:
            self._on_folder_job(title, fn)

    def _file(
        self,
        title: str,
        fn: Callable[..., Any],
        *extra: Any,
        file_filter: str = _EXCEL_FILTER,
    ) -> None:
        if self._on_file_job is not None:
            self._on_file_job(title, fn, extra, file_filter)

    def _click_pdf(self) -> None:
        if self._on_launch_pdf is not None:
            self._on_launch_pdf()

    def _click_rd_catalog(self) -> None:
        if self._on_launch_rd_catalog is not None:
            self._on_launch_rd_catalog()

    def _click_zip(self) -> None:
        if self._on_zip is not None:
            self._on_zip()

    def _click_help(self, help_key: str) -> None:
        if self._on_help is not None:
            self._on_help(help_key)
