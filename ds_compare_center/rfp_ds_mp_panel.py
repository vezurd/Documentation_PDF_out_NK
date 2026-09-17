"""RFP vs manager-matrix surnames tab for ``ds_compare_center``."""

from __future__ import annotations

from collections.abc import Callable

from PySide6.QtCore import Qt
from PySide6.QtGui import QBrush, QColor
from PySide6.QtWidgets import (
    QAbstractItemView,
    QGroupBox,
    QHeaderView,
    QLabel,
    QPushButton,
    QSizePolicy,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from RFQ.rfp_parts.ds_checklist import DEFAULT_PARTS_DIR
from RFQ.tags_rfp_compare.ds_manager_roster import (
    DsRosterCompareResult,
    get_last_ds_roster_compare,
    resolve_ds_manager_matrix_path,
    roster_status_tone,
)
from ds_compare_center.split_layout import (
    WrappingLabel,
    build_side_by_side,
    shrink_h,
)

_TABLE_STYLE = (
    "QTableWidget { font-size: 11px; }"
    "QTableWidget::item { padding: 1px 4px; }"
)
_ROW_HEIGHT = 22
_TONE_COLORS = {
    "ok": (QColor("#1a7f37"), QColor("#e6f4ea")),
    "warn": (QColor("#8a6d1b"), QColor("#fff8e1")),
    "leftover": (QColor("#5f6368"), QColor("#f1f3f4")),
    "problem": (QColor("#b42318"), QColor("#fce8e6")),
}
_BANNER_OK = "color: #287a3d; font-size: 11px;"
_BANNER_WARN = "color: #8a6d1b; font-size: 11px;"
_BANNER_BAD = "color: #b42318; font-size: 11px;"
_HEADERS = [
    "Имя ДС",
    "Фактический ДС",
    "Порядковый ДС",
    "Фамилия",
    "Статус",
    "Было на Лист2",
    "Дубликат",
    "Что делать",
]


class RfpDsMpPanel(QWidget):
    """Scan RFP part names vs surnames in «Список ДС — Фамилии МП»."""

    def __init__(
        self,
        parent: QWidget | None = None,
        *,
        on_run: Callable[[], None] | None = None,
        on_open: Callable[[], None] | None = None,
    ) -> None:
        """Create the panel.

        Args:
            parent: Qt parent.
            on_run: Start the compare job (parent owns FunctionJobRunner).
            on_open: Open the manager-matrix workbook in Excel.
        """
        super().__init__(parent)
        self._on_run = on_run
        self._on_open = on_open
        splitter, self.monitor = build_side_by_side(
            self, build_left=self._build_left, show_stop=False
        )
        self.splitter = splitter
        root = QVBoxLayout(self)
        root.setContentsMargins(8, 8, 8, 8)
        root.addWidget(splitter)

    def _build_left(self, left: QWidget, left_layout: QVBoxLayout) -> None:
        left_layout.addWidget(self._build_coverage_group(left), stretch=1)

    def _build_coverage_group(self, parent: QWidget) -> QGroupBox:
        box = shrink_h(QGroupBox("Соответствие RFP ↔ фамилии МП", parent))
        box.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        v = QVBoxLayout(box)
        v.setContentsMargins(10, 13, 10, 10)
        v.setSpacing(8)

        hint = WrappingLabel(
            "Первая строка — вердикт: «ОК», «Надо дозаполнить» или "
            "«Фамилии заполнены, проверьте лишние / дубликаты на Лист2». "
            "Кнопка сверяет файлы RFP_Зиновьев с Лист2 и дописывает справа "
            "столбцы «Статус», «Дубликат», «Что делать». "
            "Колонки «Имя ДС» и «Фамилия» не меняет — их читает Запуск RFP.",
            box,
        )
        v.addWidget(hint)

        self._paths = QLabel(
            f"Книга: {resolve_ds_manager_matrix_path()}\nRFP: {DEFAULT_PARTS_DIR}",
            box,
        )
        self._paths.setWordWrap(True)
        self._paths.setTextInteractionFlags(
            Qt.TextInteractionFlag.TextSelectableByMouse
            | Qt.TextInteractionFlag.TextSelectableByKeyboard
        )
        self._paths.setStyleSheet("color: #333; font-size: 11px;")
        v.addWidget(self._paths)

        btn_run = QPushButton("Проверить соответствие ДС ↔ фамилии МП", box)
        btn_run.setMinimumWidth(0)
        btn_run.clicked.connect(self._click_run)
        v.addWidget(btn_run)

        btn_open = QPushButton("Открыть книгу фамилий", box)
        btn_open.setMinimumWidth(0)
        btn_open.clicked.connect(self._click_open)
        v.addWidget(btn_open)

        self._banner = QLabel("Нет данных.", box)
        self._banner.setWordWrap(True)
        self._banner.setStyleSheet("color: #555; font-size: 11px;")
        self._banner.setSizePolicy(
            QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed
        )
        v.addWidget(self._banner)

        table = QTableWidget(box)
        table.setColumnCount(len(_HEADERS))
        table.setHorizontalHeaderLabels(_HEADERS)
        table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        table.setWordWrap(False)
        table.setAlternatingRowColors(False)
        table.verticalHeader().setVisible(False)
        table.verticalHeader().setDefaultSectionSize(_ROW_HEIGHT)
        table.verticalHeader().setMinimumSectionSize(_ROW_HEIGHT)
        table.verticalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Fixed)
        table.setSortingEnabled(False)
        table.setMinimumHeight(180)
        table.setSizePolicy(
            QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding
        )
        header = table.horizontalHeader()
        for col in range(len(_HEADERS) - 1):
            header.setSectionResizeMode(col, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(
            len(_HEADERS) - 1, QHeaderView.ResizeMode.Stretch
        )
        header.setStretchLastSection(True)
        table.setStyleSheet(_TABLE_STYLE)
        self._table = table
        v.addWidget(table, stretch=1)
        return box

    def _click_run(self) -> None:
        if self._on_run is None:
            return
        self._on_run()

    def _click_open(self) -> None:
        if self._on_open is None:
            return
        self._on_open()

    def fill_from_result(self, result: DsRosterCompareResult) -> None:
        """Populate the banner and table from a compare result.

        Args:
            result: Last scan outcome (may include a load error and no rows).
        """
        line = result.summary_line()
        self._banner.setText(line)
        if result.load_error or result.needs_fill:
            self._banner.setStyleSheet(_BANNER_BAD)
        elif result.is_ok:
            self._banner.setStyleSheet(_BANNER_OK)
        else:
            self._banner.setStyleSheet(_BANNER_WARN)
        self._paths.setText(f"Книга: {result.matrix_path}\nRFP: {result.parts_dir}")
        table = self._table
        rows = result.rows
        table.setRowCount(0)
        table.setRowCount(len(rows))
        v_center = Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter
        fallback_fg, fallback_bg = _TONE_COLORS["problem"]
        for row_idx, row in enumerate(rows):
            fg, bg = _TONE_COLORS.get(
                roster_status_tone(row.status), (fallback_fg, fallback_bg)
            )
            values = (
                row.copy_name,
                row.actual_label or "—",
                row.sequential_label or "—",
                row.manager or "—",
                row.status,
                row.was_on_sheet1 or "—",
                row.duplicate_label or "—",
                row.note or "—",
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
            table.setRowHeight(row_idx, _ROW_HEIGHT)

    def on_job_finished(self, success: bool, message: str) -> None:
        """Reload the last stored compare result after the job ends.

        Args:
            success: False when the worker raised or the workbook was missing.
            message: Runner summary; used only if no stored result exists.
        """
        result = get_last_ds_roster_compare()
        if result is not None:
            self.fill_from_result(result)
            return
        if not success:
            self._banner.setText(message or "Ошибка проверки.")
            self._banner.setStyleSheet(_BANNER_BAD)
