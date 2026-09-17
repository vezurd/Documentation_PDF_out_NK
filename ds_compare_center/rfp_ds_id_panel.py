"""RFP vs UL DS-number coverage tab for ``ds_compare_center``."""

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

from RFQ.ds_compare.tsd_packing_load import DEFAULT_TSD_PACKING_ROOT
from RFQ.rfp_parts.ds_checklist import DEFAULT_PARTS_DIR
from RFQ.rfp_parts.ds_id_coverage import (
    DsIdCoverageResult,
    get_last_ds_id_coverage,
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
_STATUS_COLORS = {
    "both": (QColor("#1a7f37"), QColor("#e6f4ea")),
    "rfp_only": (QColor("#b42318"), QColor("#fce8e6")),
    "ul_only": (QColor("#b42318"), QColor("#fce8e6")),
    "unparsed": (QColor("#b42318"), QColor("#fce8e6")),
    "gf_special": (QColor("#7c6f4c"), QColor("#f5f0e6")),
}
_BANNER_OK = "color: #287a3d; font-size: 11px;"
_BANNER_BAD = "color: #b42318; font-size: 11px;"
_HEADERS = [
    "Фактический ДС",
    "Статус",
    "Папка УЛ",
    "УЛ, файлов",
    "RFP",
    "Порядковый RFP",
    "Комментарий",
]


class RfpDsIdPanel(QWidget):
    """Scan RFP part names vs UL folder names; show coverage table."""

    def __init__(
        self,
        parent: QWidget | None = None,
        *,
        on_run: Callable[[], None] | None = None,
    ) -> None:
        """Create the panel.

        Args:
            parent: Qt parent.
            on_run: Start the coverage job (parent owns FunctionJobRunner).
        """
        super().__init__(parent)
        self._on_run = on_run
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
        box = shrink_h(QGroupBox("Соответствие RFP ↔ УЛ", parent))
        box.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        v = QVBoxLayout(box)
        v.setContentsMargins(10, 13, 10, 10)
        v.setSpacing(8)

        hint = WrappingLabel(
            "Папка УЛ = фактический номер сразу после ДС; корректировка RFP = "
            "ДС92_24Б (92 фактический, 24 порядковый информационный). "
            "Кнопка сканирует имена, не читает содержимое xlsx.",
            box,
        )
        v.addWidget(hint)

        paths = QLabel(
            f"УЛ: {DEFAULT_TSD_PACKING_ROOT}\nRFP: {DEFAULT_PARTS_DIR}",
            box,
        )
        paths.setWordWrap(True)
        paths.setTextInteractionFlags(
            Qt.TextInteractionFlag.TextSelectableByMouse
            | Qt.TextInteractionFlag.TextSelectableByKeyboard
        )
        paths.setStyleSheet("color: #333; font-size: 11px;")
        v.addWidget(paths)

        btn_run = QPushButton("Проверить соответствие RFP ↔ УЛ", box)
        btn_run.setMinimumWidth(0)
        btn_run.clicked.connect(self._click_run)
        v.addWidget(btn_run)

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

    def fill_from_result(self, result: DsIdCoverageResult) -> None:
        """Populate the banner and table from a coverage result.

        Args:
            result: Last scan outcome (may include a load error and no rows).
        """
        line = result.summary_line()
        self._banner.setText(line)
        self._banner.setStyleSheet(_BANNER_OK if result.is_ok else _BANNER_BAD)
        table = self._table
        rows = result.rows
        table.setRowCount(0)
        table.setRowCount(len(rows))
        v_center = Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter
        fallback_fg, fallback_bg = _STATUS_COLORS["unparsed"]
        for row_idx, row in enumerate(rows):
            fg, bg = _STATUS_COLORS.get(row.status, (fallback_fg, fallback_bg))
            rfp = ", ".join(row.rfp_labels) or "—"
            seq = ", ".join(str(n) for n in row.rfp_sequential) or "—"
            values = (
                row.actual_label,
                row.status_ru,
                row.ul_folder or "—",
                str(row.ul_xlsx),
                rfp,
                seq,
                row.note,
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
        """Reload the last stored coverage result after the job ends.

        Args:
            success: False when the worker raised or roots were missing.
            message: Runner summary; used only if no stored result exists.
        """
        result = get_last_ds_id_coverage()
        if result is not None:
            self.fill_from_result(result)
            return
        if not success:
            self._banner.setText(message or "Ошибка проверки.")
            self._banner.setStyleSheet(_BANNER_BAD)
