"""Dialog with painted examples of Комплекты MTO/revision cells."""

from __future__ import annotations

from collections.abc import Mapping, Sequence

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QAbstractItemView,
    QDialog,
    QDialogButtonBox,
    QGroupBox,
    QHeaderView,
    QLabel,
    QScrollArea,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from rd_catalog.monitor_qt import apply_monitor_cell
from rd_catalog.monitor_views import (
    KITS_PAINT_LEGEND_INTRO,
    KITS_PAINT_LEGEND_TITLE,
    PaintLegendSection,
    kits_paint_legend,
)

_HEADERS = ("Столбец", "Пример", "Что значит")


class KitsPaintLegendDialog(QDialog):
    """Show fill/bold examples from ``kits_paint_legend``."""

    def __init__(
        self,
        palette: Mapping[str, str] | None = None,
        parent: QWidget | None = None,
    ) -> None:
        """Build the dialog for the current status palette.

        Args:
            palette: Live ``status_colors.json`` map from the catalog window.
            parent: Optional Qt parent.
        """

        super().__init__(parent)
        self.setWindowTitle(KITS_PAINT_LEGEND_TITLE)
        self.setMinimumSize(760, 520)
        self.resize(860, 640)
        self._build(kits_paint_legend(palette))

    def _build(self, sections: Sequence[PaintLegendSection]) -> None:
        layout = QVBoxLayout(self)
        intro = QLabel(KITS_PAINT_LEGEND_INTRO, self)
        intro.setWordWrap(True)
        layout.addWidget(intro)

        scroll = QScrollArea(self)
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QScrollArea.Shape.NoFrame)
        body = QWidget(scroll)
        body_layout = QVBoxLayout(body)
        body_layout.setContentsMargins(0, 0, 0, 0)
        for section in sections:
            body_layout.addWidget(self._section_box(section, body))
        body_layout.addStretch(1)
        scroll.setWidget(body)
        layout.addWidget(scroll, 1)

        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Close, self
        )
        buttons.rejected.connect(self.reject)
        buttons.accepted.connect(self.accept)
        layout.addWidget(buttons)

    def _section_box(
        self, section: PaintLegendSection, parent: QWidget
    ) -> QGroupBox:
        box = QGroupBox(section.title, parent)
        layout = QVBoxLayout(box)
        if section.intro:
            hint = QLabel(section.intro, box)
            hint.setWordWrap(True)
            layout.addWidget(hint)
        table = QTableWidget(len(section.samples), 3, box)
        table.setHorizontalHeaderLabels(list(_HEADERS))
        table.setSelectionMode(QAbstractItemView.SelectionMode.NoSelection)
        table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        table.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        table.verticalHeader().setVisible(False)
        table.setWordWrap(True)
        table.setShowGrid(True)
        header = table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(2, QHeaderView.ResizeMode.Stretch)
        for row, sample in enumerate(section.samples):
            column_item = QTableWidgetItem(sample.column)
            example_item = QTableWidgetItem()
            apply_monitor_cell(example_item, sample.as_monitor_cell())
            meaning_item = QTableWidgetItem(sample.meaning)
            meaning_item.setTextAlignment(
                Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter
            )
            table.setItem(row, 0, column_item)
            table.setItem(row, 1, example_item)
            table.setItem(row, 2, meaning_item)
        table.resizeRowsToContents()
        rows_height = sum(
            table.rowHeight(row) for row in range(table.rowCount())
        )
        header_height = table.horizontalHeader().height()
        table.setFixedHeight(header_height + rows_height + 4)
        table.setHorizontalScrollBarPolicy(
            Qt.ScrollBarPolicy.ScrollBarAlwaysOff
        )
        table.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        layout.addWidget(table)
        return box
