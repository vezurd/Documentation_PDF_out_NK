"""Read-only cheat sheet: TSD packing xlsx columns (STD vs SO - PL)."""

from __future__ import annotations

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QAbstractItemView,
    QGroupBox,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QScrollArea,
    QSizePolicy,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

# (Excel letter, field, comment, required label)
_STD_ROWS: tuple[tuple[str, str, str, str], ...] = (
    ("B", "CODE", "Код РД (любой формат: BCC, ECBL, …)", "обяз."),
    ("C", "SPECIFICATION_NAME", "Имя спеки → parse Титул/Марка", "для title"),
    ("D", "TAGS", "Теги", "опц."),
    ("H", "NAME;TYPE_MARK", "Наименование;хар-ки (split по `;`)", "NAME обяз."),
    ("I", "VALUES", "Количество", "обяз."),
    ("J", "UNITS", "Ед. изм.", "обяз."),
    ("M", "VENDOR", "Поставщик (часто пуст)", "опц."),
)

_SOPL_ROWS: tuple[tuple[str, str, str, str], ...] = (
    ("C", "SPECIFICATION_NAME", "Name of document → fallback title/system", "для title"),
    ("F", "VENDOR", "Поставщик", "опц."),
    ("G", "CODE", "PO item", "обяз."),
    ("J", "TAGS", "Теги", "опц."),
    ("M", "NAME", "Russian translation (TYPE_MARK нет)", "обяз."),
    ("N", "VALUES", "Quantity", "обяз."),
    ("O", "UNITS", "Ед. изм.", "обяз."),
    ("BE", "DS_TITLE", "Титул (если пуст — из SPEC)", "опц."),
    ("BF", "DS_SYSTEM", "Марка (если пуст — из SPEC)", "опц."),
)

_HEADERS = ("Столбец", "Поле", "Комментарий", "Обяз.")


def _fill_table(table: QTableWidget, rows: tuple[tuple[str, str, str, str], ...]) -> None:
    table.setColumnCount(4)
    table.setHorizontalHeaderLabels(list(_HEADERS))
    table.setRowCount(len(rows))
    table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
    table.setSelectionMode(QAbstractItemView.SelectionMode.NoSelection)
    table.setFocusPolicy(Qt.FocusPolicy.NoFocus)
    table.verticalHeader().setVisible(False)
    table.setAlternatingRowColors(True)
    for r, (letter, field, comment, req) in enumerate(rows):
        table.setItem(r, 0, QTableWidgetItem(letter))
        table.setItem(r, 1, QTableWidgetItem(field))
        table.setItem(r, 2, QTableWidgetItem(comment))
        table.setItem(r, 3, QTableWidgetItem(req))
    header = table.horizontalHeader()
    header.setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
    header.setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
    header.setSectionResizeMode(2, QHeaderView.ResizeMode.Stretch)
    header.setSectionResizeMode(3, QHeaderView.ResizeMode.ResizeToContents)
    table.resizeRowsToContents()
    # Prefer full content height so both tables fit without inner scroll when possible.
    row_h = sum(table.rowHeight(i) for i in range(table.rowCount()))
    header_h = table.horizontalHeader().height()
    table.setMinimumHeight(header_h + row_h + 4)
    table.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Minimum)


class TsdPackingHelpPanel(QWidget):
    """Help tab: expected Excel columns for TSD packing load (manual check)."""

    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)

        intro = QLabel(
            "Сверка xlsx при замечании «0 позиций» / сдвиге столбцов.\n"
            "Позиция: обязательны CODE, NAME, VALUES, UNITS; TAGS и VENDOR — опциональны. "
            "Остальные столбцы робот не читает.\n"
            "STD (Single*): формат CODE не ограничен — отсекаются только заголовки/баннер. "
            "SO - PL: вкладка с именем «SO - PL»; шапка Name of Zip / Name of document — не позиция.\n"
            r"Master* не читаются. Свод/кэш: "
            r"\\bcc\eng\PrDoc\377_НИПИГАЗ\АГХК\КСБ\RFP_MTO_VO\_RFP\УЛ сводный файл.",
            self,
        )
        intro.setWordWrap(True)
        intro.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)

        std_box = QGroupBox("STD — вкладки Single*", self)
        std_layout = QVBoxLayout(std_box)
        std_table = QTableWidget(std_box)
        _fill_table(std_table, _STD_ROWS)
        std_layout.addWidget(std_table)

        sopl_box = QGroupBox("SO - PL — вкладка «SO - PL»", self)
        sopl_layout = QVBoxLayout(sopl_box)
        sopl_table = QTableWidget(sopl_box)
        _fill_table(sopl_table, _SOPL_ROWS)
        sopl_layout.addWidget(sopl_table)

        tables_row = QHBoxLayout()
        tables_row.setSpacing(12)
        tables_row.addWidget(std_box, stretch=1)
        tables_row.addWidget(sopl_box, stretch=1)

        inner = QWidget(self)
        inner_layout = QVBoxLayout(inner)
        inner_layout.setContentsMargins(10, 10, 10, 10)
        inner_layout.setSpacing(12)
        inner_layout.addWidget(intro)
        inner_layout.addLayout(tables_row)
        inner_layout.addStretch(1)

        scroll = QScrollArea(self)
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QScrollArea.Shape.NoFrame)
        scroll.setWidget(inner)

        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.addWidget(scroll)
