"""RFP / as-build run tab (aggregate_tags), packing-lists layout."""

from __future__ import annotations

from collections.abc import Callable
from pathlib import Path

from PySide6.QtCore import QRectF, QSize, Qt
from PySide6.QtGui import (
    QAbstractTextDocumentLayout,
    QBrush,
    QColor,
    QFont,
    QPalette,
    QTextDocument,
    QTextOption,
)
from PySide6.QtWidgets import (
    QAbstractItemView,
    QGroupBox,
    QHeaderView,
    QLabel,
    QPushButton,
    QSizePolicy,
    QStyle,
    QStyledItemDelegate,
    QStyleOptionViewItem,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from RFQ.tags_rfp_compare.rfp_tags_utils import load_config, resolve_effective_rfp_path
from ds_compare_center.rfp_progress_parser import (
    MILESTONE_IDS,
    MILESTONE_STATES,
    MilestoneEvent,
)
from ds_compare_center.split_layout import build_side_by_side, shrink_h

_MILESTONE_LABELS = {
    "prepare": "Подготовка запуска и папки результата",
    "ds_id_check": "Соответствие номеров ДС: RFP ↔ папки УЛ",
    "ds_mp_check": "Соответствие ДС: RFP ↔ фамилии МП",
    "parts_preflight": "Свежесть свода частей RFP (при необходимости пересборка)",
    "ul_preflight": "Свежесть свода УЛ (при необходимости пересборка)",
    "rfp_load": "Загрузка RFP",
    "mto_google_load": "Загрузка MTO и базы Google",
    "vo_load": "Загрузка РКД (VO)",
    "units_gate": "Проверка и преобразование единиц измерения",
    "step4_match": "Преобразование и сопоставление Step4",
    "global_checks": "Глобальные проверки тегов, баланса и сумм",
    "packing_lists": "Сопоставление с УЛ и audit",
    "save_excel": "Сохранение итогового Excel",
    "bcc_accum_matrix": "Накопительная матрица по коду BCC",
    "complete": "Завершение",
}
_STATE_TEXT = {
    "Waiting": "Ожидает",
    "Running": "Выполняется",
    "Done": "Готово",
    "Skipped": "Пропущено",
    "Error": "Ошибка",
}
_STATE_COLORS = {
    "Waiting": (QColor("#6b7280"), QColor("#f3f4f6")),
    "Running": (QColor("#1d4ed8"), QColor("#dbeafe")),
    "Done": (QColor("#287a3d"), QColor("#e6f4ea")),
    "Skipped": (QColor("#7c6f4c"), QColor("#f5f0e6")),
    "Error": (QColor("#b42318"), QColor("#fce8e6")),
}
_STATE_BY_LOWER = {state.lower(): state for state in MILESTONE_STATES}
_TABLE_STYLE = (
    "QTableWidget { font-size: 11px; }"
    "QTableWidget::item { padding: 1px 4px; }"
)
_HEADERS = ("Шаг", "Статус", "Детали")
_STEP_COL_WIDTH = 250
_ROW_MIN_HEIGHT = 22


class _WrapAnywhereDelegate(QStyledItemDelegate):
    """Paint cell text wrapping inside long tokens such as UNC paths."""

    _PAD_X = 4
    _PAD_Y = 2

    def paint(self, painter, option, index) -> None:
        opt = QStyleOptionViewItem(option)
        self.initStyleOption(opt, index)
        text = opt.text
        opt.text = ""
        widget = opt.widget
        style = widget.style() if widget is not None else None
        if style is not None:
            style.drawControl(
                QStyle.ControlElement.CE_ItemViewItem, opt, painter, widget
            )
        if not text:
            return
        width = max(opt.rect.width() - 2 * self._PAD_X, 1)
        doc = _wrap_document(str(text), opt.font, width)
        painter.save()
        painter.translate(opt.rect.left() + self._PAD_X, opt.rect.top() + self._PAD_Y)
        painter.setClipRect(
            QRectF(
                0,
                0,
                opt.rect.width() - 2 * self._PAD_X,
                opt.rect.height() - 2 * self._PAD_Y,
            )
        )
        ctx = QAbstractTextDocumentLayout.PaintContext()
        fg = index.data(Qt.ItemDataRole.ForegroundRole)
        if isinstance(fg, QBrush):
            ctx.palette.setColor(QPalette.ColorRole.Text, fg.color())
        elif isinstance(fg, QColor):
            ctx.palette.setColor(QPalette.ColorRole.Text, fg)
        else:
            ctx.palette.setColor(
                QPalette.ColorRole.Text,
                opt.palette.color(QPalette.ColorRole.Text),
            )
        doc.documentLayout().draw(painter, ctx)
        painter.restore()

    def sizeHint(self, option, index) -> QSize:
        text = index.data(Qt.ItemDataRole.DisplayRole) or ""
        width = option.rect.width()
        widget = option.widget
        if width <= 0 and isinstance(widget, QTableWidget):
            width = widget.columnWidth(index.column())
        width = max(int(width) - 2 * self._PAD_X, 40)
        doc = _wrap_document(str(text), option.font, width)
        return QSize(width, int(doc.size().height()) + 2 * self._PAD_Y)


def _wrap_document(text: str, font: QFont, width: int) -> QTextDocument:
    doc = QTextDocument()
    doc.setDefaultFont(font)
    doc.setDocumentMargin(0)
    doc.setPlainText(text)
    wrap = QTextOption()
    wrap.setWrapMode(QTextOption.WrapMode.WrapAnywhere)
    doc.setDefaultTextOption(wrap)
    doc.setTextWidth(float(width))
    return doc


class _MilestoneTable(QTableWidget):
    """Progress table that reflows wrapped rows when the pane width changes."""

    def resizeEvent(self, event) -> None:
        super().resizeEvent(event)
        self.resizeRowsToContents()


class RfpRunPanel(QWidget):
    """Run RFP↔MTO/VO and as-build checks with a live Job monitor."""

    def __init__(
        self,
        parent: QWidget | None = None,
        *,
        on_run_rfp: Callable[[], None] | None = None,
        on_run_asbuild: Callable[[], None] | None = None,
        on_open_last: Callable[[], None] | None = None,
        on_open_matrix: Callable[[], None] | None = None,
    ) -> None:
        super().__init__(parent)
        self._on_run_rfp = on_run_rfp
        self._on_run_asbuild = on_run_asbuild
        self._on_open_last = on_open_last
        self._on_open_matrix = on_open_matrix
        self._active_milestone_id: str | None = None
        self._milestone_states: dict[str, str] = {}
        self._milestone_row_by_id: dict[str, int] = {}

        splitter, self.monitor = build_side_by_side(
            self, build_left=self._build_left, show_stop=True
        )
        self.splitter = splitter
        root = QVBoxLayout(self)
        root.setContentsMargins(8, 8, 8, 8)
        root.addWidget(splitter)
        self.refresh_from_config()
        self.reset_milestones()

    def _build_left(self, left: QWidget, left_layout: QVBoxLayout) -> None:
        left_layout.addWidget(self._build_actions_group(left), stretch=0)
        left_layout.addWidget(self._build_results_group(left), stretch=0)
        left_layout.addWidget(self._build_progress_group(left), stretch=1)

    def _build_actions_group(self, parent: QWidget) -> QGroupBox:
        box = shrink_h(QGroupBox("Сопоставление RFP / MTO / РКД / УЛ", parent))
        box.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Maximum)
        v = QVBoxLayout(box)
        v.setContentsMargins(10, 13, 10, 10)
        v.setSpacing(8)

        hint = QLabel(
            "УЛ используют общий кэш упаковочных листов. Если тегов УЛ больше, "
            "чем указано количество, проверка завершается с ошибкой до записи "
            "финального Excel. Лог — справа, настройки — во вкладке «RFP · Настройки».",
            box,
        )
        hint.setWordWrap(True)
        shrink_h(hint)
        v.addWidget(hint)

        self._btn_rfp = QPushButton("", box)
        self._btn_rfp.setMinimumWidth(0)
        self._btn_rfp.clicked.connect(self._click_rfp)
        v.addWidget(self._btn_rfp)

        self._rfp_source_label = QLabel("", box)
        self._rfp_source_label.setWordWrap(True)
        self._rfp_source_label.setStyleSheet("color: #555; font-size: 10px;")
        shrink_h(self._rfp_source_label)
        v.addWidget(self._rfp_source_label)

        btn_asbuild = QPushButton("Сравнить as-build RFP ↔ MTO ↔ РКД", box)
        btn_asbuild.setMinimumWidth(0)
        btn_asbuild.clicked.connect(self._click_asbuild)
        v.addWidget(btn_asbuild)
        return box

    def _build_results_group(self, parent: QWidget) -> QGroupBox:
        box = shrink_h(QGroupBox("Результаты Step4", parent))
        box.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Maximum)
        v = QVBoxLayout(box)
        v.setContentsMargins(10, 13, 10, 10)
        v.setSpacing(8)

        hint = QLabel(
            "Открывает именно последний финальный файл Step4 из папок результатов "
            "основного и as-build запусков.",
            box,
        )
        hint.setWordWrap(True)
        shrink_h(hint)
        v.addWidget(hint)

        btn = QPushButton("Открыть последний результат Step4", box)
        btn.setMinimumWidth(0)
        btn.clicked.connect(self._click_open_last)
        v.addWidget(btn)

        btn_matrix = QPushButton("Открыть накопительную матрицу BCC", box)
        btn_matrix.setMinimumWidth(0)
        btn_matrix.clicked.connect(self._click_open_matrix)
        v.addWidget(btn_matrix)
        return box

    def _build_progress_group(self, parent: QWidget) -> QGroupBox:
        box = shrink_h(QGroupBox("Ход выполнения", parent))
        box.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        v = QVBoxLayout(box)
        v.setContentsMargins(10, 10, 10, 10)
        v.setSpacing(6)

        table = _MilestoneTable(box)
        table.setColumnCount(len(_HEADERS))
        table.setHorizontalHeaderLabels(list(_HEADERS))
        table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        table.setWordWrap(True)
        table.setTextElideMode(Qt.TextElideMode.ElideNone)
        table.setAlternatingRowColors(False)
        table.verticalHeader().setVisible(False)
        table.verticalHeader().setMinimumSectionSize(_ROW_MIN_HEIGHT)
        table.verticalHeader().setSectionResizeMode(
            QHeaderView.ResizeMode.ResizeToContents
        )
        table.setSortingEnabled(False)
        table.setMinimumHeight(180)
        table.setSizePolicy(
            QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding
        )
        table.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        table.setItemDelegate(_WrapAnywhereDelegate(table))
        header = table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeMode.Interactive)
        header.setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(2, QHeaderView.ResizeMode.Stretch)
        header.setStretchLastSection(True)
        header.setMinimumSectionSize(72)
        header.resizeSection(0, _STEP_COL_WIDTH)
        table.setStyleSheet(_TABLE_STYLE)
        table.setRowCount(len(MILESTONE_IDS))
        v_center = Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter
        selectable = Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsSelectable
        for row, milestone_id in enumerate(MILESTONE_IDS):
            self._milestone_row_by_id[milestone_id] = row
            for col in range(len(_HEADERS)):
                item = QTableWidgetItem("")
                item.setFlags(selectable)
                item.setTextAlignment(v_center)
                table.setItem(row, col, item)
        self._table = table
        v.addWidget(table, stretch=1)
        return box

    @property
    def active_milestone_id(self) -> str | None:
        """Currently running milestone, if one has been reported."""
        return self._active_milestone_id

    def refresh_from_config(self) -> None:
        """Refresh dynamic RFP action wording from the current config."""
        config = load_config()
        step4 = config.get("step4", {})
        include_packing = bool(
            step4.get("include_packing_lists", False)
            if isinstance(step4, dict)
            else False
        )
        if include_packing:
            text = "Сравнить RFP ↔ MTO ↔ РКД ↔ УЛ"
        else:
            text = "Сравнить RFP ↔ MTO ↔ РКД"
        self._btn_rfp.setText(text)
        self._refresh_rfp_source_label(config)

    def _refresh_rfp_source_label(self, config: dict) -> None:
        """Show which RFP workbook the main run button will load."""
        try:
            resolved = resolve_effective_rfp_path(config)
        except FileNotFoundError as exc:
            self._rfp_source_label.setText(str(exc))
            return
        if resolved.used_latest_net:
            name = Path(resolved.path).name
            if name.endswith("_no_tags.xlsx"):
                prefix = "Источник RFP: последний свод частей (без тегов)"
            else:
                prefix = "Источник RFP: последний свод частей"
        else:
            prefix = "Источник RFP: путь из настроек"
        self._rfp_source_label.setText(f"{prefix}\n{resolved.path}")

    def reset_milestones(self) -> None:
        """Reset every milestone to Waiting without inferring prior progress."""
        self._active_milestone_id = None
        for milestone_id in MILESTONE_IDS:
            self._set_milestone(milestone_id, "Waiting", "")

    def apply_milestone(
        self,
        event_or_id: MilestoneEvent | str,
        state: str | None = None,
        detail: str = "",
    ) -> None:
        """Apply one explicit milestone event to the progress block.

        Args:
            event_or_id: Parsed event or milestone ID.
            state: State when ``event_or_id`` is an ID.
            detail: Optional human-readable detail.

        Raises:
            ValueError: If the milestone ID or state is unknown.
        """
        if isinstance(event_or_id, MilestoneEvent):
            milestone_id = event_or_id.milestone_id
            normalized_state = _normalize_state(event_or_id.state)
            detail = event_or_id.detail
        else:
            milestone_id = event_or_id
            normalized_state = _normalize_state(state)

        if milestone_id not in self._milestone_row_by_id:
            raise ValueError(f"Unknown RFP milestone: {milestone_id!r}")
        self._set_milestone(milestone_id, normalized_state, detail)

        if normalized_state == "Running":
            self._active_milestone_id = milestone_id
        elif self._active_milestone_id == milestone_id:
            self._active_milestone_id = None

    def mark_active_error(self, detail: str = "") -> None:
        """Mark only the currently running milestone as failed."""
        if self._active_milestone_id is None:
            return
        self.apply_milestone(self._active_milestone_id, "Error", detail)

    def _set_milestone(self, milestone_id: str, state: str, detail: str) -> None:
        self._milestone_states[milestone_id] = state
        row = self._milestone_row_by_id[milestone_id]
        fg, bg = _STATE_COLORS[state]
        status_text = _STATE_TEXT[state]
        detail_text = detail.strip()
        values = (
            _MILESTONE_LABELS[milestone_id],
            status_text,
            detail_text,
        )
        v_center = Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter
        selectable = Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsSelectable
        for col, text in enumerate(values):
            item = self._table.item(row, col)
            if item is None:
                item = QTableWidgetItem()
                self._table.setItem(row, col, item)
            item.setText(text)
            item.setForeground(QBrush(fg))
            item.setBackground(QBrush(bg))
            item.setToolTip(text)
            item.setFlags(selectable)
            item.setTextAlignment(v_center)
        self._table.resizeRowsToContents()

    def _click_rfp(self) -> None:
        if self._on_run_rfp:
            self._on_run_rfp()

    def _click_asbuild(self) -> None:
        if self._on_run_asbuild:
            self._on_run_asbuild()

    def _click_open_last(self) -> None:
        if self._on_open_last:
            self._on_open_last()

    def _click_open_matrix(self) -> None:
        if self._on_open_matrix:
            self._on_open_matrix()


def _normalize_state(state: str | None) -> str:
    if not isinstance(state, str):
        raise ValueError(f"Unknown RFP milestone state: {state!r}")
    normalized = _STATE_BY_LOWER.get(state.strip().lower())
    if normalized is None:
        raise ValueError(f"Unknown RFP milestone state: {state!r}")
    return normalized
