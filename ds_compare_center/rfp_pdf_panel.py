"""RFP PDF extract tab for ``ds_compare_center``."""

from __future__ import annotations

import os
from collections.abc import Callable
from pathlib import Path

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QAbstractItemView,
    QCheckBox,
    QFileDialog,
    QGroupBox,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QSizePolicy,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from RFQ.rfp_parts.pdf_rfp_extract import PdfRfpResult, preview_ds_label
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
_ISSUES_DISPLAY_CAP = 30
_BANNER_OK = "color: #287a3d; font-size: 11px;"
_BANNER_BAD = "color: #b42318; font-size: 11px;"
_HEADERS = ("Уровень", "Сообщение")
_PDF_FILTER = "PDF (*.pdf);;All files (*.*)"


class RfpPdfPanel(QWidget):
    """Pick an RFP PDF, extract materials xlsx, show counters and issues."""

    def __init__(
        self,
        parent: QWidget | None = None,
        *,
        on_run: Callable[[str, str, bool], None] | None = None,
    ) -> None:
        """Create the panel.

        Args:
            parent: Qt parent.
            on_run: Start the extract job with ``(pdf_path, ds_label, strip_stamp)``.
        """
        super().__init__(parent)
        self._on_run = on_run
        self._last_previewed_path = ""
        self._xlsx_path: Path | None = None
        self._report_path: Path | None = None
        self._out_dir: Path | None = None
        splitter, self.monitor = build_side_by_side(
            self, build_left=self._build_left, show_stop=False
        )
        self.splitter = splitter
        root = QVBoxLayout(self)
        root.setContentsMargins(8, 8, 8, 8)
        root.addWidget(splitter)

    def _build_left(self, left: QWidget, left_layout: QVBoxLayout) -> None:
        left_layout.addWidget(self._build_extract_group(left), stretch=1)

    def _build_extract_group(self, parent: QWidget) -> QGroupBox:
        box = shrink_h(QGroupBox("RFP из PDF", parent))
        box.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        v = QVBoxLayout(box)
        v.setContentsMargins(10, 13, 10, 10)
        v.setSpacing(8)

        hint = WrappingLabel(
            "xlsx и отчёт пишутся в подпапку рядом с PDF: "
            "«результат распознавания PDF_<дата_время>». "
            "В RFP_Зиновьев ничего не пишется.",
            box,
        )
        v.addWidget(hint)

        pdf_row = QHBoxLayout()
        self._pdf_edit = QLineEdit(box)
        self._pdf_edit.setPlaceholderText("Путь к PDF")
        self._pdf_edit.setMinimumWidth(0)
        self._pdf_edit.editingFinished.connect(self._on_pdf_path_finished)
        pdf_row.addWidget(self._pdf_edit, stretch=1)
        btn_browse = QPushButton("Выбрать PDF", box)
        btn_browse.setMinimumWidth(0)
        btn_browse.clicked.connect(self._browse_pdf)
        pdf_row.addWidget(btn_browse)
        v.addLayout(pdf_row)

        ds_row = QHBoxLayout()
        ds_label = QLabel("Имя ДС", box)
        ds_row.addWidget(ds_label)
        self._ds_edit = QLineEdit(box)
        self._ds_edit.setPlaceholderText("ДС98_35Б")
        self._ds_edit.setMinimumWidth(0)
        ds_row.addWidget(self._ds_edit, stretch=1)
        v.addLayout(ds_row)

        self._chk_strip = QCheckBox(
            "Удалить печать Диадок перед распознаванием",
            box,
        )
        self._chk_strip.setChecked(True)
        self._chk_strip.setToolTip(
            "Снимает со всех листов надпись «Передан через Диадок», "
            "номер страницы этой печати и значок. Исходный PDF не меняется; "
            "очищенная копия пишется в папку результата."
        )
        v.addWidget(self._chk_strip)

        self._btn_run = QPushButton("Распознать", box)
        self._btn_run.setMinimumWidth(0)
        self._btn_run.clicked.connect(self._click_run)
        v.addWidget(self._btn_run)

        open_row = QHBoxLayout()
        self._btn_open_xlsx = QPushButton("Открыть xlsx", box)
        self._btn_open_xlsx.setMinimumWidth(0)
        self._btn_open_xlsx.setEnabled(False)
        self._btn_open_xlsx.clicked.connect(self._open_xlsx)
        open_row.addWidget(self._btn_open_xlsx)
        self._btn_open_report = QPushButton("Открыть отчёт", box)
        self._btn_open_report.setMinimumWidth(0)
        self._btn_open_report.setEnabled(False)
        self._btn_open_report.clicked.connect(self._open_report)
        open_row.addWidget(self._btn_open_report)
        self._btn_open_dir = QPushButton("Открыть папку", box)
        self._btn_open_dir.setMinimumWidth(0)
        self._btn_open_dir.setEnabled(False)
        self._btn_open_dir.clicked.connect(self._open_dir)
        open_row.addWidget(self._btn_open_dir)
        open_row.addStretch(1)
        v.addLayout(open_row)

        self._banner = QLabel("Нет данных.", box)
        self._banner.setWordWrap(True)
        self._banner.setStyleSheet("color: #555; font-size: 11px;")
        self._banner.setSizePolicy(
            QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed
        )
        v.addWidget(self._banner)

        table = QTableWidget(box)
        table.setColumnCount(len(_HEADERS))
        table.setHorizontalHeaderLabels(list(_HEADERS))
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
        header.setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        header.setStretchLastSection(True)
        table.setStyleSheet(_TABLE_STYLE)
        self._table = table
        v.addWidget(table, stretch=1)
        return box

    def _browse_pdf(self) -> None:
        start = self._pdf_edit.text().strip()
        start_dir = str(Path(start).parent) if start else ""
        path, _selected = QFileDialog.getOpenFileName(
            self,
            "Выбрать PDF",
            start_dir,
            _PDF_FILTER,
        )
        if not path:
            return
        self._pdf_edit.setText(path)
        self._fill_ds_label_from_pdf(path)

    def _on_pdf_path_finished(self) -> None:
        self._fill_ds_label_from_pdf(self._pdf_edit.text().strip())

    def _fill_ds_label_from_pdf(self, pdf_path: str) -> None:
        if not pdf_path or pdf_path == self._last_previewed_path:
            return
        self._last_previewed_path = pdf_path
        candidate = Path(pdf_path)
        if not candidate.is_file():
            return
        try:
            label = preview_ds_label(candidate)
        except Exception:
            return
        if label:
            self._ds_edit.setText(label)

    def _click_run(self) -> None:
        if self._on_run is None:
            return
        self._on_run(
            self._pdf_edit.text().strip(),
            self._ds_edit.text().strip(),
            self._chk_strip.isChecked(),
        )

    def show_result(self, result: PdfRfpResult) -> None:
        """Update banner, issues table, and stored result paths.

        Args:
            result: Last extract outcome (xlsx is always written on success).
        """
        self._xlsx_path = result.xlsx_path
        self._report_path = result.report_path
        self._out_dir = result.out_dir
        ok = result.contract_errors == 0 and result.outside_words == 0
        self._banner.setText(
            f"Строк: {result.row_count}; ошибок контракта: {result.contract_errors}; "
            f"слов вне ячеек: {result.outside_words}"
        )
        self._banner.setStyleSheet(_BANNER_OK if ok else _BANNER_BAD)
        issues = result.issues[:_ISSUES_DISPLAY_CAP]
        table = self._table
        table.setRowCount(0)
        table.setRowCount(len(issues))
        v_center = Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter
        for row_idx, issue in enumerate(issues):
            values = (issue.level, issue.message)
            for col, text in enumerate(values):
                cell = QTableWidgetItem(text)
                cell.setToolTip(text)
                cell.setFlags(
                    Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsSelectable
                )
                cell.setTextAlignment(v_center)
                table.setItem(row_idx, col, cell)
            table.setRowHeight(row_idx, _ROW_HEIGHT)
        self._btn_open_xlsx.setEnabled(True)
        self._btn_open_report.setEnabled(True)
        self._btn_open_dir.setEnabled(True)

    def _open_xlsx(self) -> None:
        self._open_stored_path(self._xlsx_path, "Открыть xlsx")

    def _open_report(self) -> None:
        self._open_stored_path(self._report_path, "Открыть отчёт")

    def _open_dir(self) -> None:
        path = self._out_dir
        if path is None or not Path(path).is_dir():
            QMessageBox.information(self, "Открыть папку", "Папка ещё не создана.")
            return
        try:
            os.startfile(str(path))  # type: ignore[attr-defined]
        except Exception as exc:
            QMessageBox.warning(self, "Открыть папку", f"Не удалось открыть:\n{exc}")

    def _open_stored_path(self, path: Path | None, title: str) -> None:
        if path is None or not Path(path).is_file():
            QMessageBox.information(self, title, "Файл ещё не создан.")
            return
        try:
            os.startfile(str(path))  # type: ignore[attr-defined]
        except Exception as exc:
            QMessageBox.warning(self, title, f"Не удалось открыть:\n{exc}")
