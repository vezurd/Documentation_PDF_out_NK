"""Tab «Письма о согласовании»: ingest ``.msg``, preview F, write Google.

Qt monitor. Parsing and F patches are domain modules; this widget does not
call the Sheets API. The parent window starts :class:`GoogleFWriteThread`
after :attr:`write_requested`.
"""

from __future__ import annotations

import hashlib
import tempfile
from collections.abc import Sequence
from dataclasses import replace
from datetime import datetime
from pathlib import Path

from PySide6.QtCore import QEvent, QObject, QPoint, QSettings, QSize, Qt, Signal, Slot
from PySide6.QtGui import (
    QAbstractTextDocumentLayout,
    QBrush,
    QColor,
    QDragEnterEvent,
    QDropEvent,
    QGuiApplication,
    QKeyEvent,
    QKeySequence,
    QPainter,
    QShortcut,
    QTextDocument,
)
from PySide6.QtWidgets import (
    QAbstractItemView,
    QCheckBox,
    QComboBox,
    QCompleter,
    QDialog,
    QDialogButtonBox,
    QFileDialog,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QMenu,
    QMessageBox,
    QPlainTextEdit,
    QPushButton,
    QStyledItemDelegate,
    QStyle,
    QStyleOptionViewItem,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)
from shiboken6 import isValid

from rd_catalog.approval_mail import (
    ApprovalMail,
    KitLookup,
    apply_mail_parts,
    mail_field_choices,
    mail_line_transmittal,
    mail_mto_text,
    parse_msg_file,
)
from rd_catalog.approval_mail_revisions import (
    load_packaged_revisions,
    merge_revision_choices,
)
from rd_catalog.context_menu_qt import exec_tracked_menu
from rd_catalog.context_menu_usage import MENU_APPROVAL_MAIL
from rd_catalog.f_journal import (
    JOURNAL_STAGE_LABELS,
    journal_diff_html,
    journal_highlight_spans,
    journal_stage_key,
)
from rd_catalog.approval_mail_dump import (
    build_mail_text_preview,
    build_outlook_dump,
)
from rd_catalog.approval_mail_log import append_approval_mail_ingest, ingest_log_path
from rd_catalog.path_actions import open_path
from rd_catalog.approval_mail_preview import (
    CommentLookup,
    MailPreviewRow,
    build_mail_previews,
    mail_dedupe_key,
    writable_jobs,
)
from rd_catalog.google_f_write import GoogleWriteResult, JournalWriteJob
from rd_catalog.sheet_de_sync_dialog import SHEET_DE_SYNC_BUTTON
from rd_catalog.outlook_item import save_outlook_item
from rd_catalog.outlook_drop_qt import (
    DroppedMsg,
    extract_dropped_messages,
    mime_has_approval_mail,
)

_ROLE_ROW = Qt.ItemDataRole.UserRole
_ROLE_DIFF_SPANS = Qt.ItemDataRole.UserRole + 5
_F_MARK = "#FFE082"
_F_SELECTED = QColor("#D6EAF8")
_F_PAD_X = 3
_F_PAD_Y = 2
_HEADERS = (
    "Файл",
    "Вид",
    "Титул",
    "Марка",
    "Дата",
    "Стадия",
    "Рев.",
    "TRM",
    "MTO",
    "Коды",
    "D/E",
    "Запись",
    "F до",
    "F после",
    "Ошибка",
)
_COL_FILE = 0
_COL_KIND = 1
_COL_TITLE = 2
_COL_MARK = 3
_COL_DATE = 4
_COL_STAGE = 5
_COL_REV = 6
_COL_TRM = 7
_COL_MTO = 8
_COL_CODES = 9
_COL_DE = 10
_COL_WRITE = 11
_COL_F_BEFORE = 12
_COL_F_AFTER = 13
_COL_ERROR = 14
_COLUMN_WIDTHS = {
    _COL_FILE: 150,
    _COL_KIND: 72,
    _COL_TITLE: 72,
    _COL_MARK: 72,
    _COL_DATE: 88,
    _COL_STAGE: 150,
    _COL_REV: 80,
    _COL_TRM: 150,
    _COL_MTO: 72,
    _COL_CODES: 56,
    _COL_DE: 40,
    _COL_WRITE: 80,
    _COL_ERROR: 140,
}
_EDIT_FIELDS = (
    (_COL_TITLE, "title"),
    (_COL_MARK, "mark"),
    (_COL_DATE, "date"),
    (_COL_STAGE, "stage"),
    (_COL_REV, "od_revision"),
    (_COL_TRM, "transmittal"),
    (_COL_MTO, "mto_text"),
)
_COLUMN_TIPS = {
    _COL_TITLE: "Титул строки F. Список — из письма и имён вложений.",
    _COL_MARK: "Марка строки F. Можно выбрать или ввести.",
    _COL_DATE: "Дата события DD.MM.YYYY.",
    _COL_STAGE: "Стадия журнала F (как в строке записи).",
    _COL_REV: (
        "Ревизия OD в строке F. Список — из письма и файла ревизий, "
        "собранного по строкам F с пометкой auto."
    ),
    _COL_TRM: "TRM в строке F.",
    _COL_MTO: "Ревизия MTO, «Нет», или пусто — без суффикса MTO.",
}
_SETTINGS_SUBFOLDERS = "window/approval_mail_subfolders"
_KIND_LABELS = {
    "review_codes": "Коды",
    "tdo_reply": "ТДО",
    "cover_letter": "Сопровод.",
}
_WRITTEN = QColor("#2E7D32")
_FAILED = QColor("#8B1A1A")


class _FDiffDelegate(QStyledItemDelegate):
    """Paint F до/F после with a yellow marker on the character diff."""

    def paint(
        self,
        painter: QPainter,
        option: QStyleOptionViewItem,
        index,
    ) -> None:
        opt = QStyleOptionViewItem(option)
        self.initStyleOption(opt, index)
        painter.save()
        painter.setClipRect(opt.rect)
        selected = bool(opt.state & QStyle.StateFlag.State_Selected)
        if selected:
            painter.fillRect(opt.rect, _F_SELECTED)
        else:
            bg = _role_color(index.data(Qt.ItemDataRole.BackgroundRole))
            painter.fillRect(opt.rect, bg or opt.palette.base())
        doc = self._document(opt, index)
        painter.translate(opt.rect.left() + _F_PAD_X, opt.rect.top() + _F_PAD_Y)
        ctx = QAbstractTextDocumentLayout.PaintContext()
        doc.documentLayout().draw(painter, ctx)
        painter.restore()

    def sizeHint(self, option: QStyleOptionViewItem, index) -> QSize:
        opt = QStyleOptionViewItem(option)
        self.initStyleOption(opt, index)
        doc = self._document(opt, index)
        size = doc.size()
        return QSize(
            int(size.width()) + 2 * _F_PAD_X,
            int(size.height()) + 2 * _F_PAD_Y,
        )

    def _document(self, option: QStyleOptionViewItem, index) -> QTextDocument:
        text = str(index.data(Qt.ItemDataRole.DisplayRole) or "")
        spans = _spans_from_role(index.data(_ROLE_DIFF_SPANS))
        fg = _role_color(index.data(Qt.ItemDataRole.ForegroundRole))
        html = journal_diff_html(
            text,
            spans,
            mark_color=_F_MARK,
            fg=fg.name() if fg is not None else None,
        )
        doc = QTextDocument()
        doc.setDefaultFont(option.font)
        doc.setDocumentMargin(0)
        width = option.rect.width() - 2 * _F_PAD_X
        parent = self.parent()
        if width < 40 and isinstance(parent, QTableWidget):
            width = parent.columnWidth(index.column()) - 2 * _F_PAD_X
        doc.setTextWidth(max(width, 40))
        doc.setHtml(html)
        return doc


def _role_color(value: object) -> QColor | None:
    if value is None:
        return None
    if isinstance(value, QColor):
        return value if value.isValid() else None
    if isinstance(value, QBrush):
        color = value.color()
        return color if color.isValid() else None
    color = QColor(value)
    return color if color.isValid() else None


def _spans_from_role(value: object) -> tuple[tuple[int, int], ...]:
    if not value:
        return ()
    spans: list[tuple[int, int]] = []
    for item in value:
        if isinstance(item, (list, tuple)) and len(item) >= 2:
            spans.append((int(item[0]), int(item[1])))
    return tuple(spans)


class ApprovalMailTab(QWidget):
    """Drop zone and F preview for customer/TDO letters."""

    write_requested = Signal(object)
    de_sync_requested = Signal()
    kit_activated = Signal(str, str)

    def __init__(
        self,
        parent: QWidget | None = None,
        *,
        runtime_dir: str | Path | None = None,
    ) -> None:
        """Build an empty tab.

        Args:
            parent: Optional Qt parent.
            runtime_dir: Folder for temporary Outlook ``.msg`` dumps.
        """

        super().__init__(parent)
        self._comment_lookup: CommentLookup = lambda _t, _m: None
        self._kit_lookup: KitLookup | None = None
        self._dropped: list[DroppedMsg] = []
        self._write_marks: list[tuple[str, str]] = []
        self._parsed_mails: list[ApprovalMail | None] = []
        self._overrides: list[dict[str, str]] = []
        self._pending_indexes: tuple[int, ...] = ()
        self._rows: tuple[MailPreviewRow, ...] = ()
        self._catalog_busy = False
        self._filling = False
        self._skipped_dupes = 0
        self._owned_temp: tempfile.TemporaryDirectory[str] | None = None
        self._runtime_dir = Path(runtime_dir) if runtime_dir is not None else None
        if self._runtime_dir is not None:
            self._temp_dir = self._runtime_dir / "approval_mail_drop"
            self._temp_dir.mkdir(parents=True, exist_ok=True)
        else:
            self._owned_temp = tempfile.TemporaryDirectory(
                prefix="rd_catalog_mail_"
            )
            self._temp_dir = Path(self._owned_temp.name)
        self.setAcceptDrops(True)
        self._build()

    def set_lookups(
        self,
        *,
        comment_lookup: CommentLookup,
        kit_from_transmittal: KitLookup | None,
    ) -> None:
        """Install F and issuance lookups, then rebuild the preview.

        Args:
            comment_lookup: ``None`` means the kit is absent from КСБ ИД.
            kit_from_transmittal: Send-TRM lookup for TDO replies.
        """

        if kit_from_transmittal is not self._kit_lookup:
            self._parsed_mails = [None] * len(self._dropped)
        self._comment_lookup = comment_lookup
        self._kit_lookup = kit_from_transmittal
        self._rebuild()

    def set_catalog_busy(self, busy: bool) -> None:
        """Disable Google write while a catalog worker is running."""

        self._catalog_busy = busy
        self._sync_buttons()

    def restore_settings(self, settings: QSettings) -> None:
        """Load the subfolder-walk checkbox.

        Args:
            settings: Catalog QSettings.
        """

        self._subfolders_box.setChecked(
            _settings_flag(settings, _SETTINGS_SUBFOLDERS, False)
        )

    def save_settings(self, settings: QSettings) -> None:
        """Persist the subfolder-walk checkbox.

        Args:
            settings: Catalog QSettings.
        """

        settings.setValue(
            _SETTINGS_SUBFOLDERS, self._subfolders_box.isChecked()
        )

    def _walk_subfolders(self) -> bool:
        box = getattr(self, "_subfolders_box", None)
        return bool(box is not None and box.isChecked())

    def preview_rows(self) -> tuple[MailPreviewRow, ...]:
        """Return the current preview rows."""

        return self._rows

    def ingest_paths(self, paths: Sequence[str | Path]) -> None:
        """Parse Explorer files or a folder of ``.msg``."""

        from PySide6.QtCore import QMimeData, QUrl

        mime = QMimeData()
        mime.setUrls([QUrl.fromLocalFile(str(Path(path))) for path in paths])
        self.ingest_mime(mime)

    def ingest_mime(self, mime) -> None:
        """Ingest a drop from Explorer or Outlook."""

        dropped = extract_dropped_messages(
            mime, recursive=self._walk_subfolders()
        )
        if not dropped:
            return
        self.ingest_dropped(dropped)

    def ingest_dropped(self, dropped: Sequence[DroppedMsg]) -> None:
        """Ingest already extracted drop items (window-level DnD)."""

        if not dropped:
            return
        if self._catalog_busy:
            QMessageBox.information(
                self,
                "Письма о согласовании",
                "Дождитесь окончания записи в КСБ ИД.",
            )
            return
        accepted, skipped = _filter_identity_dupes(self._dropped, dropped)
        self._skipped_dupes += skipped
        if accepted:
            self._dropped.extend(accepted)
            self._write_marks.extend(("", "") for _ in accepted)
            self._parsed_mails.extend(None for _ in accepted)
            self._overrides.extend({} for _ in accepted)
        self._rebuild()
        if accepted:
            self._log_ingested(accepted)

    def dragEnterEvent(self, event: QDragEnterEvent) -> None:
        if self._catalog_busy:
            event.ignore()
            return
        if mime_has_approval_mail(event.mimeData()):
            event.acceptProposedAction()
        else:
            event.ignore()

    def dropEvent(self, event: QDropEvent) -> None:
        self.ingest_mime(event.mimeData())
        event.acceptProposedAction()

    def _build(self) -> None:
        layout = QVBoxLayout(self)
        hint = QLabel(
            "Перетащите письма Outlook или файлы .msg (и на окно каталога). "
            "Одно письмо = один титул–марка. Несколько марок — выберите одну "
            "в колонках Титул/Марка (парсер старых писем не учим: значения "
            "берём из разбора и имён вложений). "
            "Титул, марка, дата, стадия, рев., TRM и MTO — список найденных "
            "или ручной ввод; F после собирается из этих частей. "
            "Повтор того же письма пропускается. Если строка F уже есть в КСБ ИД — "
            "помечается «уже в F»; D и E дописываются только если они ещё не совпали. "
            "«Папка с .msg» берёт только выбранный каталог; галка "
            "«Обходить подпапки» включает вложенные (например 17_Даты TRM). "
            "Несколько писем на один комплект дописывают F по дате "
            "(уже существующая строка — «уже в F»). "
            "ПКМ → «Текст письма» — тема и тело, чтобы смотреть ошибку. "
            "Записанные строки помечаются и не уходят повторно. "
            "«Убрать записанные» снимает и только что записанные, "
            "и «уже в F» (строка уже была в КСБ ИД). "
            "После успешной записи в Google это происходит само. "
            "Delete / «Удалить выбранные» убирают лишние. "
            "Попавшие письма дописываются в журнал; кнопка «Открыть лог» "
            "показывает тему и текст. "
            "Outlook и каталог должны быть с одинаковыми правами Windows.",
            self,
        )
        hint.setWordWrap(True)
        layout.addWidget(hint)

        buttons = QHBoxLayout()
        folder_btn = QPushButton("Папка с .msg…", self)
        folder_btn.clicked.connect(self._pick_folder)
        files_btn = QPushButton("Файлы .msg…", self)
        files_btn.clicked.connect(self._pick_files)
        self._folder_button = folder_btn
        self._files_button = files_btn
        self._subfolders_box = QCheckBox("Обходить подпапки", self)
        self._subfolders_box.setToolTip(
            "Искать .msg во вложенных каталогах при «Папка с .msg…» "
            "и при перетаскивании папки из Проводника."
        )
        self._write_button = QPushButton("Записать в Google", self)
        write_font = self._write_button.font()
        write_font.setBold(True)
        self._write_button.setFont(write_font)
        self._write_button.setMinimumWidth(160)
        self._write_button.setToolTip(
            "Основное действие вкладки. Остальные кнопки сдвинуты, "
            "чтобы эта оставалась заметной."
        )
        self._write_button.clicked.connect(self._start_write)
        self._write_button.setEnabled(False)
        self._de_sync_button = QPushButton(SHEET_DE_SYNC_BUTTON, self)
        self._de_sync_button.setToolTip(
            "Комплекты, где последнее событие F совпадает с ревизией РД, "
            "а столбцы D/E ещё нет. Запись тем же движком, что письма."
        )
        self._de_sync_button.clicked.connect(self.de_sync_requested.emit)
        self._delete_button = QPushButton("Удалить выбранные", self)
        self._delete_button.clicked.connect(self._delete_selected)
        self._clear_written_button = QPushButton("Убрать записанные", self)
        self._clear_written_button.setToolTip(
            "Убрать из таблицы письма, которые уже в F "
            "(в том числе с прошлого раза) или только что записанные. "
            "Строки с ошибкой и «только D/E» остаются."
        )
        self._clear_written_button.clicked.connect(self._clear_written)
        clear_btn = QPushButton("Очистить", self)
        clear_btn.clicked.connect(self._clear)
        self._clear_button = clear_btn
        self._open_log_button = QPushButton("Открыть лог…", self)
        self._open_log_button.setToolTip(
            "Журнал писем, которые попали на эту вкладку "
            "(тема и текст): runtime_dir/approval_mail_ingest.log."
        )
        self._open_log_button.clicked.connect(self._open_ingest_log)
        self._open_log_button.setEnabled(self._runtime_dir is not None)
        buttons.addWidget(folder_btn)
        buttons.addWidget(self._subfolders_box)
        buttons.addWidget(files_btn)
        buttons.addSpacing(12)
        buttons.addWidget(self._write_button)
        buttons.addSpacing(36)
        buttons.addWidget(self._de_sync_button)
        buttons.addWidget(self._delete_button)
        buttons.addWidget(self._clear_written_button)
        buttons.addWidget(clear_btn)
        buttons.addWidget(self._open_log_button)
        buttons.addStretch(1)
        layout.addLayout(buttons)

        self._status = QLabel("Писем нет.", self)
        layout.addWidget(self._status)

        self._table = QTableWidget(0, len(_HEADERS), self)
        self._table.setHorizontalHeaderLabels(_HEADERS)
        self._table.setSelectionBehavior(
            QAbstractItemView.SelectionBehavior.SelectRows
        )
        self._table.setSelectionMode(
            QAbstractItemView.SelectionMode.ExtendedSelection
        )
        self._table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self._table.verticalHeader().setVisible(False)
        self._table.setWordWrap(True)
        self._table.setTextElideMode(Qt.TextElideMode.ElideNone)
        header = self._table.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        header.setSectionResizeMode(_COL_F_BEFORE, QHeaderView.ResizeMode.Stretch)
        header.setSectionResizeMode(_COL_F_AFTER, QHeaderView.ResizeMode.Stretch)
        header.setStretchLastSection(False)
        header.setMinimumSectionSize(36)
        for column, width in _COLUMN_WIDTHS.items():
            self._table.setColumnWidth(column, width)
        for column, tip in _COLUMN_TIPS.items():
            item = self._table.horizontalHeaderItem(column)
            if item is not None:
                item.setToolTip(tip)
        self._table.cellDoubleClicked.connect(self._on_row_activated)
        self._table.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self._table.customContextMenuRequested.connect(self._show_row_menu)
        diff_delegate = _FDiffDelegate(self._table)
        self._table.setItemDelegateForColumn(_COL_F_BEFORE, diff_delegate)
        self._table.setItemDelegateForColumn(_COL_F_AFTER, diff_delegate)
        delete_shortcut = QShortcut(QKeySequence.StandardKey.Delete, self._table)
        delete_shortcut.activated.connect(self._delete_selected)
        layout.addWidget(self._table, 1)

    @Slot()
    def _pick_folder(self) -> None:
        folder = QFileDialog.getExistingDirectory(self, "Папка с письмами .msg")
        if folder:
            self.ingest_paths([folder])

    @Slot()
    def _pick_files(self) -> None:
        files, _ok = QFileDialog.getOpenFileNames(
            self,
            "Письма Outlook",
            "",
            "Письма Outlook (*.msg);;Все файлы (*.*)",
        )
        if files:
            self.ingest_paths(files)

    @Slot()
    def _clear(self) -> None:
        if self._catalog_busy:
            return
        self._dropped.clear()
        self._write_marks.clear()
        self._parsed_mails.clear()
        self._overrides.clear()
        self._pending_indexes = ()
        self._skipped_dupes = 0
        self._rebuild()

    @Slot()
    def _open_ingest_log(self) -> None:
        if self._runtime_dir is None:
            QMessageBox.information(
                self,
                "Лог писем",
                "Папка журнала не задана.",
            )
            return
        path = ingest_log_path(self._runtime_dir)
        if not path.is_file():
            QMessageBox.information(
                self,
                "Лог писем",
                "Журнал ещё пуст. Он появится после первого письма.\n"
                f"{path}",
            )
            return
        ok, message = open_path(str(path))
        if not ok:
            QMessageBox.warning(self, "Лог писем", message)

    def apply_write_results(self, results: Sequence[GoogleWriteResult]) -> None:
        """Mark preview rows that were sent in the last Google batch.

        Args:
            results: One result per job, same order as the last write request.
        """

        indexes = self._pending_indexes
        self._pending_indexes = ()
        self._ensure_write_marks()
        any_ok = False
        for index, result in zip(indexes, results, strict=False):
            if index < 0 or index >= len(self._write_marks):
                continue
            if result.error:
                self._write_marks[index] = ("failed", result.error)
            else:
                any_ok = True
                self._write_marks[index] = (
                    "written",
                    f"стр. {result.row_index}",
                )
        self._rebuild()
        if any_ok:
            self._remove_written_rows(ignore_busy=True)

    @Slot()
    def _start_write(self) -> None:
        writable = tuple(row for row in self._rows if row.writable)
        if not writable:
            QMessageBox.information(
                self,
                "Письма о согласовании",
                "Нет писем, которые можно записать в КСБ ИД.",
            )
            return
        jobs: tuple[JournalWriteJob, ...] = writable_jobs(writable)
        self._pending_indexes = tuple(
            index for index, row in enumerate(self._rows) if row.writable
        )
        self.write_requested.emit(jobs)

    @Slot()
    def _delete_selected(self) -> None:
        if self._catalog_busy:
            return
        indexes = sorted(
            {item.row() for item in self._table.selectedItems() if item is not None}
        )
        self._delete_indexes(indexes)

    def _delete_indexes(self, indexes: Sequence[int]) -> None:
        drop = {index for index in indexes if 0 <= index < len(self._dropped)}
        if not drop:
            return
        self._dropped = [
            item for index, item in enumerate(self._dropped) if index not in drop
        ]
        self._write_marks = [
            mark for index, mark in enumerate(self._write_marks) if index not in drop
        ]
        self._parsed_mails = [
            mail
            for index, mail in enumerate(self._parsed_mails)
            if index not in drop
        ]
        self._overrides = [
            item
            for index, item in enumerate(self._overrides)
            if index not in drop
        ]
        self._pending_indexes = ()
        self._rebuild()

    @Slot()
    def _clear_written(self) -> None:
        self._remove_written_rows(ignore_busy=False)

    def _remove_written_rows(self, *, ignore_busy: bool = False) -> None:
        if self._catalog_busy and not ignore_busy:
            return
        indexes = [
            index
            for index, row in enumerate(self._rows)
            if row.already_recorded
        ]
        self._delete_indexes(indexes)

    def _rebuild(self) -> None:
        self._ensure_write_marks()
        parsed = self._mails_for_dropped()
        dropped, marks, mails, overrides, letter_skipped = _compact_letter_dupes(
            self._dropped, self._write_marks, parsed, self._overrides
        )
        self._dropped = dropped
        self._write_marks = marks
        self._parsed_mails = list(mails)
        self._overrides = overrides
        self._skipped_dupes += letter_skipped
        built = build_mail_previews(
            self._effective_mails(), comment_lookup=self._comment_lookup
        )
        self._rows = self._with_write_marks(built)
        self._fill_table()
        self._paint_status()
        self._sync_buttons()
        if self._rows:
            current = self._table.currentRow()
            if current < 0 or current >= len(self._rows):
                current = 0
            self._table.selectRow(current)

    def _paint_status(self) -> None:
        writable = sum(1 for row in self._rows if row.writable)
        written = sum(1 for row in self._rows if row.write_status == "written")
        present = sum(
            1
            for row in self._rows
            if row.already_complete and row.write_status != "written"
        )
        errors = sum(
            1
            for row in self._rows
            if row.error or row.write_status == "failed"
        )
        parts = [
            f"Писем: {len(self._rows)}",
            f"к записи: {writable}",
            f"записано: {written}",
            f"ошибок: {errors}",
        ]
        if present:
            parts.insert(3, f"уже в F: {present}")
        if self._skipped_dupes:
            parts.append(f"пропущено дублей: {self._skipped_dupes}")
        self._status.setText(" · ".join(parts))

    def _log_ingested(self, accepted: Sequence[DroppedMsg]) -> None:
        if self._runtime_dir is None or not accepted:
            return
        by_obj = {
            id(dropped): (dropped, mail)
            for dropped, mail in zip(
                self._dropped, self._parsed_mails, strict=False
            )
        }
        stamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        blocks: list[str] = []
        for item in accepted:
            pair = by_obj.get(id(item))
            if pair is None:
                continue
            dropped, mail = pair
            preview = build_mail_text_preview(dropped, mail)
            blocks.append(f"======== {stamp} ========\n{preview.rstrip()}\n")
        append_approval_mail_ingest(self._runtime_dir, blocks)

    def _ensure_write_marks(self) -> None:
        while len(self._write_marks) < len(self._dropped):
            self._write_marks.append(("", ""))
        del self._write_marks[len(self._dropped) :]
        while len(self._overrides) < len(self._dropped):
            self._overrides.append({})
        del self._overrides[len(self._dropped) :]

    def _effective_mails(self) -> list[ApprovalMail]:
        """Return parsed mails with the user's F-part overrides applied."""

        mails = self._mails_for_dropped()
        self._ensure_write_marks()
        return [
            apply_mail_parts(mail, **override) if override else mail
            for mail, override in zip(mails, self._overrides, strict=False)
        ]

    def _with_write_marks(
        self, rows: Sequence[MailPreviewRow]
    ) -> tuple[MailPreviewRow, ...]:
        marked: list[MailPreviewRow] = []
        for row, (status, detail) in zip(rows, self._write_marks, strict=False):
            if status == "written":
                marked.append(
                    replace(
                        row,
                        error="",
                        write_status=status,
                        write_detail=detail,
                    )
                )
            elif status == "failed":
                marked.append(
                    replace(row, write_status=status, write_detail=detail)
                )
            else:
                marked.append(row)
        return tuple(marked)

    def _sync_buttons(self) -> None:
        writable = any(row.writable for row in self._rows)
        has_rows = bool(self._rows)
        has_written = any(row.already_recorded for row in self._rows)
        busy = self._catalog_busy
        self._write_button.setEnabled(writable and not busy)
        self._de_sync_button.setEnabled(not busy)
        self._delete_button.setEnabled(has_rows and not busy)
        self._clear_written_button.setEnabled(has_written and not busy)
        self._clear_button.setEnabled(has_rows and not busy)
        if hasattr(self, "_folder_button"):
            self._folder_button.setEnabled(not busy)
            self._files_button.setEnabled(not busy)
        if hasattr(self, "_subfolders_box"):
            self._subfolders_box.setEnabled(not busy)
        if hasattr(self, "_open_log_button"):
            self._open_log_button.setEnabled(self._runtime_dir is not None)
        if hasattr(self, "_table"):
            for row_index in range(self._table.rowCount()):
                for column, _field in _EDIT_FIELDS:
                    widget = self._table.cellWidget(row_index, column)
                    if widget is not None:
                        widget.setEnabled(not busy and not _row_written(self._rows, row_index))

    def _sync_write_enabled(self) -> None:
        self._sync_buttons()

    def _mails_for_dropped(self) -> list[ApprovalMail]:
        """Return cached parses for ``_dropped``, filling missing slots.

        Returns:
            One ``ApprovalMail`` per dropped item, in the same order.
        """

        n = len(self._dropped)
        while len(self._parsed_mails) < n:
            self._parsed_mails.append(None)
        del self._parsed_mails[n:]
        mails: list[ApprovalMail] = []
        for index, item in enumerate(self._dropped):
            cached = self._parsed_mails[index]
            if cached is None:
                cached = self._parse_dropped(item)
                self._parsed_mails[index] = cached
            mails.append(cached)
        return mails

    def _parse_dropped(self, dropped: DroppedMsg) -> ApprovalMail:
        if dropped.error:
            return _error_mail(dropped.name, dropped.error)
        try:
            if dropped.path:
                return parse_msg_file(
                    dropped.path,
                    kit_from_transmittal=self._kit_lookup,
                )
            if dropped.outlook_entry_id and dropped.outlook_store_id:
                path = self._next_temp_path(dropped.name)
                save_outlook_item(
                    entry_id=dropped.outlook_entry_id,
                    store_id=dropped.outlook_store_id,
                    destination=path,
                )
                return parse_msg_file(
                    path, kit_from_transmittal=self._kit_lookup
                )
            if not dropped.payload:
                return _error_mail(dropped.name, "Пустое содержимое письма.")
            path = self._write_temp(dropped.name, dropped.payload)
            return parse_msg_file(
                path, kit_from_transmittal=self._kit_lookup
            )
        except Exception as exc:
            return _error_mail(
                dropped.name or dropped.path,
                f"{type(exc).__name__}: {exc}",
            )

    def _write_temp(self, name: str, payload: bytes) -> str:
        path = self._next_temp_path(name)
        path.write_bytes(payload)
        return str(path)

    def _next_temp_path(self, name: str) -> Path:
        safe = Path(name or "letter.msg").name
        if not safe.casefold().endswith(".msg"):
            safe += ".msg"
        index = len(list(self._temp_dir.glob("*.msg")))
        return self._temp_dir / f"{index:04d}_{safe}"

    def _fill_table(self) -> None:
        self._filling = True
        try:
            self._table.setRowCount(len(self._rows))
            originals = self._mails_for_dropped()
            for row_index, row in enumerate(self._rows):
                self._write_row_items(row_index, row)
                source = (
                    originals[row_index]
                    if row_index < len(originals)
                    else row.mail
                )
                self._install_part_combos(row_index, source, row.mail)
        finally:
            self._filling = False
        self._table.resizeRowsToContents()

    def _fill_derived_columns(self) -> None:
        self._filling = True
        try:
            for row_index, row in enumerate(self._rows):
                self._write_row_items(row_index, row)
                self._sync_part_combos(row_index, row.mail)
        finally:
            self._filling = False
        self._table.resizeRowsToContents()

    def _write_row_items(self, row_index: int, row: MailPreviewRow) -> None:
        mail = row.mail
        kind = _KIND_LABELS.get(mail.kind, mail.kind or "—")
        de = ""
        if row.patch is not None:
            if row.write_d or row.write_e:
                de = "да"
            elif row.patch.update_de:
                de = "уже"
            else:
                de = "нет"
        if row.write_status == "written":
            write_text = (
                f"записано · {row.write_detail}" if row.write_detail else "записано"
            )
        elif row.write_status == "failed":
            write_text = "не записано"
        elif row.already_complete or row.write_status == "present":
            write_text = "уже в F"
        elif row.already_in_f and (row.write_d or row.write_e):
            write_text = "только D/E"
        else:
            write_text = ""
        f_before = row.comment_before
        f_after = row.patch.comment_after if row.patch is not None else ""
        before_spans, after_spans = journal_highlight_spans(f_before, f_after)
        values = {
            _COL_FILE: Path(mail.source_path).name or mail.subject or "письмо",
            _COL_KIND: kind,
            _COL_TITLE: mail.title,
            _COL_MARK: mail.mark,
            _COL_DATE: mail.date,
            _COL_STAGE: JOURNAL_STAGE_LABELS.get(mail.stage, mail.stage),
            _COL_REV: mail.od_revision,
            _COL_TRM: mail_line_transmittal(mail),
            _COL_MTO: mail_mto_text(mail),
            _COL_CODES: row.letter_counts_text,
            _COL_DE: de,
            _COL_WRITE: write_text,
            _COL_F_BEFORE: f_before,
            _COL_F_AFTER: f_after,
            _COL_ERROR: row.error
            or (row.write_detail if row.write_status == "failed" else ""),
        }
        for column, text in values.items():
            item = self._table.item(row_index, column)
            if item is None:
                item = QTableWidgetItem()
                item.setTextAlignment(
                    Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignLeft
                )
                self._table.setItem(row_index, column, item)
            item.setText(text)
            if column == _COL_FILE:
                item.setData(_ROLE_ROW, row_index)
                item.setToolTip(text)
            if column in {_COL_F_BEFORE, _COL_F_AFTER}:
                item.setToolTip(text)
                spans = before_spans if column == _COL_F_BEFORE else after_spans
                item.setData(
                    _ROLE_DIFF_SPANS,
                    [list(span) for span in spans] if text else [],
                )
            if (
                row.write_status == "written"
                or row.already_complete
                or row.write_status == "present"
            ):
                item.setForeground(_WRITTEN)
            elif row.error or row.write_status == "failed":
                item.setForeground(_FAILED)
            else:
                item.setForeground(self._table.palette().text().color())

    def _install_part_combos(
        self,
        row_index: int,
        source: ApprovalMail,
        current: ApprovalMail,
    ) -> None:
        choices = mail_field_choices(source)
        by_field = {
            "title": (choices.titles, current.title),
            "mark": (choices.marks, current.mark),
            "date": (choices.dates, current.date),
            "stage": (choices.stages, current.stage),
            "od_revision": (choices.revisions, current.od_revision),
            "transmittal": (choices.transmittals, mail_line_transmittal(current)),
            "mto_text": (choices.mto_values, mail_mto_text(current)),
        }
        for column, field in _EDIT_FIELDS:
            options, value = by_field[field]
            if field == "od_revision":
                options = merge_revision_choices(options, load_packaged_revisions())
            combo = _make_part_combo(
                self._table,
                options,
                value,
                stage=field == "stage",
            )
            combo.activated.connect(
                lambda _i, r=row_index, f=field, c=combo: self._on_part_edited(
                    r, f, c
                )
            )
            edit = combo.lineEdit()
            if edit is not None:
                edit.editingFinished.connect(
                    lambda r=row_index, f=field, c=combo: self._on_part_edited(
                        r, f, c
                    )
                )
            combo.setEnabled(
                not self._catalog_busy
                and not _row_written(self._rows, row_index)
            )
            self._table.setCellWidget(row_index, column, combo)

    def _sync_part_combos(self, row_index: int, mail: ApprovalMail) -> None:
        current = {
            "title": mail.title,
            "mark": mail.mark,
            "date": mail.date,
            "stage": mail.stage,
            "od_revision": mail.od_revision,
            "transmittal": mail_line_transmittal(mail),
            "mto_text": mail_mto_text(mail),
        }
        for column, field in _EDIT_FIELDS:
            combo = self._table.cellWidget(row_index, column)
            if not isinstance(combo, QComboBox):
                continue
            _set_combo_value(combo, current[field], stage=field == "stage")

    def _on_part_edited(
        self, row_index: int, field: str, combo: QComboBox
    ) -> None:
        if self._filling or self._catalog_busy:
            return
        if row_index < 0 or row_index >= len(self._dropped):
            return
        if _row_written(self._rows, row_index):
            return
        value = _combo_value(combo, field=field)
        self._ensure_write_marks()
        previous = self._overrides[row_index].get(field)
        if previous == value:
            return
        self._overrides[row_index] = {**self._overrides[row_index], field: value}
        built = build_mail_previews(
            self._effective_mails(), comment_lookup=self._comment_lookup
        )
        self._rows = self._with_write_marks(built)
        self._fill_derived_columns()
        self._paint_status()
        self._sync_buttons()

    def _on_row_activated(self, row_index: int, column: int) -> None:
        if row_index < 0 or row_index >= len(self._rows):
            return
        if column in {item[0] for item in _EDIT_FIELDS}:
            return
        row = self._rows[row_index]
        if row.error or not row.mail.title:
            self._show_text_preview_dialog(row_index)
            return
        if row.mail.title and row.mail.mark:
            self.kit_activated.emit(row.mail.title, row.mail.mark)

    def dump_text_for_row(self, row_index: int) -> str:
        """Return the Outlook dump for one preview row.

        Args:
            row_index: Table row, same order as ingest.

        Returns:
            Empty string when the index is out of range.
        """

        if row_index < 0 or row_index >= len(self._rows):
            return ""
        dropped = (
            self._dropped[row_index]
            if row_index < len(self._dropped)
            else DroppedMsg(name="")
        )
        return build_outlook_dump(dropped, self._rows[row_index].mail)

    def text_preview_for_row(self, row_index: int) -> str:
        """Return the simple letter text for one preview row.

        Args:
            row_index: Table row, same order as ingest.

        Returns:
            Empty string when the index is out of range.
        """

        if row_index < 0 or row_index >= len(self._rows):
            return ""
        dropped = (
            self._dropped[row_index]
            if row_index < len(self._dropped)
            else DroppedMsg(name="")
        )
        return build_mail_text_preview(dropped, self._rows[row_index].mail)

    def _show_row_menu(self, pos: QPoint) -> None:
        item = self._table.itemAt(pos)
        if item is None:
            return
        row_index = item.row()
        if row_index < 0 or row_index >= len(self._rows):
            return
        menu = QMenu(self)
        text_action = menu.addAction("Текст письма…")
        dump_action = menu.addAction("Дамп Outlook…")
        copy_action = menu.addAction("Копировать дамп")
        save_action = menu.addAction("Сохранить дамп…")
        menu.addSeparator()
        delete_action = menu.addAction("Удалить строку")
        delete_action.setEnabled(not self._catalog_busy)
        chosen = exec_tracked_menu(
            menu, MENU_APPROVAL_MAIL, self._table.viewport().mapToGlobal(pos)
        )
        if chosen is text_action:
            self._show_text_preview_dialog(row_index)
        elif chosen is dump_action:
            self._show_dump_dialog(row_index)
        elif chosen is copy_action:
            self._copy_dump(row_index)
        elif chosen is save_action:
            self._save_dump(row_index)
        elif chosen is delete_action:
            self._delete_indexes((row_index,))

    def _show_text_preview_dialog(self, row_index: int) -> None:
        text = self.text_preview_for_row(row_index)
        if not text:
            return
        dialog = _TextViewDialog(
            "Текст письма",
            "Тема и тело так, как их видит разбор. Удобно смотреть ошибку.",
            text,
            self,
        )
        dialog.exec()

    def _show_dump_dialog(self, row_index: int) -> None:
        text = self.dump_text_for_row(row_index)
        if not text:
            return
        dialog = _TextViewDialog(
            "Дамп письма Outlook",
            "То, что каталог получил от Outlook/Проводника и что увидел "
            "extract_msg. Можно вставить в чат.",
            text,
            self,
        )
        dialog.exec()

    def _copy_dump(self, row_index: int) -> None:
        text = self.dump_text_for_row(row_index)
        if not text:
            return
        clipboard = QGuiApplication.clipboard()
        if clipboard is not None:
            clipboard.setText(text)

    def _save_dump(self, row_index: int) -> None:
        text = self.dump_text_for_row(row_index)
        if not text:
            return
        mail = self._rows[row_index].mail
        suggested = Path(mail.source_path or mail.subject or "letter").name
        stem = Path(suggested).stem or "letter"
        default = str(self._temp_dir / f"{stem}_outlook_dump.txt")
        path, _ok = QFileDialog.getSaveFileName(
            self,
            "Сохранить дамп Outlook",
            default,
            "Текст (*.txt);;Все файлы (*.*)",
        )
        if not path:
            return
        Path(path).write_text(text, encoding="utf-8")


class _TextViewDialog(QDialog):
    """Read-only text with copy-to-clipboard."""

    def __init__(
        self,
        title: str,
        hint_text: str,
        text: str,
        parent: QWidget | None = None,
    ) -> None:
        super().__init__(parent)
        self.setWindowTitle(title)
        self.setMinimumSize(720, 520)
        self.resize(860, 640)
        layout = QVBoxLayout(self)
        hint = QLabel(hint_text, self)
        hint.setWordWrap(True)
        layout.addWidget(hint)
        view = QPlainTextEdit(self)
        view.setReadOnly(True)
        view.setPlainText(text)
        layout.addWidget(view, 1)
        buttons = QDialogButtonBox(self)
        copy_btn = buttons.addButton(
            "Копировать", QDialogButtonBox.ButtonRole.ActionRole
        )
        buttons.addButton("Закрыть", QDialogButtonBox.ButtonRole.RejectRole)
        copy_btn.clicked.connect(lambda: _copy_text(text))
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)


def _copy_text(text: str) -> None:
    clipboard = QGuiApplication.clipboard()
    if clipboard is not None:
        clipboard.setText(text)


def _error_mail(label: str, error: str) -> ApprovalMail:
    return ApprovalMail(
        kind="",
        subject=label,
        date="",
        stage="",
        send_transmittal="",
        incoming_transmittal="",
        title="",
        mark="",
        od_revision="",
        kit_code="",
        letter_counts=(),
        f_line="",
        sheet_revision="",
        status_sheet="",
        error=error,
        documents=(),
        source_path=label,
    )


def _drop_identity(dropped: DroppedMsg) -> str:
    """Stable identity of one drop: Outlook IDs, path, or payload hash."""

    store = (dropped.outlook_store_id or "").casefold()
    entry = (dropped.outlook_entry_id or "").casefold()
    if store and entry:
        return f"ol|{store}|{entry}"
    if dropped.path:
        try:
            resolved = str(Path(dropped.path).resolve())
        except OSError:
            resolved = dropped.path
        return f"path|{resolved.casefold()}"
    if dropped.payload:
        digest = hashlib.sha256(dropped.payload).hexdigest()
        return f"sha|{digest}"
    return ""


def _filter_identity_dupes(
    existing: Sequence[DroppedMsg],
    incoming: Sequence[DroppedMsg],
) -> tuple[list[DroppedMsg], int]:
    """Keep incoming drops whose Outlook/path/payload identity is new."""

    seen = {key for item in existing if (key := _drop_identity(item))}
    accepted: list[DroppedMsg] = []
    skipped = 0
    for item in incoming:
        key = _drop_identity(item)
        if key and key in seen:
            skipped += 1
            continue
        if key:
            seen.add(key)
        accepted.append(item)
    return accepted, skipped


def _compact_letter_dupes(
    dropped: Sequence[DroppedMsg],
    marks: Sequence[tuple[str, str]],
    mails: Sequence[ApprovalMail],
    overrides: Sequence[dict[str, str]],
) -> tuple[
    list[DroppedMsg],
    list[tuple[str, str]],
    list[ApprovalMail],
    list[dict[str, str]],
    int,
]:
    """Drop later letters with the same title+mark+F line."""

    kept_dropped: list[DroppedMsg] = []
    kept_marks: list[tuple[str, str]] = []
    kept_mails: list[ApprovalMail] = []
    kept_overrides: list[dict[str, str]] = []
    seen: set[str] = set()
    skipped = 0
    for index, (item, mark, mail) in enumerate(
        zip(dropped, marks, mails, strict=False)
    ):
        override = overrides[index] if index < len(overrides) else {}
        key = mail_dedupe_key(mail)
        if key and key in seen:
            skipped += 1
            continue
        if key:
            seen.add(key)
        kept_dropped.append(item)
        kept_marks.append(mark)
        kept_mails.append(mail)
        kept_overrides.append(dict(override))
    return kept_dropped, kept_marks, kept_mails, kept_overrides, skipped


def _row_written(rows: Sequence[MailPreviewRow], row_index: int) -> bool:
    if row_index < 0 or row_index >= len(rows):
        return False
    return rows[row_index].write_status == "written"


def _combo_value(combo: QComboBox, *, field: str) -> str:
    text = combo.currentText().strip()
    if field != "stage":
        return text
    index = combo.currentIndex()
    if 0 <= index < combo.count() and combo.itemText(index) == combo.currentText():
        data = combo.itemData(index)
        if data:
            return str(data)
        return ""
    return journal_stage_key(text) or text


def _set_combo_value(combo: QComboBox, value: str, *, stage: bool) -> None:
    combo.blockSignals(True)
    completer = combo.completer()
    if completer is not None:
        completer.setCompletionMode(QCompleter.CompletionMode.InlineCompletion)
    try:
        if stage:
            pos = combo.findData(value)
            if pos >= 0:
                combo.setCurrentIndex(pos)
            else:
                combo.setEditText(JOURNAL_STAGE_LABELS.get(value, value))
        elif combo.currentText() != value:
            combo.setEditText(value)
    finally:
        if completer is not None:
            completer.setCompletionMode(QCompleter.CompletionMode.PopupCompletion)
        combo.blockSignals(False)


_NO_ARROW_STYLE = (
    "QComboBox { border: none; padding-right: 2px; background: palette(base); }"
    "QComboBox::drop-down { width: 0px; border: none; }"
    "QComboBox::down-arrow { image: none; width: 0px; height: 0px; }"
)
_COMBO_LIST_TIP = "Двойной щелчок или ↓ — список. Можно ввести своё."


class _ComboListOpener(QObject):
    """Open a combo list that has no drop-down arrow.

    Double-click and Down show the full list. Typing still uses the
    completer popup; Down is left to that popup while it is open.
    """

    def __init__(self, combo: QComboBox) -> None:
        super().__init__(combo)
        self._combo = combo

    def eventFilter(self, watched: QObject, event: QEvent) -> bool:
        del watched
        combo = self._combo
        if not isValid(combo) or not combo.isEnabled():
            return False
        if event.type() == QEvent.Type.MouseButtonDblClick:
            _open_combo_list(combo)
            return True
        if (
            isinstance(event, QKeyEvent)
            and event.key() == Qt.Key.Key_Down
            and event.modifiers() == Qt.KeyboardModifier.NoModifier
        ):
            completer = combo.completer()
            popup = completer.popup() if completer is not None else None
            if popup is not None and popup.isVisible():
                return False
            _open_combo_list(combo)
            return True
        return False


def _open_combo_list(combo: QComboBox) -> None:
    completer = combo.completer()
    if completer is not None:
        completer.popup().hide()
    view = combo.view()
    hint = view.sizeHintForColumn(0)
    view.setMinimumWidth(max(combo.width(), hint + 28))
    combo.showPopup()


def _arm_combo_list(combo: QComboBox) -> None:
    """Hide the arrow and filter the list while the user types."""

    combo.setStyleSheet(_NO_ARROW_STYLE)
    combo.setCursor(Qt.CursorShape.IBeamCursor)
    combo.setToolTip(_COMBO_LIST_TIP)
    edit = combo.lineEdit()
    if edit is None:
        return
    completer = edit.completer()
    if completer is None:
        completer = QCompleter(combo.model(), combo)
        edit.setCompleter(completer)
    completer.setFilterMode(Qt.MatchFlag.MatchContains)
    completer.setCaseSensitivity(Qt.CaseSensitivity.CaseInsensitive)
    completer.setCompletionMode(QCompleter.CompletionMode.PopupCompletion)
    completer.activated.connect(lambda _text: edit.editingFinished.emit())
    opener = _ComboListOpener(combo)
    setattr(combo, "_list_opener", opener)
    edit.installEventFilter(opener)


def _make_part_combo(
    parent: QWidget,
    choices: Sequence[str],
    current: str,
    *,
    stage: bool,
) -> QComboBox:
    combo = QComboBox(parent)
    combo.setEditable(True)
    combo.setInsertPolicy(QComboBox.InsertPolicy.NoInsert)
    combo.setSizeAdjustPolicy(
        QComboBox.SizeAdjustPolicy.AdjustToMinimumContentsLengthWithIcon
    )
    combo.setMinimumContentsLength(10 if len(current) > 8 else 4)
    combo.setFrame(False)
    combo.setMaxVisibleItems(16)
    if stage:
        combo.addItem("", "")
        for key in JOURNAL_STAGE_LABELS:
            combo.addItem(JOURNAL_STAGE_LABELS[key], key)
        pos = combo.findData(current)
        if pos >= 0:
            combo.setCurrentIndex(pos)
        elif current:
            combo.setEditText(JOURNAL_STAGE_LABELS.get(current, current))
    else:
        combo.addItem("")
        seen = {""}
        if current:
            combo.addItem(current)
            seen.add(current.casefold())
        for value in choices:
            key = value.casefold()
            if not value or key in seen:
                continue
            seen.add(key)
            combo.addItem(value)
        combo.setEditText(current)
    _arm_combo_list(combo)
    return combo


def _settings_flag(settings: QSettings, key: str, default: bool) -> bool:
    value = settings.value(key, default)
    if isinstance(value, bool):
        return value
    if isinstance(value, (int, float)):
        return bool(value)
    text = str(value).strip().casefold()
    if text in {"1", "true", "yes", "on"}:
        return True
    if text in {"0", "false", "no", "off", ""}:
        return False
    return default
