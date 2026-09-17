"""Preview dialog: F до / editable F после for RD-folder legalize.

Qt monitor. Does not call the Sheets API. The parent starts
:class:`~rd_catalog.google_f_write_thread.GoogleFWriteThread` with the
returned :class:`~rd_catalog.google_f_write.JournalWriteJob`.
"""

from __future__ import annotations

from PySide6.QtWidgets import (
    QDialog,
    QDialogButtonBox,
    QLabel,
    QMessageBox,
    QPlainTextEdit,
    QTextBrowser,
    QVBoxLayout,
    QWidget,
)

from rd_catalog.f_journal import JournalPatch, journal_diff_html, journal_highlight_spans
from rd_catalog.f_legalize import (
    LEGALIZE_APPROVAL_STAGE,
    LEGALIZE_APPROVAL_TITLE,
    LEGALIZE_APPROVAL_TOKEN,
    job_from_edited_comment,
)
from rd_catalog.google_f_write import JournalWriteJob
from rd_catalog.kits import parse_history_line

_F_MARK = "#FFE082"


class FLegalizeDialog(QDialog):
    """Show F до read-only and F после editable, then return a write job."""

    def __init__(
        self,
        *,
        title: str,
        mark: str,
        revision: str,
        patch: JournalPatch,
        parent: QWidget | None = None,
    ) -> None:
        """Prefill F до / F после from a journal patch.

        Args:
            title: Four-digit title.
            mark: Latin AGCC mark.
            revision: Folder filename revision for column D.
            patch: Preview from :func:`build_journal_patch`.
            parent: Optional Qt parent (the catalog window).
        """

        super().__init__(parent)
        self._title = title
        self._mark = mark
        self._revision = revision
        self._patch = patch
        parsed = parse_history_line(patch.f_line)
        self._stage = (
            parsed.stage
            if parsed.stage and parsed.stage != "other"
            else LEGALIZE_APPROVAL_STAGE
        )
        self._job: JournalWriteJob | None = None
        self.setWindowTitle(LEGALIZE_APPROVAL_TITLE)
        self.setMinimumSize(720, 520)
        self.resize(860, 640)
        layout = QVBoxLayout(self)
        hint = QLabel(
            f"{title}-{mark} · рев. {revision}\n"
            "Пишет код А в столбец F листа «Контроль выдачи» "
            f"(вместо TRM: {LEGALIZE_APPROVAL_TOKEN}). "
            "Править можно только «F после». Строку комплекта не создаём.",
            self,
        )
        hint.setWordWrap(True)
        layout.addWidget(hint)
        de_parts: list[str] = []
        if patch.sheet_revision:
            de_parts.append(f"D: {patch.sheet_revision}")
        if patch.status_sheet:
            de_parts.append(f"E: {patch.status_sheet}")
        if de_parts:
            de_hint = QLabel(" · ".join(de_parts), self)
            de_hint.setWordWrap(True)
            layout.addWidget(de_hint)
        before_label = QLabel("F до", self)
        before_label.setStyleSheet("font-weight: bold;")
        layout.addWidget(before_label)
        before_spans, _after_spans = journal_highlight_spans(
            patch.comment_before, patch.comment_after
        )
        self._before = QTextBrowser(self)
        self._before.setReadOnly(True)
        self._before.setOpenExternalLinks(False)
        self._before.setHtml(
            journal_diff_html(
                patch.comment_before or "—",
                before_spans,
                mark_color=_F_MARK,
            )
        )
        layout.addWidget(self._before, 1)
        after_label = QLabel("F после", self)
        after_label.setStyleSheet("font-weight: bold;")
        layout.addWidget(after_label)
        self._after = QPlainTextEdit(self)
        self._after.setPlainText(patch.comment_after)
        self._after.setPlaceholderText("F после")
        layout.addWidget(self._after, 1)
        buttons = QDialogButtonBox(self)
        buttons.addButton(
            "Записать в Google", QDialogButtonBox.ButtonRole.AcceptRole
        )
        buttons.addButton("Отмена", QDialogButtonBox.ButtonRole.RejectRole)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def job(self) -> JournalWriteJob | None:
        """Return the confirmed write job after Accept, else ``None``."""

        return self._job

    def accept(self) -> None:
        """Validate the edited F после, then close."""

        try:
            self._job = job_from_edited_comment(
                title=self._title,
                mark=self._mark,
                comment_before=self._patch.comment_before,
                comment_after=self._after.toPlainText(),
                fallback_line=self._patch.f_line,
                fallback_revision=self._revision,
                fallback_stage=self._stage,
            )
        except ValueError as exc:
            QMessageBox.warning(self, LEGALIZE_APPROVAL_TITLE, str(exc))
            return
        super().accept()
