"""Confirmation dialog for moving an SQ kit folder into RD."""

from __future__ import annotations

from PySide6.QtWidgets import (
    QDialog,
    QDialogButtonBox,
    QLabel,
    QPlainTextEdit,
    QVBoxLayout,
    QWidget,
)

from rd_catalog.sq_to_rd import SqToRdPlan, plan_summary_text


class SqToRdDialog(QDialog):
    """Show source SQ folder and the new RD transfer path before moving."""

    def __init__(self, plan: SqToRdPlan, parent: QWidget | None = None) -> None:
        """Build the confirmation dialog.

        Args:
            plan: Move plan for one title+mark kit.
            parent: Optional Qt parent.
        """

        super().__init__(parent)
        self._plan = plan
        self.setWindowTitle("Перенести SQ в РД")
        self.setMinimumSize(720, 420)
        self.resize(820, 520)
        self._build()

    def _build(self) -> None:
        layout = QVBoxLayout(self)
        hint = QLabel(
            "Папка SQ будет перемещена в новую передачу РД "
            "«NN_рев.<ревизия>_от_ГГГГ.ММ.ДД» (дата самого свежего файла). "
            "Это запись в РД и SQ; отмена после подтверждения невозможна.",
            self,
        )
        hint.setWordWrap(True)
        layout.addWidget(hint)
        details = QPlainTextEdit(self)
        details.setReadOnly(True)
        details.setPlainText(plan_summary_text(self._plan))
        layout.addWidget(details, 1)
        buttons = QDialogButtonBox(self)
        buttons.addButton(
            "Подтвердить перенос", QDialogButtonBox.ButtonRole.AcceptRole
        )
        buttons.addButton("Отмена", QDialogButtonBox.ButtonRole.RejectRole)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)
