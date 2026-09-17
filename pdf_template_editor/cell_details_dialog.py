"""Read-only dialog that shows the full contents of a field-list row."""

from __future__ import annotations

from collections.abc import Sequence

from PySide6.QtWidgets import (
    QDialog,
    QDialogButtonBox,
    QFormLayout,
    QLineEdit,
    QPlainTextEdit,
    QScrollArea,
    QVBoxLayout,
    QWidget,
)


class CellDetailsDialog(QDialog):
    """Display the current cells-panel row in copy-friendly read-only editors."""

    def __init__(
        self,
        title: str,
        rows: Sequence[tuple[str, str, bool]],
        parent=None,
    ) -> None:
        super().__init__(parent)
        self.setWindowTitle(title)
        self.resize(760, 360)

        layout = QVBoxLayout(self)
        scroll = QScrollArea(self)
        scroll.setWidgetResizable(True)
        layout.addWidget(scroll)

        container = QWidget(scroll)
        form = QFormLayout(container)
        form.setFieldGrowthPolicy(QFormLayout.FieldGrowthPolicy.ExpandingFieldsGrow)
        form.setLabelAlignment(form.labelAlignment())
        scroll.setWidget(container)

        for label, value, is_multiline in rows:
            editor = self._build_editor(value, is_multiline)
            form.addRow(f"{label}:", editor)

        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Close, parent=self)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def _build_editor(self, value: str, is_multiline: bool):
        """Create a read-only editor suitable for copying text."""
        if is_multiline:
            editor = QPlainTextEdit(self)
            editor.setReadOnly(True)
            editor.setPlainText(value)
            editor.setMinimumHeight(self._multiline_height(editor, lines=3))
            return editor
        editor = QLineEdit(value, self)
        editor.setReadOnly(True)
        return editor

    @staticmethod
    def _multiline_height(editor: QPlainTextEdit, lines: int) -> int:
        """Return a comfortable default height for a multiline text box."""
        metrics = editor.fontMetrics()
        margins = editor.contentsMargins()
        frame = editor.frameWidth() * 2
        return (metrics.lineSpacing() * lines) + margins.top() + margins.bottom() + frame + 12
