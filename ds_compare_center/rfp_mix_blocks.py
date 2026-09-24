"""Collapsible mix/collect pickers for RFP parts and launch tabs."""

from __future__ import annotations

from PySide6.QtCore import Qt, Signal
from PySide6.QtWidgets import (
    QAbstractButton,
    QButtonGroup,
    QFrame,
    QHBoxLayout,
    QLabel,
    QMessageBox,
    QRadioButton,
    QSizePolicy,
    QVBoxLayout,
    QWidget,
)

from RFQ.rfp_parts.ds_rfp_tag_placement import MIX_SEPARATE
from RFQ.tags_rfp_compare.rfp_tags_utils import load_config, save_config
from ds_compare_center.rfp_mix_settings import (
    COLLECT_JOB_CHOICES,
    COLLECT_JOB_HYBRID,
    KEY_COLLECT_JOB,
    KEY_COLLECT_MIX_MODE,
    KEY_LAUNCH_MIX_MODE,
    MIX_CHOICES,
    format_collect_block_title,
    format_mix_block_title,
    read_collect_job,
    read_collect_mix_mode,
    read_launch_mix_mode,
    resolve_collect_job,
    resolve_mix_mode,
    set_rfp_parts_key,
)

_FRAME_STYLE = (
    "QFrame#rfpMixBlock {"
    " border: 1px solid #c8c8c8;"
    " border-radius: 4px;"
    " background: #fafafa;"
    "}"
)
_HINT_STYLE = "color: #555; margin-left: 22px;"
_SECTION_STYLE = "font-weight: 600;"


class _CollapseHeader(QWidget):
    """Clickable header row for a collapsible picker."""

    clicked = Signal()

    def mouseReleaseEvent(self, event) -> None:  # noqa: N802
        self.clicked.emit()
        event.accept()


class _RfpCollapsibleFrame(QFrame):
    """Shared chrome: bordered frame, arrow header, hidden body."""

    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setObjectName("rfpMixBlock")
        self.setStyleSheet(_FRAME_STYLE)
        self.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Maximum)
        self._expanded = False

        root = QVBoxLayout(self)
        root.setContentsMargins(8, 6, 8, 6)
        root.setSpacing(6)

        self._header = _CollapseHeader(self)
        self._header.setCursor(Qt.CursorShape.PointingHandCursor)
        header_row = QHBoxLayout(self._header)
        header_row.setContentsMargins(0, 0, 0, 0)
        header_row.setSpacing(6)
        self._arrow = QLabel("▸", self._header)
        self._arrow.setStyleSheet("font-weight: 600;")
        self._title = QLabel("", self._header)
        self._title.setWordWrap(True)
        self._title.setStyleSheet("font-weight: 600;")
        header_row.addWidget(self._arrow, stretch=0)
        header_row.addWidget(self._title, stretch=1)
        self._header.clicked.connect(self._toggle)
        root.addWidget(self._header)

        self._body = QWidget(self)
        self._body_layout = QVBoxLayout(self._body)
        self._body_layout.setContentsMargins(18, 0, 0, 0)
        self._body_layout.setSpacing(6)
        self._body.hide()
        root.addWidget(self._body)

    def is_expanded(self) -> bool:
        """Return whether the body is visible."""
        return self._expanded

    def _toggle(self) -> None:
        self._expanded = not self._expanded
        self._body.setVisible(self._expanded)
        self._arrow.setText("▾" if self._expanded else "▸")


def _add_choice_radios(
    parent: QWidget,
    layout: QVBoxLayout,
    group: QButtonGroup,
    choices: tuple[tuple[str, str, str], ...],
    property_name: str,
) -> tuple[dict[str, QRadioButton], list[QLabel]]:
    radios: dict[str, QRadioButton] = {}
    hints: list[QLabel] = []
    for value, label, description in choices:
        radio = QRadioButton(label, parent)
        radio.setProperty(property_name, value)
        radios[value] = radio
        group.addButton(radio)
        layout.addWidget(radio)
        hint = QLabel(description, parent)
        hint.setWordWrap(True)
        hint.setStyleSheet(_HINT_STYLE)
        layout.addWidget(hint)
        hints.append(hint)
    return radios, hints


def _persist_parts_key(owner: QWidget, title: str, key: str, value: str) -> bool:
    try:
        config = load_config()
        if not isinstance(config, dict):
            config = {}
        set_rfp_parts_key(config, key, value)
        saved = save_config(config)
    except Exception as exc:
        QMessageBox.critical(owner, title, f"Не удалось сохранить выбор:\n{exc}")
        return False
    if not saved:
        QMessageBox.critical(
            owner,
            title,
            "Функция сохранения сообщила об ошибке. Выбор не сохранён.",
        )
    return bool(saved)


class RfpMixModeBlock(_RfpCollapsibleFrame):
    """Collapsible mix-mode radios that write one JSON key.

    Used on the launch tab (``launch_mix_mode``). The same radios and
    one-line explanations appear inside :class:`RfpCollectJobBlock`.
    """

    mix_changed = Signal(str)

    def __init__(
        self,
        parent: QWidget | None = None,
        *,
        json_key: str = KEY_LAUNCH_MIX_MODE,
        persist_title: str = "Посадка RFP",
    ) -> None:
        super().__init__(parent)
        self._json_key = json_key
        self._persist_title = persist_title
        self._group = QButtonGroup(self)
        self._group.setExclusive(True)
        self._radios, self._hints = _add_choice_radios(
            self._body, self._body_layout, self._group, MIX_CHOICES, "mix_mode"
        )
        self._group.blockSignals(True)
        self._radios[MIX_SEPARATE].setChecked(True)
        self._sync_title(MIX_SEPARATE)
        self._group.blockSignals(False)
        self._group.buttonClicked.connect(self._on_picked)

    def current_mix(self) -> str:
        """Return the selected mix mode."""
        for mode, radio in self._radios.items():
            if radio.isChecked():
                return mode
        return MIX_SEPARATE

    def apply_config(self, config: dict) -> None:
        """Select the mix stored under this block's JSON key without writing."""
        if self._json_key == KEY_COLLECT_MIX_MODE:
            mix = read_collect_mix_mode(config)
        else:
            mix = read_launch_mix_mode(config)
        self._group.blockSignals(True)
        self._radios[mix].setChecked(True)
        self._group.blockSignals(False)
        self._sync_title(mix)

    def set_mix_enabled(self, enabled: bool) -> None:
        """Show mix radios even when disabled (non-hybrid collect job)."""
        for radio in self._radios.values():
            radio.setEnabled(enabled)
        for hint in self._hints:
            hint.setEnabled(enabled)

    def _sync_title(self, mix: str) -> None:
        title = format_mix_block_title(mix)
        self._title.setText(title)
        self._title.setToolTip(title)

    def _on_picked(self, button: QAbstractButton) -> None:
        mix = resolve_mix_mode(button.property("mix_mode"))
        self._sync_title(mix)
        if _persist_parts_key(self, self._persist_title, self._json_key, mix):
            self.mix_changed.emit(mix)


class RfpCollectJobBlock(_RfpCollapsibleFrame):
    """Collapsible collect-job + mix picker for the parts tab.

    Writes ``collect_job`` and ``collect_mix_mode`` independently.
    Mix radios stay visible and are enabled only when the job is hybrid.
    """

    selection_changed = Signal(str, str)

    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        job_caption = QLabel("Сбор", self._body)
        job_caption.setStyleSheet(_SECTION_STYLE)
        self._body_layout.addWidget(job_caption)
        self._job_group = QButtonGroup(self)
        self._job_group.setExclusive(True)
        self._job_radios, _job_hints = _add_choice_radios(
            self._body,
            self._body_layout,
            self._job_group,
            COLLECT_JOB_CHOICES,
            "collect_job",
        )

        mix_caption = QLabel("Посадка RFP", self._body)
        mix_caption.setStyleSheet(_SECTION_STYLE)
        self._body_layout.addWidget(mix_caption)
        self._mix_group = QButtonGroup(self)
        self._mix_group.setExclusive(True)
        self._mix_radios, self._mix_hints = _add_choice_radios(
            self._body,
            self._body_layout,
            self._mix_group,
            MIX_CHOICES,
            "mix_mode",
        )

        self._job_group.blockSignals(True)
        self._mix_group.blockSignals(True)
        self._job_radios[COLLECT_JOB_HYBRID].setChecked(True)
        self._mix_radios[MIX_SEPARATE].setChecked(True)
        self._sync_title(COLLECT_JOB_HYBRID, MIX_SEPARATE)
        self._set_mix_enabled(True)
        self._job_group.blockSignals(False)
        self._mix_group.blockSignals(False)
        self._job_group.buttonClicked.connect(self._on_job_picked)
        self._mix_group.buttonClicked.connect(self._on_mix_picked)

    def current_job(self) -> str:
        """Return the selected collect job."""
        for job, radio in self._job_radios.items():
            if radio.isChecked():
                return job
        return COLLECT_JOB_HYBRID

    def current_mix(self) -> str:
        """Return the selected collect mix mode."""
        for mode, radio in self._mix_radios.items():
            if radio.isChecked():
                return mode
        return MIX_SEPARATE

    def apply_config(self, config: dict) -> None:
        """Select stored collect job/mix without writing the file."""
        job = read_collect_job(config)
        mix = read_collect_mix_mode(config)
        self._job_group.blockSignals(True)
        self._mix_group.blockSignals(True)
        self._job_radios[job].setChecked(True)
        self._mix_radios[mix].setChecked(True)
        self._job_group.blockSignals(False)
        self._mix_group.blockSignals(False)
        self._set_mix_enabled(job == COLLECT_JOB_HYBRID)
        self._sync_title(job, mix)

    def _set_mix_enabled(self, enabled: bool) -> None:
        for radio in self._mix_radios.values():
            radio.setEnabled(enabled)
        for hint in self._mix_hints:
            hint.setEnabled(enabled)

    def _sync_title(self, job: str, mix: str) -> None:
        title = format_collect_block_title(job, mix)
        self._title.setText(title)
        self._title.setToolTip(title)

    def _on_job_picked(self, button: QAbstractButton) -> None:
        job = resolve_collect_job(button.property("collect_job"))
        mix = self.current_mix()
        self._set_mix_enabled(job == COLLECT_JOB_HYBRID)
        self._sync_title(job, mix)
        if _persist_parts_key(self, "Сбор", KEY_COLLECT_JOB, job):
            self.selection_changed.emit(job, mix)

    def _on_mix_picked(self, button: QAbstractButton) -> None:
        mix = resolve_mix_mode(button.property("mix_mode"))
        job = self.current_job()
        self._sync_title(job, mix)
        if _persist_parts_key(self, "Посадка RFP", KEY_COLLECT_MIX_MODE, mix):
            self.selection_changed.emit(job, mix)
