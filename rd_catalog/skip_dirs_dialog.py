"""Dialog to view and edit RD/SQ skip-dir tokens."""

from __future__ import annotations

from PySide6.QtCore import Signal, Slot
from PySide6.QtWidgets import (
    QAbstractItemView,
    QDialog,
    QDialogButtonBox,
    QHBoxLayout,
    QInputDialog,
    QLabel,
    QListWidget,
    QMessageBox,
    QPushButton,
    QVBoxLayout,
    QWidget,
)

from rd_catalog.skip_dirs import SkipDirsStore, validate_skip_token


class SkipDirsDialog(QDialog):
    """Edit the persisted skip-dir list; save writes JSON, not UNC."""

    changed = Signal()

    def __init__(
        self,
        store: SkipDirsStore,
        parent: QWidget | None = None,
    ) -> None:
        """Build the dialog bound to an existing store.

        Args:
            store: Skip-dir persistence object.
            parent: Optional Qt parent.
        """

        super().__init__(parent)
        self._store = store
        self._draft = list(store.tokens())
        self.setWindowTitle("Skip-папки")
        self.setMinimumSize(480, 420)
        self.resize(560, 520)
        self._build()
        self._reload_list()

    def _build(self) -> None:
        layout = QVBoxLayout(self)
        hint = QLabel(
            "Фрагменты имён папок: скан UNC в них не заходит (подстрока, "
            "без учёта регистра). Список — локальный skip_dirs.json, не сеть. "
            "После сохранения дерево прячет совпадения сразу; "
            "«Убрать skip из дерева» помечает их в базе без скана UNC.",
            self,
        )
        hint.setWordWrap(True)
        layout.addWidget(hint)

        self._list = QListWidget(self)
        self._list.setSelectionMode(
            QAbstractItemView.SelectionMode.SingleSelection
        )
        self._list.setAlternatingRowColors(True)
        layout.addWidget(self._list, 1)

        row = QHBoxLayout()
        self._edit_button = QPushButton("Изменить", self)
        self._remove_button = QPushButton("Удалить", self)
        self._up_button = QPushButton("Вверх", self)
        self._down_button = QPushButton("Вниз", self)
        add_button = QPushButton("Добавить", self)
        sort_button = QPushButton("А→Я", self)
        add_button.clicked.connect(self._on_add)
        self._edit_button.clicked.connect(self._on_edit)
        self._remove_button.clicked.connect(self._on_remove)
        self._up_button.clicked.connect(self._on_up)
        self._down_button.clicked.connect(self._on_down)
        sort_button.clicked.connect(self._on_sort)
        for button in (
            add_button,
            self._edit_button,
            self._remove_button,
            self._up_button,
            self._down_button,
            sort_button,
        ):
            row.addWidget(button)
        layout.addLayout(row)

        box = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Save
            | QDialogButtonBox.StandardButton.Cancel,
            self,
        )
        box.accepted.connect(self._on_save)
        box.rejected.connect(self.reject)
        layout.addWidget(box)
        self._list.itemSelectionChanged.connect(self._update_row_actions)
        self._update_row_actions()

    def _reload_list(self, select: int | None = None) -> None:
        current = self._list.currentRow()
        self._list.clear()
        self._list.addItems(self._draft)
        if self._draft:
            index = current if select is None else select
            index = max(0, min(index, len(self._draft) - 1))
            self._list.setCurrentRow(index)
        self._update_row_actions()

    def _selected_index(self) -> int:
        return self._list.currentRow()

    @Slot()
    def _update_row_actions(self) -> None:
        index = self._selected_index()
        has_row = index >= 0
        self._edit_button.setEnabled(has_row)
        self._remove_button.setEnabled(has_row)
        self._up_button.setEnabled(index > 0)
        self._down_button.setEnabled(has_row and index < len(self._draft) - 1)

    def _prompt_token(self, title: str, initial: str = "") -> str | None:
        text, ok = QInputDialog.getText(
            self, title, "Фрагмент имени папки:", text=initial
        )
        if not ok:
            return None
        return text

    @Slot()
    def _on_add(self) -> None:
        raw = self._prompt_token("Добавить skip")
        if raw is None:
            return
        try:
            token = validate_skip_token(raw, self._draft)
        except ValueError as exc:
            QMessageBox.warning(self, "Skip-папки", str(exc))
            return
        self._draft.append(token)
        self._reload_list(select=len(self._draft) - 1)

    @Slot()
    def _on_edit(self) -> None:
        index = self._selected_index()
        if index < 0:
            return
        raw = self._prompt_token("Изменить skip", self._draft[index])
        if raw is None:
            return
        others = [item for i, item in enumerate(self._draft) if i != index]
        try:
            token = validate_skip_token(raw, others)
        except ValueError as exc:
            QMessageBox.warning(self, "Skip-папки", str(exc))
            return
        self._draft[index] = token
        self._reload_list(select=index)

    @Slot()
    def _on_remove(self) -> None:
        index = self._selected_index()
        if index < 0:
            return
        del self._draft[index]
        self._reload_list()

    @Slot()
    def _on_up(self) -> None:
        index = self._selected_index()
        if index <= 0:
            return
        self._draft[index - 1], self._draft[index] = (
            self._draft[index],
            self._draft[index - 1],
        )
        self._reload_list(select=index - 1)

    @Slot()
    def _on_down(self) -> None:
        index = self._selected_index()
        if index < 0 or index >= len(self._draft) - 1:
            return
        self._draft[index + 1], self._draft[index] = (
            self._draft[index],
            self._draft[index + 1],
        )
        self._reload_list(select=index + 1)

    @Slot()
    def _on_sort(self) -> None:
        self._draft.sort(key=lambda item: item.casefold())
        self._reload_list(select=0 if self._draft else None)

    @Slot()
    def _on_save(self) -> None:
        try:
            self._store.set_tokens(self._draft)
            self._store.save()
        except (ValueError, OSError) as exc:
            QMessageBox.warning(self, "Skip-папки", str(exc))
            return
        self.changed.emit()
        self.accept()
