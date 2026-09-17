"""User-editable skip-dir tokens for RD/SQ scans and fast DB prune.

The working list is always ``skip_dirs.json`` under the catalog runtime
directory (never UNC). Packaged tokens in ``default_config.json`` are
only a heal template: if the runtime file is missing or corrupt, it is
rewritten from that template. Matching is a case-insensitive substring
of a directory name, same as ``os.walk`` pruning during UNC scans.
"""

from __future__ import annotations

import json
from collections.abc import Iterable, Sequence
from pathlib import Path
from typing import Any

SKIP_DIRS_FILENAME = "skip_dirs.json"
SKIP_DIRS_VERSION = 1

_GATE_FOLDER_NAMES = (
    "Для передачи",
    "Для_передачи",
    "На_отправку",
    "На_передачу",
)

_PACKAGED_CONFIG_PATH = Path(__file__).with_name("default_config.json")


def is_skipped_dir_name(name: str, skip_dirs: Iterable[str]) -> bool:
    """Return whether a directory name matches any skip token.

    Args:
        name: Single path segment (folder name).
        skip_dirs: Substring tokens; comparison is case-insensitive.

    Returns:
        ``True`` when any token is a substring of ``name``.
    """

    folded = name.casefold()
    return any(skip.casefold() in folded for skip in skip_dirs)


def path_has_skipped_dir(path: str | Path, skip_dirs: Iterable[str]) -> bool:
    """Return whether any parent directory of ``path`` matches skip tokens.

    The filename itself is not tested, matching scan-time ``os.walk`` prune
    (only directory names are skipped).

    Args:
        path: File path (local or UNC).
        skip_dirs: Skip tokens.

    Returns:
        ``True`` when a parent folder should have been pruned.
    """

    tokens = tuple(skip_dirs)
    if not tokens:
        return False
    for part in Path(path).parent.parts:
        if is_skipped_dir_name(part, tokens):
            return True
    return False


def validate_skip_token(token: str, existing: Iterable[str]) -> str:
    """Normalize a skip token and reject empty, duplicate, or gate-hitting values.

    Args:
        token: Raw dialog/config text.
        existing: Tokens already in the list (the token being edited omitted).

    Returns:
        Stripped token.

    Raises:
        ValueError: Russian message when the token is not usable.
    """

    text = str(token or "").strip()
    if not text:
        raise ValueError("Пустой фрагмент skip.")
    needle = text.casefold()
    if any(str(item).casefold() == needle for item in existing):
        raise ValueError(f"«{text}» уже в списке.")
    if any(is_skipped_dir_name(gate, (text,)) for gate in _GATE_FOLDER_NAMES):
        raise ValueError(
            f"«{text}» скрыл бы папку передачи "
            "(Для передачи / На_отправку)."
        )
    return text


def packaged_skip_dirs() -> tuple[str, ...]:
    """Return factory skip tokens from the packaged default config.

    Returns:
        Ordered skip tokens.

    Raises:
        ValueError: If the packaged JSON is missing a valid list.
    """

    raw = json.loads(_PACKAGED_CONFIG_PATH.read_text(encoding="utf-8"))
    tokens = raw.get("skip_dirs") if isinstance(raw, dict) else None
    if not isinstance(tokens, list) or not all(
        isinstance(item, str) and item.strip() for item in tokens
    ):
        raise ValueError("Packaged skip_dirs must be a non-empty string list")
    return tuple(item.strip() for item in tokens)


def load_runtime_skip_dirs(runtime_dir: str | Path) -> tuple[str, ...] | None:
    """Load skip tokens from ``skip_dirs.json`` if present and valid.

    An empty ``tokens`` list is valid (skip nothing). Missing or corrupt
    files return ``None``.

    Args:
        runtime_dir: Catalog runtime directory.

    Returns:
        Tokens, or ``None`` when the file is absent or unusable.
    """

    path = Path(runtime_dir) / SKIP_DIRS_FILENAME
    if not path.is_file():
        return None
    try:
        raw = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError):
        return None
    if not isinstance(raw, dict):
        return None
    tokens = raw.get("tokens")
    if not isinstance(tokens, list):
        return None
    if not all(isinstance(item, str) and item.strip() for item in tokens):
        return None
    return tuple(item.strip() for item in tokens)


def write_skip_dirs_file(path: str | Path, tokens: Sequence[str]) -> None:
    """Atomically write ``skip_dirs.json``.

    Args:
        path: Target JSON path.
        tokens: Ordered skip tokens.
    """

    payload: dict[str, Any] = {
        "version": SKIP_DIRS_VERSION,
        "tokens": [str(item) for item in tokens],
    }
    target = Path(path)
    target.parent.mkdir(parents=True, exist_ok=True)
    text = json.dumps(payload, ensure_ascii=False, indent=2) + "\n"
    tmp_path = target.with_name(target.name + ".tmp")
    tmp_path.write_text(text, encoding="utf-8")
    tmp_path.replace(target)


class SkipDirsStore:
    """Load, edit, and persist the ordered skip-dir token list."""

    def __init__(
        self,
        path: str | Path,
        *,
        packaged: Sequence[str] | None = None,
    ) -> None:
        """Bind to a JSON path and load or heal the runtime file.

        Args:
            path: Target JSON under the catalog runtime directory.
            packaged: Heal template when the file is missing or corrupt.
        """

        self.path = Path(path)
        self._packaged = tuple(packaged) if packaged is not None else packaged_skip_dirs()
        self._tokens: list[str] = list(self._packaged)
        self.load_error: str | None = None
        self.load()

    @classmethod
    def from_runtime_dir(
        cls,
        runtime_dir: str | Path,
        *,
        packaged: Sequence[str] | None = None,
    ) -> SkipDirsStore:
        """Return a store bound to ``skip_dirs.json`` in ``runtime_dir``.

        Args:
            runtime_dir: Catalog runtime directory.
            packaged: Heal template when the file is missing or corrupt.

        Returns:
            Store instance (loads immediately; may write the JSON).
        """

        return cls(
            Path(runtime_dir) / SKIP_DIRS_FILENAME,
            packaged=packaged,
        )

    def tokens(self) -> tuple[str, ...]:
        """Return the current ordered skip tokens."""

        return tuple(self._tokens)

    def load(self) -> None:
        """Load tokens from disk, rewriting factory defaults if needed.

        A valid file, including an empty token list, is kept as-is. A
        missing or corrupt file is replaced with the packaged template.
        """

        self.load_error = None
        loaded = load_runtime_skip_dirs(self.path.parent)
        if loaded is not None:
            self._tokens = list(loaded)
            return
        damaged = self.path.exists()
        self._tokens = list(self._packaged)
        try:
            self.save()
        except OSError as exc:
            self.load_error = f"Не удалось записать skip_dirs.json: {exc}"
            return
        if damaged:
            self.load_error = (
                "skip_dirs.json повреждён — записан заводской список."
            )

    def set_tokens(self, tokens: Sequence[str]) -> None:
        """Replace the list after validating each token.

        Args:
            tokens: New ordered tokens.

        Raises:
            ValueError: If any token is invalid or duplicated.
        """

        normalized: list[str] = []
        for token in tokens:
            normalized.append(validate_skip_token(token, normalized))
        self._tokens = normalized

    def add(self, token: str) -> str:
        """Append a validated token (in memory; call :meth:`save`).

        Args:
            token: Raw token text.

        Returns:
            Normalized token.

        Raises:
            ValueError: If the token is not usable.
        """

        text = validate_skip_token(token, self._tokens)
        self._tokens.append(text)
        return text

    def replace_at(self, index: int, token: str) -> str:
        """Replace one token (in memory; call :meth:`save`).

        Args:
            index: Existing index.
            token: Raw replacement text.

        Returns:
            Normalized token.

        Raises:
            IndexError: If ``index`` is out of range.
            ValueError: If the token is not usable.
        """

        others = [item for i, item in enumerate(self._tokens) if i != index]
        text = validate_skip_token(token, others)
        self._tokens[index] = text
        return text

    def remove_at(self, index: int) -> None:
        """Delete one token (in memory; call :meth:`save`).

        Args:
            index: Existing index.

        Raises:
            IndexError: If ``index`` is out of range.
        """

        del self._tokens[index]

    def move(self, index: int, delta: int) -> int:
        """Move a token by ``delta`` positions.

        Args:
            index: Current index.
            delta: ``-1`` up, ``+1`` down.

        Returns:
            New index.

        Raises:
            IndexError: If ``index`` is out of range.
        """

        new_index = index + delta
        if new_index < 0 or new_index >= len(self._tokens):
            return index
        token = self._tokens.pop(index)
        self._tokens.insert(new_index, token)
        return new_index

    def sort_casefold(self) -> None:
        """Sort tokens case-insensitively (in memory; call :meth:`save`)."""

        self._tokens.sort(key=lambda item: item.casefold())

    def save(self) -> None:
        """Write the current list atomically as UTF-8 JSON."""

        write_skip_dirs_file(self.path, self._tokens)
