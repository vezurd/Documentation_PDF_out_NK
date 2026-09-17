"""User-managed ban list of title+mark pairs hidden from catalog views.

Persists under the catalog runtime directory (never UNC). Identity is the
same case-insensitive ``(title, mark)`` key as the Google kits matrix.
"""

from __future__ import annotations

import json
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

from rd_catalog.kits import kit_identity_key
from rd_catalog.parse import normalize_unicode_dashes

BAN_FILTER_FILENAME = "banned_title_marks.json"
BAN_FILTER_VERSION = 1


@dataclass(frozen=True, slots=True)
class BannedTitleMark:
    """One hidden title+mark pair."""

    title: str
    mark: str
    comment: str = ""
    added_at: str = ""

    @property
    def identity(self) -> tuple[str, str]:
        """Return the case-insensitive kit identity."""

        return kit_identity_key(self.title, self.mark)

    @property
    def label(self) -> str:
        """Return ``TITLE-MARK`` for logs and dialogs."""

        return f"{self.title}-{self.mark}"


def parse_title_mark(title: str, mark: str = "") -> tuple[str, str]:
    """Normalize a title/mark pair from dialog or context-menu input.

    Accepts two fields, or a combined ``5850-SKUD`` (including Unicode dashes)
    in ``title`` when ``mark`` is empty.

    Args:
        title: Title digits, or ``title-mark`` when mark is omitted.
        mark: Latin AGCC mark; optional if ``title`` already contains it.

    Returns:
        Stripped ``(title, mark)``.

    Raises:
        ValueError: If either side is empty after parsing.
    """

    title_text = normalize_unicode_dashes(str(title or "")).strip()
    mark_text = normalize_unicode_dashes(str(mark or "")).strip()
    if not mark_text and "-" in title_text:
        left, right = title_text.split("-", 1)
        title_text, mark_text = left.strip(), right.strip()
    if not title_text or not mark_text:
        raise ValueError("Нужны и титул, и марка (например 5850 и SKUD).")
    return title_text, mark_text


def title_mark_from_row(
    title: Any = None,
    mark: Any = None,
    title_system: Any = None,
) -> tuple[str, str]:
    """Extract title and mark from a kits/MTO row payload.

    Args:
        title: Title field, if present.
        mark: Mark field, if present.
        title_system: Combined ``####-MARK`` fallback.

    Returns:
        ``(title, mark)``, possibly empty strings when unknown.
    """

    title_text = str(title or "").strip()
    mark_text = str(mark or "").strip()
    if title_text and mark_text:
        return title_text, mark_text
    if title_system:
        try:
            parsed_title, parsed_mark = parse_title_mark(str(title_system), "")
        except ValueError:
            parsed_title, parsed_mark = "", ""
        return title_text or parsed_title, mark_text or parsed_mark
    return title_text, mark_text


class BanFilterStore:
    """Load, query, and persist banned title+mark pairs as JSON."""

    def __init__(self, path: str | Path) -> None:
        """Bind to a JSON path and load existing pairs if the file is present.

        Args:
            path: Target JSON file under the catalog runtime directory.
        """

        self.path = Path(path)
        self._pairs: dict[tuple[str, str], BannedTitleMark] = {}
        self._revision = 0
        self.load_error: str | None = None
        self.load()

    @classmethod
    def from_runtime_dir(cls, runtime_dir: str | Path) -> BanFilterStore:
        """Return a store bound to ``banned_title_marks.json`` in ``runtime_dir``.

        Args:
            runtime_dir: Catalog runtime directory.

        Returns:
            Store instance (loads immediately).
        """

        return cls(Path(runtime_dir) / BAN_FILTER_FILENAME)

    @property
    def revision(self) -> int:
        """Monotonic counter; bumps when the in-memory pair set changes."""

        return self._revision

    def load(self) -> None:
        """Replace in-memory pairs from disk. Missing file yields an empty list."""

        self._revision += 1
        self.load_error = None
        self._pairs = {}
        if not self.path.exists():
            return
        try:
            raw = json.loads(self.path.read_text(encoding="utf-8"))
        except (OSError, UnicodeError, json.JSONDecodeError) as exc:
            self.load_error = f"Не удалось прочитать бан-фильтр: {exc}"
            return
        items: list[Any]
        if isinstance(raw, dict):
            items = list(raw.get("pairs") or ())
        elif isinstance(raw, list):
            items = raw
        else:
            self.load_error = "Бан-фильтр: неожиданная форма JSON."
            return
        for item in items:
            pair = _pair_from_json(item)
            if pair is not None:
                self._pairs[pair.identity] = pair

    def pairs(self) -> tuple[BannedTitleMark, ...]:
        """Return banned pairs sorted by title, then mark."""

        return tuple(
            sorted(self._pairs.values(), key=lambda pair: (pair.title, pair.mark))
        )

    def contains(self, title: str, mark: str) -> bool:
        """Return whether ``(title, mark)`` is banned.

        Args:
            title: Title text.
            mark: Mark text.

        Returns:
            ``True`` when the pair is on the list.
        """

        title_text = str(title or "").strip()
        mark_text = str(mark or "").strip()
        if not title_text or not mark_text:
            return False
        return kit_identity_key(title_text, mark_text) in self._pairs

    def add(
        self, title: str, mark: str, comment: str = ""
    ) -> tuple[BannedTitleMark, bool]:
        """Add a pair and persist. Duplicate identity is a no-op.

        Args:
            title: Title text.
            mark: Mark text.
            comment: Optional reason (cancelled title, etc.).

        Returns:
            The stored pair and whether it was newly inserted.

        Raises:
            ValueError: If title or mark is empty.
            OSError: If the JSON file cannot be written.
        """

        title_text, mark_text = parse_title_mark(title, mark)
        key = kit_identity_key(title_text, mark_text)
        existing = self._pairs.get(key)
        if existing is not None:
            return existing, False
        pair = BannedTitleMark(
            title=title_text,
            mark=mark_text,
            comment=str(comment or "").strip(),
            added_at=_utc_now(),
        )
        self._pairs[key] = pair
        self._revision += 1
        self.save()
        return pair, True

    def remove(self, title: str, mark: str) -> bool:
        """Remove a pair and persist.

        Args:
            title: Title text.
            mark: Mark text.

        Returns:
            ``True`` if the pair was present.

        Raises:
            OSError: If the JSON file cannot be written.
        """

        title_text = str(title or "").strip()
        mark_text = str(mark or "").strip()
        if not title_text or not mark_text:
            return False
        key = kit_identity_key(title_text, mark_text)
        if key not in self._pairs:
            return False
        del self._pairs[key]
        self._revision += 1
        self.save()
        return True

    def save(self) -> None:
        """Write the current list atomically as UTF-8 JSON."""

        payload = {
            "version": BAN_FILTER_VERSION,
            "pairs": [
                {
                    "title": pair.title,
                    "mark": pair.mark,
                    "comment": pair.comment,
                    "added_at": pair.added_at,
                }
                for pair in self.pairs()
            ],
        }
        self.path.parent.mkdir(parents=True, exist_ok=True)
        text = json.dumps(payload, ensure_ascii=False, indent=2) + "\n"
        tmp_path = self.path.with_name(self.path.name + ".tmp")
        tmp_path.write_text(text, encoding="utf-8")
        tmp_path.replace(self.path)


def _utc_now() -> str:
    return datetime.now(timezone.utc).isoformat()


def _pair_from_json(item: Any) -> BannedTitleMark | None:
    if not isinstance(item, dict):
        return None
    title = str(item.get("title") or "").strip()
    mark = str(item.get("mark") or "").strip()
    if not title or not mark:
        return None
    return BannedTitleMark(
        title=title,
        mark=mark,
        comment=str(item.get("comment") or "").strip(),
        added_at=str(item.get("added_at") or "").strip(),
    )
