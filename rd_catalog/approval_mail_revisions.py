"""OD revision vocabulary for the approval-mail «Рев.» list.

The packaged JSON is a one-time extract of column F lines that end with
``auto`` (catalog writes). Hand-typed F lines stay out, so a typo in the
sheet does not become a suggested revision. The mail tab reads the file;
it does not rescan the journal.
"""

from __future__ import annotations

import json
from collections.abc import Iterable
from functools import lru_cache
from pathlib import Path

from rd_catalog.kits import format_revision, parse_history_line, parse_sheet_revision
from rd_catalog.overlay import revision_rank

_PACKAGED = Path(__file__).with_name("approval_mail_revisions.json")


def revisions_from_auto_lines(lines: Iterable[str]) -> tuple[str, ...]:
    """Return unique OD revisions from F lines marked ``auto``.

    Args:
        lines: Raw column-F lines (one event each).

    Returns:
        Revision texts such as ``02-AN01``, ordered like
        :func:`merge_revision_choices`. Lines without ``auto``, and auto
        lines with no OD revision, are skipped.
    """

    texts: list[str] = []
    for line in lines:
        event = parse_history_line(line)
        if not event.from_robot_auto:
            continue
        text = format_revision(event.revision, event.appendix)
        if text:
            texts.append(text)
    return merge_revision_choices(texts)


def merge_revision_choices(*groups: Iterable[str]) -> tuple[str, ...]:
    """Return unique revision texts, numeric ranks first.

    Args:
        groups: Candidate strings. Empty values are dropped. The first
            spelling of a case-fold duplicate is kept.

    Returns:
        Texts ordered by revision rank. Numeric tokens (``0``, ``0-AN01``,
        ``01``, …) come before letter tokens such as ``A``.
    """

    seen: set[str] = set()
    texts: list[str] = []
    for group in groups:
        for value in group:
            text = (value or "").strip()
            key = text.casefold()
            if not text or key in seen:
                continue
            seen.add(key)
            texts.append(text)
    texts.sort(key=_revision_sort_key)
    return tuple(texts)


def _revision_sort_key(text: str) -> tuple[int, tuple[int, int, str]]:
    revision, appendix = parse_sheet_revision(text)
    rank = revision_rank(revision, appendix)
    numeric_first = 0 if rank[0] >= 0 else 1
    return (numeric_first, rank)


@lru_cache(maxsize=1)
def load_packaged_revisions() -> tuple[str, ...]:
    """Return the packaged OD revision list.

    Returns:
        Revision texts from ``approval_mail_revisions.json``. Missing or
        unreadable file yields an empty tuple.
    """

    try:
        payload = json.loads(_PACKAGED.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError, UnicodeError):
        return ()
    raw = payload.get("revisions") if isinstance(payload, dict) else None
    if not isinstance(raw, list):
        return ()
    return merge_revision_choices(str(item) for item in raw if isinstance(item, str))
