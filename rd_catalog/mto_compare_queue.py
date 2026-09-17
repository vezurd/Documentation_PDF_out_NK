"""Build an incremental MTO compare plan from overlay, robot files, and cache."""

from __future__ import annotations

from collections.abc import Iterable
from dataclasses import dataclass

from rd_catalog.db import CatalogDatabase
from rd_catalog.models import ParsedFile
from rd_catalog.mto_diff import MtoPair, pair_current_mto
from rd_catalog.overlay import OverlayResult

MtoDocumentKey = tuple[str, str]


@dataclass(frozen=True, slots=True)
class MtoComparePlan:
    """Ordered pairs that still need XLSX content comparison."""

    pairs: tuple[MtoPair, ...]
    skipped_cache_hits: int
    scope_limited: bool

    @property
    def document_keys(self) -> tuple[MtoDocumentKey, ...]:
        return tuple(pair.key for pair in self.pairs)


def build_mto_compare_plan(
    database: CatalogDatabase,
    rd_overlay: OverlayResult,
    robot_files: Iterable[ParsedFile],
    *,
    priority_keys: Iterable[MtoDocumentKey] = (),
    scope_keys: Iterable[MtoDocumentKey] | None = None,
) -> MtoComparePlan:
    """Return pairs that need work, priority keys first.

    Args:
        database: Catalog SQLite (files already persisted).
        rd_overlay: Current RD MTO overlay.
        robot_files: Present robot MTO files.
        priority_keys: Keys from the latest scan walk; sorted first among
            work items.
        scope_keys: If not None, only these MTO keys (subtree scan).

    Returns:
        Plan with cache-valid pairs omitted.
    """

    scope_set = None if scope_keys is None else set(scope_keys)
    pairing = pair_current_mto(rd_overlay, robot_files)
    in_scope: list[tuple[int, MtoPair]] = []
    for index, pair in enumerate(pairing.pairs):
        if scope_set is not None and pair.key not in scope_set:
            continue
        in_scope.append((index, pair))

    needs_work: list[tuple[int, MtoPair]] = []
    skipped_cache_hits = 0
    for index, pair in in_scope:
        if database.mto_pair_cache_valid(pair):
            skipped_cache_hits += 1
        else:
            needs_work.append((index, pair))

    priority_set = set(priority_keys)
    needs_work.sort(
        key=lambda item: (0 if item[1].key in priority_set else 1, item[0])
    )
    return MtoComparePlan(
        pairs=tuple(pair for _, pair in needs_work),
        skipped_cache_hits=skipped_cache_hits,
        scope_limited=scope_keys is not None,
    )
