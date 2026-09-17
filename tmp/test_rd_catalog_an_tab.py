"""Offscreen checks for АН tab kit-jump date sort."""

from __future__ import annotations

import os
import sys
from datetime import datetime
from pathlib import Path

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from PySide6.QtCore import Qt
from PySide6.QtWidgets import QApplication

from rd_catalog.an_compare import AN_HEADERS
from rd_catalog.an_index import AnMtoFile, KitAnTargets
from rd_catalog.an_tab import AnTab
from rd_catalog.kits import kit_identity_key
from rd_catalog.models import make_path_key

_COL_DATE = AN_HEADERS.index("Дата")
_COL_REV = AN_HEADERS.index("Ревизия АН")


def _mtime_ns(year: int, month: int, day: int) -> int:
    return int(datetime(year, month, day).timestamp() * 1_000_000_000)


def _an_file(
    *,
    title: str,
    mark: str,
    revision_text: str,
    mtime_ns: int,
) -> AnMtoFile:
    name = f"AGCC.287-{title}-{mark}.MTO-0001_{revision_text}_RU.xlsx"
    path = rf"\\bcc\eng\AN\{title}\{name}"
    return AnMtoFile(
        path=path,
        path_key=make_path_key(path),
        title=title,
        mark=mark,
        revision_text=revision_text,
        core_stem=f"AGCC.287-{title}-{mark}.MTO-0001",
        discipline_block="MTO-0001",
        name=name,
        parent_dir=rf"\\bcc\eng\AN\{title}",
        mtime_ns=mtime_ns,
        size=1,
    )


def _visible_revisions(tab: AnTab) -> list[str]:
    table = tab.table()
    out: list[str] = []
    for index in range(table.rowCount()):
        if table.isRowHidden(index):
            continue
        item = table.item(index, _COL_REV)
        out.append("" if item is None else item.text())
    return out


def main() -> None:
    """Jump from Комплекты must sort visible АН rows by file date."""

    instance = QApplication.instance()
    if instance is None:
        QApplication([])

    newest = _an_file(
        title="2873",
        mark="SOS",
        revision_text="01-AN01",
        mtime_ns=_mtime_ns(2026, 6, 16),
    )
    oldest = _an_file(
        title="2873",
        mark="SOS",
        revision_text="0-AN01",
        mtime_ns=_mtime_ns(2025, 6, 16),
    )
    middle = _an_file(
        title="2873",
        mark="SOS",
        revision_text="02",
        mtime_ns=_mtime_ns(2025, 11, 26),
    )
    decoy = _an_file(
        title="1600",
        mark="POS",
        revision_text="99",
        mtime_ns=_mtime_ns(2024, 1, 1),
    )

    tab = AnTab()
    sos_key = kit_identity_key("2873", "SOS")
    pos_key = kit_identity_key("1600", "POS")
    tab.set_rows(
        (newest, oldest, middle, decoy),
        is_banned=lambda _title, _mark: False,
        kit_targets={sos_key: KitAnTargets(), pos_key: KitAnTargets()},
        allowed_kits={sos_key, pos_key},
    )
    table = tab.table()
    table.sortByColumn(_COL_REV, Qt.SortOrder.AscendingOrder)
    tab._not_in_kits.setChecked(True)

    tab.set_kit_filter("2873", "SOS")

    header = table.horizontalHeader()
    assert header.sortIndicatorSection() == _COL_DATE
    assert header.sortIndicatorOrder() == Qt.SortOrder.AscendingOrder
    assert tab._not_in_kits.isChecked() is False
    assert tab._filter.text() == "2873-SOS"
    assert _visible_revisions(tab) == ["0-AN01", "02", "01-AN01"]

    print("RD catalog AN tab: OK")


if __name__ == "__main__":
    main()
