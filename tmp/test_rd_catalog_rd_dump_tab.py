"""Offscreen checks for the РД dump tab: non-canonical filter and date sort."""

from __future__ import annotations

import os
import sys
from datetime import datetime
from pathlib import Path

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from PySide6.QtCore import Qt
from PySide6.QtWidgets import QApplication

from rd_catalog.an_index import AnMtoFile, KitAnTargets
from rd_catalog.kits import kit_identity_key
from rd_catalog.models import make_path_key
from rd_catalog.rd_dump_index import CANON_HEADER, KIND_HEADER, RD_DUMP_HEADERS
from rd_catalog.rd_dump_tab import RdDumpTab

_COL_DATE = RD_DUMP_HEADERS.index("Дата")
_COL_REV = RD_DUMP_HEADERS.index("Ревизия")
_COL_KIND = RD_DUMP_HEADERS.index(KIND_HEADER)
_COL_CANON = RD_DUMP_HEADERS.index(CANON_HEADER)
_RD_ROOT = r"\\bcc\eng\PrDoc\377_НИПИГАЗ\АГХК\КСБ\РД"


def _mtime_ns(year: int, month: int, day: int) -> int:
    return int(datetime(year, month, day).timestamp() * 1_000_000_000)


def _dump_file(
    *,
    title: str,
    mark: str,
    revision_text: str,
    mtime_ns: int,
    canonical: bool,
    kind: str = "MTO",
) -> AnMtoFile:
    if kind == "OD":
        name = f"AGCC.287-{title}-{mark}.OD-0001_{revision_text}_RU.docx"
        media = "DWG"
        discipline = "OD-0001"
        stem = f"AGCC.287-{title}-{mark}.OD-0001"
    else:
        name = f"AGCC.287-{title}-{mark}.MTO-0001_{revision_text}_RU.xlsx"
        media = "PDF"
        discipline = "MTO-0001"
        stem = f"AGCC.287-{title}-{mark}.MTO-0001"
    if canonical:
        path = (
            rf"{_RD_ROOT}\{title}\{mark}\Для передачи\01_рев.{revision_text}"
            rf"\{media}\{name}"
        )
    else:
        path = rf"{_RD_ROOT}\{title}\{mark}\{name}"
    return AnMtoFile(
        path=path,
        path_key=make_path_key(path),
        title=title,
        mark=mark,
        revision_text=revision_text,
        core_stem=stem,
        discipline_block=discipline,
        name=name,
        parent_dir=str(Path(path).parent),
        mtime_ns=mtime_ns,
        size=1,
    )


def _visible_revisions(tab: RdDumpTab) -> list[str]:
    table = tab.table()
    out: list[str] = []
    for index in range(table.rowCount()):
        if table.isRowHidden(index):
            continue
        item = table.item(index, _COL_REV)
        out.append("" if item is None else item.text())
    return out


def _visible_kinds(tab: RdDumpTab) -> list[str]:
    table = tab.table()
    out: list[str] = []
    for index in range(table.rowCount()):
        if table.isRowHidden(index):
            continue
        item = table.item(index, _COL_KIND)
        out.append("" if item is None else item.text())
    return out


def _visible_canon(tab: RdDumpTab) -> list[str]:
    table = tab.table()
    out: list[str] = []
    for index in range(table.rowCount()):
        if table.isRowHidden(index):
            continue
        item = table.item(index, _COL_CANON)
        out.append("" if item is None else item.text())
    return out


def main() -> None:
    """Jump from Комплекты sorts by date; non-canonical filter hides issued paths."""

    instance = QApplication.instance()
    if instance is None:
        QApplication([])

    newest = _dump_file(
        title="2873",
        mark="SOS",
        revision_text="01-AN01",
        mtime_ns=_mtime_ns(2026, 6, 16),
        canonical=True,
    )
    oldest = _dump_file(
        title="2873",
        mark="SOS",
        revision_text="01",
        mtime_ns=_mtime_ns(2025, 11, 3),
        canonical=False,
    )
    other = _dump_file(
        title="1600",
        mark="POS",
        revision_text="02",
        mtime_ns=_mtime_ns(2026, 1, 1),
        canonical=True,
    )
    od_file = _dump_file(
        title="2873",
        mark="SOS",
        revision_text="01-AN01",
        mtime_ns=_mtime_ns(2026, 3, 1),
        canonical=False,
        kind="OD",
    )

    tab = RdDumpTab()
    tab.set_rows(
        (newest, oldest, other, od_file),
        is_banned=lambda _t, _m: False,
        kit_targets={
            kit_identity_key("2873", "SOS"): KitAnTargets(rd_mto="01-AN01"),
        },
        rd_root=_RD_ROOT,
        allowed_kits={kit_identity_key("2873", "SOS"), kit_identity_key("1600", "POS")},
    )
    tab.show()
    tab.set_kit_filter("2873", "SOS")
    visible = _visible_revisions(tab)
    assert visible == ["01", "01-AN01", "01-AN01"], visible
    table = tab.table()
    assert table.horizontalHeader().sortIndicatorSection() == _COL_DATE
    assert table.horizontalHeader().sortIndicatorOrder() == Qt.SortOrder.AscendingOrder

    tab._noncanonical.setChecked(True)
    canon = _visible_canon(tab)
    assert canon == ["нет", "нет"], canon
    assert set(_visible_kinds(tab)) == {"MTO", "OD"}

    tab._only_od.setChecked(True)
    assert _visible_kinds(tab) == ["OD"]
    assert _visible_revisions(tab) == ["01-AN01"]

    print("RD catalog RD dump tab: OK")


if __name__ == "__main__":
    main()
