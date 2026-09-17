"""Snapshot a ``QTableWidget`` into :class:`~rd_catalog.table_xlsx.ExportedTable`.

This module imports PySide6. Do not import it from the Qt-free
``rd_catalog`` package init.
"""

from __future__ import annotations

from collections.abc import Mapping, Sequence

from PySide6.QtWidgets import QTableWidget, QTableWidgetItem

from rd_catalog.monitor_views import MonitorCell
from rd_catalog.table_xlsx import (
    ExportedCell,
    ExportedColumn,
    ExportedTable,
    cell_from_monitor,
)


def _row_item(table: QTableWidget, row: int) -> QTableWidgetItem | None:
    item = table.item(row, 0)
    if item is not None:
        return item
    for column in range(1, table.columnCount()):
        item = table.item(row, column)
        if item is not None:
            return item
    return None


def _column_header(
    table: QTableWidget,
    logical: int,
    header_names: Sequence[str] | None,
) -> str:
    if header_names is not None and 0 <= logical < len(header_names):
        return header_names[logical]
    item = table.horizontalHeaderItem(logical)
    if item is not None:
        return item.text()
    return f"Column {logical}"


def _item_text(table: QTableWidget, row: int, logical: int) -> str:
    item = table.item(row, logical)
    if item is None:
        return ""
    return item.text() or ""


def snapshot_qtable(
    table: QTableWidget,
    *,
    visible_only: bool = True,
    monitor_role: int | None = None,
    header_names: Sequence[str] | None = None,
    sheet_name: str = "Лист",
) -> ExportedTable:
    """Capture visual columns and visible rows of ``table``.

    Args:
        table: Source widget. Sort order is the current Qt row order.
        visible_only: Skip ``table.isRowHidden`` rows (filter-hidden).
        monitor_role: Qt role holding a payload with ``.cells`` mapping
            header → ``MonitorCell``. ``None`` reads item text only.
        header_names: Logical-index names (kits: ``_KITS_HEADERS``).
        sheet_name: Workbook sheet title.

    Returns:
        Frozen table. Hidden header sections are omitted. Paint comes from
        ``MonitorCell``, never ``item.background()``.
    """

    header = table.horizontalHeader()
    columns: list[ExportedColumn] = []
    for visual in range(header.count()):
        logical = header.logicalIndex(visual)
        if header.isSectionHidden(logical):
            continue
        columns.append(
            ExportedColumn(
                header=_column_header(table, logical, header_names),
                width_px=int(header.sectionSize(logical)),
                logical_index=logical,
            )
        )
    exported_columns = tuple(columns)

    rows: list[tuple[ExportedCell, ...]] = []
    for row in range(table.rowCount()):
        if visible_only and table.isRowHidden(row):
            continue
        payload = None
        if monitor_role is not None:
            item = _row_item(table, row)
            if item is not None:
                payload = item.data(monitor_role)
        cells_map = getattr(payload, "cells", None)
        if isinstance(cells_map, Mapping):
            row_cells = tuple(
                cell_from_monitor(
                    cells_map.get(column.header, MonitorCell(text="—"))
                )
                for column in exported_columns
            )
        else:
            row_cells = tuple(
                ExportedCell(text=_item_text(table, row, column.logical_index))
                for column in exported_columns
            )
        rows.append(row_cells)

    return ExportedTable(
        columns=exported_columns,
        rows=tuple(rows),
        sheet_name=sheet_name,
    )
