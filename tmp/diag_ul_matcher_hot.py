"""Read-only micro-benchmarks of the two heaviest Step4 UL matcher paths.

1. ``_allocate_queue_phase`` on a realistically large per-key unit queue.
2. ``_merge_packing_units_trace`` when the units gate has stamped
   ``UNITS_CHECK_STATUS`` on every UL row (the production case).
"""

from __future__ import annotations

import gc
import sys
import time
from collections import deque
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from base.base_classes import RowStd, TableComments  # noqa: E402
from base.tables_columns import CODE, DS_TITLE, TAGS, UNITS, VALUES  # noqa: E402
from RFQ.tags_rfp_compare.step4.step4_packing_compare import (  # noqa: E402
    _PackingUnit,
    _RowAllocationState,
    _allocate_queue_phase,
    _fill_unit_details,
    _merge_packing_units_trace,
)


def _source_row() -> RowStd:
    row = RowStd.get_std_check_row({}, TableComments())
    row.el[DS_TITLE].value = "8445"
    row.el[CODE].value = "BCC0003052"
    row.el[UNITS].value = "шт"
    row.el[VALUES].value = 1
    row.el[TAGS].value = []
    return row


def _units(count: int, *, with_status: bool) -> list[_PackingUnit]:
    src = _source_row()
    return [
        _PackingUnit(
            key=("1", "8445", "sot", "bcc0003052"),
            source_row=src,
            tag="",
            units="шт",
            source="УЛ ДС1 · Лист1 · строка 5",
            units_check_status="identity" if with_status else "",
            units_conversion_trace="шт->шт (k=1)" if with_status else "",
            qty=1.0,
        )
        for _ in range(count)
    ]


def bench_allocate(sizes: tuple[int, ...]) -> None:
    print("--- _allocate_queue_phase(blind): одна строка забирает всю очередь ---")
    for size in sizes:
        queue = deque(_units(size, with_status=False))
        target = RowStd.get_std_check_row({}, TableComments())
        state = _RowAllocationState(
            row=target, key=("1", "8445", "sot", "bcc0003052"), fallback="",
            ordered=float(size),
        )
        t0 = time.perf_counter()
        _allocate_queue_phase(queue, state, "blind")
        dt = time.perf_counter() - t0
        print(
            f"  qty={size:>6}: {dt:7.2f}s  "
            f"({dt / size * 1e6:6.1f} us/слот, allocated={len(state.allocated)})"
        )


def bench_trace(sizes: tuple[int, ...]) -> None:
    print("\n--- _merge_packing_units_trace: pseudo RowStd на каждый слот ---")
    for size in sizes:
        units = _units(size, with_status=True)
        target = RowStd.get_std_check_row({}, TableComments())
        t0 = time.perf_counter()
        _merge_packing_units_trace(target, units)
        dt = time.perf_counter() - t0
        print(f"  слотов={size:>6}: {dt:7.2f}s ({dt / size * 1e6:6.1f} us/слот)")


def bench_fill_details(size: int) -> None:
    print("\n--- _fill_unit_details целиком (включая trace) ---")
    units = _units(size, with_status=True)
    target = RowStd.get_std_check_row({}, TableComments())
    t0 = time.perf_counter()
    _fill_unit_details(target, units)
    print(f"  слотов={size}: {time.perf_counter() - t0:.2f}s")


if __name__ == "__main__":
    print(f"gc enabled: {gc.isenabled()}")
    bench_allocate((1000, 4000, 10000, 22000))
    bench_trace((1000, 5000, 20000))
    bench_fill_details(20000)
