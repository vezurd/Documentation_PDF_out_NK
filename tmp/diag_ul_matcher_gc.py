"""Read-only check: how much of the UL matcher cost is cyclic-GC overhead.

Step4 disables GC only around the TM loop; the UL step runs with GC enabled
while it allocates ~1.25M unit objects.
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
from RFQ.tags_rfp_compare.step4.step4_packing_compare import (  # noqa: E402
    _PackingUnit,
    _RowAllocationState,
    _allocate_queue_phase,
    _delivered_qty,
    _merge_packing_units_trace,
)

KEY = ("1", "8445", "sot", "bcc0003052")


def _units(count: int, *, with_status: bool) -> list[_PackingUnit]:
    src = RowStd.get_std_check_row({}, TableComments())
    return [
        _PackingUnit(
            key=KEY,
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


def run(label: str, *, gc_on: bool) -> None:
    if gc_on:
        gc.enable()
    else:
        gc.disable()
    size = 10000
    queue = deque(_units(size, with_status=False))
    state = _RowAllocationState(
        row=RowStd.get_std_check_row({}, TableComments()),
        key=KEY,
        fallback="",
        ordered=float(size),
    )
    t0 = time.perf_counter()
    _allocate_queue_phase(queue, state, "blind")
    alloc = time.perf_counter() - t0

    units = _units(size, with_status=True)
    target = RowStd.get_std_check_row({}, TableComments())
    t0 = time.perf_counter()
    _merge_packing_units_trace(target, units)
    trace = time.perf_counter() - t0
    print(f"{label}: allocate(10000)={alloc:.2f}s  merge_trace(10000)={trace:.2f}s")
    gc.enable()


def isolate_remaining() -> None:
    units = _units(22000, with_status=False)
    t0 = time.perf_counter()
    _delivered_qty(units)
    one = time.perf_counter() - t0
    print(
        f"_delivered_qty(22000 слотов) = {one * 1e3:.2f} ms за один вызов; "
        f"при 22000 вызовах в среднем по половине списка -> {one * 22000 / 2:.1f} s"
    )


if __name__ == "__main__":
    run("GC ON  (как сейчас в шаге УЛ)", gc_on=True)
    run("GC OFF (как в TM-цикле)      ", gc_on=False)
    isolate_remaining()
