"""Read-only cost diagnostics for the Step4 UL matcher (no pipeline changes).

Loads the shared packing cache, reports the unit-queue shape, and micro-benchmarks
the hot operations used inside ``compare_rfp_rows_with_packing``.
"""

from __future__ import annotations

import math
import sys
import time
from collections import Counter, defaultdict
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from base.base_classes import CheckElement, RowStd, TableComments  # noqa: E402
from base.tables_columns import (  # noqa: E402
    UNITS,
    UNITS_CHECK_STATUS,
    UNITS_CONVERSION_TRACE,
    VALUES,
)
from RFQ.packing_list_provider import (  # noqa: E402
    default_packing_cache_path,
    load_packing_dataset,
    packing_row_actual_ds,
    packing_row_key,
    rfp_packing_match_key,
)
from RFQ.ds_compare.ds_quantity_parse import parse_quantity_strict, read_raw_quantity  # noqa: E402
from RFQ.tags_rfp_compare.step4.step4_packing_compare import (  # noqa: E402
    _row_leftover_code_key,
    parse_composite_title,
)


def _q(row) -> float | None:
    try:
        raw = read_raw_quantity(row, VALUES)
        value = parse_quantity_strict(raw)
    except (ValueError, TypeError):
        return None
    if not math.isfinite(value) or value < 0:
        return None
    return float(value)


def report_dataset() -> None:
    path = default_packing_cache_path()
    print(f"cache: {path}")
    print(f"exists: {path.is_file()}")
    if not path.is_file():
        print("UNC недоступен из этой сессии — пропускаю раздел с реальными данными")
        return
    t0 = time.perf_counter()
    dataset = load_packing_dataset()
    print(f"load_packing_dataset: {time.perf_counter() - t0:.2f}s")
    print(f"quality={dataset.quality.value} rows={len(dataset.rows)}")

    total_units = 0
    queue_len: dict[tuple[str, str, str, str], int] = defaultdict(int)
    per_row_qty: list[float] = []
    with_status = 0
    with_trace = 0
    for row in dataset.rows:
        quantity = _q(row)
        if quantity is None:
            continue
        key = rfp_packing_match_key(
            packing_row_actual_ds(row) or "", *packing_row_key(row)
        )
        if not all(key[1:]):
            continue
        tags = [str(t).strip() for t in row.get_tags_list() if str(t).strip()]
        if len(tags) > quantity:
            continue
        whole = int(math.floor(quantity + 1e-9))
        slots = whole + (1 if quantity - whole > 1e-9 else 0)
        queue_len[key] += slots
        total_units += slots
        per_row_qty.append(quantity)
        if str(row.el.get(UNITS_CHECK_STATUS, CheckElement(None)).value or "").strip():
            with_status += 1
        if str(
            row.el.get(UNITS_CONVERSION_TRACE, CheckElement(None)).value or ""
        ).strip():
            with_trace += 1

    lengths = sorted(queue_len.values(), reverse=True)
    print(f"queued rows={len(per_row_qty)}  total unit slots={total_units}")
    print(f"distinct A keys={len(lengths)}")
    print(f"top-15 queue lengths: {lengths[:15]}")
    print(f"rows with UNITS_CHECK_STATUS={with_status}, TRACE={with_trace}")
    buckets = Counter()
    for quantity in per_row_qty:
        if quantity <= 1:
            buckets["<=1"] += 1
        elif quantity <= 10:
            buckets["2-10"] += 1
        elif quantity <= 100:
            buckets["11-100"] += 1
        elif quantity <= 1000:
            buckets["101-1000"] += 1
        else:
            buckets[">1000"] += 1
    print(f"qty buckets per UL row: {dict(buckets)}")
    print(f"max single-row qty={max(per_row_qty):.2f}")
    # sum of k^2 over per-key queues: cost model of remaining_qty() recomputation
    quad = sum(n * n for n in lengths)
    print(f"sum(queue_len^2) = {quad:.3e}  (модель стоимости remaining_qty)")


def bench_ops() -> None:
    print("\n--- микробенчмарки горячих операций ---")
    t_com = TableComments()
    n = 2000
    t0 = time.perf_counter()
    for _ in range(n):
        RowStd.get_std_check_row({}, TableComments())
    per = (time.perf_counter() - t0) / n
    print(f"RowStd.get_std_check_row({{}}, TableComments()): {per * 1e6:.1f} us/шт")

    row = RowStd.get_std_check_row({}, t_com)
    row.el["DS_TITLE" if "DS_TITLE" in row.el else list(row.el)[0]]
    from base.tables_columns import CODE_MTO, DS_TITLE

    row.el[DS_TITLE].value = "8445-SOT1"
    row.el[CODE_MTO].value = "BCC0003052"
    n = 200_000
    t0 = time.perf_counter()
    for _ in range(n):
        _row_leftover_code_key(row, CODE_MTO)
    per = (time.perf_counter() - t0) / n
    print(f"_row_leftover_code_key: {per * 1e6:.2f} us/шт")

    t0 = time.perf_counter()
    for _ in range(n):
        parse_composite_title("8445-SOT1")
    per = (time.perf_counter() - t0) / n
    print(f"parse_composite_title: {per * 1e6:.2f} us/шт")

    t0 = time.perf_counter()
    for _ in range(n):
        packing_row_actual_ds(row)
    per = (time.perf_counter() - t0) / n
    print(f"packing_row_actual_ds: {per * 1e6:.2f} us/шт")

    # sibling scan: cost per inner iteration of _reserved_tags_for_state
    class _S:
        __slots__ = ("key", "allocated")

        def __init__(self, key):
            self.key = key
            self.allocated = []

    states = [_S(("1", "8445", "sot", f"bcc{i % 7000:07d}")) for i in range(20_000)]
    probe = states[0]
    t0 = time.perf_counter()
    hits = 0
    for other in states:
        if other is probe or other.key != probe.key:
            continue
        hits += 1
    dt = time.perf_counter() - t0
    print(
        f"один проход по 20000 siblings: {dt * 1e3:.2f} ms "
        f"({dt / len(states) * 1e9:.0f} ns/итерация)"
    )
    print(
        "  => полный O(N^2) для N=20000: "
        f"{dt * len(states):.1f} s на один проход резервирования"
    )
    print(f"  (для справки: {UNITS} / {VALUES} не участвуют)")


if __name__ == "__main__":
    report_dataset()
    bench_ops()
