"""IN_CABINET-aware MTO quantity allocation for DS vs MTO compare."""

from __future__ import annotations

from dataclasses import dataclass, field

from base.base_classes import RowStd
from base.tables_columns import CODE, IN_CABINET, VALUES


def norm_in_cabinet(value: object) -> str:
    return str(value or "").strip()


def merge_in_cabinet(current: str, new: str) -> str:
    """Merge cabinet tags; different values joined with ``/`` (step4 style)."""
    new_value = norm_in_cabinet(new)
    if not new_value:
        return norm_in_cabinet(current)
    current_value = norm_in_cabinet(current)
    if not current_value:
        return new_value
    if current_value == new_value:
        return current_value
    parts = [p.strip() for p in current_value.split("/") if p.strip()]
    if new_value in parts:
        return current_value
    return f"{current_value}/{new_value}"


@dataclass
class MtoTakeSlice:
    """One portion taken from a single MTO row."""

    row: RowStd
    take_amount: float
    in_cabinet: str = ""

    @property
    def has_cabinet(self) -> bool:
        return bool(norm_in_cabinet(self.in_cabinet))


@dataclass
class CabinetAllocationResult:
    """Result of multi-row MTO allocation for one DS row."""

    slices: list[MtoTakeSlice] = field(default_factory=list)
    primary_row: RowStd | None = None
    total_taken: float = 0.0

    @property
    def in_cabinet_value(self) -> str:
        value = ""
        for sl in self.slices:
            if sl.has_cabinet:
                value = merge_in_cabinet(value, sl.in_cabinet)
        return value

    @property
    def in_cabinet_comment(self) -> str:
        """Note for IN_CABINET cell when allocation mixes cabinet and non-cabinet MTO."""
        qty_cabinet = sum(sl.take_amount for sl in self.slices if sl.has_cabinet)
        qty_no = sum(sl.take_amount for sl in self.slices if not sl.has_cabinet)
        parts: list[str] = []
        if qty_no > 0 and qty_cabinet > 0:
            q = int(qty_no) if qty_no == int(qty_no) else qty_no
            parts.append(f"часть {q} шт. без шкафа (из MTO)")
        if len({sl.in_cabinet for sl in self.slices if sl.has_cabinet}) > 1:
            parts.append("сопоставление из нескольких шкафов MTO")
        if qty_no > 0 and qty_cabinet == 0:
            q = int(qty_no) if qty_no == int(qty_no) else qty_no
            parts.append(f"все {q} шт. без шкафа в MTO")
        return "; ".join(parts)


def _mto_sort_key(row: RowStd, ds_row: RowStd | None) -> tuple:
    """Prefer cabinet rows, then smaller qty (consume narrow rows first)."""
    ic = norm_in_cabinet(row.get_value(IN_CABINET))
    has_ic = 0 if ic else 1
    try:
        qty = float(row.get_value(VALUES) or 0)
    except (TypeError, ValueError):
        qty = 0.0
    code_prio = 0
    if ds_row is not None:
        ds_code = ds_row.get_value(CODE)
        if str(row.get_value(CODE) or "") == str(ds_code or ""):
            code_prio = 0
        else:
            code_prio = 1
    return (code_prio, has_ic, qty)


def allocate_mto_slices(
    mto_rows: list[RowStd],
    required_amount: float,
    ds_row: RowStd | None = None,
) -> CabinetAllocationResult:
    """
    Take ``required_amount`` from ``mto_rows`` in order: cabinet-tagged rows first,
    then rows without cabinet; within each group smaller VALUES first.
    """
    result = CabinetAllocationResult()
    need = float(required_amount or 0)
    if need <= 0:
        return result

    available = [
        r
        for r in mto_rows
        if float(r.get_value(VALUES) or 0) > 0
    ]
    if not available:
        return result

    ordered = sorted(available, key=lambda r: _mto_sort_key(r, ds_row))

    for mto_row in ordered:
        if need <= 0:
            break
        mto_qty = float(mto_row.get_value(VALUES) or 0)
        if mto_qty <= 0:
            continue
        take = min(need, mto_qty)
        if take <= 0:
            continue
        result.slices.append(
            MtoTakeSlice(
                row=mto_row,
                take_amount=take,
                in_cabinet=norm_in_cabinet(mto_row.get_value(IN_CABINET)),
            )
        )
        need -= take

    result.total_taken = sum(sl.take_amount for sl in result.slices)
    if result.slices:
        result.primary_row = result.slices[0].row
    return result


def cabinet_alloc_from_used_rows(
    used_rows_info: list[dict],
    primary_row: RowStd | None = None,
) -> CabinetAllocationResult:
    """Build allocation result from aggregated-row ``used_rows_info`` entries."""
    slices: list[MtoTakeSlice] = []
    for info in used_rows_info:
        take = float(info.get("take_amount") or 0)
        if take <= 0:
            continue
        row = info["row"]
        slices.append(
            MtoTakeSlice(
                row=row,
                take_amount=take,
                in_cabinet=norm_in_cabinet(row.get_value(IN_CABINET)),
            )
        )
    total = sum(sl.take_amount for sl in slices)
    return CabinetAllocationResult(
        slices=slices,
        primary_row=primary_row or (slices[0].row if slices else None),
        total_taken=total,
    )
