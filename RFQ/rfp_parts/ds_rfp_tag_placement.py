"""Cluster-level RFP tag placement for a legal shared-file DS group.

Hybrid overlay stays atomic for a number that is not in a cluster. A legal
cluster (same nonempty RFP file set on two or more active source_ids) is
placed here: either each key lands on one member, or every summary row of
the cluster uses the slash label. Tag split reuses ``_rfp_net_rows`` /
``_build_unit_counters``; this module does not invent a second splitter.
"""

from __future__ import annotations

from collections.abc import Mapping, Sequence
from dataclasses import dataclass, replace
from decimal import Decimal

from RFQ.ds_compare.ds_units_normalize import normalize_units_text
from RFQ.rfp_parts.analyze_rfp_parts import NetSummaryRow, RfpRecord
from RFQ.rfp_parts.ds_baseline import DsBaselinePosition, _composite_title
from RFQ.rfp_parts.ds_registry import canonical_supply_group_id
from RFQ.tags_rfp_compare.rfp_supply_status import is_canonical_excluded_from_supply
from RFQ.units_convert.models import normalize_code

MIX_SEPARATE = "separate"  # без смешения
MIX_MIXED = "mixed"  # со смешением

PlacementKey = tuple[str, str, str]
_ZERO = Decimal("0")


@dataclass(frozen=True, slots=True)
class ClusterPlacement:
    """Summary rows and report fields for one legal RFP-file cluster."""

    rows: tuple[NetSummaryRow, ...]
    group_id: str
    group_label: str
    status: str
    selected_source: str
    reason: str
    overlay_blocked: bool
    fallback: bool
    split_ok: bool
    ds_source_ids: tuple[str, ...]
    rfp_files: tuple[str, ...]
    ds_position_count: int
    rfp_record_count: int
    failed_key: PlacementKey | None = None


def sorted_source_ids(source_ids: Sequence[str]) -> tuple[str, ...]:
    """Return cluster members: pure-digit ids numerically, then other strings."""

    unique: list[str] = []
    seen: set[str] = set()
    for raw in source_ids:
        text = str(raw).strip()
        if not text or text in seen:
            continue
        seen.add(text)
        unique.append(text)
    digits = [item for item in unique if item.isdigit()]
    others = [item for item in unique if not item.isdigit()]
    digits.sort(key=int)
    others.sort()
    return tuple(digits + others)


def mixed_ds_label(source_ids: Sequence[str]) -> str:
    """Return the slash cluster label, e.g. ``ДС15/61`` or ``ДС13/47/102``.

    Args:
        source_ids: Active registry numbers that share one RFP file set.

    Returns:
        ``ДС`` plus ids joined by ``/``. Digit ids sort numerically ascending;
        other ids sort after them as strings. This is not the filename form
        ``ДС15_61``.
    """

    return "ДС" + "/".join(sorted_source_ids(source_ids))


def cluster_report_id(source_ids: Sequence[str], *, mix_mode: str) -> str:
    """Stable report group id: slash label in mixed mode, else ``cluster:15+61``."""

    ordered = sorted_source_ids(source_ids)
    if mix_mode == MIX_MIXED:
        return mixed_ds_label(ordered)
    return "cluster:" + "+".join(ordered)


def place_cluster_summary(
    *,
    source_ids: Sequence[str],
    mix_mode: str,
    rfp_records: Sequence[RfpRecord],
    positions_by_source: Mapping[str, Sequence[DsBaselinePosition]],
    rfp_files: Sequence[str] = (),
    blocked: bool = False,
    block_reason: str = "",
    fallback: bool = False,
) -> ClusterPlacement:
    """Place RFP tags for one legal cluster under ``mix_mode``.

    Args:
        source_ids: Active members of the cluster.
        mix_mode: ``MIX_SEPARATE`` or ``MIX_MIXED``.
        rfp_records: Converted records of the cluster files (read once).
        positions_by_source: Baseline positions keyed by source_id.
        rfp_files: Relpaths of the cluster RFP workbooks, for the report.
        blocked: Overlay/fallback/extract/units error on any member. When True,
            placement does not run: each member's DS rows, no slash name.
        block_reason: Text appended when ``blocked`` is True.
        fallback: Grouping-fallback flag stored on the group result.

    Returns:
        Summary rows plus one cluster report identity. A key mismatch does not
        set a global blocker; the caller must not stop the workbook for it.
    """

    ordered = sorted_source_ids(source_ids)
    slash = mixed_ds_label(ordered)
    ds_count = sum(len(positions_by_source.get(sid, ())) for sid in ordered)
    rfp_count = len(rfp_records)
    files = tuple(rfp_files)

    if blocked:
        rows = _member_ds_rows(ordered, positions_by_source)
        reason = "без смешения"
        if block_reason:
            reason = f"{reason}; {block_reason}"
        return ClusterPlacement(
            rows=tuple(rows),
            group_id=cluster_report_id(ordered, mix_mode=MIX_SEPARATE),
            group_label=cluster_report_id(ordered, mix_mode=MIX_SEPARATE),
            status="BLOCKED",
            selected_source="ДС" if rows else "",
            reason=reason,
            overlay_blocked=True,
            fallback=fallback,
            split_ok=False,
            ds_source_ids=ordered,
            rfp_files=files,
            ds_position_count=ds_count,
            rfp_record_count=rfp_count,
        )

    rfp_bag = _rfp_qty_bag(rfp_records)
    rfp_nz = {key: qty for key, qty in rfp_bag.items() if qty != _ZERO}
    member_bags = {
        sid: _ds_qty_bag(positions_by_source.get(sid, ())) for sid in ordered
    }

    if mix_mode == MIX_MIXED:
        return _place_mixed(
            ordered=ordered,
            slash=slash,
            rfp_records=rfp_records,
            rfp_bag=rfp_bag,
            member_bags=member_bags,
            positions_by_source=positions_by_source,
            rfp_files=files,
            ds_count=ds_count,
            rfp_count=rfp_count,
            fallback=fallback,
        )
    return _place_separate(
        ordered=ordered,
        rfp_records=rfp_records,
        rfp_nz=rfp_nz,
        member_bags=member_bags,
        positions_by_source=positions_by_source,
        rfp_files=files,
        ds_count=ds_count,
        rfp_count=rfp_count,
        fallback=fallback,
    )


def _place_separate(
    *,
    ordered: tuple[str, ...],
    rfp_records: Sequence[RfpRecord],
    rfp_nz: dict[PlacementKey, Decimal],
    member_bags: Mapping[str, dict[PlacementKey, Decimal]],
    positions_by_source: Mapping[str, Sequence[DsBaselinePosition]],
    rfp_files: tuple[str, ...],
    ds_count: int,
    rfp_count: int,
    fallback: bool,
) -> ClusterPlacement:
    assigned, failed_key, fail_detail = _unique_assignments(rfp_nz, ordered, member_bags)
    report_id = cluster_report_id(ordered, mix_mode=MIX_SEPARATE)
    if failed_key is not None:
        rows = _member_ds_rows(ordered, positions_by_source)
        reason = f"без смешения; {fail_detail}"
        return ClusterPlacement(
            rows=tuple(rows),
            group_id=report_id,
            group_label=report_id,
            status="MISMATCH",
            selected_source="ДС" if rows else "",
            reason=reason,
            overlay_blocked=False,
            fallback=fallback,
            split_ok=False,
            ds_source_ids=ordered,
            rfp_files=rfp_files,
            ds_position_count=ds_count,
            rfp_record_count=rfp_count,
            failed_key=failed_key,
        )

    rows: list[NetSummaryRow] = []
    for key in sorted(assigned):
        owner = assigned[key]
        label = canonical_supply_group_id(owner)
        rows.extend(_emit_rfp_rows(_rfp_records_for_key(rfp_records, key), label))
    ds_only = _ds_only_keys(ordered, member_bags, rfp_nz)
    for key in ds_only:
        for sid in ordered:
            if member_bags[sid].get(key, _ZERO) == _ZERO:
                continue
            label = canonical_supply_group_id(sid)
            rows.extend(
                _emit_ds_rows(
                    _positions_for_key(positions_by_source.get(sid, ()), key),
                    label,
                )
            )
    reason = "без смешения; ключи RFP на своих номерах"
    return ClusterPlacement(
        rows=tuple(rows),
        group_id=report_id,
        group_label=report_id,
        status="MATCH",
        selected_source="RFP",
        reason=reason,
        overlay_blocked=False,
        fallback=fallback,
        split_ok=True,
        ds_source_ids=ordered,
        rfp_files=rfp_files,
        ds_position_count=ds_count,
        rfp_record_count=rfp_count,
    )


def _place_mixed(
    *,
    ordered: tuple[str, ...],
    slash: str,
    rfp_records: Sequence[RfpRecord],
    rfp_bag: dict[PlacementKey, Decimal],
    member_bags: Mapping[str, dict[PlacementKey, Decimal]],
    positions_by_source: Mapping[str, Sequence[DsBaselinePosition]],
    rfp_files: tuple[str, ...],
    ds_count: int,
    rfp_count: int,
    fallback: bool,
) -> ClusterPlacement:
    ds_total = _sum_member_bags(ordered, member_bags)
    keys = sorted(set(ds_total) | set(rfp_bag))
    rows: list[NetSummaryRow] = []
    notes: list[str] = []
    rfp_only: list[PlacementKey] = []
    mismatch = False
    used_rfp = False
    for key in keys:
        ds_qty = ds_total.get(key, _ZERO)
        rfp_qty = rfp_bag.get(key, _ZERO)
        if rfp_qty != _ZERO and ds_qty == rfp_qty:
            rows.extend(_emit_rfp_rows(_rfp_records_for_key(rfp_records, key), slash))
            used_rfp = True
            continue
        if ds_qty != _ZERO and ds_qty != rfp_qty:
            mismatch = True
            notes.append(
                f"ключ {_format_key(key)}: ДС={_qty_text(ds_qty)} "
                f"RFP={_qty_text(rfp_qty)}"
            )
            rows.extend(
                _emit_ds_rows(
                    _cluster_positions_for_key(ordered, positions_by_source, key),
                    slash,
                )
            )
            continue
        if rfp_qty != _ZERO and ds_qty == _ZERO:
            mismatch = True
            rfp_only.append(key)
            continue
        if ds_qty != _ZERO and rfp_qty == _ZERO:
            rows.extend(
                _emit_ds_rows(
                    _cluster_positions_for_key(ordered, positions_by_source, key),
                    slash,
                )
            )
    reason = f"смешение {slash}"
    if notes:
        reason = f"{reason}; " + "; ".join(notes)
    if rfp_only:
        listed = ", ".join(_format_key(key) for key in rfp_only)
        reason = f"{reason}; только RFP без ДС: {listed}"
    status = "MISMATCH" if mismatch else "MATCH"
    if used_rfp and not mismatch:
        selected = "RFP"
    elif rows:
        selected = "RFP" if used_rfp else "ДС"
    else:
        selected = ""
    failed_key: PlacementKey | None = None
    for key in keys:
        ds_qty = ds_total.get(key, _ZERO)
        rfp_qty = rfp_bag.get(key, _ZERO)
        if ds_qty != rfp_qty:
            failed_key = key
            break
    return ClusterPlacement(
        rows=tuple(rows),
        group_id=slash,
        group_label=slash,
        status=status,
        selected_source=selected,
        reason=reason,
        overlay_blocked=False,
        fallback=fallback,
        split_ok=not mismatch,
        ds_source_ids=ordered,
        rfp_files=rfp_files,
        ds_position_count=ds_count,
        rfp_record_count=rfp_count,
        failed_key=failed_key,
    )


def _unique_assignments(
    rfp_nz: Mapping[PlacementKey, Decimal],
    ordered: Sequence[str],
    member_bags: Mapping[str, Mapping[PlacementKey, Decimal]],
) -> tuple[dict[PlacementKey, str], PlacementKey | None, str]:
    assigned: dict[PlacementKey, str] = {}
    for key in sorted(rfp_nz):
        rfp_qty = rfp_nz[key]
        owners = [
            sid
            for sid in ordered
            if member_bags[sid].get(key, _ZERO) != _ZERO
        ]
        if len(owners) != 1:
            if not owners:
                return {}, key, f"ключ {_format_key(key)} нет владельца ДС"
            labels = "/".join(canonical_supply_group_id(sid) for sid in owners)
            return (
                {},
                key,
                f"ключ {_format_key(key)} у нескольких номеров ({labels})",
            )
        owner = owners[0]
        ds_qty = member_bags[owner][key]
        if ds_qty != rfp_qty:
            return (
                {},
                key,
                (
                    f"ключ {_format_key(key)} количества не равны: "
                    f"ДС={_qty_text(ds_qty)} RFP={_qty_text(rfp_qty)}"
                ),
            )
        assigned[key] = owner
    return assigned, None, ""


def _ds_only_keys(
    ordered: Sequence[str],
    member_bags: Mapping[str, Mapping[PlacementKey, Decimal]],
    rfp_nz: Mapping[PlacementKey, Decimal],
) -> tuple[PlacementKey, ...]:
    keys: set[PlacementKey] = set()
    for sid in ordered:
        for key, qty in member_bags[sid].items():
            if qty != _ZERO and key not in rfp_nz:
                keys.add(key)
    return tuple(sorted(keys))


def _sum_member_bags(
    ordered: Sequence[str],
    member_bags: Mapping[str, Mapping[PlacementKey, Decimal]],
) -> dict[PlacementKey, Decimal]:
    total: dict[PlacementKey, Decimal] = {}
    for sid in ordered:
        for key, qty in member_bags[sid].items():
            total[key] = total.get(key, _ZERO) + qty
    return total


def _placement_key_ds(position: DsBaselinePosition) -> PlacementKey | None:
    if position.qty is None:
        return None
    code = normalize_code(position.code or position.code_normalized)
    if not code:
        return None
    title = _composite_title(position.title, position.system)
    unit = normalize_units_text(position.units)
    return (title, code, unit)


def _placement_key_rfp(record: RfpRecord) -> PlacementKey | None:
    if is_canonical_excluded_from_supply(record.rfp_supply_status):
        return None
    code = normalize_code(record.code)
    if not code:
        return None
    title = (record.ds_title or "").strip()
    unit = normalize_units_text(record.units)
    return (title, code, unit)


def _ds_qty_bag(
    positions: Sequence[DsBaselinePosition],
) -> dict[PlacementKey, Decimal]:
    bag: dict[PlacementKey, Decimal] = {}
    for item in positions:
        key = _placement_key_ds(item)
        if key is None or item.qty is None:
            continue
        bag[key] = bag.get(key, _ZERO) + item.qty
    return bag


def _rfp_qty_bag(records: Sequence[RfpRecord]) -> dict[PlacementKey, Decimal]:
    bag: dict[PlacementKey, Decimal] = {}
    for item in records:
        key = _placement_key_rfp(item)
        if key is None:
            continue
        bag[key] = bag.get(key, _ZERO) + item.values
    return bag


def _rfp_records_for_key(
    records: Sequence[RfpRecord],
    key: PlacementKey,
) -> list[RfpRecord]:
    return [item for item in records if _placement_key_rfp(item) == key]


def _positions_for_key(
    positions: Sequence[DsBaselinePosition],
    key: PlacementKey,
) -> list[DsBaselinePosition]:
    return [item for item in positions if _placement_key_ds(item) == key]


def _cluster_positions_for_key(
    ordered: Sequence[str],
    positions_by_source: Mapping[str, Sequence[DsBaselinePosition]],
    key: PlacementKey,
) -> list[DsBaselinePosition]:
    out: list[DsBaselinePosition] = []
    for sid in ordered:
        out.extend(_positions_for_key(positions_by_source.get(sid, ()), key))
    return out


def _member_ds_rows(
    ordered: Sequence[str],
    positions_by_source: Mapping[str, Sequence[DsBaselinePosition]],
) -> list[NetSummaryRow]:
    rows: list[NetSummaryRow] = []
    for sid in ordered:
        label = canonical_supply_group_id(sid)
        rows.extend(_emit_ds_rows(positions_by_source.get(sid, ()), label))
    return rows


def _relabel_net_rows(
    rows: Sequence[NetSummaryRow],
    ds_name: str,
) -> list[NetSummaryRow]:
    out: list[NetSummaryRow] = []
    for item in rows:
        out.append(
            NetSummaryRow(
                record=replace(item.record, ds_name=ds_name),
                source=item.source,
                flag_unmatched_decrease=item.flag_unmatched_decrease,
                flag_negative_net=item.flag_negative_net,
                flag_units_mismatch=item.flag_units_mismatch,
                flag_via_replacement=item.flag_via_replacement,
            )
        )
    return out


def _emit_rfp_rows(
    records: Sequence[RfpRecord],
    ds_name: str,
) -> list[NetSummaryRow]:
    if not records:
        return []
    from RFQ.rfp_parts.ds_rfp_hybrid import _rfp_net_rows

    return _relabel_net_rows(_rfp_net_rows(records, ds_name), ds_name)


def _emit_ds_rows(
    positions: Sequence[DsBaselinePosition],
    ds_name: str,
) -> list[NetSummaryRow]:
    if not positions:
        return []
    from RFQ.rfp_parts.ds_rfp_hybrid import _ds_net_rows

    return _relabel_net_rows(_ds_net_rows(positions), ds_name)


def _format_key(key: PlacementKey) -> str:
    return " ".join(part for part in key if part)


def _qty_text(value: Decimal) -> str:
    if value == value.to_integral_value():
        return str(int(value))
    return format(value, "f")
