"""Trace IN_CABINET through DS vs MTO compare (debug txt for analysis).

Filters: ``ds_compare_config.json`` → ``in_cabinet_debug`` or
``analyze_ds_specification(..., compare_debug_log=True)``.
Empty filter field = all values on that axis.
"""

from __future__ import annotations

from datetime import datetime
from pathlib import Path
from typing import Iterable, Optional, TextIO

from base.base_classes import RowStd, RowType
from base.tables_columns import (
    CODE,
    CODE_2,
    DS_NUMBER,
    NUMBERS,
    DS_SPECIFICATION,
    DS_SYSTEM,
    DS_TITLE,
    IN_CABINET,
    TAGS,
    TAGS_2,
    TYPE_MARK,
    TYPE_MARK_2,
    VALUES,
    VALUES_2,
)
from pdf_parsing_v2_engine.document import V2Document
from RFQ.ds_compare.ds_compare_config import InCabinetWatchFilters
from RFQ.ds_compare.load_mto import (
    MTO_PICKLE_EXTRA_KEY,
    MtoLoadAudit,
    reload_mto_positions_fresh,
)


def _norm(s: object) -> str:
    return str(s or "").strip()


def _row_spec_short(row: RowStd) -> str:
    try:
        path = row.t_com.file_name if row.t_com else row.get_value(DS_SPECIFICATION)
        if path:
            return _norm(V2Document.from_file_path(path).doc_Short_Title)
    except Exception:
        pass
    return ""


def _row_codes(row: RowStd) -> set[str]:
    out: set[str] = set()
    for attr in (CODE, CODE_2):
        v = _norm(row.get_value(attr))
        if v:
            out.add(v)
    return out


def _row_title(row: RowStd) -> str:
    return _norm(row.get_value(DS_TITLE))


def _row_mark(row: RowStd, *, mto_side: bool) -> str:
    if mto_side:
        return _norm(row.get_value(TYPE_MARK_2)) or _norm(row.get_value(TYPE_MARK))
    return _norm(row.get_value(TYPE_MARK))


def _mark_haystacks(row: RowStd, *, mto_side: bool) -> list[str]:
    """Strings used to match the «марка» filter (TYPE_MARK, spec key, DS_SYSTEM, …)."""
    parts = [
        _row_mark(row, mto_side=mto_side),
        _row_spec_short(row),
        _row_title(row),
    ]
    if not mto_side:
        parts.append(_norm(row.get_value(DS_SYSTEM)))
    seen: set[str] = set()
    out: list[str] = []
    for p in parts:
        if p and p not in seen:
            seen.add(p)
            out.append(p)
    return out


def _mark_filter_matches(row: RowStd, filters: InCabinetWatchFilters, *, mto_side: bool) -> bool:
    if not filters.marks:
        return True
    haystacks = [h.upper() for h in _mark_haystacks(row, mto_side=mto_side)]
    if not haystacks:
        return False
    return any(m.upper() in h for m in filters.marks for h in haystacks)


def spec_key_matches_watch(spec_key: str, filters: InCabinetWatchFilters) -> bool:
    """Match MTO dict key (e.g. ``8529-SOS``) against title/mark filters only."""
    key = _norm(spec_key)
    if not key:
        return not (filters.titles or filters.marks)
    if filters.titles:
        if not any(t in key for t in filters.titles):
            return False
    if filters.marks:
        ku = key.upper()
        if not any(m.upper() in ku for m in filters.marks):
            return False
    return True


def row_matches_watch(
    row: RowStd,
    filters: InCabinetWatchFilters,
    *,
    mto_side: bool = False,
) -> bool:
    """True if row matches configured filters (empty filter axis = all)."""
    if not isinstance(row, RowStd) or row.row_type != RowType.position_row:
        return False

    if filters.codes:
        if not (_row_codes(row) & filters.codes):
            return False

    if filters.titles:
        title = _row_title(row)
        spec = _row_spec_short(row)
        if not any(t in title or t in spec for t in filters.titles):
            return False

    if not _mark_filter_matches(row, filters, mto_side=mto_side):
        return False

    if filters.cabinets:
        ic = _norm(row.get_value(IN_CABINET)).upper()
        if not ic:
            return False
        if not any(c.upper() in ic for c in filters.cabinets):
            return False

    return True


def format_row_snapshot(row: RowStd, *, label: str = "row") -> str:
    """One-line snapshot for trace file."""
    mto_side = bool(_norm(row.get_value(CODE_2)) and not _norm(row.get_value(CODE)))
    return (
        f"{label}: spec={_row_spec_short(row)!r} "
        f"title={_row_title(row)!r} mark={_row_mark(row, mto_side=mto_side)!r} "
        f"ds_system={_norm(row.get_value(DS_SYSTEM))!r} "
        f"CODE={_norm(row.get_value(CODE))!r} CODE_2={_norm(row.get_value(CODE_2))!r} "
        f"NUMBERS={_norm(row.get_value(NUMBERS))!r} "
        f"VALUES={row.get_value(VALUES)!r} VALUES_2={row.get_value(VALUES_2)!r} "
        f"IN_CABINET={_norm(row.get_value(IN_CABINET))!r} "
        f"IN_CABINET_in_el={IN_CABINET in row.el} "
        f"TAGS={row.get_value(TAGS)!r} TAGS_2={row.get_value(TAGS_2)!r} "
        f"DS#={_norm(row.get_value(DS_NUMBER))!r}"
    )


class CabinetInCabinetTrace:
    """Append-only UTF-8 trace log for IN_CABINET investigation."""

    def __init__(
        self,
        path: str | Path,
        filters: InCabinetWatchFilters,
    ) -> None:
        self.path = Path(path)
        self.filters = filters
        self._fp: Optional[TextIO] = None

    def open(self) -> None:
        self.path.parent.mkdir(parents=True, exist_ok=True)
        self._fp = open(self.path, "w", encoding="utf-8", newline="\n")
        self.write(
            "=== DS vs MTO — trace IN_CABINET (Оборудование относящиеся к шкафам) ===\n"
        )
        self.write(f"Started: {datetime.now().isoformat(timespec='seconds')}\n")
        self.write(f"Log file: {self.path.resolve()}\n")
        self.write(f"Filters: {self.filters.describe()}\n")
        self.write(
            "(пустое поле фильтра = все по оси; марка: TYPE_MARK, DS_SYSTEM, ключ спецификации 8529-SOS)\n\n"
        )

    def close(self) -> None:
        if self._fp is not None:
            self.write(f"\nFinished: {datetime.now().isoformat(timespec='seconds')}\n")
            self._fp.close()
            self._fp = None

    def write(self, text: str) -> None:
        if self._fp is None:
            return
        self._fp.write(text if text.endswith("\n") else text + "\n")
        self._fp.flush()

    def section(self, title: str) -> None:
        self.write(f"\n{'=' * 72}\n{title}\n{'=' * 72}")

    def event(self, message: str) -> None:
        self.write(f"  >> {message}")

    def snapshot(self, row: RowStd, *, label: str) -> None:
        self.write(f"  {format_row_snapshot(row, label=label)}")

    def log_mto_load_audit(
        self,
        spec_dict: dict[str, str],
        load_audit: dict[str, MtoLoadAudit],
    ) -> None:
        self.section("MTO load — paths and cache (DS compare pipeline)")
        self.write(
            "  Тот же разбор, что open_file_att(MTO) → google_sheets.start: "
            "get_mto_std_from_file + add_cabinet_composition_info().\n"
            "  DS compare дополнительно: pickle в RFQ/ds_compare/cache "
            f"(ключ {MTO_PICKLE_EXTRA_KEY!r}).\n"
        )
        shown = 0
        for spec_name in sorted(spec_dict.keys()):
            if not spec_key_matches_watch(spec_name, self.filters):
                continue
            path = spec_dict.get(spec_name, "")
            audit = load_audit.get(spec_name)
            self.write(f"\n  Spec: {spec_name!r}")
            self.write(f"    xlsx path: {path!r}")
            if audit:
                self.write(f"    from_cache: {audit['from_cache']}")
                if audit.get("cache_file"):
                    self.write(f"    cache_file: {audit['cache_file']}")
                self.write(f"    position_rows: {audit.get('position_row_count', 0)}")
            shown += 1
        if shown == 0:
            self.write("  (no spec keys matched title/mark filters)")

    def log_cache_vs_fresh_in_cabinet(
        self,
        spec_name: str,
        mto_path: str,
        cached_rows: list[RowStd],
        *,
        from_cache: bool,
    ) -> None:
        """If cache was used, reload from Excel and compare IN_CABINET for watched CODE rows."""
        if not from_cache or mto_path == "Файл МТО не найден":
            return
        self.section(f"Cache vs fresh Excel — {spec_name!r}")
        self.event("Повторная загрузка без pickle (как при проверке MTO в main)...")
        fresh_rows = reload_mto_positions_fresh(mto_path)
        if fresh_rows is None:
            self.write("  ERROR: fresh reload returned None")
            return

        def _code_rows(rows: list[RowStd]) -> list[RowStd]:
            out = [
                r
                for r in rows
                if row_matches_watch(r, self.filters, mto_side=True)
            ]
            if self.filters.codes:
                out = [r for r in out if _row_codes(r) & self.filters.codes]
            return out

        cached_m = _code_rows(cached_rows)
        fresh_m = _code_rows(fresh_rows)
        self.write(f"  Watched CODE rows: cache={len(cached_m)}, fresh={len(fresh_m)}")

        def _summarize(label: str, rows: list[RowStd]) -> dict[str, float]:
            by_ic: dict[str, float] = {}
            for r in rows:
                ic = _norm(r.get_value(IN_CABINET)) or "(empty)"
                by_ic[ic] = by_ic.get(ic, 0.0) + float(r.get_value(VALUES) or 0)
            self.write(f"  {label} summary IN_CABINET×VALUES: {dict(sorted(by_ic.items()))}")
            return by_ic

        c_sum = _summarize("cache", cached_m)
        f_sum = _summarize("fresh", fresh_m)
        if c_sum != f_sum:
            self.event(
                "РАСХОЖДЕНИЕ cache vs fresh: вероятно устаревший pickle "
                "(включите mto_force_update в ds_compare_config или удалите cache/*.cache)"
            )
        elif not f_sum or all(k == "(empty)" for k in f_sum):
            self.event(
                "И cache, и fresh: IN_CABINET пуст для отфильтрованных строк — "
                "add_cabinet_composition_info не пометил позиции (см. NUMBERS / cabinet_title_row)"
            )
        else:
            self.event("cache и fresh совпадают по IN_CABINET — pickle не причина пустого шкафа")

        for r in fresh_m[:20]:
            self.snapshot(r, label="fresh")
        if len(fresh_m) > 20:
            self.write(f"  ... ещё {len(fresh_m) - 20} fresh row(s)")

    def log_mto_inventory(self, mto_dict: dict[str, list[RowStd]]) -> None:
        self.section("MTO after load — rows matching filters (before compare)")
        total = 0
        for spec_name, rows in sorted(mto_dict.items()):
            if not spec_key_matches_watch(spec_name, self.filters):
                continue
            matched = [
                r for r in rows if row_matches_watch(r, self.filters, mto_side=True)
            ]
            if not matched:
                continue
            self.write(f"\n  Spec key: {spec_name!r} ({len(matched)} row(s))")
            by_ic: dict[str, float] = {}
            for r in matched:
                ic = _norm(r.get_value(IN_CABINET)) or "(empty)"
                by_ic[ic] = by_ic.get(ic, 0.0) + float(r.get_value(VALUES) or 0)
                self.snapshot(r, label="mto")
                total += 1
            self.write(f"  Summary by IN_CABINET: {dict(sorted(by_ic.items()))}")
        if total == 0:
            self.write("  (no MTO position_row matched filters)")

    def log_ds_inventory(self, ds_rows: list[RowStd], *, stage: str) -> None:
        self.section(f"DS rows — {stage}")
        matched = [
            r for r in ds_rows if row_matches_watch(r, self.filters, mto_side=False)
        ]
        if not matched:
            self.write("  (no DS position_row matched filters)")
            return
        by_ic: dict[str, float] = {}
        for r in matched:
            ic = _norm(r.get_value(IN_CABINET)) or "(empty)"
            by_ic[ic] = by_ic.get(ic, 0.0) + float(r.get_value(VALUES) or 0)
            self.snapshot(r, label="ds")
        self.write(f"  Summary by IN_CABINET: {dict(sorted(by_ic.items()))}")
        self.write(f"  Total DS qty (VALUES): {sum(by_ic.values())}")

    def log_allocation(
        self,
        *,
        phase: str,
        code: str,
        ds_row: RowStd,
        mto_row: RowStd,
        allocated: float,
        extra: str = "",
    ) -> None:
        if not row_matches_watch(ds_row, self.filters, mto_side=False):
            return
        self.write(
            f"\n  --- allocation [{phase}] code={code!r} allocated={allocated} {extra} ---"
        )
        self.snapshot(ds_row, label="ds")
        self.snapshot(mto_row, label="mto_chosen")
        if hasattr(mto_row, "is_aggregated") and mto_row.is_aggregated:
            self.event(
                "mto_row is AGGREGATED (IN_CABINET taken from first source row only in code)"
            )

    def log_aggregation_sources(
        self,
        *,
        code: str,
        ds_row: RowStd,
        used_rows_info: Iterable[dict],
        aggregated_row: RowStd,
    ) -> None:
        if not row_matches_watch(ds_row, self.filters, mto_side=False):
            return
        self.write(f"\n  --- aggregation code={code!r} ---")
        self.snapshot(ds_row, label="ds_target")
        cabinets: list[str] = []
        for info in used_rows_info:
            mto_row = info["row"]
            take = info.get("take_amount", 0)
            ic = _norm(mto_row.get_value(IN_CABINET))
            cabinets.append(ic or "(empty)")
            self.write(
                f"    source: take={take} VALUES_rem={mto_row.get_value(VALUES)!r} "
                f"IN_CABINET={ic!r} {format_row_snapshot(mto_row, label='src')}"
            )
        self.write(f"    IN_CABINET values in sources: {cabinets}")
        self.write(
            f"    aggregated_row IN_CABINET (from mto_rows[0], not merged): "
            f"{_norm(aggregated_row.get_value(IN_CABINET))!r}"
        )
        if len({c for c in cabinets if c != "(empty)"}) > 1 or (
            any(c != "(empty)" for c in cabinets) and "(empty)" in cabinets
        ):
            self.event(
                "LIKELY BUG: mixed cabinet/empty sources but aggregation does not merge IN_CABINET"
            )
