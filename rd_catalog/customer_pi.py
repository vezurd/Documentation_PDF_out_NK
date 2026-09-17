"""Customer EDMS / PI report dump (xlsb → pickle) for BCC MTO/BOM rows.

Accepts two xlsb layouts and always stores 49 canonical columns:

- ``full``: two-row header (dates План/Прогноз/Факт), 49 columns (``TDSheet``).
- ``identity``: one-row header, identity + «По РД» only (BCC ``*_ВСС.xlsb``).
  Missing columns (dates, ПИ qty, ЗИП, «Центр закупки») are stored empty.

This module does not join RD overlay or write UNC sources.
"""

from __future__ import annotations

import argparse
import os
import pickle
import re
import sys
import tempfile
from collections import defaultdict
from collections.abc import Iterator, Mapping, Sequence
from dataclasses import dataclass
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Any

from utils.file_name_converts import AgccFilenamePatterns

SCHEMA_VERSION = 1
FORMAT_ID = "rd_catalog.customer_pi"

DUMP_DIR = Path(__file__).resolve().parent / "База заказчика"
DEFAULT_PICKLE_NAME = "customer_pi_bcc.pkl"
DEFAULT_XLSB = Path(
    r"\\bcc\root\CurProjects\di_manegers\Пятницкий_П"
    r"\Амурский ГХК\Поставки\Отчет систем ПИ"
    r"\Отчет для заказчика АГХК 07.09.2026.xlsb"
)

SUPPLIER_KEEP = "БИ.СИ.СИ., ООО"
_DISCIPLINE_RE = re.compile(r"\.(MTO|BOM|BOE|DS|SP|OD|VO)-", re.I)

N_COLS = 49
DATE_COL_INDEXES = frozenset(range(27, 36)) | frozenset(range(38, 44))
_EXCEL_EPOCH = datetime(1899, 12, 30)
LAYOUT_FULL = "full"
LAYOUT_IDENTITY = "identity"
# Identity dumps pad the used range with hundreds of thousands of empty rows.
_EMPTY_TAIL_STOP = 128

COLUMN_KEYS: tuple[str, ...] = (
    "installation",
    "purchase_center",
    "ul_no",
    "supplier",
    "contract",
    "lot",
    "mtr_group",
    "nomenclature_group",
    "title",
    "mark",
    "spec",
    "rfq_no",
    "tags",
    "code_1c",
    "code_rd",
    "mapping_code",
    "rd_revision",
    "name",
    "type_mark",
    "analog",
    "note",
    "closing_request",
    "units_rd",
    "qty_rd",
    "qty_pi",
    "qty_pi_involved",
    "qty_pi_remain",
    "rfq_issue_plan",
    "rfq_issue_forecast",
    "rfq_issue_actual",
    "gp_issue_plan",
    "gp_issue_forecast",
    "gp_issue_actual",
    "contract_sign_plan",
    "contract_sign_forecast",
    "contract_sign_actual",
    "contracted_qty_rd",
    "contracted_qty_vendor",
    "ready_to_ship_plan",
    "ready_to_ship_forecast",
    "ready_to_ship_actual",
    "site_delivery_plan",
    "site_delivery_forecast",
    "site_delivery_actual",
    "delivered_qty_vendor",
    "vendor_units",
    "complectation_comment",
    "tech_eval_comment",
    "is_spare",
)

COLUMN_CAPTIONS: tuple[str, ...] = (
    "Установка",
    "Центр закупки",
    "№ УЛ",
    "Поставщик",
    "Договор с поставщиком",
    "№ лота",
    "Группа МТР",
    "Группа номенклатуры",
    "Титул",
    "Раздел",
    "Спецификация",
    "№ RFQ",
    "TAG / Линия",
    "Код",
    "Код РД",
    "Код мэппинга",
    "№ ревизии РД",
    "Описание",
    "Тех.характеристики",
    "Аналог",
    "Примечание",
    "Закрывающая заявка",
    "Единица измерения по РД",
    "По РД",
    "ПИ",
    "Вовлечено ПИ",
    "Остаток ПИ",
    "Выпуск RFQ / План",
    "Выпуск RFQ / Прогноз",
    "Выпуск RFQ / Факт",
    "Размещение заказа (Выпуск ГП) / План",
    "Размещение заказа (Выпуск ГП) / Прогноз",
    "Размещение заказа (Выпуск ГП) / Факт",
    "Подписание договора / План",
    "Подписание договора / Прогноз",
    "Подписание договора / Факт",
    "Законтрактовано в ЕИ РД",
    "Законтрактовано в ЕИ поставщика",
    "Готовность к отгрузке / План",
    "Готовность к отгрузке / Прогноз",
    "Готовность к отгрузке / Факт",
    "Поставка на площадку / План",
    "Поставка на площадку / Прогноз",
    "Поставка на площадку / Факт",
    "Поставлено в ЕИ Поставщика",
    "ЕИ Поставщика",
    "Комментарий комплектации поставок",
    "Комментарий технической оценки",
    "ЗИП",
)

# One-row BCC dumps drop «Центр закупки» and everything after «По РД».
IDENTITY_CAPTIONS: tuple[str, ...] = (COLUMN_CAPTIONS[0],) + COLUMN_CAPTIONS[2:24]

_REQUIRED_CAPTIONS: tuple[str, ...] = (
    "Поставщик",
    "Титул",
    "Раздел",
    "Спецификация",
    "TAG / Линия",
    "Код РД",
    "№ ревизии РД",
    "Описание",
    "Тех.характеристики",
    "Единица измерения по РД",
    "По РД",
)
_IDENTITY_MARKERS = frozenset({"Спецификация", "№ ревизии РД", "По РД"})
_SUBHEAD_TOKENS = frozenset({"План", "Прогноз", "Факт"})

if len(COLUMN_KEYS) != N_COLS or len(COLUMN_CAPTIONS) != N_COLS:
    raise RuntimeError("COLUMN_KEYS / COLUMN_CAPTIONS must have N_COLS entries")
if len(IDENTITY_CAPTIONS) != 23:
    raise RuntimeError("IDENTITY_CAPTIONS must list the 23-column BCC layout")

IDX_SUPPLIER = 3
IDX_CONTRACT = 4
IDX_TITLE = 8
IDX_MARK = 9
IDX_SPEC = 10
IDX_TAGS = 12
IDX_CODE_1C = 13
IDX_CODE_RD = 14
IDX_RD_REV = 16
IDX_NAME = 17
IDX_TYPE_MARK = 18
IDX_UNITS = 22
IDX_QTY_RD = 23

_KEY_INDEX = {name: i for i, name in enumerate(COLUMN_KEYS)}


def default_pickle_path() -> Path:
    """Return the packaged BCC pickle path under ``rd_catalog/База заказчика``."""

    return DUMP_DIR / DEFAULT_PICKLE_NAME


def normalize_text(value: object) -> str:
    """Collapse NBSP and trim; keep inner spaces."""

    if value is None:
        return ""
    text = str(value).replace("\xa0", " ").replace("\u202f", " ").strip()
    return text


def is_bcc_supplier(supplier: str) -> bool:
    """Return True when the supplier cell is BCC."""

    compact = normalize_text(supplier).replace(" ", "")
    return "БИ.СИ.СИ." in compact


def _cell_text(value: object, *, col: int) -> str:
    if value is None:
        return ""
    if isinstance(value, datetime):
        return value.strftime("%Y-%m-%d")
    if isinstance(value, bool):
        return "Да" if value else "Нет"
    if isinstance(value, float):
        if col in DATE_COL_INDEXES and 20000.0 < value < 80000.0:
            dt = _EXCEL_EPOCH + timedelta(days=value)
            return dt.strftime("%Y-%m-%d")
        if value.is_integer():
            return str(int(value))
        return format(value, "g")
    if isinstance(value, int):
        return str(value)
    return normalize_text(value)


def combine_headers(row0: Sequence[str], row1: Sequence[str], n: int) -> list[str]:
    """Merge the two-row xlsb header into 49 stable captions."""

    last = ""
    out: list[str] = []
    for i in range(n):
        a = row0[i] if i < len(row0) else ""
        b = row1[i] if i < len(row1) else ""
        if a:
            last = a
        if a and b:
            out.append(f"{a} / {b}")
        elif a:
            out.append(a)
        elif b:
            out.append(f"{last} / {b}" if last else b)
        else:
            out.append("")
    return out


def _header_src_for_dest(headers: Sequence[str]) -> tuple[int | None, ...]:
    """Map each canonical caption to a source column index (or None if absent)."""

    by_name: dict[str, int] = {}
    for i, name in enumerate(headers):
        caption = normalize_text(name)
        if caption:
            by_name[caption] = i
    missing = [cap for cap in _REQUIRED_CAPTIONS if cap not in by_name]
    if missing:
        raise ValueError(
            f"PI headers missing {missing}; got {list(headers)[:24]!r}"
        )
    return tuple(by_name.get(cap) for cap in COLUMN_CAPTIONS)


@dataclass(frozen=True, slots=True)
class PiXlsbLayout:
    """Detected PI xlsb header: full 49-col sheet or identity-only BCC dump."""

    kind: str
    headers: tuple[str, ...]
    header_row_count: int
    data_start_row0: int
    src_for_dest: tuple[int | None, ...]
    row0: tuple[str, ...]
    row1: tuple[str, ...]
    source_n_cols: int


def detect_pi_layout(row0: Sequence[str], row1: Sequence[str]) -> PiXlsbLayout:
    """Choose ``full`` vs ``identity`` from the first two physical rows.

    ``full``: row 0 has «Центр закупки» and/or row 1 is План/Прогноз/Факт.
    ``identity``: row 0 already has Спецификация / № ревизии РД / По РД and
    row 1 is data (no date sub-headers).

    Args:
        row0: First sheet row (header texts).
        row1: Second sheet row (sub-header or first data row).

    Returns:
        Layout with a 49-wide ``src_for_dest`` map.

    Raises:
        ValueError: required AutoMTO captions are missing.
    """

    n = max(len(row0), len(row1), 1)
    r0 = tuple(normalize_text(row0[i]) if i < len(row0) else "" for i in range(n))
    r1 = tuple(normalize_text(row1[i]) if i < len(row1) else "" for i in range(n))
    set0 = {text for text in r0 if text}
    set1 = {text for text in r1 if text}
    identity_in_row0 = _IDENTITY_MARKERS <= set0
    two_row = ("Центр закупки" in set0) or bool(_SUBHEAD_TOKENS & set1)
    if identity_in_row0 and not two_row:
        kind = LAYOUT_IDENTITY
        headers = r0
        header_row_count = 1
        data_start_row0 = 1
    else:
        kind = LAYOUT_FULL
        headers = tuple(combine_headers(r0, r1, n))
        header_row_count = 2
        data_start_row0 = 2
    src_for_dest = _header_src_for_dest(headers)
    return PiXlsbLayout(
        kind=kind,
        headers=headers,
        header_row_count=header_row_count,
        data_start_row0=data_start_row0,
        src_for_dest=src_for_dest,
        row0=r0,
        row1=r1,
        source_n_cols=n,
    )


def canonical_pi_row(
    cells: Mapping[int, Any],
    src_for_dest: Sequence[int | None],
) -> list[str]:
    """Project a source xlsb row onto the 49 canonical columns.

    Args:
        cells: pyxlsb column index → raw value.
        src_for_dest: source column for each canonical index (``None`` = empty).

    Returns:
        49 cells; dates converted using the canonical (destination) index.
    """

    if len(src_for_dest) != N_COLS:
        raise ValueError(f"expected {N_COLS} source-index slots, got {len(src_for_dest)}")
    out: list[str] = []
    for dest_i, src_i in enumerate(src_for_dest):
        raw = cells.get(src_i) if src_i is not None else None
        out.append(_cell_text(raw, col=dest_i))
    return out


def parse_spec(spec: str) -> tuple[str, str, str]:
    """Return ``(title, mark, discipline)`` from the Спецификация stem.

    Discipline is ``MTO`` / ``BOM`` / ``BOE`` / ``DS`` / ``other``. Title/mark come from
    :class:`AgccFilenamePatterns` when the stem parses; otherwise empty.
    """

    text = normalize_text(spec)
    kind_m = _DISCIPLINE_RE.search(text)
    kind = kind_m.group(1).upper() if kind_m else "other"
    parts = AgccFilenamePatterns.parse_strict(text) or AgccFilenamePatterns.parse_loose(
        text
    )
    if not parts or not parts.title_system or "-" not in parts.title_system:
        return "", "", kind
    title, mark = parts.title_system.split("-", 1)
    return title.strip(), mark.strip(), kind


@dataclass(frozen=True, slots=True)
class CustomerPiRecord:
    """One filtered PI row with raw cells and parsed AGCC identity."""

    excel_row: int
    raw: tuple[str, ...]
    parsed_title: str
    parsed_mark: str
    discipline: str

    def get(self, key: str) -> str:
        """Return a cell by English ``COLUMN_KEYS`` name."""

        return self.raw[_KEY_INDEX[key]]

    @property
    def spec(self) -> str:
        return self.raw[IDX_SPEC]

    @property
    def rd_revision(self) -> str:
        return self.raw[IDX_RD_REV]

    @property
    def title(self) -> str:
        return self.raw[IDX_TITLE]

    @property
    def mark(self) -> str:
        return self.raw[IDX_MARK]

    @property
    def code_rd(self) -> str:
        return self.raw[IDX_CODE_RD]

    @property
    def code_1c(self) -> str:
        return self.raw[IDX_CODE_1C]

    @property
    def tags(self) -> str:
        return self.raw[IDX_TAGS]

    @property
    def name(self) -> str:
        return self.raw[IDX_NAME]

    @property
    def type_mark(self) -> str:
        return self.raw[IDX_TYPE_MARK]

    @property
    def units_rd(self) -> str:
        return self.raw[IDX_UNITS]

    @property
    def qty_rd(self) -> str:
        return self.raw[IDX_QTY_RD]

    @property
    def supplier(self) -> str:
        return self.raw[IDX_SUPPLIER]

    @property
    def contract(self) -> str:
        return self.raw[IDX_CONTRACT]

    def as_mto_fields(self) -> dict[str, str]:
        """Map PI cells onto catalog canonical MTO field names.

        ``VENDOR`` is left empty: PI «Поставщик» is BCC, not the equipment vendor.
        """

        return {
            "CODE": self.code_rd,
            "UNITS": self.units_rd,
            "VALUES": self.qty_rd,
            "TAGS": self.tags,
            "NAME": self.name,
            "VENDOR": "",
            "TYPE_MARK": self.type_mark,
        }


class CustomerPiStore:
    """In-memory PI dump: rows, meta, lazy indexes."""

    def __init__(
        self,
        *,
        meta: dict[str, Any],
        columns: Sequence[str],
        rows: list[tuple[str, ...]],
        excel_rows: list[int],
        parsed: list[tuple[str, str, str]],
    ) -> None:
        if len(columns) != N_COLS:
            raise ValueError(f"expected {N_COLS} columns, got {len(columns)}")
        if not (len(rows) == len(excel_rows) == len(parsed)):
            raise ValueError("rows / excel_rows / parsed length mismatch")
        self.meta = meta
        self.columns = tuple(columns)
        self.column_keys = COLUMN_KEYS
        self.rows = rows
        self.excel_rows = excel_rows
        self.parsed = parsed
        self._idx_spec: dict[str, list[int]] | None = None
        self._idx_spec_rev: dict[tuple[str, str], list[int]] | None = None
        self._idx_title_mark: dict[tuple[str, str], list[int]] | None = None
        self._idx_code_rd: dict[str, list[int]] | None = None

    def __len__(self) -> int:
        return len(self.rows)

    def record(self, index: int) -> CustomerPiRecord:
        """Return record ``index`` (0-based in the filtered dump)."""

        p_title, p_mark, discipline = self.parsed[index]
        return CustomerPiRecord(
            excel_row=self.excel_rows[index],
            raw=self.rows[index],
            parsed_title=p_title,
            parsed_mark=p_mark,
            discipline=discipline,
        )

    def records(self) -> Iterator[CustomerPiRecord]:
        """Yield all filtered records in dump order."""

        for i in range(len(self.rows)):
            yield self.record(i)

    def by_spec(self) -> Mapping[str, list[int]]:
        """Index: Спецификация stem → row indexes."""

        if self._idx_spec is None:
            idx: dict[str, list[int]] = defaultdict(list)
            for i, raw in enumerate(self.rows):
                spec = raw[IDX_SPEC]
                if spec:
                    idx[spec].append(i)
            self._idx_spec = dict(idx)
        return self._idx_spec

    def by_spec_rev(self) -> Mapping[tuple[str, str], list[int]]:
        """Index: (Спецификация, № ревизии РД) → row indexes."""

        if self._idx_spec_rev is None:
            idx: dict[tuple[str, str], list[int]] = defaultdict(list)
            for i, raw in enumerate(self.rows):
                idx[(raw[IDX_SPEC], raw[IDX_RD_REV])].append(i)
            self._idx_spec_rev = dict(idx)
        return self._idx_spec_rev

    def by_title_mark(self) -> Mapping[tuple[str, str], list[int]]:
        """Index: sheet Титул + Раздел (not parsed stem) → row indexes."""

        if self._idx_title_mark is None:
            idx: dict[tuple[str, str], list[int]] = defaultdict(list)
            for i, raw in enumerate(self.rows):
                idx[(raw[IDX_TITLE], raw[IDX_MARK])].append(i)
            self._idx_title_mark = dict(idx)
        return self._idx_title_mark

    def by_code_rd(self) -> Mapping[str, list[int]]:
        """Index: Код РД → row indexes."""

        if self._idx_code_rd is None:
            idx: dict[str, list[int]] = defaultdict(list)
            for i, raw in enumerate(self.rows):
                code = raw[IDX_CODE_RD]
                if code:
                    idx[code].append(i)
            self._idx_code_rd = dict(idx)
        return self._idx_code_rd

    def records_for_spec(self, spec: str) -> tuple[CustomerPiRecord, ...]:
        """Return dump rows for one «Спецификация» stem, pickle order.

        Args:
            spec: PI stem (``AGCC.287-8445-SOT.MTO-0001``). Empty matches
                rows with an empty spec cell.

        Returns:
            Records whose normalized spec equals ``spec``.
        """

        wanted = normalize_text(spec)
        indexes: list[int] = []
        if wanted:
            for key, group in self.by_spec().items():
                if normalize_text(key) == wanted:
                    indexes.extend(group)
        else:
            indexes = [
                index
                for index, raw in enumerate(self.rows)
                if not normalize_text(raw[IDX_SPEC])
            ]
        indexes.sort()
        return tuple(self.record(index) for index in indexes)

    def to_pickle_payload(self) -> dict[str, Any]:
        """Return the serializable dict written to disk."""

        return {
            "schema_version": SCHEMA_VERSION,
            "format": FORMAT_ID,
            "meta": self.meta,
            "columns": list(self.columns),
            "column_keys": list(COLUMN_KEYS),
            "rows": self.rows,
            "excel_rows": self.excel_rows,
            "parsed": self.parsed,
        }


def load_customer_pi(path: str | Path | None = None) -> CustomerPiStore:
    """Load a dump pickle. Default: ``rd_catalog/База заказчика/customer_pi_bcc.pkl``."""

    pickle_path = Path(path) if path else default_pickle_path()
    with pickle_path.open("rb") as fh:
        payload = pickle.load(fh)
    if not isinstance(payload, dict):
        raise ValueError(f"unexpected pickle type {type(payload)!r} in {pickle_path}")
    if payload.get("format") != FORMAT_ID:
        raise ValueError(
            f"not a customer PI dump: format={payload.get('format')!r} in {pickle_path}"
        )
    if int(payload.get("schema_version", 0)) != SCHEMA_VERSION:
        raise ValueError(
            f"unsupported schema_version={payload.get('schema_version')!r} "
            f"(need {SCHEMA_VERSION}) in {pickle_path}"
        )
    return CustomerPiStore(
        meta=payload["meta"],
        columns=payload["columns"],
        rows=payload["rows"],
        excel_rows=payload["excel_rows"],
        parsed=[tuple(p) for p in payload["parsed"]],
    )


def save_customer_pi(store: CustomerPiStore, path: str | Path | None = None) -> Path:
    """Atomically pickle ``store`` (protocol 4)."""

    pickle_path = Path(path) if path else default_pickle_path()
    pickle_path.parent.mkdir(parents=True, exist_ok=True)
    payload = store.to_pickle_payload()
    fd, tmp_name = tempfile.mkstemp(
        prefix=pickle_path.stem + ".",
        suffix=".tmp",
        dir=str(pickle_path.parent),
    )
    try:
        with os.fdopen(fd, "wb") as fh:
            pickle.dump(payload, fh, protocol=4)
        os.replace(tmp_name, pickle_path)
    except Exception:
        try:
            os.unlink(tmp_name)
        except OSError:
            pass
        raise
    return pickle_path


def _cell_dict(row: Any) -> dict[int, Any]:
    return {c.c: c.v for c in row}


def _row_texts(cells: Mapping[int, Any]) -> list[str]:
    if not cells:
        return []
    n = max(cells) + 1
    return [_cell_text(cells.get(i), col=0) for i in range(n)]


def _keep_bcc_row(
    values: Sequence[str],
    *,
    excel_row: int,
    rows: list[tuple[str, ...]],
    excel_rows: list[int],
    parsed: list[tuple[str, str, str]],
    discipline_counts: dict[str, int],
    spec_revs: dict[str, set[str]],
) -> tuple[int, int, int]:
    """Append one BCC row; return (bcc_delta, title_mismatch_delta, code_bcc_delta)."""

    supplier = values[IDX_SUPPLIER]
    if not is_bcc_supplier(supplier):
        return 0, 0, 0
    spec = values[IDX_SPEC]
    rev = values[IDX_RD_REV]
    p_title, p_mark, discipline = parse_spec(spec)
    mismatch = 0
    if (
        p_title
        and p_mark
        and (p_title != values[IDX_TITLE] or p_mark != values[IDX_MARK])
    ):
        mismatch = 1
    code_bcc = 1 if values[IDX_CODE_RD].upper().startswith("BCC") else 0
    discipline_counts[discipline] += 1
    if spec:
        spec_revs[spec].add(rev)
    rows.append(tuple(values))
    excel_rows.append(excel_row)
    parsed.append((p_title, p_mark, discipline))
    return 1, mismatch, code_bcc


def build_from_xlsb(
    xlsb_path: str | Path,
    *,
    pickle_path: str | Path | None = None,
) -> CustomerPiStore:
    """Stream the xlsb, keep BCC supplier rows, pickle. Contracts are not filtered.

    Accepts the full 49-column two-row export and the 23-column identity dump.
    The pickle always stores 49 canonical columns; absent source fields are empty.

    Args:
        xlsb_path: Customer ``*.xlsb`` export.
        pickle_path: Output path; default packaged dump.

    Returns:
        Loaded :class:`CustomerPiStore` (same object that was written).

    Raises:
        FileNotFoundError: ``xlsb_path`` is missing.
        ImportError: ``pyxlsb`` is not installed.
        ValueError: header is not a known PI layout, or the sheet is empty.
    """

    try:
        from pyxlsb import open_workbook
    except ImportError as exc:
        raise ImportError("pyxlsb is required to parse the customer xlsb") from exc

    src = Path(xlsb_path)
    if not src.exists():
        raise FileNotFoundError(src)
    st = src.stat()
    rows: list[tuple[str, ...]] = []
    excel_rows: list[int] = []
    parsed: list[tuple[str, str, str]] = []
    n_source = 0
    n_bcc = 0
    discipline_counts: dict[str, int] = defaultdict(int)
    n_title_mismatch = 0
    n_code_bcc = 0
    spec_revs: dict[str, set[str]] = defaultdict(set)

    with open_workbook(str(src)) as wb:
        sheet_name = wb.sheets[0]
        with wb.get_sheet(sheet_name) as sh:
            it = sh.rows()
            try:
                r0 = _cell_dict(next(it))
            except StopIteration as exc:
                raise ValueError(f"empty PI sheet in {src}") from exc
            try:
                r1 = _cell_dict(next(it))
            except StopIteration:
                r1 = {}
            layout = detect_pi_layout(_row_texts(r0), _row_texts(r1))
            src_for_dest = layout.src_for_dest

            def _ingest(row0_index: int, cells: Mapping[int, Any]) -> bool:
                nonlocal n_source, n_bcc, n_title_mismatch, n_code_bcc
                values = canonical_pi_row(cells, src_for_dest)
                if not any(values):
                    return False
                n_source += 1
                bcc, mismatch, code_bcc = _keep_bcc_row(
                    values,
                    excel_row=row0_index + 1,
                    rows=rows,
                    excel_rows=excel_rows,
                    parsed=parsed,
                    discipline_counts=discipline_counts,
                    spec_revs=spec_revs,
                )
                n_bcc += bcc
                n_title_mismatch += mismatch
                n_code_bcc += code_bcc
                return True

            if layout.kind == LAYOUT_IDENTITY:
                _ingest(layout.data_start_row0, r1)
                loop_start = layout.data_start_row0 + 1
            else:
                loop_start = layout.data_start_row0
            n_empty_tail = 0
            seen_data = n_source > 0
            for i, row in enumerate(it, start=loop_start):
                ingested = _ingest(i, _cell_dict(row))
                if ingested:
                    n_empty_tail = 0
                    seen_data = True
                    continue
                if seen_data:
                    n_empty_tail += 1
                    if n_empty_tail >= _EMPTY_TAIL_STOP:
                        break

    n_multi_rev = sum(1 for revs in spec_revs.values() if len(revs) > 1)
    header_rows = list(range(layout.header_row_count))
    meta = {
        "format": FORMAT_ID,
        "schema_version": SCHEMA_VERSION,
        "source_path": str(src),
        "source_name": src.name,
        "source_size": st.st_size,
        "source_mtime_ns": st.st_mtime_ns,
        "sheet": sheet_name,
        "layout": layout.kind,
        "source_n_cols": layout.source_n_cols,
        "header_rows": header_rows,
        "data_start_row0": layout.data_start_row0,
        "parsed_at": datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
        "filter": {
            "supplier": SUPPLIER_KEEP,
            "contracts_excluded_prefixes": [],
            "note": "Keep every BCC supplier row; contract filters are applied later by the consumer.",
        },
        "counts": {
            "source_rows": n_source,
            "bcc_rows": n_bcc,
            "kept_rows": len(rows),
            "discipline": dict(discipline_counts),
            "specs": len(spec_revs),
            "specs_with_multiple_rd_rev": n_multi_rev,
            "title_mark_vs_spec_mismatch": n_title_mismatch,
            "code_rd_bcc_prefix": n_code_bcc,
        },
        "row0": list(layout.row0),
        "row1": list(layout.row1),
    }
    out = Path(pickle_path) if pickle_path else default_pickle_path()
    meta["pickle_path"] = str(out)
    store = CustomerPiStore(
        meta=meta,
        columns=list(COLUMN_CAPTIONS),
        rows=rows,
        excel_rows=excel_rows,
        parsed=parsed,
    )
    save_customer_pi(store, out)
    return store


def _cli(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="Parse customer PI xlsb (BCC) into rd_catalog/База заказчика pickle."
    )
    parser.add_argument(
        "--from-xlsb",
        type=Path,
        default=None,
        help="Source xlsb: 49-col TDSheet or 23-col BCC identity dump",
    )
    parser.add_argument(
        "--out",
        type=Path,
        default=None,
        help="Pickle path (default: rd_catalog/База заказчика/customer_pi_bcc.pkl)",
    )
    parser.add_argument(
        "--load",
        action="store_true",
        help="Only load the existing pickle and print counts",
    )
    args = parser.parse_args(argv)
    if args.load:
        store = load_customer_pi(args.out)
        print("loaded", len(store), "rows")
        print("meta.counts", store.meta.get("counts"))
        rec = store.record(0)
        print("sample0", rec.spec, rec.rd_revision, rec.code_rd)
        return 0
    src = args.from_xlsb or DEFAULT_XLSB
    store = build_from_xlsb(src, pickle_path=args.out)
    print("kept", len(store), "rows →", store.meta.get("pickle_path"))
    print("layout", store.meta.get("layout"), "source_n_cols", store.meta.get("source_n_cols"))
    print("counts", store.meta.get("counts"))
    return 0


if __name__ == "__main__":
    sys.exit(_cli())
