"""Conversion matrix xlsx load, reconcile, and atomic write."""

from __future__ import annotations

import gc
import hashlib
import os
import time
import uuid
from io import BytesIO
from dataclasses import dataclass
from decimal import ROUND_HALF_UP, Decimal, InvalidOperation
from pathlib import Path
from typing import Callable

from openpyxl import Workbook, load_workbook
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

from RFQ.ds_compare.ds_units_normalize import normalize_units_text
from RFQ.units_convert.models import (
    MATRIX_CORE_LEGACY,
    MATRIX_HEADER,
    PAIR_COEF_PREFIX,
    PAIR_FORMULA_PREFIX,
    PAIR_SOURCE_PREFIX,
    STATUS_COLLISION,
    STATUS_NEEDS_CLARIFICATION_RU,
    STATUS_NO_GOOGLE_RU,
    GoogleUnitsIndex,
    UnitsConversionError,
    normalize_code,
)

YELLOW_FILL = PatternFill(fill_type="solid", fgColor="FFFF00")
FORMULA_GRAY_FILL = PatternFill(fill_type="solid", fgColor="D9D9D9")
PAIR_COLUMN_STRIDE = 3
GOOGLE_UNIT_COL = 4
PAIR_START_COL = 5
_MAX_WRITE_RETRIES = 2
_MAX_PERMISSION_RETRIES = 3

# Persisted production matrix formatting (sheet Matrix, columns A onward).
MATRIX_FIXED_COLUMN_WIDTHS: dict[int, float] = {
    1: 30.42578125,
    2: 127.0,
    3: 13.28515625,
    4: 9.140625,
}
PAIR_1_COLUMN_WIDTHS: tuple[float, float, float] = (13.0, 13.0, 23.0)
PAIR_2_PLUS_COLUMN_WIDTHS: tuple[float, float, float] = (9.140625, 13.0, 23.28515625)
HEADER_ROW_HEIGHT = 30.0
MATRIX_WIDTH_TOLERANCE = 1e-6

HEADER_ALIGNMENT = Alignment(horizontal="left", vertical="center", wrap_text=True)
DATA_ALIGNMENT = Alignment(horizontal="left", wrap_text=True)

RECIPROCAL_DENOM_MIN = 2
RECIPROCAL_DENOM_MAX = 100
RECIPROCAL_SNAP_TOLERANCE = Decimal("0.005")
SOURCE_NAME_COMMENT_WIDTH = 640
SOURCE_NAME_COMMENT_HEIGHT = 320
SOURCE_NAME_COMMENT_AUTHOR = "матрица ЕИ"
_MAX_SOURCE_NAMES = 12

# Test hook: invoked after temp workbook is written, before signature check/replace.
_before_replace_hook: Callable[[Path, Path], None] | None = None


def format_source_name_comment(names: list[str]) -> str:
    """Join unique position names for an Excel comment, with a short overflow line."""
    unique: list[str] = []
    seen: set[str] = set()
    for raw in names:
        text = " ".join(str(raw or "").split())
        if not text or text.startswith("… и ещё") or text in seen:
            continue
        seen.add(text)
        unique.append(text)
    extra = len(unique) - _MAX_SOURCE_NAMES
    shown = unique[:_MAX_SOURCE_NAMES]
    body = "\n".join(shown)
    if extra > 0:
        body = f"{body}\n… и ещё {extra}"
    return body


def merge_source_names(existing: str, incoming: str) -> str:
    """Merge stored comment lines with newly observed position names."""
    lines: list[str] = []
    for block in (existing, incoming):
        lines.extend(str(block or "").splitlines())
    return format_source_name_comment(lines)


def _comment_text_from_cell(cell: object) -> str:
    comment = getattr(cell, "comment", None) if cell is not None else None
    if comment is None:
        return ""
    return str(getattr(comment, "text", "") or "").strip()


def _source_unit_comment(text: str) -> Comment | None:
    body = str(text or "").strip()
    if not body:
        return None
    comment = Comment(body, SOURCE_NAME_COMMENT_AUTHOR)
    comment.width = SOURCE_NAME_COMMENT_WIDTH
    comment.height = SOURCE_NAME_COMMENT_HEIGHT
    return comment


def _apply_source_names_to_row(
    row: MatrixRow,
    *,
    in_google: bool,
    names_for_units: dict[str, str],
) -> None:
    if in_google:
        for pair in row.pairs:
            pair.source_names = ""
        return
    for pair in row.pairs:
        pair.source_names = merge_source_names(
            pair.source_names,
            names_for_units.get(pair.source_normalized, ""),
        )


@dataclass
class MatrixPair:
    """One source-unit to coefficient mapping in the matrix."""

    source_display: str
    source_normalized: str
    coefficient_raw: str
    coefficient: Decimal | None
    is_placeholder: bool
    source_names: str = ""


@dataclass
class MatrixRow:
    """Mutable in-memory matrix row during reconcile."""

    code_display: str
    code_normalized: str
    google_name: str
    code_status: str
    google_display: str
    google_normalized: str
    pairs: list[MatrixPair]

    def clone(self) -> MatrixRow:
        """Return a shallow copy of this row."""
        return MatrixRow(
            code_display=self.code_display,
            code_normalized=self.code_normalized,
            google_name=self.google_name,
            code_status=self.code_status,
            google_display=self.google_display,
            google_normalized=self.google_normalized,
            pairs=[MatrixPair(**pair.__dict__) for pair in self.pairs],
        )


@dataclass(frozen=True)
class MatrixDocument:
    """In-memory matrix workbook content."""

    rows: tuple[MatrixRow, ...]
    layout_current: bool = True


def _pair_sort_key(pair: MatrixPair, *, google_normalized: str) -> tuple[int, str, str]:
    is_identity = int(pair.source_normalized == google_normalized)
    return (0 if is_identity else 1, pair.source_normalized, pair.source_display)


def normalize_reciprocal_coefficient(value: Decimal) -> Decimal:
    """Snap near-reciprocal decimals to exact ``1/n`` for small positive ``n``.

    For positive coefficients below 1, if the entered value is within
    ``RECIPROCAL_SNAP_TOLERANCE`` of ``1/d`` for integer ``d`` in
    ``[RECIPROCAL_DENOM_MIN, RECIPROCAL_DENOM_MAX]`` (denominator chosen via
    ``ROUND_HALF_UP`` on ``1/value``), return the high-precision Decimal
    reciprocal. Otherwise return ``value`` unchanged.

    Args:
        value: Parsed matrix coefficient.

    Returns:
        Effective coefficient for arithmetic and dependency snapshots.
    """
    if value <= 0 or value >= 1:
        return value
    denominator = (Decimal("1") / value).quantize(Decimal("1"), rounding=ROUND_HALF_UP)
    denom_int = int(denominator)
    if RECIPROCAL_DENOM_MIN <= denom_int <= RECIPROCAL_DENOM_MAX:
        reciprocal = Decimal("1") / denominator
        if abs(value - reciprocal) <= RECIPROCAL_SNAP_TOLERANCE:
            return reciprocal
    return value


def _parse_decimal_literal(fragment: str) -> Decimal:
    """Parse one coefficient literal fragment (comma or dot decimal)."""
    compact = "".join(fragment.split()).replace(",", ".")
    return Decimal(compact)


def _try_parse_fraction_coefficient(text: str) -> Decimal | None:
    """Return effective coefficient for ``numerator/denominator`` syntax, else None."""
    parse_text = text[1:].strip() if text.startswith("=") else text
    if "/" not in parse_text:
        return None
    parts = parse_text.split("/")
    if len(parts) != 2:
        return None
    try:
        numerator = _parse_decimal_literal(parts[0])
        denominator = _parse_decimal_literal(parts[1])
    except (InvalidOperation, ValueError):
        return None
    if denominator == 0:
        return None
    return numerator / denominator


def _parse_coefficient(raw: object) -> tuple[str, Decimal | None, bool]:
    text = str(raw or "").strip()
    if not text:
        return "", None, True
    if text == "?":
        return "?", None, True
    fraction_value = _try_parse_fraction_coefficient(text)
    if fraction_value is not None:
        return text, normalize_reciprocal_coefficient(fraction_value), False
    if "/" in text:
        return text, None, True
    compact = "".join(text.split()).replace(",", ".")
    try:
        value = Decimal(compact)
    except (InvalidOperation, ValueError):
        return text, None, True
    return text, normalize_reciprocal_coefficient(value), False


@dataclass(frozen=True)
class _MatrixCoreLayout:
    """1-based worksheet columns for the fixed matrix prefix."""

    code_col: int
    name_col: int | None
    status_col: int
    google_col: int
    pair_start_col: int
    is_current: bool

    @property
    def length(self) -> int:
        """Number of fixed prefix columns."""
        return self.pair_start_col - 1


def _parse_core_header(header: tuple[str, ...]) -> _MatrixCoreLayout:
    if len(header) >= len(MATRIX_HEADER) and tuple(header[: len(MATRIX_HEADER)]) == MATRIX_HEADER:
        return _MatrixCoreLayout(
            code_col=1,
            name_col=2,
            status_col=3,
            google_col=4,
            pair_start_col=PAIR_START_COL,
            is_current=True,
        )
    if (
        len(header) >= len(MATRIX_CORE_LEGACY)
        and tuple(header[: len(MATRIX_CORE_LEGACY)]) == MATRIX_CORE_LEGACY
    ):
        return _MatrixCoreLayout(
            code_col=1,
            name_col=None,
            status_col=2,
            google_col=3,
            pair_start_col=len(MATRIX_CORE_LEGACY) + 1,
            is_current=False,
        )
    raise UnitsConversionError("Строка заголовка матрицы должна соответствовать требуемому шаблону")


@dataclass(frozen=True)
class _PairHeaderColumns:
    """1-based worksheet columns for one source/coef/(optional formula) group."""

    source_col: int
    coef_col: int
    formula_col: int | None


def _parse_header_pairs(
    header: tuple[str, ...],
    core: _MatrixCoreLayout,
) -> tuple[tuple[_PairHeaderColumns, ...], bool]:
    pairs: list[_PairHeaderColumns] = []
    index = core.length
    all_triple = True
    while index + 1 < len(header):
        source_title = header[index]
        coef_title = header[index + 1]
        if not source_title.startswith(PAIR_SOURCE_PREFIX):
            break
        if not coef_title.startswith(PAIR_COEF_PREFIX):
            break
        source_col = index + 1
        coef_col = index + 2
        formula_col: int | None = None
        if index + 2 < len(header) and header[index + 2].startswith(PAIR_FORMULA_PREFIX):
            formula_col = index + 3
            index += PAIR_COLUMN_STRIDE
        else:
            all_triple = False
            index += 2
        pairs.append(
            _PairHeaderColumns(
                source_col=source_col,
                coef_col=coef_col,
                formula_col=formula_col,
            )
        )
    if not pairs:
        layout_current = core.is_current and len(header) == core.length
    else:
        layout_current = (
            core.is_current
            and all_triple
            and all(pair.formula_col is not None for pair in pairs)
        )
    return tuple(pairs), layout_current


def _build_header(pair_count: int) -> tuple[str, ...]:
    header = list(MATRIX_HEADER)
    for index in range(1, pair_count + 1):
        header.append(f"{PAIR_SOURCE_PREFIX}{index}")
        header.append(f"{PAIR_COEF_PREFIX}{index}")
        header.append(f"{PAIR_FORMULA_PREFIX}{index}")
    return tuple(header)


def _pair_conversion_formula(
    *,
    row_index: int,
    source_col: int,
    coef_col: int,
    google_col: int = GOOGLE_UNIT_COL,
) -> str:
    source_ref = f"{get_column_letter(source_col)}{row_index}"
    coef_ref = f"{get_column_letter(coef_col)}{row_index}"
    google_ref = f"${get_column_letter(google_col)}{row_index}"
    return (
        f'=IF(OR({source_ref}="",{coef_ref}="",{coef_ref}="?",{google_ref}=""),"",'
        f'{coef_ref}&"/"&{google_ref}&" = 1/"&{source_ref})'
    )


def _row_cell(row_cells: tuple, col: int):
    """Return worksheet cell from a 1-based column index in an ``iter_rows`` tuple."""
    if col < 1 or col > len(row_cells):
        return None
    return row_cells[col - 1]


def _coef_cell_needs_text_format(cell) -> bool:
    """Return True when a coefficient cell is not Excel Text (``@``) formatted."""
    if cell is None:
        return False
    return getattr(cell, "number_format", "@") != "@"


def expected_matrix_column_width(col: int, *, pair_count: int) -> float:
    """Return persisted column width for one 1-based matrix column."""
    fixed = MATRIX_FIXED_COLUMN_WIDTHS.get(col)
    if fixed is not None:
        return fixed
    offset = col - PAIR_START_COL
    if offset < 0:
        raise ValueError(f"Matrix column {col} is outside the fixed prefix")
    pair_index = offset // PAIR_COLUMN_STRIDE + 1
    position_in_pair = offset % PAIR_COLUMN_STRIDE
    if pair_index > pair_count:
        raise ValueError(
            f"Matrix column {col} exceeds pair_count={pair_count}"
        )
    widths = PAIR_1_COLUMN_WIDTHS if pair_index == 1 else PAIR_2_PLUS_COLUMN_WIDTHS
    return widths[position_in_pair]


def _width_matches(actual: float | None, expected: float) -> bool:
    if actual is None:
        return False
    return abs(actual - expected) <= MATRIX_WIDTH_TOLERANCE


def _alignment_matches_header(cell) -> bool:
    if cell is None:
        return False
    alignment = getattr(cell, "alignment", None)
    if alignment is None:
        return False
    return (
        alignment.horizontal == "left"
        and alignment.vertical == "center"
        and alignment.wrap_text is True
    )


def _alignment_matches_data(cell) -> bool:
    if cell is None:
        return True
    alignment = getattr(cell, "alignment", None)
    if alignment is None:
        return False
    vertical = alignment.vertical
    return (
        alignment.horizontal == "left"
        and (vertical is None or vertical == "bottom")
        and alignment.wrap_text is True
    )


def _matrix_formatting_current(ws: Worksheet, *, pair_count: int, last_col: int) -> bool:
    """Return True when persisted widths, header height, and header alignment match."""
    header_height = ws.row_dimensions[1].height
    if not _width_matches(header_height, HEADER_ROW_HEIGHT):
        return False
    for col in range(1, last_col + 1):
        letter = get_column_letter(col)
        dimension = ws.column_dimensions.get(letter)
        actual_width = dimension.width if dimension is not None else None
        expected_width = expected_matrix_column_width(col, pair_count=pair_count)
        if not _width_matches(actual_width, expected_width):
            return False
        if not _alignment_matches_header(ws.cell(row=1, column=col)):
            return False
    return True


def _read_header(ws: Worksheet) -> tuple[tuple[str, ...], _MatrixCoreLayout]:
    values: list[str] = []
    min_core = len(MATRIX_CORE_LEGACY)
    for col in range(1, ws.max_column + 1):
        cell_value = ws.cell(row=1, column=col).value
        if cell_value is None and col > min_core:
            break
        values.append(str(cell_value or "").strip())
    if len(values) < min_core:
        raise UnitsConversionError("Строка заголовка матрицы неполная")
    core = _parse_core_header(tuple(values))
    return tuple(values), core


def _rows_by_code(document: MatrixDocument) -> dict[str, MatrixRow]:
    rows: dict[str, MatrixRow] = {}
    for row in document.rows:
        if row.code_normalized in rows:
            raise UnitsConversionError(
                f"Дублирующийся нормализованный код матрицы {row.code_normalized!r}"
            )
        rows[row.code_normalized] = row.clone()
    return rows


def load_matrix(path: Path) -> MatrixDocument:
    """Load matrix xlsx from ``path`` or return an empty document if missing."""
    if not path.exists():
        return MatrixDocument(rows=())
    try:
        payload = path.read_bytes()
    except PermissionError as exc:
        raise _matrix_busy_error(path, action="прочитать") from exc
    wb = load_workbook(BytesIO(payload), data_only=True, read_only=True)
    try:
        ws = wb.active
        header, core = _read_header(ws)
        pair_columns, header_layout_current = _parse_header_pairs(header, core)
        pair_count = len(pair_columns) if pair_columns else 0
        last_col = len(header)
        layout_current = header_layout_current
        check_data_alignment = header_layout_current
        rows: list[MatrixRow] = []
        seen_codes: set[str] = set()
        for row_idx, row_cells in enumerate(ws.iter_rows(min_row=2, values_only=False), start=2):
            code_cell = _row_cell(row_cells, core.code_col)
            code_raw = code_cell.value if code_cell is not None else None
            if code_raw is None or str(code_raw).strip() == "":
                continue
            if check_data_alignment:
                for col in range(1, last_col + 1):
                    if not _alignment_matches_data(_row_cell(row_cells, col)):
                        layout_current = False
                        check_data_alignment = False
                        break
            code_norm = normalize_code(code_raw)
            if code_norm in seen_codes:
                raise UnitsConversionError(
                    f"Дублирующийся нормализованный код матрицы {code_norm!r} в строке {row_idx}"
                )
            seen_codes.add(code_norm)
            status_cell = _row_cell(row_cells, core.status_col)
            status = str((status_cell.value if status_cell is not None else "") or "").strip()
            google_cell = _row_cell(row_cells, core.google_col)
            google_raw = google_cell.value if google_cell is not None else None
            google_norm = normalize_units_text(google_raw)
            google_name = ""
            if core.name_col is not None:
                name_cell = _row_cell(row_cells, core.name_col)
                google_name = str(
                    (name_cell.value if name_cell is not None else "") or ""
                ).strip()
            pairs: list[MatrixPair] = []
            for pair_cols in pair_columns:
                coef_cell = _row_cell(row_cells, pair_cols.coef_col)
                if _coef_cell_needs_text_format(coef_cell):
                    layout_current = False
                source_cell = _row_cell(row_cells, pair_cols.source_col)
                source_raw = source_cell.value if source_cell is not None else None
                coef_raw = coef_cell.value if coef_cell is not None else None
                source_display = str(source_raw or "").strip()
                source_norm = normalize_units_text(source_raw)
                coef_text, coef_value, is_placeholder = _parse_coefficient(coef_raw)
                if not source_display and not coef_text:
                    continue
                pairs.append(
                    MatrixPair(
                        source_display=source_display,
                        source_normalized=source_norm,
                        coefficient_raw=coef_text,
                        coefficient=coef_value,
                        is_placeholder=is_placeholder,
                        source_names=_comment_text_from_cell(source_cell),
                    )
                )
            rows.append(
                MatrixRow(
                    code_display=str(code_raw).strip(),
                    code_normalized=code_norm,
                    google_name=google_name,
                    code_status=status,
                    google_display=str(google_raw or "").strip(),
                    google_normalized=google_norm,
                    pairs=pairs,
                )
            )
        rows.sort(key=lambda item: item.code_normalized)
        if header_layout_current and pair_count > 0:
            wb_fmt = load_workbook(BytesIO(payload), data_only=True, read_only=False)
            try:
                ws_fmt = wb_fmt.active
                if not _matrix_formatting_current(
                    ws_fmt,
                    pair_count=pair_count,
                    last_col=last_col,
                ):
                    layout_current = False
            finally:
                wb_fmt.close()
        return MatrixDocument(rows=tuple(rows), layout_current=layout_current)
    finally:
        wb.close()


def _ensure_pair(
    row: MatrixRow,
    *,
    source_display: str,
    source_normalized: str,
    coefficient_raw: str,
    coefficient: Decimal | None,
    is_placeholder: bool,
) -> None:
    for pair in row.pairs:
        if pair.source_normalized == source_normalized:
            pair.source_display = source_display or pair.source_display
            pair.coefficient_raw = coefficient_raw
            pair.coefficient = coefficient
            pair.is_placeholder = is_placeholder
            return
    row.pairs.append(
        MatrixPair(
            source_display=source_display,
            source_normalized=source_normalized,
            coefficient_raw=coefficient_raw,
            coefficient=coefficient,
            is_placeholder=is_placeholder,
        )
    )


def _sort_pairs(row: MatrixRow) -> None:
    row.pairs.sort(
        key=lambda pair: _pair_sort_key(pair, google_normalized=row.google_normalized)
    )


def _has_question_placeholder(row: MatrixRow) -> bool:
    return any(pair.coefficient_raw == "?" for pair in row.pairs)


def resolved_conversion_target(
    code_normalized: str,
    *,
    google_index: GoogleUnitsIndex,
    matrix_row: MatrixRow | None,
) -> tuple[str, str]:
    """Return ``(normalized, display)`` target unit for a code.

    Google base wins. If Google has no unit, column D of the matrix is the
    local canonical unit so mixed source units can be converted without
    adding the code to Google.

    Args:
        code_normalized: Normalized material code.
        google_index: Google units lookup.
        matrix_row: Loaded matrix row for the code, if any.

    Returns:
        Empty strings when neither Google nor matrix D provides a target.
    """
    google_norm = google_index.google_unit(code_normalized) or ""
    if google_norm:
        display = google_index.google_unit_display(code_normalized) or google_norm
        return google_norm, display
    if matrix_row is None:
        return "", ""
    display = (matrix_row.google_display or "").strip()
    norm = matrix_row.google_normalized or normalize_units_text(display)
    if not norm:
        return "", ""
    return norm, display or norm


def _bind_row_to_target(
    row: MatrixRow,
    *,
    target_norm: str,
    target_display: str,
    active_units: dict[str, str],
    target_changed: bool,
) -> None:
    if target_changed:
        for pair in row.pairs:
            if pair.source_normalized != target_norm:
                pair.coefficient_raw = "?"
                pair.coefficient = None
                pair.is_placeholder = True
    _ensure_pair(
        row,
        source_display=target_display,
        source_normalized=target_norm,
        coefficient_raw="1",
        coefficient=Decimal("1"),
        is_placeholder=False,
    )
    for source_norm, source_display in sorted(active_units.items()):
        if source_norm == target_norm:
            continue
        existing = next(
            (pair for pair in row.pairs if pair.source_normalized == source_norm),
            None,
        )
        if existing is None or target_changed:
            _ensure_pair(
                row,
                source_display=source_display,
                source_normalized=source_norm,
                coefficient_raw="?",
                coefficient=None,
                is_placeholder=True,
            )
        elif existing.coefficient is None:
            existing.source_display = source_display or existing.source_display
            existing.coefficient_raw = "?"
            existing.is_placeholder = True


def _record_active_source_units(row: MatrixRow, active_units: dict[str, str]) -> None:
    """Keep every observed source unit on a collision row (inventory, coef 1)."""
    for source_norm, source_display in sorted(active_units.items()):
        existing = next(
            (pair for pair in row.pairs if pair.source_normalized == source_norm),
            None,
        )
        if existing is None:
            _ensure_pair(
                row,
                source_display=source_display,
                source_normalized=source_norm,
                coefficient_raw="1",
                coefficient=Decimal("1"),
                is_placeholder=False,
            )
        else:
            existing.source_display = source_display or existing.source_display


def reconcile_matrix(
    document: MatrixDocument,
    *,
    google_index: GoogleUnitsIndex,
    active_by_code: dict[str, dict[str, str]],
    code_display_by_code: dict[str, str] | None = None,
    names_by_code: dict[str, dict[str, str]] | None = None,
) -> MatrixDocument:
    """Reconcile matrix rows with Google lookup for observed/history codes only."""
    code_display_by_code = code_display_by_code or {}
    names_by_code = names_by_code or {}
    rows = _rows_by_code(document)
    all_codes = set(rows) | set(active_by_code)
    for code_norm in sorted(all_codes):
        row = rows.get(code_norm)
        google_norm = google_index.google_unit(code_norm)
        google_display = google_index.google_unit_display(code_norm) or ""
        google_name = google_index.google_name(code_norm) or ""
        active_units = active_by_code.get(code_norm, {})
        canonical_code_display = (
            code_display_by_code.get(code_norm)
            or google_index.display_code_by_code.get(code_norm)
            or code_norm
        )
        if row is None:
            row = MatrixRow(
                code_display=canonical_code_display,
                code_normalized=code_norm,
                google_name=google_name if google_norm else "",
                code_status="",
                google_display=google_display if google_norm else "",
                google_normalized=google_norm or "",
                pairs=[],
            )
            rows[code_norm] = row
        elif code_display_by_code.get(code_norm):
            row.code_display = code_display_by_code[code_norm]
        row.google_name = google_name if google_norm else ""
        previous_google = row.google_normalized
        previous_status = row.code_status
        bound_target = False
        if google_norm:
            row.google_display = google_display or google_norm
            row.google_normalized = google_norm
            row.code_status = ""
            _bind_row_to_target(
                row,
                target_norm=google_norm,
                target_display=row.google_display,
                active_units=active_units,
                target_changed=previous_google != google_norm,
            )
            bound_target = True
        else:
            local_display = (row.google_display or "").strip()
            local_norm = row.google_normalized or normalize_units_text(local_display)
            if local_norm:
                row.google_display = local_display or local_norm
                row.google_normalized = local_norm
                row.code_status = ""
                promoted = previous_status == STATUS_COLLISION
                _bind_row_to_target(
                    row,
                    target_norm=local_norm,
                    target_display=row.google_display,
                    active_units=active_units,
                    target_changed=previous_google != local_norm or promoted,
                )
                bound_target = True
            else:
                active_norms = sorted(active_units)
                if len(active_norms) > 1:
                    row.code_status = STATUS_COLLISION
                    row.google_display = ""
                    row.google_normalized = ""
                    _record_active_source_units(row, active_units)
                elif len(active_norms) == 1:
                    only_norm = active_norms[0]
                    only_display = active_units[only_norm]
                    row.code_status = STATUS_NO_GOOGLE_RU
                    row.google_display = ""
                    row.google_normalized = ""
                    _ensure_pair(
                        row,
                        source_display=only_display,
                        source_normalized=only_norm,
                        coefficient_raw="1",
                        coefficient=Decimal("1"),
                        is_placeholder=False,
                    )
                elif not row.pairs and row.google_normalized:
                    _ensure_pair(
                        row,
                        source_display=row.google_display or row.google_normalized,
                        source_normalized=row.google_normalized,
                        coefficient_raw="1",
                        coefficient=Decimal("1"),
                        is_placeholder=False,
                    )
        _apply_source_names_to_row(
            row,
            in_google=bool(google_norm),
            names_for_units=names_by_code.get(code_norm, {}),
        )
        _sort_pairs(row)
        if bound_target and _has_question_placeholder(row):
            row.code_status = STATUS_NEEDS_CLARIFICATION_RU
    ordered = [rows[key] for key in sorted(rows)]
    return MatrixDocument(rows=tuple(ordered), layout_current=document.layout_current)


def _required_pair_count(document: MatrixDocument) -> int:
    max_pairs = 0
    for row in document.rows:
        max_pairs = max(max_pairs, len(row.pairs))
    return max(max_pairs, 1)


def _document_semantic_key(document: MatrixDocument) -> bytes:
    """Hash matrix semantics for skip-write; includes display-only google_name."""
    parts: list[str] = []
    for row in document.rows:
        pair_chunks = []
        for pair in row.pairs:
            pair_chunks.append(
                "|".join(
                    [
                        pair.source_normalized,
                        pair.source_display,
                        pair.coefficient_raw,
                        str(pair.coefficient) if pair.coefficient is not None else "",
                        pair.source_names,
                    ]
                )
            )
        parts.append(
            "\t".join(
                [
                    row.code_normalized,
                    row.code_display,
                    row.google_name,
                    row.code_status,
                    row.google_normalized,
                    row.google_display,
                    ";".join(pair_chunks),
                ]
            )
        )
    payload = "\n".join(parts).encode("utf-8")
    return hashlib.sha256(payload).digest()


def _file_signature(path: Path) -> tuple[int, int]:
    stat = path.stat()
    return stat.st_mtime_ns, stat.st_size


def _signature_matches(path: Path, *, expected_existed: bool, expected_signature: tuple[int, int] | None) -> bool:
    exists = path.exists()
    if exists != expected_existed:
        return False
    if not expected_existed:
        return True
    assert expected_signature is not None
    return _file_signature(path) == expected_signature


def _safe_unlink(path: Path) -> None:
    try:
        if path.exists():
            path.unlink()
    except OSError:
        pass


def _matrix_busy_error(path: Path, *, action: str) -> UnitsConversionError:
    """Build a fatal error when the matrix xlsx is locked or inaccessible."""
    return UnitsConversionError(
        f"Не удалось {action} матрицу ед. изм.: файл занят другой программой.\n"
        f"{path}\n"
        "\nЧто нужно сделать:\n"
        f"1. Закройте файл матрицы в Excel («{path.name}»).\n"
        "2. Повторите сбор частей или запуск RFP."
    )


def _apply_matrix_column_widths(ws: Worksheet, *, pair_count: int) -> None:
    last_col = GOOGLE_UNIT_COL + pair_count * PAIR_COLUMN_STRIDE
    for col in range(1, last_col + 1):
        letter = get_column_letter(col)
        ws.column_dimensions[letter].width = expected_matrix_column_width(
            col,
            pair_count=pair_count,
        )


def _apply_matrix_cell_formatting(ws: Worksheet, *, last_col: int, last_row: int) -> None:
    ws.row_dimensions[1].height = HEADER_ROW_HEIGHT
    for col in range(1, last_col + 1):
        ws.cell(row=1, column=col).alignment = HEADER_ALIGNMENT
    for row_index in range(2, last_row + 1):
        for col in range(1, last_col + 1):
            ws.cell(row=row_index, column=col).alignment = DATA_ALIGNMENT


def _write_temp_workbook(tmp_path: Path, document: MatrixDocument) -> None:
    pair_count = _required_pair_count(document)
    header = _build_header(pair_count)
    wb = Workbook()
    try:
        ws = wb.active
        ws.title = "Matrix"
        for col_index, title in enumerate(header, start=1):
            header_cell = ws.cell(row=1, column=col_index, value=title)
            if title.startswith(PAIR_FORMULA_PREFIX):
                header_cell.fill = FORMULA_GRAY_FILL
        for row_index, row in enumerate(document.rows, start=2):
            ws.cell(row=row_index, column=1, value=row.code_display)
            ws.cell(row=row_index, column=2, value=row.google_name)
            status_cell = ws.cell(row=row_index, column=3, value=row.code_status)
            if row.code_status == STATUS_NEEDS_CLARIFICATION_RU:
                status_cell.fill = YELLOW_FILL
            ws.cell(row=row_index, column=4, value=row.google_display)
            for pair_index, pair in enumerate(row.pairs):
                source_col = PAIR_START_COL + pair_index * PAIR_COLUMN_STRIDE
                coef_col = source_col + 1
                formula_col = source_col + 2
                source_cell = ws.cell(row=row_index, column=source_col, value=pair.source_display)
                source_comment = _source_unit_comment(pair.source_names)
                if source_comment is not None:
                    source_cell.comment = source_comment
                coef_cell = ws.cell(
                    row=row_index,
                    column=coef_col,
                    value=pair.coefficient_raw,
                )
                coef_cell.number_format = "@"
                if pair.is_placeholder and pair.coefficient_raw == "?":
                    coef_cell.fill = YELLOW_FILL
                formula_cell = ws.cell(
                    row=row_index,
                    column=formula_col,
                    value=_pair_conversion_formula(
                        row_index=row_index,
                        source_col=source_col,
                        coef_col=coef_col,
                        google_col=GOOGLE_UNIT_COL,
                    ),
                )
                formula_cell.fill = FORMULA_GRAY_FILL
            for pair_index in range(len(row.pairs), pair_count):
                source_col = PAIR_START_COL + pair_index * PAIR_COLUMN_STRIDE
                coef_col = source_col + 1
                formula_col = source_col + 2
                blank_coef_cell = ws.cell(row=row_index, column=coef_col)
                blank_coef_cell.number_format = "@"
                formula_cell = ws.cell(
                    row=row_index,
                    column=formula_col,
                    value=_pair_conversion_formula(
                        row_index=row_index,
                        source_col=source_col,
                        coef_col=coef_col,
                        google_col=GOOGLE_UNIT_COL,
                    ),
                )
                formula_cell.fill = FORMULA_GRAY_FILL
        last_col = len(header)
        last_row = max(1, len(document.rows) + 1)
        _apply_matrix_column_widths(ws, pair_count=pair_count)
        _apply_matrix_cell_formatting(ws, last_col=last_col, last_row=last_row)
        ws.auto_filter.ref = f"A1:{get_column_letter(last_col)}{last_row}"
        ws.freeze_panes = "A2"
        tmp_path.parent.mkdir(parents=True, exist_ok=True)
        buffer = BytesIO()
        wb.save(buffer)
        tmp_path.write_bytes(buffer.getvalue())
    finally:
        wb.close()


def _atomic_replace_with_permission_retry(tmp_path: Path, path: Path) -> None:
    for attempt in range(_MAX_PERMISSION_RETRIES):
        try:
            os.replace(tmp_path, path)
            return
        except PermissionError:
            gc.collect()
            if attempt == _MAX_PERMISSION_RETRIES - 1:
                raise
            time.sleep(0.05)


def save_matrix_atomic(
    path: Path,
    document: MatrixDocument,
    *,
    google_index: GoogleUnitsIndex,
    active_by_code: dict[str, dict[str, str]],
    code_display_by_code: dict[str, str] | None = None,
    names_by_code: dict[str, dict[str, str]] | None = None,
    previous_signature: tuple[int, int] | None,
    previous_semantic: bytes | None,
) -> None:
    """Write matrix xlsx atomically with bounded concurrent-change retries.

    Raises:
        UnitsConversionError: If another process changed the file during write
            retries, or if the destination is locked (for example open in Excel).
    """
    code_display_by_code = code_display_by_code or {}
    names_by_code = names_by_code or {}
    semantic = _document_semantic_key(document)
    if (
        path.exists()
        and document.layout_current
        and previous_semantic == semantic
        and _signature_matches(
            path,
            expected_existed=previous_signature is not None,
            expected_signature=previous_signature,
        )
    ):
        return

    expected_existed = previous_signature is not None
    expected_signature = previous_signature
    current_document = document
    attempt = 0

    while attempt <= _MAX_WRITE_RETRIES:
        tmp_path = path.with_name(f"{path.stem}.{os.getpid()}.{uuid.uuid4().hex}.tmp.xlsx")
        try:
            _write_temp_workbook(tmp_path, current_document)
            if _before_replace_hook is not None:
                _before_replace_hook(tmp_path, path)

            if not _signature_matches(
                path,
                expected_existed=expected_existed,
                expected_signature=expected_signature,
            ):
                _safe_unlink(tmp_path)
                if path.exists():
                    reloaded = load_matrix(path)
                    expected_signature = _file_signature(path)
                    expected_existed = True
                else:
                    reloaded = MatrixDocument(rows=())
                    expected_signature = None
                    expected_existed = False
                current_document = reconcile_matrix(
                    reloaded,
                    google_index=google_index,
                    active_by_code=active_by_code,
                    code_display_by_code=code_display_by_code,
                    names_by_code=names_by_code,
                )
                attempt += 1
                continue

            _atomic_replace_with_permission_retry(tmp_path, path)
            return
        except PermissionError as exc:
            _safe_unlink(tmp_path)
            raise _matrix_busy_error(path, action="записать") from exc
        except Exception:
            _safe_unlink(tmp_path)
            raise

    raise UnitsConversionError(
        f"Файл матрицы изменился параллельно и не удалось сохранить: {path}"
    )


def lookup_coefficient(
    document: MatrixDocument,
    *,
    code_normalized: str,
    source_normalized: str,
) -> MatrixPair | None:
    """Return matrix pair for code/source unit, if present."""
    for row in document.rows:
        if row.code_normalized != code_normalized:
            continue
        for pair in row.pairs:
            if pair.source_normalized == source_normalized:
                return pair
    return None


def matrix_row_for_code(document: MatrixDocument, code_normalized: str) -> MatrixRow | None:
    """Return matrix row for normalized code."""
    for row in document.rows:
        if row.code_normalized == code_normalized:
            return row
    return None
