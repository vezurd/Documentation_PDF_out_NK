"""Strict parsing for quantity columns (VALUES / VALUES_2) in DS compare."""

from __future__ import annotations

import math
import sys
from dataclasses import dataclass

from prettytable import PrettyTable

from base.base_classes import RowStd, RowType
from base.tables_columns import CODE, DS_SYSTEM, DS_TITLE, NUMBERS, VALUES, VALUES_2


@dataclass(frozen=True)
class QuantityValidationError:
    """One row/column that could not be parsed as a number."""

    row_index: int
    column: str
    raw_value: object
    row_type: str
    numbers: str
    title: str
    system: str
    code: str


def read_raw_quantity(row: RowStd, column: str) -> object:
    """Cell value before ``RowStd.get_value`` float coercion."""
    el = row.el.get(column)
    if el is None:
        return None
    return el.value


def _compact_quantity_text(text: str) -> str:
    """Drop whitespace (space, NBSP ``\\xa0``, etc.) — Excel thousands separator."""
    return "".join(str(text).split())


def try_parse_quantity(value: object) -> tuple[bool, float]:
    """Parse quantity without mutating the source cell.

    Strings like ``2 982`` or ``2\\xa0982`` (Excel ru-RU thousands) become ``2982.0``.
    Formulas (``=...``), non-numeric text, NaN, and infinities fail.

    Args:
        value: Raw cell value.

    Returns:
        ``(True, float)`` on success; ``(False, 0.0)`` if not a plain number.
    """
    if value is None or value == "":
        return True, 0.0
    if isinstance(value, bool):
        return False, 0.0
    if isinstance(value, (int, float)):
        parsed = float(value)
        return (True, parsed) if math.isfinite(parsed) else (False, 0.0)
    if isinstance(value, str):
        text = value.strip()
        if not text:
            return True, 0.0
        if text.startswith("="):
            return False, 0.0
        compact = _compact_quantity_text(text)
        normalized = compact.replace(",", ".")
        try:
            parsed = float(normalized)
            return (True, parsed) if math.isfinite(parsed) else (False, 0.0)
        except ValueError:
            return False, 0.0
    return False, 0.0


def normalize_row_quantity(row: RowStd, column: str) -> None:
    """Write parsed float back to ``row.el[column].value`` after validation."""
    raw = read_raw_quantity(row, column)
    ok, parsed = try_parse_quantity(raw)
    if not ok:
        return
    if column not in row.el:
        from base.base_classes import CheckElement

        row.el[column] = CheckElement(None)
    row.el[column].value = parsed


def collect_quantity_errors(
    rows: list[RowStd],
    *,
    source_label: str,
    columns: tuple[str, ...] = (VALUES,),
    only_position_row: bool = True,
) -> list[QuantityValidationError]:
    """Collect unparsable quantity values from loaded rows."""
    del source_label  # reserved for future log sections
    errors: list[QuantityValidationError] = []
    for idx, row in enumerate(rows, start=1):
        if only_position_row and row.row_type != RowType.position_row:
            continue
        for column in columns:
            raw = read_raw_quantity(row, column)
            ok, _ = try_parse_quantity(raw)
            if ok:
                continue
            errors.append(
                QuantityValidationError(
                    row_index=idx,
                    column=column,
                    raw_value=raw,
                    row_type=str(row.row_type or ""),
                    numbers=str(row.get_value(NUMBERS) or "").strip(),
                    title=str(read_raw_quantity(row, DS_TITLE) or "").strip(),
                    system=str(read_raw_quantity(row, DS_SYSTEM) or "").strip(),
                    code=str(read_raw_quantity(row, CODE) or "").strip(),
                )
            )
    return errors


def print_quantity_errors_and_exit(
    errors: list[QuantityValidationError],
    *,
    headline: str,
    file_hint: str = "",
) -> None:
    """Print error table and terminate the process with exit code 1."""
    if not errors:
        return

    table = PrettyTable()
    table.field_names = [
        "№",
        "Строка",
        "Колонка",
        "Тип строки",
        "№ п/п",
        "Титул",
        "Раздел",
        "Код",
        "Значение",
    ]
    for col in table.field_names:
        table.align[col] = "l"

    for n, err in enumerate(errors, start=1):
        table.add_row(
            [
                n,
                err.row_index,
                err.column,
                err.row_type,
                err.numbers,
                err.title[:40] if err.title else "",
                err.system[:20] if err.system else "",
                err.code,
                repr(err.raw_value),
            ]
        )

    print("\n" + "=" * 80)
    print(headline)
    if file_hint:
        print(f"Файл: {file_hint}")
    print(
        "Количество должно быть числом (целым или дробным). "
        "Пробелы и неразрывный пробел (\\xa0) внутри числа — разделитель тысяч, как в Excel "
        "(«2 982» → 2982). Не допускаются: формулы (=...), произвольный текст."
    )
    print("=" * 80)
    print(table)
    print("=" * 80)
    print("Исправьте исходный файл и повторите запуск.")
    print("=" * 80 + "\n")
    sys.exit(1)


def validate_rows_quantities_or_exit(
    rows: list[RowStd],
    *,
    headline: str,
    file_hint: str = "",
    columns: tuple[str, ...] = (VALUES,),
    only_position_row: bool = True,
    normalize: bool = True,
) -> None:
    """Validate quantities; on success optionally normalize cells to float."""
    errors = collect_quantity_errors(
        rows,
        source_label=file_hint,
        columns=columns,
        only_position_row=only_position_row,
    )
    print_quantity_errors_and_exit(errors, headline=headline, file_hint=file_hint)
    if normalize:
        for row in rows:
            if only_position_row and row.row_type != RowType.position_row:
                continue
            for column in columns:
                normalize_row_quantity(row, column)


def parse_quantity_strict(value: object, *, context: str = "") -> float:
    """Parse quantity or raise ``ValueError`` (no silent zero on bad input)."""
    ok, parsed = try_parse_quantity(value)
    if not ok:
        raise ValueError(f"Invalid quantity {value!r}{': ' + context if context else ''}")
    return parsed


def quantity_as_float(value: object) -> float:
    """Same as ``parse_quantity_strict`` for internal sums after validation."""
    return parse_quantity_strict(value)
