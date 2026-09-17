"""Immutable models and helpers for RFQ units conversion."""

from __future__ import annotations

from collections.abc import Iterable, Mapping
from dataclasses import dataclass, field
from decimal import Decimal
from enum import Enum
from typing import Any

from base.base_classes import RowStd
from base.tables_columns import CODE, NAME, UNITS
from RFQ.ds_compare.ds_units_normalize import normalize_units_text

# v5: parts tagged counters keep leftover lot qty on <NO_TAG> when lot > tags.
ALGORITHM_VERSION = "v5"
RESIDUAL_TOLERANCE = Decimal("1e-9")

MATRIX_CORE_LEGACY = (
    "Код",
    "Статус кода",
    "Ед. изм. Google",
)
MATRIX_HEADER = (
    "Код",
    "Наименование Google",
    "Статус кода",
    "Ед. изм. Google",
)
PAIR_SOURCE_PREFIX = "Найденная ЕИ "
PAIR_COEF_PREFIX = "Коэффициент "
PAIR_FORMULA_PREFIX = "Формула "

STATUS_IDENTITY = "identity"
STATUS_CONVERTED = "converted"
STATUS_NO_GOOGLE = "no_google"
STATUS_NO_GOOGLE_RU = "Нет в базе"
STATUS_NEEDS_CLARIFICATION_RU = "Требует уточнения"
STATUS_COLLISION = "collision"


class UnitsConversionError(Exception):
    """Fatal units conversion or matrix validation error."""


class ConversionInvariant(Enum):
    """Quantity invariant enforced when a row is actually changed."""

    NONE = "none"
    TAGS_EQUAL = "tags_equal"
    TAGS_NOT_EXCEED_QUANTITY = "tags_not_exceed_quantity"


def normalize_code(value: object) -> str:
    """Strip, remove all whitespace, and upper-case a code value.

    Args:
        value: Raw code from RowStd or matrix cell.

    Returns:
        Normalized uppercase code without whitespace.
    """
    text = str(value or "")
    return "".join(text.split()).upper()


def parse_decimal_quantity(value: object, *, location: str = "") -> Decimal:
    """Parse quantity text into ``Decimal`` without float rounding.

    Args:
        value: Numeric value or string with optional comma/inner whitespace.
        location: Optional location hint for error messages.

    Raises:
        UnitsConversionError: If value is empty or not numeric.
    """
    if isinstance(value, Decimal):
        return value
    if value is None:
        raise UnitsConversionError(_qty_error_message(value, location))
    text = str(value).strip()
    if not text:
        raise UnitsConversionError(_qty_error_message(value, location))
    compact = "".join(text.split()).replace(",", ".")
    try:
        return Decimal(compact)
    except Exception as exc:  # noqa: BLE001 - wrap as domain error
        raise UnitsConversionError(_qty_error_message(value, location)) from exc


def _qty_error_message(value: object, location: str) -> str:
    suffix = f" источник={location}" if location else ""
    return f"Некорректное количество {value!r}{suffix}"


@dataclass(frozen=True)
class GoogleUnitsIndex:
    """Immutable Google units lookup keyed by normalized code."""

    units_by_code: Mapping[str, str]
    display_units_by_code: Mapping[str, str]
    display_code_by_code: Mapping[str, str]
    display_names_by_code: Mapping[str, str] = field(default_factory=dict)

    def google_unit(self, code_normalized: str) -> str | None:
        """Return normalized Google units for a code, if present."""
        return self.units_by_code.get(code_normalized)

    def google_unit_display(self, code_normalized: str) -> str | None:
        """Return display Google units text for a code, if present."""
        return self.display_units_by_code.get(code_normalized)

    def google_name(self, code_normalized: str) -> str | None:
        """Return display Google item name for a code, if present."""
        return self.display_names_by_code.get(code_normalized)


def build_google_units_index(rows: Iterable[RowStd]) -> GoogleUnitsIndex:
    """Build an immutable Google units index from ``RowStd`` CODE/UNITS rows.

    Args:
        rows: Iterable of Google base ``RowStd`` objects.

    Returns:
        ``GoogleUnitsIndex`` with one normalized unit per code.

    Raises:
        UnitsConversionError: If one code maps to multiple distinct Google units.
    """
    units_by_code: dict[str, str] = {}
    display_units_by_code: dict[str, str] = {}
    display_code_by_code: dict[str, str] = {}
    display_names_by_code: dict[str, str] = {}
    for row in rows:
        code_raw = row.get_value(CODE)
        units_raw = row.get_value(UNITS)
        name_element = row.el.get(NAME)
        name_raw = name_element.value if name_element is not None else None
        code_norm = normalize_code(code_raw)
        units_norm = normalize_units_text(units_raw)
        if not code_norm or not units_norm:
            continue
        display_code_by_code.setdefault(code_norm, str(code_raw or "").strip())
        name_display = str(name_raw or "").strip()
        if name_display:
            display_names_by_code.setdefault(code_norm, name_display)
        existing = units_by_code.get(code_norm)
        if existing is not None and existing != units_norm:
            display_a = display_units_by_code.get(code_norm, existing)
            display_b = str(units_raw or "").strip() or units_norm
            raise UnitsConversionError(
                "Конфликт единиц измерения Google для кода "
                f"{display_code_by_code.get(code_norm, code_norm)!r}: "
                f"{display_a!r} и {display_b!r}"
            )
        units_by_code[code_norm] = units_norm
        display_units_by_code[code_norm] = str(units_raw or "").strip() or units_norm
    return GoogleUnitsIndex(
        units_by_code=units_by_code,
        display_units_by_code=display_units_by_code,
        display_code_by_code=display_code_by_code,
        display_names_by_code=display_names_by_code,
    )


@dataclass(frozen=True)
class ConversionRequest:
    """Single immutable conversion request."""

    request_id: str
    contour: str
    code: str
    source_unit: str
    quantity: Any
    tags_count: int
    invariant: ConversionInvariant
    location: str
    item_name: str = ""


@dataclass(frozen=True)
class ConversionAction:
    """Resolved conversion action for one request."""

    request_id: str
    code_raw: str
    code_normalized: str
    original_unit: str
    source_unit: str
    target_unit: str
    original_quantity: Decimal
    coefficient: Decimal
    result_quantity: Decimal
    status: str
    trace: str
    quantity_changed: bool
    units_changed: bool
    machine_residual_adjusted: bool
    fractional_result: bool = False


@dataclass(frozen=True)
class FractionalConversionIssue:
    """One real (non-machine) fractional conversion kept without rounding."""

    contour: str
    location: str
    request_id: str
    code: str
    original_quantity: Decimal
    source_unit: str
    target_unit: str
    coefficient: Decimal
    result_quantity: Decimal
    nearest_int: Decimal
    residual: Decimal
    tags_count: int
    invariant_skipped: bool


@dataclass(frozen=True)
class ConversionDependency:
    """Matrix dependency entry for parts preflight."""

    code: str
    source_unit: str
    target_unit: str
    coefficient: Decimal


@dataclass(frozen=True)
class ConversionPlan:
    """Validated immutable conversion plan."""

    actions: tuple[ConversionAction, ...]
    warnings: tuple[str, ...]
    dependencies: tuple[ConversionDependency, ...]
    fractional_issues: tuple[FractionalConversionIssue, ...] = ()


@dataclass(frozen=True)
class RowStdBinding:
    """Binding between one plan action and mutable ``RowStd`` columns."""

    request_id: str
    row: RowStd
    quantity_column: str
    units_column: str
    status_column: str
    trace_column: str


def action_by_id(plan: ConversionPlan, request_id: str) -> ConversionAction | None:
    """Return the plan action for ``request_id``, if present."""
    for action in plan.actions:
        if action.request_id == request_id:
            return action
    return None


def converted_quantity(action: ConversionAction) -> Decimal:
    """Return converted quantity for ``dataclasses.replace`` helpers."""
    return action.result_quantity


def converted_unit(action: ConversionAction) -> str:
    """Return converted unit text for ``dataclasses.replace`` helpers."""
    return action.target_unit
