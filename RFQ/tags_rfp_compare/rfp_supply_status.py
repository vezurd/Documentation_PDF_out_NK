"""RFP supply / position status from parts (not Step4 ``POSITION_STATUS``).

Known cell values:

* empty
* canonical ``Исключен из поставки`` (skip match / packing)
* ``{old} тег заменен на {new}`` — replace ``TAGS`` with *new* before compare

Any other non-empty text is fatal (``RfpSupplyStatusError``).
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Sequence

from base.tables_columns import CODE, RFP_SUPPLY_STATUS, TAGS

CANONICAL_EXCLUDED_FROM_SUPPLY = "Исключен из поставки"
KIND_EMPTY = "empty"
KIND_EXCLUDED = "excluded"
KIND_TAG_REPLACED = "tag_replaced"
KIND_UNKNOWN = "unknown"

_WS_RE = re.compile(r"\s+")
_CANONICAL_FOLD = CANONICAL_EXCLUDED_FROM_SUPPLY.casefold()
_TAG_REPLACED_RE = re.compile(
    r"^(?P<old>\S.*?)\s+тег\s+замен[её]н\s+на\s+(?P<new>\S.*?)$",
    re.IGNORECASE,
)
_TAG_SPLIT_RE = re.compile(r"[,;]")
_HYPHEN_SPACE_RE = re.compile(r"\s*-\s*")

KNOWN_STATUSES_HINT = (
    "пустая ячейка; "
    f"«{CANONICAL_EXCLUDED_FROM_SUPPLY}»; "
    "«<старый тег> тег заменен на <новый тег>»"
)


@dataclass(frozen=True)
class ParsedRfpSupplyStatus:
    """Normalized classification of one status cell."""

    kind: str
    text: str
    old_tag: str = ""
    new_tag: str = ""


@dataclass(frozen=True)
class RfpSupplyStatusIssue:
    """One illegal or inconsistent status cell."""

    file_name: str
    excel_row: int
    code: str
    raw: str
    detail: str = ""

    def as_warning_message(self) -> str:
        """Parts diagnostics line (``строка N: …``)."""
        suffix = f"; {self.detail}" if self.detail else ""
        code = self.code or "—"
        return (
            f"строка {self.excel_row}: некорректный статус позиции; "
            f"код {code}; прочитано: {self.raw!r}{suffix}"
        )


class RfpSupplyStatusError(Exception):
    """Fatal: status cell is unknown or tag replacement does not match TAGS."""

    def __init__(self, issues: Sequence[RfpSupplyStatusIssue]):
        self.issues = list(issues)
        super().__init__(format_rfp_supply_status_error(self.issues))


def normalize_rfp_supply_status(raw: object) -> str:
    """Strip, collapse inner whitespace, keep original case.

    Args:
        raw: Cell value or already-normalized text.

    Returns:
        Normalized status text, or ``""`` when empty.
    """
    text = str(raw or "").replace("\n", " ").replace("\r", " ")
    text = _WS_RE.sub(" ", text).strip()
    return text


def format_tag_replaced_status(old_tag: str, new_tag: str) -> str:
    """Canonical «тег заменен» cell stored on the net / Step4 row."""
    return f"{old_tag} тег заменен на {new_tag}"


def parse_rfp_supply_status(raw: object) -> ParsedRfpSupplyStatus:
    """Classify a status cell as empty, excluded, tag replacement, or unknown.

    Args:
        raw: Cell value from «MTO, Статус позиции» / «Статус поставки».

    Returns:
        Parsed status. ``KIND_UNKNOWN`` when non-empty text matches nothing.
    """
    text = normalize_rfp_supply_status(raw)
    if not text:
        return ParsedRfpSupplyStatus(kind=KIND_EMPTY, text="")
    if text.casefold() == _CANONICAL_FOLD:
        return ParsedRfpSupplyStatus(
            kind=KIND_EXCLUDED,
            text=CANONICAL_EXCLUDED_FROM_SUPPLY,
        )
    match = _TAG_REPLACED_RE.fullmatch(text)
    if match:
        old_tag = _normalize_tag_token(match.group("old"))
        new_tag = _normalize_tag_token(match.group("new"))
        if old_tag and new_tag:
            return ParsedRfpSupplyStatus(
                kind=KIND_TAG_REPLACED,
                text=format_tag_replaced_status(old_tag, new_tag),
                old_tag=old_tag,
                new_tag=new_tag,
            )
    return ParsedRfpSupplyStatus(kind=KIND_UNKNOWN, text=text)


def is_canonical_excluded_from_supply(raw: object) -> bool:
    """True when *raw* matches the canonical exclusion text (casefold)."""
    return parse_rfp_supply_status(raw).kind == KIND_EXCLUDED


def is_tag_replaced_status(raw: object) -> bool:
    """True when *raw* is a parsed tag-replacement status."""
    return parse_rfp_supply_status(raw).kind == KIND_TAG_REPLACED


def is_excluded_from_supply(row: Any) -> bool:
    """True when a Step4 / parts ``RowStd`` is excluded from MTO/VO/UL allocation.

    Args:
        row: ``RowStd``-like object with ``el`` mapping, or ``None``.

    Returns:
        ``True`` only for the canonical «Исключен из поставки» value.
    """
    if row is None:
        return False
    cells = getattr(row, "el", None)
    if not cells:
        return False
    cell = cells.get(RFP_SUPPLY_STATUS)
    raw = cell.value if cell is not None else None
    return is_canonical_excluded_from_supply(raw)


def apply_tag_replacement(
    tags: object,
    parsed: ParsedRfpSupplyStatus,
) -> tuple[str, str | None]:
    """Replace *old_tag* in ``TAGS`` with *new_tag*.

    Args:
        tags: Current Tag cell (string, list, or empty).
        parsed: Must be ``KIND_TAG_REPLACED``.

    Returns:
        ``(new_tags, None)`` on success, or ``(original, error_detail)``.
    """
    if parsed.kind != KIND_TAG_REPLACED:
        return _tags_as_text(tags), "статус не является заменой тега"
    old_fold = parsed.old_tag.casefold()
    new_fold = parsed.new_tag.casefold()
    tokens = _tag_tokens(tags)
    if not tokens:
        return parsed.new_tag, None
    replaced = False
    out: list[str] = []
    for token in tokens:
        if token.casefold() == old_fold:
            out.append(parsed.new_tag)
            replaced = True
        else:
            out.append(token)
    if replaced:
        return _join_tag_tokens(out), None
    if len(tokens) == 1 and tokens[0].casefold() == new_fold:
        return parsed.new_tag, None
    original = _tags_as_text(tags)
    return original, (
        f"колонка Tag {original!r} не содержит заменяемый тег {parsed.old_tag!r}"
    )


def apply_rfp_supply_status_on_rows(
    rows: Sequence[Any],
    *,
    file_name: str = "",
) -> None:
    """Apply known statuses on loaded RFP rows; raise on unknown / inconsistent.

    Tag replacement mutates ``TAGS`` and stores the canonical replacement text
    on ``RFP_SUPPLY_STATUS``. Exclusion is left as-is. Empty cells are skipped.

    Args:
        rows: ``RowStd`` list from Step1 / net load.
        file_name: Workbook name for the error text (net path basename is fine).

    Raises:
        RfpSupplyStatusError: At least one cell is not a known status, or
            replacement does not match ``TAGS``.
    """
    issues: list[RfpSupplyStatusIssue] = []
    source = file_name or _row_file_name(rows[0] if rows else None)
    for row in rows:
        if getattr(row, "row_type", None) != "position_row":
            continue
        cells = getattr(row, "el", None)
        if not cells:
            continue
        status_cell = cells.get(RFP_SUPPLY_STATUS)
        raw = status_cell.value if status_cell is not None else None
        parsed = parse_rfp_supply_status(raw)
        if parsed.kind in (KIND_EMPTY, KIND_EXCLUDED):
            if parsed.kind == KIND_EXCLUDED and status_cell is not None:
                status_cell.value = CANONICAL_EXCLUDED_FROM_SUPPLY
            continue
        excel_row = _row_excel_row(row)
        code = _cell_text(cells.get(CODE))
        if parsed.kind == KIND_UNKNOWN:
            issues.append(
                RfpSupplyStatusIssue(
                    file_name=source,
                    excel_row=excel_row,
                    code=code,
                    raw=parsed.text,
                )
            )
            continue
        tags_cell = cells.get(TAGS)
        new_tags, error = apply_tag_replacement(
            tags_cell.value if tags_cell is not None else "",
            parsed,
        )
        if error:
            issues.append(
                RfpSupplyStatusIssue(
                    file_name=source,
                    excel_row=excel_row,
                    code=code,
                    raw=parsed.text,
                    detail=error,
                )
            )
            continue
        if tags_cell is not None:
            tags_cell.value = new_tags
        if status_cell is not None:
            status_cell.value = parsed.text
    raise_if_rfp_supply_status_issues(issues)


def raise_if_rfp_supply_status_issues(
    issues: Sequence[RfpSupplyStatusIssue],
) -> None:
    """Raise ``RfpSupplyStatusError`` when *issues* is non-empty.

    Args:
        issues: Collected illegal status cells.

    Raises:
        RfpSupplyStatusError: *issues* is not empty.
    """
    if issues:
        raise RfpSupplyStatusError(issues)


def format_rfp_supply_status_error(issues: Sequence[RfpSupplyStatusIssue]) -> str:
    """Russian fatal message: file, Excel row, code, raw cell, optional detail."""
    lines = [
        "Неизвестный или некорректный статус в колонке «MTO, Статус позиции»:",
    ]
    for item in issues:
        code = item.code or "—"
        extra = f"; {item.detail}" if item.detail else ""
        lines.append(
            f"  файл {item.file_name}, строка {item.excel_row}, "
            f"код {code}, прочитано: {item.raw!r}{extra}"
        )
    lines.append(f"Допустимые значения: {KNOWN_STATUSES_HINT}.")
    return "\n".join(lines)


def _normalize_tag_token(value: str) -> str:
    text = str(value or "").replace("\xa0", " ").strip()
    return _HYPHEN_SPACE_RE.sub("-", text)


def _tags_as_text(raw: object) -> str:
    if raw is None:
        return ""
    if isinstance(raw, list):
        return ", ".join(
            _normalize_tag_token(str(item)) for item in raw if str(item).strip()
        )
    return _normalize_tag_token(str(raw))


def _tag_tokens(raw: object) -> list[str]:
    if isinstance(raw, list):
        return [_normalize_tag_token(str(item)) for item in raw if str(item).strip()]
    text = _tags_as_text(raw)
    if not text:
        return []
    parts = [_normalize_tag_token(item) for item in _TAG_SPLIT_RE.split(text)]
    parts = [item for item in parts if item]
    return parts or [text]


def _join_tag_tokens(tokens: list[str]) -> str:
    if len(tokens) == 1:
        return tokens[0]
    return ", ".join(tokens)


def _cell_text(cell: Any) -> str:
    if cell is None:
        return ""
    value = getattr(cell, "value", cell)
    if value is None:
        return ""
    return str(value).strip()


def _row_excel_row(row: Any) -> int:
    for attr in ("_xlsx_source_row", "_orig_idx"):
        raw = getattr(row, attr, None)
        if raw is None:
            continue
        try:
            return int(raw)
        except (TypeError, ValueError):
            continue
    return 0


def _row_file_name(row: Any) -> str:
    t_com = getattr(row, "t_com", None)
    if t_com is None:
        return ""
    name = str(getattr(t_com, "file_name", "") or "").strip()
    if name:
        return name
    full = str(getattr(t_com, "file_full_path", "") or "").strip()
    return Path(full).name if full else ""
