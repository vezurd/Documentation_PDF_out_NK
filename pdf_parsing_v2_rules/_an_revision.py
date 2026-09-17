from __future__ import annotations

import re
from typing import Literal

_STAMP_FULL_RD_AN = re.compile(r"^\s*(\d+)\s*-\s*(AN\d+)\s*$", re.I)
_STAMP_AN_TOKEN = re.compile(r"(AN\d+)", re.I)
_RE_NUMERIC_ONLY = re.compile(r"^\d+$")
_RE_AN_ONLY = re.compile(r"^AN\d+$", re.I)
_RE_AN_SUFFIX = re.compile(r"AN(\d+)", re.I)

_CellKind = Literal["empty", "letter_only", "numeric_only", "an_only", "combined", "other"]


def _first_cell_text(temp_p: dict, att: str) -> str | None:
    raw = temp_p.get(att, [])
    if not raw:
        return None
    s = str(raw[0]).strip()
    return s if s else None


def _classify_kind(text: str, rev_literal: tuple[str, ...]) -> _CellKind:
    if not text:
        return "empty"
    if _STAMP_FULL_RD_AN.match(text):
        return "combined"
    if rev_literal and text in rev_literal:
        return "letter_only"
    if _RE_NUMERIC_ONLY.match(text):
        return "numeric_only"
    if _RE_AN_ONLY.match(text):
        return "an_only"
    return "other"


def _parse_combined_rd_an(text: str) -> tuple[str, str] | None:
    m = _STAMP_FULL_RD_AN.match(text.strip())
    if not m:
        return None
    return m.group(1), m.group(2).upper()


def _format_stamp_parse_error(temp_p: dict, row_atts: list[str]) -> str:
    lines = ["Ревизии некорректны:"]
    for att in row_atts:
        raw = temp_p.get(att, [])
        if not raw:
            val = "пусто"
        else:
            val = str(raw[0]).strip() or "пусто"
        lines.append(f"    {att}: {val}")
    return "\n".join(lines)


def _geometry_violation(
    kinds: list[_CellKind],
) -> bool:
    """True if layout breaks stamp rules (bottom=index 0)."""
    if kinds and kinds[0] == "an_only":
        return True
    if len(kinds) < 3:
        return False
    bottom, mid, top = kinds[0], kinds[1], kinds[2]
    if bottom == "numeric_only" and mid == "an_only" and top == "numeric_only":
        return True
    return False


def _row_numeric_rd_value(temp_p: dict, att: str) -> int | None:
    """RD number from one ``18_1_*`` cell: digits-only or ``NN`` from ``NN-ANxx``."""
    text = _first_cell_text(temp_p, att)
    if not text:
        return None
    combined = _parse_combined_rd_an(text)
    if combined:
        return int(combined[0])
    if _RE_NUMERIC_ONLY.match(text):
        return int(text)
    return None


def c18_2_field_id_for_max_numeric_revision(
    temp_p: dict,
    row_atts: list[str] | None = None,
    c18_2_keys: list[str] | None = None,
) -> str | None:
    """``18_2`` row field id matching the largest RD number in ``18_1_*`` rows.

    Rows are bottom-to-top (``18_1_1`` … ``18_1_3``). If several rows share the
    maximum RD, the **topmost** row (largest index) wins. Returns ``None`` if
    no row contributes a numeric RD fragment.

    Args:
        temp_p: ``dict_attributes``-like map.
        row_atts: Keys for ``18_1_1`` … (default: three stamp rows).
        c18_2_keys: Parallel keys for ``18_2_1`` … (default: three date rows).
    """
    if row_atts is None:
        from pdf_parsing_v2_engine.stamp_fields import c_18_1_1, c_18_1_2, c_18_1_3

        row_atts = [c_18_1_1, c_18_1_2, c_18_1_3]
    if c18_2_keys is None:
        from pdf_parsing_v2_engine.stamp_fields import c_18_2_1, c_18_2_2, c_18_2_3

        c18_2_keys = [c_18_2_1, c_18_2_2, c_18_2_3]
    if len(row_atts) != len(c18_2_keys):
        return None
    pairs: list[tuple[int, int]] = []
    for i, att in enumerate(row_atts):
        v = _row_numeric_rd_value(temp_p, att)
        if v is not None:
            pairs.append((i, v))
    if not pairs:
        return None
    max_v = max(v for _, v in pairs)
    chosen_i = max(i for i, v in pairs if v == max_v)
    return c18_2_keys[chosen_i]


def _row_literal_index(temp_p: dict, att: str, rev_literal: tuple[str, ...]) -> int | None:
    """Index of a single-letter revision in ``rev_literal``, or None."""
    text = _first_cell_text(temp_p, att)
    if not text or text not in rev_literal:
        return None
    return rev_literal.index(text)


def c18_2_field_id_for_max_literal_revision(
    temp_p: dict,
    row_atts: list[str] | None = None,
    c18_2_keys: list[str] | None = None,
    rev_literal: tuple[str, ...] | None = None,
) -> str | None:
    """``18_2`` row field id for the **largest** literal revision in ``18_1_*``.

    Same tie-break as numeric: among rows with the max literal order, the
    **topmost** row (largest index in ``18_1_1`` … ``18_1_3``) wins.

    Args:
        temp_p: ``dict_attributes``-like map.
        row_atts: Keys for ``18_1_1`` … (default: three stamp rows).
        c18_2_keys: Parallel keys for ``18_2_1`` … (default: three date rows).
        rev_literal: Allowed single-letter revisions (e.g. ``A`` … ``G``).
    """
    if not rev_literal:
        return None
    if row_atts is None:
        from pdf_parsing_v2_engine.stamp_fields import c_18_1_1, c_18_1_2, c_18_1_3

        row_atts = [c_18_1_1, c_18_1_2, c_18_1_3]
    if c18_2_keys is None:
        from pdf_parsing_v2_engine.stamp_fields import c_18_2_1, c_18_2_2, c_18_2_3

        c18_2_keys = [c_18_2_1, c_18_2_2, c_18_2_3]
    if len(row_atts) != len(c18_2_keys):
        return None
    pairs: list[tuple[int, int]] = []
    for i, att in enumerate(row_atts):
        v = _row_literal_index(temp_p, att, rev_literal)
        if v is not None:
            pairs.append((i, v))
    if not pairs:
        return None
    max_v = max(v for _, v in pairs)
    chosen_i = max(i for i, v in pairs if v == max_v)
    return c18_2_keys[chosen_i]


def _letter_rungs_bottom_to_top(
    temp_p: dict, row_atts: list, rev_literal: tuple[str, ...]
) -> list[int]:
    """Literal revision indices (``rev_literal`` order), bottom to top."""
    out: list[int] = []
    for att in row_atts:
        text = _first_cell_text(temp_p, att)
        if not text or text not in rev_literal:
            continue
        out.append(rev_literal.index(text))
    return out


def _numeric_rungs_bottom_to_top(temp_p: dict, row_atts: list) -> list[int]:
    """RD revision numbers only, in stamp order (bottom to top): numeric cells
    and the digits part of ``NN-ANxx`` combined cells. Skips AN-only and other.
    """
    out: list[int] = []
    for att in row_atts:
        text = _first_cell_text(temp_p, att)
        if not text:
            continue
        combined = _parse_combined_rd_an(text)
        if combined:
            out.append(int(combined[0]))
            continue
        if _RE_NUMERIC_ONLY.match(text):
            out.append(int(text))
            continue
    return out


def _an_suffix_int(an_token: str) -> int:
    m = _RE_AN_SUFFIX.search(an_token)
    return int(m.group(1)) if m else 0


def _an_rungs_bottom_to_top(temp_p: dict, row_atts: list) -> list[int]:
    """AN revision indices in stamp order: ``ANxx``-only cells and AN part of combined."""
    out: list[int] = []
    for att in row_atts:
        text = _first_cell_text(temp_p, att)
        if not text:
            continue
        combined = _parse_combined_rd_an(text)
        if combined:
            _rd, an_tok = combined
            out.append(_an_suffix_int(an_tok))
            continue
        if _RE_AN_ONLY.match(text):
            out.append(_an_suffix_int(text))
            continue
    return out


def _sequence_not_strictly_increasing(values: list[int]) -> bool:
    """True if not strictly increasing: duplicate or decreasing adjacent step."""
    if len(values) <= 1:
        return False
    return any(values[i] >= values[i + 1] for i in range(len(values) - 1))


def _an_revision_long_from_18_1_rows(
    temp_p: dict,
    row_atts: list,
    *,
    rev_literal: tuple[str, ...] = (),
) -> tuple[str, str | None, str | None, str | None]:
    """Build full revision string and parts from three 18_1_* cells (bottom to top).

    ``row_atts`` must be ``[c_18_1_1, c_18_1_2, c_18_1_3]``: ``18_1_1`` is the
    lowest row on the sheet, ``18_1_3`` the highest. Walking in that order, the
    last purely numeric cell (``0``, ``01``, …) sets the RD part; the last
    ``ANxx``-only cell sets the AN part; a single cell ``digits-ANxx`` sets both
    at that row.

    When ``rev_literal`` is non-empty (e.g. ``A`` … ``G``), cells whose text
    **exactly** matches one of those strings are **literal-only** revisions.
    Literal rungs (bottom to top) must **strictly increase** in profile order
    (same rule as numeric rungs: no duplicates, no decrease). A column cannot
    mix literal cells with numeric / ``ANxx`` / combined ``NN-ANxx`` cells.

    Successful forms: ``NN`` (digits only), ``NN-ANxx`` when an AN component
    exists, or a single literal (e.g. ``C``) when the column is literal-only.
    Unknown cell text does not update the running ``last_*`` values (numeric/AN
    walk).

    ``stamp_parse_error`` is ``None`` on success. Otherwise it is a full message
    starting with ``Ревизии некорректны:`` and listing all three field ids with
    values (``пусто`` if missing). Non-``None`` is returned when:

    - literal cells are mixed with numeric / ``ANxx`` / combined / other
      non-empty cells;
    - literal rungs are not strictly increasing;
    - the bottom cell is ``ANxx`` only (never allowed);
    - the column has purely numeric, then ``ANxx`` only, then purely numeric
      bottom-to-top (invalid layout);
    - there is at least one ``ANxx`` cell but no numeric revision anywhere in
      the three cells (including no ``digits-ANxx`` combined cell);
    - numeric-only cells and the RD part of combined cells, taken bottom to
      top, do not **strictly increase** (forbidden: decrease, e.g. ``01`` then
      ``0``; or duplicate, e.g. ``0`` then ``0``);
    - ``ANxx``-only cells and the AN part of combined cells, bottom to top,
      do not **strictly increase** in the AN index (e.g. ``AN02`` then ``AN01``,
      or two identical ``AN01``).

    Returns:
        ``(rev_full, rev_number, rev_an, stamp_parse_error)``.

        - ``rev_full``: full stamp revision string (``NN``, ``NN-ANxx``, or a
          literal letter).
        - ``rev_number``: RD digits as in the stamp (``01``, ``0``, …) or the
          same literal letter as ``rev_full`` when the column is letter-only;
          ``None`` if empty / error.
        - ``rev_an``: ``ANxx`` when present, else ``None``. Combined rows set
          both ``rev_number`` and ``rev_an``.
        On error, ``rev_full`` is empty and ``rev_number`` / ``rev_an`` are
        ``None``.
    """
    kinds: list[_CellKind] = []
    for att in row_atts:
        t = _first_cell_text(temp_p, att)
        kinds.append(_classify_kind(t, rev_literal) if t is not None else "empty")

    err_msg = _format_stamp_parse_error(temp_p, row_atts)

    has_letter = any(k == "letter_only" for k in kinds)
    has_non_letter_nonempty = any(
        k in ("numeric_only", "combined", "an_only", "other") for k in kinds
    )
    if has_letter and has_non_letter_nonempty:
        return "", None, None, err_msg

    if has_letter:
        letter_rungs = _letter_rungs_bottom_to_top(temp_p, row_atts, rev_literal)
        if _sequence_not_strictly_increasing(letter_rungs):
            return "", None, None, err_msg
        last_letter: str | None = None
        for att in row_atts:
            text = _first_cell_text(temp_p, att)
            if text and text in rev_literal:
                last_letter = text
        if last_letter is not None:
            return last_letter, last_letter, None, None
        return "", None, None, err_msg

    if _geometry_violation(kinds):
        return "", None, None, err_msg

    rungs = _numeric_rungs_bottom_to_top(temp_p, row_atts)
    if _sequence_not_strictly_increasing(rungs):
        return "", None, None, err_msg

    an_rungs = _an_rungs_bottom_to_top(temp_p, row_atts)
    if _sequence_not_strictly_increasing(an_rungs):
        return "", None, None, err_msg

    last_numeric: str | None = None
    last_an: str | None = None

    for att in row_atts:
        text = _first_cell_text(temp_p, att)
        if not text:
            continue
        combined = _parse_combined_rd_an(text)
        if combined:
            last_numeric, last_an = combined[0], combined[1]
            continue
        if _RE_NUMERIC_ONLY.match(text):
            last_numeric = text
            continue
        if _RE_AN_ONLY.match(text):
            last_an = _RE_AN_ONLY.match(text).group(0).upper()
            continue

    if last_an is not None and last_numeric is None:
        return "", None, None, err_msg

    if last_an is not None and last_numeric is not None:
        return f"{last_numeric}-{last_an}", last_numeric, last_an, None

    if last_numeric is not None:
        return last_numeric, last_numeric, None, None

    return "", None, None, None


if __name__ == "__main__":
    import sys
    from pathlib import Path

    _repo_root = Path(__file__).resolve().parent.parent
    if str(_repo_root) not in sys.path:
        sys.path.insert(0, str(_repo_root))

    from pdf_parsing_v2_engine.stamp_fields import c_18_1_1, c_18_1_2, c_18_1_3

    _REV_LIT = ("A", "B", "C", "D", "E", "F", "G")

    # Bottom (18_1_1) -> middle (18_1_2) -> top (18_1_3). None = missing field.
    CELL_18_1_1: str | None = "AN02"
    CELL_18_1_2: str | None = "01"
    CELL_18_1_3: str | None = None

    def _cell_to_raw(value: str | None) -> list[str]:
        if value is None:
            return []
        return [value]

    row_atts = [c_18_1_1, c_18_1_2, c_18_1_3]
    temp_p = {
        c_18_1_1: _cell_to_raw(CELL_18_1_1),
        c_18_1_2: _cell_to_raw(CELL_18_1_2),
        c_18_1_3: _cell_to_raw(CELL_18_1_3),
    }

    rf, rn, ra, err = _an_revision_long_from_18_1_rows(temp_p, row_atts)
    print("Input 18_1_1 (bottom) -> 18_1_2 -> 18_1_3 (top):")
    for k in row_atts:
        print(f"  {k}: {temp_p.get(k, [])!r}")
    print(f"=> rev_full={rf!r}, rev_number={rn!r}, rev_an={ra!r}, err={err!r}")

    _PRESETS: list[tuple[str, dict[str, list[str]]]] = [
        ("0, empty, empty -> 0", {c_18_1_1: ["0"], c_18_1_2: [], c_18_1_3: []}),
        ("0, 01, empty -> 01", {c_18_1_1: ["0"], c_18_1_2: ["01"], c_18_1_3: []}),
        ("0, 01, 02 -> 02", {c_18_1_1: ["0"], c_18_1_2: ["01"], c_18_1_3: ["02"]}),
        (
            "invalid: AN order AN02 then AN01",
            {c_18_1_1: ["0"], c_18_1_2: ["AN02"], c_18_1_3: ["AN01"]},
        ),
        (
            "0, AN01, AN02 -> 0-AN02",
            {c_18_1_1: ["0"], c_18_1_2: ["AN01"], c_18_1_3: ["AN02"]},
        ),
        (
            "invalid: duplicate numeric 0 then 0 then 01",
            {c_18_1_1: ["0"], c_18_1_2: ["0"], c_18_1_3: ["01"]},
        ),
        (
            "invalid: duplicate AN AN01 then AN01",
            {c_18_1_1: ["0"], c_18_1_2: ["AN01"], c_18_1_3: ["AN01"]},
        ),
        (
            "invalid: numeric order 01 then 0 then AN02",
            {c_18_1_1: ["01"], c_18_1_2: ["0"], c_18_1_3: ["AN02"]},
        ),
        (
            "combined bottom 01-AN02",
            {c_18_1_1: ["01-AN02"], c_18_1_2: [], c_18_1_3: []},
        ),
        (
            "invalid: AN-only in bottom",
            {c_18_1_1: ["AN01"], c_18_1_2: ["01"], c_18_1_3: []},
        ),
        (
            "invalid: 01, AN01, 02",
            {c_18_1_1: ["01"], c_18_1_2: ["AN01"], c_18_1_3: ["02"]},
        ),
        (
            "invalid: only AN rows (no numeric)",
            {c_18_1_1: [], c_18_1_2: ["AN01"], c_18_1_3: ["AN02"]},
        ),
    ]

    _PRESETS_LETTER: list[tuple[str, dict[str, list[str]]]] = [
        ("A, B, C -> C", {c_18_1_1: ["A"], c_18_1_2: ["B"], c_18_1_3: ["C"]}),
        ("A, empty, C -> C", {c_18_1_1: ["A"], c_18_1_2: [], c_18_1_3: ["C"]}),
        (
            "invalid: letter order C then A",
            {c_18_1_1: ["C"], c_18_1_2: [], c_18_1_3: ["A"]},
        ),
        (
            "invalid: duplicate B",
            {c_18_1_1: ["A"], c_18_1_2: ["B"], c_18_1_3: ["B"]},
        ),
        (
            "invalid: mix letter A and numeric 01",
            {c_18_1_1: ["A"], c_18_1_2: ["01"], c_18_1_3: []},
        ),
        (
            "invalid: mix letter and AN01",
            {c_18_1_1: ["A"], c_18_1_2: ["AN01"], c_18_1_3: []},
        ),
    ]

    print("\n--- presets ---")
    for title, p in _PRESETS:
        rf, rn, ra, e = _an_revision_long_from_18_1_rows(p, row_atts)
        print(f"{title}\n  => full={rf!r}, num={rn!r}, an={ra!r}, err={'set' if e else None}")

    print("\n--- presets (rev_literal A…G) ---")
    for title, p in _PRESETS_LETTER:
        rf, rn, ra, e = _an_revision_long_from_18_1_rows(p, row_atts, rev_literal=_REV_LIT)
        print(f"{title}\n  => full={rf!r}, num={rn!r}, an={ra!r}, err={'set' if e else None}")
