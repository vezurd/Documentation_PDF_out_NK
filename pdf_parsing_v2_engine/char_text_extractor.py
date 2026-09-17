"""
Char-level text extraction from PDF fields.

Instead of ``get_textbox()`` (intersection-based), uses per-character center-point
containment — the same approach PyMuPDF uses internally in ``Table.extract()``.

The key insight: ``get_textbox()`` includes any character whose bbox *intersects*
the query rectangle.  In dense stamp tables, text spans often cross cell boundaries,
causing adjacent-cell text to leak into the target field.  Filtering by the
**center point** of each character assigns each character unambiguously to exactly
one cell, matching the geometric intent.

Coordinate space
----------------
``get_text("rawdict")`` returns character bboxes in **unrotated** fitz coordinates
(same space as ``get_textbox()`` expects for rotation=90/270 pages).
Therefore ``extract_text()`` expects a *query rect in unrotated fitz coordinates*.

In ``stamp_extractor`` this is already the ``query_rect`` computed via
``fitz_displayed_to_unrotated()``, so no extra conversion is needed at the call
site for **hit testing**.

**Reading order (rotation 90°/270°):** after selecting characters, line clustering
and left-to-right order use **displayed** fitz coordinates (via
``fitz_unrotated_to_displayed(..., fitz_page)``) when *fitz_page* is passed and
``fitz_page.rotation`` is 90 or 270.  Otherwise the legacy unrotated ``(y0, x0)``
sort is used (0°/180°).

Typical usage::

    tp = fitz_page.get_textpage(flags=fitz.TEXTFLAGS_TEXT)
    index = PageCharIndex.build(fitz_page, textpage=tp)

    # query_rect is already in unrotated coords; pass page for displayed-order sort
    text = index.extract_text(query_rect, frame=frame, fitz_page=fitz_page)
"""

from __future__ import annotations

import fitz

from pdf_parsing_v2_engine.coord_transform import fitz_unrotated_to_displayed

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

_DEFAULT_Y_TOLERANCE: float = 3.0  # pts — tolerance for grouping chars into lines

# Flags used for TextPage creation — same standard set as get_text() defaults.
# TEXT_MEDIABOX_CLIP avoids artefacts from chars technically outside the page.
_TEXTPAGE_FLAGS: int = fitz.TEXTFLAGS_TEXT | fitz.TEXT_MEDIABOX_CLIP


# ---------------------------------------------------------------------------
# Public class
# ---------------------------------------------------------------------------

class PageCharIndex:
    """Per-page character index for precise cell-level text extraction.

    Characters are stored with their bboxes as delivered by
    ``get_text("rawdict")`` — i.e. in *unrotated* fitz coordinates.

    Build once per page, call ``extract_text()`` for each field rect.
    """

    __slots__ = ("_chars",)

    def __init__(self, chars: list[dict]) -> None:
        self._chars = chars

    # ------------------------------------------------------------------
    # Factory
    # ------------------------------------------------------------------

    @classmethod
    def build(
        cls,
        fitz_page: fitz.Page,
        textpage: fitz.TextPage | None = None,
    ) -> "PageCharIndex":
        """Build the index from *fitz_page*.

        Args:
            fitz_page: The page to extract characters from.
            textpage:  Optional pre-built ``TextPage``.  When supplied, no
                       second extraction is performed; the caller is
                       responsible for the flags used when creating it.
                       When ``None``, a new ``TextPage`` is created with
                       ``_TEXTPAGE_FLAGS``.
        """
        if textpage is None:
            textpage = fitz_page.get_textpage(flags=_TEXTPAGE_FLAGS)

        raw = fitz_page.get_text("rawdict", textpage=textpage)
        chars: list[dict] = []
        for block in raw.get("blocks", []):
            if block.get("type") != 0:  # skip image blocks
                continue
            for line in block.get("lines", []):
                for span in line.get("spans", []):
                    for ch in span.get("chars", []):
                        bb = ch["bbox"]  # (x0, y0, x1, y1) in unrotated fitz coords
                        c = ch.get("c", "")
                        if not c:
                            continue
                        chars.append({
                            "c": c,
                            "x0": float(bb[0]),
                            "y0": float(bb[1]),
                            "x1": float(bb[2]),
                            "y1": float(bb[3]),
                        })
        return cls(chars)

    # ------------------------------------------------------------------
    # Extraction
    # ------------------------------------------------------------------

    def extract_text(
        self,
        rect: "fitz.Rect | tuple[float, float, float, float]",
        mode: str = "center",
        y_tolerance: float = _DEFAULT_Y_TOLERANCE,
        frame: "FrameInfo | None" = None,
        fitz_page: "fitz.Page | None" = None,
    ) -> str:
        """Extract text whose characters fall within *rect*.

        Args:
            rect: Target rectangle in **unrotated** fitz coordinates.
                  In ``stamp_extractor`` this is the ``query_rect``
                  (already converted via ``fitz_displayed_to_unrotated``).
            mode: Inclusion criterion for each character.

                  ``"center"`` *(default)* — character center point must lie
                  inside the rect (half-open: ``x0 ≤ cx < x1, y0 ≤ cy < y1``).
                  This matches PyMuPDF's own ``Table.extract()`` logic and
                  assigns every character to exactly one cell with no overlap.

                  ``"intersect"`` — character bbox must intersect the rect
                  (equivalent to legacy ``get_textbox`` behaviour; more
                  inclusive, may produce cross-cell noise).

            y_tolerance: Tolerance in pts for clustering characters into visual
                         lines.  Default 3 pts is sufficient for typical CAD
                         stamp fonts.

            frame: Passed through for API symmetry; optional metadata.
            fitz_page: If set and ``rotation`` is 90 or 270, sort/cluster reading
                   order in **displayed** fitz space via ``page.rotation_matrix``.
                   Does not change which characters are selected (still unrotated *rect*).

        Returns:
            Extracted text string; visual lines separated by ``\\n``.
            Returns ``""`` when no characters match.
        """
        if isinstance(rect, fitz.Rect):
            x0, y0, x1, y1 = rect.x0, rect.y0, rect.x1, rect.y1
        else:
            x0, y0, x1, y1 = float(rect[0]), float(rect[1]), float(rect[2]), float(rect[3])

        if mode == "center":
            selected = [ch for ch in self._chars if _center_in_rect(ch, x0, y0, x1, y1)]
        elif mode == "intersect":
            selected = [ch for ch in self._chars if _bbox_intersects_rect(ch, x0, y0, x1, y1)]
        else:
            raise ValueError(f"Unknown extraction mode {mode!r}; expected 'center' or 'intersect'")

        if not selected:
            return ""

        return _chars_to_text(selected, y_tolerance=y_tolerance, fitz_page=fitz_page)

    def __len__(self) -> int:
        return len(self._chars)

    def __repr__(self) -> str:
        return f"PageCharIndex({len(self._chars)} chars)"


# ---------------------------------------------------------------------------
# Internal helpers
# ---------------------------------------------------------------------------

def _center_in_rect(
    ch: dict,
    x0: float, y0: float, x1: float, y1: float,
) -> bool:
    """True when the character's center point lies within [x0, x1) × [y0, y1)."""
    cx = (ch["x0"] + ch["x1"]) * 0.5
    cy = (ch["y0"] + ch["y1"]) * 0.5
    return x0 <= cx < x1 and y0 <= cy < y1


def _bbox_intersects_rect(
    ch: dict,
    x0: float, y0: float, x1: float, y1: float,
) -> bool:
    """True when the character's bbox has a non-empty intersection with the rect."""
    return not (ch["x1"] <= x0 or ch["x0"] >= x1 or ch["y1"] <= y0 or ch["y0"] >= y1)


def _char_center_displayed(ch: dict, fitz_page: fitz.Page) -> tuple[float, float]:
    """Character bbox centre in **displayed** fitz coordinates."""
    r = fitz.Rect(ch["x0"], ch["y0"], ch["x1"], ch["y1"])
    rd = fitz_unrotated_to_displayed(r, fitz_page)
    return (rd.x0 + rd.x1) * 0.5, (rd.y0 + rd.y1) * 0.5


def _chars_to_text(
    chars: list[dict],
    y_tolerance: float = _DEFAULT_Y_TOLERANCE,
    fitz_page: "fitz.Page | None" = None,
) -> str:
    """Reconstruct a text string from a list of character dicts.

    For rotation 0/180 (or *fitz_page* omitted / not 90-270), sort by unrotated
    ``(y0, x0)`` and cluster by unrotated y-centre — legacy behaviour.

    For rotation 90/270 with *fitz_page*, sort and cluster by **displayed** centres
    (``page.rotation_matrix``) so on-screen lines read left-to-right.
    """
    if fitz_page is not None and fitz_page.rotation % 360 in (90, 270):
        return _chars_to_text_displayed_order(chars, y_tolerance, fitz_page)
    return _chars_to_text_unrotated_order(chars, y_tolerance)


def _chars_to_text_unrotated_order(chars: list[dict], y_tolerance: float) -> str:
    """Sort by (y0, x0); cluster lines by unrotated y-centre; within line by x0."""
    chars_sorted = sorted(chars, key=lambda c: (c["y0"], c["x0"]))

    lines: list[list[dict]] = []
    current_line: list[dict] = []
    line_cy: float | None = None

    for ch in chars_sorted:
        cy = (ch["y0"] + ch["y1"]) * 0.5
        if line_cy is None or abs(cy - line_cy) <= y_tolerance:
            current_line.append(ch)
            if line_cy is None:
                line_cy = cy
        else:
            lines.append(sorted(current_line, key=lambda c: c["x0"]))
            current_line = [ch]
            line_cy = cy

    if current_line:
        lines.append(sorted(current_line, key=lambda c: c["x0"]))

    return "\n".join("".join(c["c"] for c in line) for line in lines)


def _chars_to_text_displayed_order(
    chars: list[dict], y_tolerance: float, fitz_page: fitz.Page,
) -> str:
    """Sort/cluster by displayed (cy, cx); within line by displayed x (left-to-right)."""
    with_keys: list[tuple[dict, float, float]] = []
    for ch in chars:
        cx_d, cy_d = _char_center_displayed(ch, fitz_page)
        with_keys.append((ch, cx_d, cy_d))

    with_keys.sort(key=lambda t: (t[2], t[1]))

    lines: list[list[tuple[dict, float, float]]] = []
    current: list[tuple[dict, float, float]] = []
    line_cy: float | None = None

    for item in with_keys:
        ch, cx_d, cy_d = item
        if line_cy is None or abs(cy_d - line_cy) <= y_tolerance:
            current.append(item)
            if line_cy is None:
                line_cy = cy_d
        else:
            lines.append(sorted(current, key=lambda t: t[1]))
            current = [item]
            line_cy = cy_d

    if current:
        lines.append(sorted(current, key=lambda t: t[1]))

    return "\n".join("".join(t[0]["c"] for t in line) for line in lines)
