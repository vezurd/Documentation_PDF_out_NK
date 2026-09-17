"""Low-level string normalization for stamp text (no semantic junk filtering)."""

from __future__ import annotations

import re

# Unicode hyphen used in legacy cleaners (U+2010 HYPHEN).
_UNICODE_HYPHEN = "\u2010"


def remove_dates(input_text: str) -> str:
    """Strip ``DD.MM.YY`` / ``DD.MM.YYYY``-style date tokens (legacy behavior)."""
    return re.sub(r"(\d\d.\d\d.\d{2,4})", "", input_text).strip()


def remove_newlines(input_text: str) -> str:
    """Remove newline characters and strip (legacy ``string_Remove_New_Lines``)."""
    output_text = input_text
    output_text = re.sub(r"\n", "", output_text)
    return output_text.strip()


def remove_newlines_with_space(input_text: str) -> str:
    """Join lines with spaces (legacy ``string_Remove_New_Lines_With_Space``)."""
    if "\n" in input_text:
        parts = []
        for line in input_text.split("\n"):
            v = line.strip()
            parts.append(v)
        return " ".join(parts).strip()
    return input_text.strip()


def normalize_line_text(line: str) -> str:
    """NBSP, degree, hyphen variants — per-line step from legacy junk cleaner."""
    s = line.replace("\xa0", " ")
    s = s.replace("°", " ")
    s = s.replace(_UNICODE_HYPHEN, "-")
    return s


def split_lines(text: str) -> list[str]:
    """Split on ``\\n`` and normalize each line (first half of legacy junk cleaner)."""
    return [normalize_line_text(x) for x in text.split("\n")]


def list_flatter(input_list: list) -> list:
    """Flatten one level of nested lists (legacy ``list_flatter``)."""
    out_list: list = []
    for x in input_list:
        if isinstance(x, list):
            out_list.extend(x)
        else:
            out_list.append(x)
    return out_list

##########################################

def v2_normalize_line_text(line: str) -> str:
    """Normalize line text for v2."""
    s = line.strip("\r\n").strip()
    s = s.replace("\xa0", " ")    # NBSP
    s = s.replace(_UNICODE_HYPHEN, "-")  # U+2010 HYPHEN
    return s.strip()