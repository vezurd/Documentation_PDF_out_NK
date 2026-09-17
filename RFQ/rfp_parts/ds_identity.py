"""Parse sequential vs actual DS numbers from RFP file names and UL folders.

Naming contract (after ID prep):

- The number immediately after ``ДС`` is **actual**. The token after ``_``
  is informational **sequential** only.
- UL folder: ``согл УЛ ДС{actual}`` (the DS the packing belongs to).
  A ``ГФ`` / госфин label does not cancel a ``ДСn`` already in the name
  (``согл УЛ ДС1 ГФ 5титулов`` → actual 1). ``kind='gf'`` only when there
  is no DS number.
- Ordinary RFP: ``ДС{n}. …`` (n is both sequential and actual).
- Correction RFP: ``ДС{actual}_{sequential}{letter}. …``
  e.g. ``ДС92_24Б. AGCC…`` — actual 92, sequential 24, revision Б.

``parse_ds_name_from_file_name`` stays the legacy prefix string (``ДС92_24Б``)
for parts aggregation. New consumers use ``parse_rfp_ds_identity`` /
``parse_ul_folder_ds_identity``.
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path
from typing import Literal

Kind = Literal["simple", "compound", "gf", "bare", "unparsed"]
Source = Literal["rfp_file", "ul_folder"]

_CYR_LOOKALIKES = str.maketrans(
    {
        "A": "А",
        "B": "В",
        "E": "Е",
        "K": "К",
        "M": "М",
        "H": "Н",
        "O": "О",
        "P": "Р",
        "C": "С",
        "T": "Т",
        "X": "Х",
    }
)
# Real corrections: «ДС92_24Б. AGCC» / «ДС48_14. » — digit/letter then dot+space or space.
# Reject «ДС4905_1.xlsx» (copy suffix before extension).
_RFP_COMPOUND_RE = re.compile(
    r"ДС\s*(\d+)\s*[_\-]\s*(\d+)([А-ЯA-Zа-яa-z]*)(?=\.\s|\s)",
    re.IGNORECASE,
)
_DS_SIMPLE_RE = re.compile(r"ДС\s*(\d+)", re.IGNORECASE)
_UL_COMPOUND_RE = re.compile(
    r"ДС\s*(\d+)\s*[_\-]\s*(\d+)([А-ЯA-Zа-яa-z]*)",
    re.IGNORECASE,
)
_UL_BARE_NUM_RE = re.compile(r"УЛ\s+(\d+)\s*$", re.IGNORECASE)
_GF_RE = re.compile(r"ГФ|ГОСФИН|ГОС\.ФИН", re.IGNORECASE)


def _norm_letter(raw: str) -> str:
    return (raw or "").upper().translate(_CYR_LOOKALIKES)


def _label(seq: int | None, orig: int | None, letter: str, *, compound: bool) -> str:
    if orig is None and seq is None:
        return ""
    letter = _norm_letter(letter)
    if compound and orig is not None and seq is not None:
        return f"ДС{orig}_{seq}{letter}"
    n = orig if orig is not None else seq
    return f"ДС{n}{letter}"


@dataclass(frozen=True, slots=True)
class DsIdentity:
    """Structured DS identity parsed from one folder or file name.

    Attributes:
        raw: Source text (folder or file name).
        sequential: Informational registry number after ``_`` (24 in
            ``ДС92_24Б``). Same as actual for a simple name.
        actual: DS number immediately after ``ДС`` (92 in ``ДС92_24Б``).
            Same as sequential for a simple name.
        letter: Optional revision letter on a correction (``А`` / ``Б``).
        compound: True when sequential and actual were written separately.
        kind: Parse class.
        source: Whether the text was an RFP file or a UL folder.
    """

    raw: str
    sequential: int | None
    actual: int | None
    letter: str
    compound: bool
    kind: Kind
    source: Source

    @property
    def match_key(self) -> int | None:
        """Join key: actual DS number (token immediately after ``ДС``)."""
        return self.actual

    @property
    def label(self) -> str:
        if self.kind == "gf":
            return "ГФ"
        return _label(
            self.sequential,
            self.actual,
            self.letter,
            compound=self.compound,
        )

    @property
    def match_keys(self) -> list[int]:
        """Keys that may join this item to the other side.

        Join is actual-only. The number after ``_`` is informational.
        """
        if self.kind == "gf":
            return []
        if self.actual is not None:
            return [self.actual]
        if self.sequential is not None:
            return [self.sequential]
        return []


def parse_rfp_ds_identity(file_name: str) -> DsIdentity:
    """Parse a workbook name in ``RFP_Зиновьев``.

    Args:
        file_name: File name or stem (path basename is fine).

    Returns:
        Identity. ``kind='unparsed'`` when no ``ДС{n}`` token is present.
    """
    text = Path(file_name).name
    compound = _RFP_COMPOUND_RE.search(text)
    if compound:
        actual = int(compound.group(1))
        seq = int(compound.group(2))
        letter = _norm_letter(compound.group(3))
        return DsIdentity(
            raw=text,
            sequential=seq,
            actual=actual,
            letter=letter,
            compound=True,
            kind="compound",
            source="rfp_file",
        )
    simple = _DS_SIMPLE_RE.search(text)
    if simple:
        n = int(simple.group(1))
        return DsIdentity(
            raw=text,
            sequential=n,
            actual=n,
            letter="",
            compound=False,
            kind="simple",
            source="rfp_file",
        )
    return DsIdentity(
        raw=text,
        sequential=None,
        actual=None,
        letter="",
        compound=False,
        kind="unparsed",
        source="rfp_file",
    )


def parse_ul_folder_ds_identity(folder_name: str) -> DsIdentity:
    """Parse a first-level TSD packing folder name.

    Args:
        folder_name: Folder basename (e.g. ``согл УЛ ДС24``).

    Returns:
        Identity. ``kind='gf'`` only when the name has госфин and no ``ДСn``.
        If both are present, ``ДСn`` wins (same as a simple folder).
    """
    text = Path(folder_name).name
    compound = _UL_COMPOUND_RE.search(text)
    if compound:
        actual = int(compound.group(1))
        seq = int(compound.group(2))
        letter = _norm_letter(compound.group(3))
        return DsIdentity(
            raw=text,
            sequential=seq,
            actual=actual,
            letter=letter,
            compound=True,
            kind="compound",
            source="ul_folder",
        )
    simple = _DS_SIMPLE_RE.search(text)
    if simple:
        n = int(simple.group(1))
        return DsIdentity(
            raw=text,
            sequential=n,
            actual=n,
            letter="",
            compound=False,
            kind="simple",
            source="ul_folder",
        )
    bare = _UL_BARE_NUM_RE.search(text)
    if bare:
        n = int(bare.group(1))
        return DsIdentity(
            raw=text,
            sequential=n,
            actual=n,
            letter="",
            compound=False,
            kind="bare",
            source="ul_folder",
        )
    if _GF_RE.search(text):
        return DsIdentity(
            raw=text,
            sequential=None,
            actual=None,
            letter="",
            compound=False,
            kind="gf",
            source="ul_folder",
        )
    return DsIdentity(
        raw=text,
        sequential=None,
        actual=None,
        letter="",
        compound=False,
        kind="unparsed",
        source="ul_folder",
    )
