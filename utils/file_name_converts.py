"""Project file-name parsing helpers with a project-aware facade.

This module has two layers:

1. Project handlers such as ``AgccFilenamePatterns``.
   They own regex fragments, ``find_*`` / ``scan_*`` search helpers,
   and ``parse_*`` methods for structured parsing of one file name.
2. ``ProjectFileName`` facade.
   It resolves a handler by ``project_name`` and exposes ``find_*`` / ``scan_*`` / ``parse_*``.

The naming convention is intentional:

- ``find_*`` / ``scan_*``: search in arbitrary text
- ``parse_*``: parse a single file name into structured parts
"""

from __future__ import annotations

import re
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import ClassVar, Protocol

_DASH_TRANSLATION = str.maketrans(
    {
        "\u2010": "-",
        "\u2011": "-",
        "\u2012": "-",
        "\u2013": "-",
        "\u2014": "-",
        "\u2015": "-",
        "\u2043": "-",
        "\u2212": "-",
        "\ufe58": "-",
        "\ufe63": "-",
        "\uff0d": "-",
    }
)


def normalize_unicode_dashes(value: str) -> str:
    """Replace Unicode dash variants with the ASCII hyphen.

    AGCC tokens (``01-AN02``, ``####-MARK``, TRM ids) use ASCII ``-``.
    Sheets, Word, and Windows paths often insert U+2010 / U+2013 instead.

    Args:
        value: Source text.

    Returns:
        Text with normalized dashes.
    """

    return value.translate(_DASH_TRANSLATION)


# ---------------------------------------------------------------------------
# Layout (читаемая схема имени; не используется в regex напрямую)
#
#   [CONTRACT]  [SEP_CT]  [TITLE_SYSTEM]  [SEP_TD]  [DISCIPLINE_BODY]  [опц. лист .N]  [хвост ревизии]
#   AGCC.287    -         2879-SOT        .         OD-0001            (редко .7)       _01-AN02_RU.doc
#
# Опционально между телом дисциплины и хвостом: ``.{1..2}`` — номер листа/страницы документа в комплекте
# (тот же символ точки, что ``SEP_TD``, но другая семантика; см. ``OPTIONAL_SHEET_PAGE_SUFFIX_RX``).
#
# SEP_CT  = между контрактом и титулом (дефис).
# SEP_TD  = между титулом и дисциплиной (точка); в regex — экранированная «\.»,
#           в обычных строках — символ "." (см. AgccFilenamePatterns.SEP_TITLE_DISCIPLINE_STR).
#
# Хвост ревизии (после ствола):
#   [SEP_SR] [REVISION_BODY] [SEP_RL] [LANGUAGE] [SEP_LF] [FILE_TYPE]
#   _        01-AN02          _        RU         .        doc
#
# SEP_SR = подчёркивание между стволом и телом ревизии.
# SEP_RL = подчёркивание между ревизией и кодом языка.
# SEP_LF = точка между языком и расширением (doc, dwg, docx, xls, xlsx, …).
# ---------------------------------------------------------------------------


@dataclass(frozen=True)
class FilenameBreakdownRow:
    """One line in :meth:`AgccFilenamePatterns.explain_breakdown` (coarse → fine).

    Attributes:
        label_ru: Human-readable name of the slice.
        value: Extracted text, or ``None`` if missing / no match.
        pick_hint_ru: Short guidance on when to use this value.
        call_ru: How to obtain the value in code (assume ``p = AgccFilenamePatterns.parse_strict(s)``).
    """

    label_ru: str
    value: str | None
    pick_hint_ru: str
    call_ru: str


@dataclass(frozen=True)
class AgccFilenameParts:
    """Structured pieces of a matched AGCC file name."""

    contract: str
    """e.g. ``AGCC.287``."""
    title_system: str
    """e.g. ``2879-SOT`` (title + system code)."""
    discipline_block: str
    """Discipline body only: ``OD-0001`` or ``OD`` — без ведущей точки (точка — :attr:`AgccFilenamePatterns.SEP_TITLE_DISCIPLINE_STR`)."""
    drawing_sheet_page: str | None = None
    """Optional sheet/page index after discipline (digits only), e.g. ``7`` for ``.7`` in the file name."""
    revision_tail: str | None = None
    """e.g. ``_01-AN02_RU.doc``; set by strict :meth:`AgccFilenamePatterns.parse_strict`."""
    simple_ext: str | None = None
    """Bare ``.doc`` / ``.pdf`` after the stem when using :meth:`AgccFilenamePatterns.parse_loose`."""

    @property
    def core_stem(self) -> str:
        """Stem: contract, title, discipline, optional ``.{drawing_sheet_page}``; без хвоста ревизии и без ``simple_ext``."""
        # Точка между титулом и дисциплиной — тот же символ, что в SEP_TITLE_DISCIPLINE_STR у класса ниже.
        sep_td = AgccFilenamePatterns.SEP_TITLE_DISCIPLINE_STR
        stem = f"{self.contract}-{self.title_system}{sep_td}{self.discipline_block}"
        if self.drawing_sheet_page is not None:
            stem = f"{stem}.{self.drawing_sheet_page}"
        return stem

    @property
    def full_name(self) -> str:
        """Strict revision tail wins; else optional ``simple_ext``; else stem only."""
        if self.revision_tail is not None:
            return f"{self.core_stem}{self.revision_tail}"
        if self.simple_ext:
            return f"{self.core_stem}{self.simple_ext}"
        return self.core_stem


@dataclass(frozen=True)
class RevisionTailParts:
    """Pieces of ``_01-AN02_RU.doc``-style tail from :meth:`AgccFilenamePatterns.split_revision_tail`."""

    rev_sheet: str
    an: str | None
    lang: str
    file_type: str
    """Extension without dot: doc, dwg, docx, xls, xlsx, pdf, …"""

    @property
    def ext(self) -> str:
        """Alias of :attr:`file_type` (backward compatible name)."""
        return self.file_type


class FilenamePatternHandler(Protocol):
    """Project-specific filename handler contract used by ``ProjectFileName``."""

    @classmethod
    def pattern_core(cls) -> str:
        """Regex for core stem without revision tail."""

    @classmethod
    def pattern_revision_with_ext(cls) -> str:
        """Regex for revision tail with file extension."""

    @classmethod
    def find_stem_matches(cls, text: str) -> list[str]:
        """Find all stem matches in arbitrary text."""

    @classmethod
    def find_full_matches(cls, text: str) -> list[str]:
        """Find all full file-name matches in arbitrary text."""

    @classmethod
    def scan_title_system(cls, text: str) -> str | None:
        """Find first title-system token in arbitrary text."""

    @classmethod
    def scan_title(cls, text: str) -> str | None:
        """Find first title number in arbitrary text."""

    @classmethod
    def parse_strict(cls, file_name: str) -> object | None:
        """Parse one file name using strict rules."""

    @classmethod
    def parse_loose(cls, file_name: str) -> object | None:
        """Parse one file name using relaxed rules."""


class AgccFilenamePatterns:
    """Regex fragments and parsing helpers for AGCC file names.

    The class is intentionally split into visual sections:

    - atomic regex tokens without outer separators
    - separator constants in regex and string form
    - composite patterns such as ``pattern_core`` / ``pattern_full``
    - search helpers (``find_*`` / ``scan_*``)
    - parsing helpers (``parse_*``)
    - breakdown helpers for human-readable debugging
    """

    # --- Токены regex (без дефиса контракт–титул и без точки титул–дисциплина снаружи) ---

    CONTRACT: ClassVar[str] = r"AGCC\.\d{3,4}"
    # Mark: 2–5 letters, optional 1–2 trailing digits (KSB1, SS30, PD21), optional .N (SKUD.1).
    TITLE_SYSTEM: ClassVar[str] = r"\d{4}-[a-zA-Zа-яА-Я]{2,5}\d{0,2}(?:\.\d{1}){0,1}"
    # Тело дисциплины БЕЗ ведущей точки — точка задаётся отдельно (SEP_TITLE_DISCIPLINE_RX).
    DISCIPLINE_BODY_STRICT: ClassVar[str] = r"[a-zA-Zа-яА-Я0-9]{2,5}-\d{4}"
    DISCIPLINE_BODY_LOOSE: ClassVar[str] = r"[a-zA-Zа-яА-Я0-9]{2,5}(?:-\d{4})?"
    # Хвост ревизии: токены без внешних «рамочных» разделителей (подчёркивания и точка задаются отдельно).
    REVISION_BODY: ClassVar[str] = r"(?:\d{1,2}(?:-AN\d{1,2})?|[VS])"
    """Лист/ревизия: ``01``, ``01-AN02`` или однобуквенная ``V`` / ``S``."""
    LANGUAGE: ClassVar[str] = r"(?:RU|EN|ER)"
    FILE_TYPE: ClassVar[str] = r"[a-zA-Z]{3,5}"
    """Расширение без точки: doc, dwg, docx, pdf, xls, xlsx, …"""

    # --- Разделители хвоста ревизии (regex) ---

    SEP_STEM_REVISION_RX: ClassVar[str] = r"_"
    """Между стволом документа и ``REVISION_BODY``."""
    SEP_REVISION_LANGUAGE_RX: ClassVar[str] = r"_"
    """Между ``REVISION_BODY`` и ``LANGUAGE``."""
    SEP_LANGUAGE_FILE_RX: ClassVar[str] = r"\."
    """Между ``LANGUAGE`` и ``FILE_TYPE``."""

    # --- Разделители хвоста ревизии (литералы для строк и подсказок) ---

    SEP_STEM_REVISION_STR: ClassVar[str] = "_"
    SEP_REVISION_LANGUAGE_STR: ClassVar[str] = "_"
    SEP_LANGUAGE_FILE_STR: ClassVar[str] = "."

    # Составной хвост: _ + ревизия + _ + язык + . + тип файла
    REVISION_WITH_EXT: ClassVar[str] = (
        f"{SEP_STEM_REVISION_RX}{REVISION_BODY}"
        f"{SEP_REVISION_LANGUAGE_RX}{LANGUAGE}"
        f"{SEP_LANGUAGE_FILE_RX}{FILE_TYPE}"
    )

    # Разбор хвоста (именованные группы = те же токены и разделители, что в REVISION_WITH_EXT)
    _REVISION_DECOMP: ClassVar[str] = (
        f"{SEP_STEM_REVISION_RX}(?P<rev_sheet>(?:\\d{{1,2}}(?:-AN(?P<an>\\d{{1,2}}))?|[VS]))"
        f"{SEP_REVISION_LANGUAGE_RX}(?P<lang>RU|EN|ER)"
        f"{SEP_LANGUAGE_FILE_RX}(?P<file_type>[a-zA-Z]{{3,5}})"
    )

    # --- Разделители между токенами в regex (подставляются в f/rf-строки шаблона) ---

    SEP_CONTRACT_TITLE: ClassVar[str] = r"-"
    SEP_TITLE_DISCIPLINE_RX: ClassVar[str] = r"\."
    """Между ``title_system`` и ``discipline_block`` в шаблоне (экранированная точка)."""

    # --- Те же разделители как обычные символы (склейка человекочитаемого ствола) ---

    SEP_CONTRACT_TITLE_STR: ClassVar[str] = "-"
    SEP_TITLE_DISCIPLINE_STR: ClassVar[str] = "."
    """Литерал точки; должен соответствовать смыслу ``SEP_TITLE_DISCIPLINE_RX``."""

    # После тела дисциплины, до хвоста ревизии: лист/страница документа (точка + 1–2 цифры).
    # Та же литеральная точка, что ``SEP_TITLE_DISCIPLINE_STR``, но отдельная семантика в разборе.
    SEP_DISCIPLINE_SHEET_PAGE_RX: ClassVar[str] = r"\."
    SHEET_PAGE_INDEX: ClassVar[str] = r"\d{1,2}"
    OPTIONAL_SHEET_PAGE_SUFFIX_RX: ClassVar[str] = r"(?:\.\d{1,2})?"

    # То же, что ``SEP_LANGUAGE_FILE_*``, для простого расширения после ствола (parse_loose)
    SIMPLE_EXT: ClassVar[str] = f"{SEP_LANGUAGE_FILE_RX}{FILE_TYPE}"

    @classmethod
    def pattern_core(cls) -> str:
        """Return the strict stem regex (feeds module-level ``dict_doc_name``).

        Returns:
            Regex for ``AGCC.287-2879-SOT.OD-0001``-style stems, optionally with ``.7`` (sheet page) before any revision tail.
        """
        return (
            f"{cls.CONTRACT}{cls.SEP_CONTRACT_TITLE}{cls.TITLE_SYSTEM}"
            f"{cls.SEP_TITLE_DISCIPLINE_RX}{cls.DISCIPLINE_BODY_STRICT}"
            f"{cls.OPTIONAL_SHEET_PAGE_SUFFIX_RX}"
        )

    @classmethod
    def pattern_revision_with_ext(cls) -> str:
        """Return the revision-tail regex (feeds module-level ``dict_doc_name``).

        Returns:
            Regex for ``_01-AN02_RU.doc``-style tails.
        """
        return cls.REVISION_WITH_EXT

    @classmethod
    def pattern_full(cls) -> str:
        """Return the strict full-name regex.

        Returns:
            Regex for ``AGCC.287-2879-SOT.OD-0001_01-AN02_RU.doc``-style names; stem may include optional ``.{1..2}`` before ``_…`` (sheet page).
        """
        return cls.pattern_core() + cls.pattern_revision_with_ext()

    # --- Search methods: scan/find in arbitrary text ------------------------

    @classmethod
    def find_stem_matches(cls, text: str) -> list[str]:
        """Find all strict core stem matches in arbitrary text.

        Args:
            text: Source text that may contain one or more file names.

        Returns:
            List of matched stems in left-to-right order.
        """
        return re.findall(cls.pattern_core(), normalize_unicode_dashes(text))

    @classmethod
    def find_full_matches(cls, text: str) -> list[str]:
        """Find all strict full-name matches in arbitrary text.

        Args:
            text: Source text that may contain one or more file names.

        Returns:
            List of matched full file names in left-to-right order.
        """
        return re.findall(cls.pattern_full(), normalize_unicode_dashes(text))

    @classmethod
    def scan_title_system(cls, text: str) -> str | None:
        """Return the first title-system token found in arbitrary text.

        Args:
            text: Source text to scan.

        Returns:
            First ``####-XXX``-style token, or ``None`` when absent.
        """
        match = re.findall(cls.TITLE_SYSTEM, normalize_unicode_dashes(text))
        return match[0].strip() if match else None

    @classmethod
    def scan_title(cls, text: str) -> str | None:
        """Return the numeric title part from the first title-system token.

        Args:
            text: Source text to scan.

        Returns:
            Numeric title, e.g. ``7417``, or ``None`` when absent.
        """
        title_system = cls.scan_title_system(text)
        if title_system:
            return title_system.split("-", 1)[0]
        return None

    @classmethod
    def pattern_named_full(cls) -> str:
        """Один regex с именованными группами; ``discipline_block`` — только тело (``OD-0001``)."""
        return (
            rf"(?P<contract>{cls.CONTRACT}){cls.SEP_CONTRACT_TITLE}"
            rf"(?P<title_system>{cls.TITLE_SYSTEM}){cls.SEP_TITLE_DISCIPLINE_RX}"
            rf"(?P<discipline_block>{cls.DISCIPLINE_BODY_STRICT})"
            rf"(?:\.(?P<drawing_sheet_page>{cls.SHEET_PAGE_INDEX}))?"
            rf"(?P<revision_tail>{cls.REVISION_WITH_EXT})?"
        )

    @classmethod
    def compile_full(cls) -> re.Pattern[str]:
        return re.compile(cls.pattern_named_full())

    @classmethod
    def parse_strict(cls, file_name: str) -> AgccFilenameParts | None:
        """Parse the first strict AGCC file-name match in the string.

        Args:
            file_name: String containing one file name or arbitrary text around it.

        Returns:
            ``AgccFilenameParts`` for the first strict match, or ``None``.
        """
        m = cls.compile_full().search(normalize_unicode_dashes(file_name))
        if not m:
            return None
        rev = m.group("revision_tail")
        return AgccFilenameParts(
            contract=m.group("contract"),
            title_system=m.group("title_system"),
            discipline_block=m.group("discipline_block"),
            drawing_sheet_page=m.group("drawing_sheet_page"),
            revision_tail=rev,
            simple_ext=None,
        )

    @classmethod
    def pattern_named_loose(cls) -> str:
        """Return the relaxed named pattern.

        Returns:
            Regex where discipline may be ``OD`` without ``-####`` and a simple
            extension like ``.doc`` is optional after the stem.
        """
        return (
            rf"(?P<contract>{cls.CONTRACT}){cls.SEP_CONTRACT_TITLE}"
            rf"(?P<title_system>{cls.TITLE_SYSTEM}){cls.SEP_TITLE_DISCIPLINE_RX}"
            rf"(?P<discipline_block>{cls.DISCIPLINE_BODY_LOOSE})"
            rf"(?:\.(?P<drawing_sheet_page>{cls.SHEET_PAGE_INDEX}))?"
            rf"(?P<simple_ext>{cls.SIMPLE_EXT})?"
            rf"(?P<revision_tail>{cls.REVISION_WITH_EXT})?"
        )

    @classmethod
    def compile_loose(cls) -> re.Pattern[str]:
        return re.compile(cls.pattern_named_loose())

    @classmethod
    def parse_loose(cls, file_name: str) -> AgccFilenameParts | None:
        """Parse a relaxed AGCC file-name match.

        Args:
            file_name: String containing one file name or arbitrary text around it.

        Returns:
            ``AgccFilenameParts`` for the first relaxed match, or ``None``.
        """
        m = cls.compile_loose().search(normalize_unicode_dashes(file_name))
        if not m:
            return None
        rev = m.group("revision_tail")
        ext = m.group("simple_ext")
        return AgccFilenameParts(
            contract=m.group("contract"),
            title_system=m.group("title_system"),
            discipline_block=m.group("discipline_block"),
            drawing_sheet_page=m.group("drawing_sheet_page"),
            revision_tail=rev,
            simple_ext=ext,
        )

    @classmethod
    def split_revision_tail(cls, tail: str | None) -> RevisionTailParts | None:
        """Split a revision tail into structured parts.

        Args:
            tail: Tail like ``_01-AN02_RU.doc``.

        Returns:
            ``RevisionTailParts`` when the tail matches the expected structure,
            otherwise ``None``.
        """
        if not tail:
            return None
        m = re.match(cls._REVISION_DECOMP, normalize_unicode_dashes(tail))
        if not m:
            return None
        return RevisionTailParts(
            rev_sheet=m.group("rev_sheet"),
            an=m.group("an"),
            lang=m.group("lang"),
            file_type=m.group("file_type"),
        )

    # --- Вспомогательные разбиения для explain_breakdown (не regex) ---

    @classmethod
    def _split_contract(cls, contract: str) -> tuple[str | None, str | None]:
        if "." not in contract:
            return None, None
        prefix, num = contract.split(".", 1)
        return prefix.strip() or None, num.strip() or None

    @classmethod
    def _split_discipline(cls, discipline_block: str) -> tuple[str | None, str | None]:
        # discipline_block хранится без ведущей точки; lstrip оставлен на случай старых данных.
        inner = discipline_block.lstrip(".")
        if "-" not in inner:
            return inner or None, None
        code, serial = inner.split("-", 1)
        return code.strip() or None, serial.strip() or None

    @classmethod
    def _split_title_system(cls, title_system: str) -> tuple[str | None, str | None]:
        if "-" not in title_system:
            return title_system or None, None
        num, rest = title_system.split("-", 1)
        return num.strip() or None, rest.strip() or None

    @classmethod
    def explain_breakdown(cls, text: str) -> list[FilenameBreakdownRow]:
        """Разбор от крупного к мелкому; подписи и подсказки по использованию."""
        parts = cls.parse_strict(text) or cls.parse_loose(text)
        if parts is None:
            return [
                FilenameBreakdownRow(
                    label_ru="Совпадение с шаблоном AGCC",
                    value=None,
                    pick_hint_ru=(
                        "Строго: тело дисциплины ``OD-####`` после точки; полное ``.doc`` часто в хвосте "
                        "``_##_…_RU.ext``. Смягчённо: ``parse_loose`` (``OD`` и опционально ``.doc``)."
                    ),
                    call_ru=(
                        "`AgccFilenamePatterns.parse_strict(s)` и `AgccFilenamePatterns.parse_loose(s)` → `None`. "
                        "`print(AgccFilenamePatterns.format_breakdown(s))`."
                    ),
                )
            ]

        sep_td = cls.SEP_TITLE_DISCIPLINE_STR
        rows: list[FilenameBreakdownRow] = []
        rows.append(
            FilenameBreakdownRow(
                label_ru="Полное имя (ствол + ревизия + расширение)",
                value=parts.full_name,
                pick_hint_ru="Берите для сравнения с именем файла на диске или в архиве.",
                call_ru=(
                    "`p = AgccFilenamePatterns.parse_strict(s) or AgccFilenamePatterns.parse_loose(s); p.full_name` "
                    "(строгий файл: ещё `ProjectFileName.find_full_matches(s)[0]`)."
                ),
            )
        )
        same_as_full = parts.revision_tail is None and not parts.simple_ext
        rows.append(
            FilenameBreakdownRow(
                label_ru="Ствол документа (без _##_…_RU.ext и без простого .ext)",
                value=parts.core_stem,
                pick_hint_ru=(
                    "Устойчивый ключ чертежа без языка/ревизии; часто достаточно для МТО/справочников."
                    + (
                        " Совпадает с полным именем — нет хвоста ревизии и отдельного ``.ext``."
                        if same_as_full
                        else ""
                    )
                ),
                call_ru=(
                    "`p.core_stem` (контракт + `-` + титул + `"
                    + sep_td
                    + "` + дисциплина + при необходимости `.{p.drawing_sheet_page}`) или "
                    "`ProjectFileName.find_stem_matches(s)[0]`."
                ),
            )
        )
        rows.append(
            FilenameBreakdownRow(
                label_ru="Контракт (префикс проекта + номер)",
                value=parts.contract,
                pick_hint_ru="Например сверка с кодом договора / пакета AGCC.###.",
                call_ru="`p.contract`",
            )
        )

        proj, contract_no = cls._split_contract(parts.contract)
        if proj:
            rows.append(
                FilenameBreakdownRow(
                    label_ru="Литера проекта (префикс до точки)",
                    value=proj,
                    pick_hint_ru="Обычно AGCC; редко нужен отдельно от полного контракта.",
                    call_ru="`p.contract.split('.', 1)[0]`",
                )
            )
        if contract_no:
            rows.append(
                FilenameBreakdownRow(
                    label_ru="Номер контракта (после точки)",
                    value=contract_no,
                    pick_hint_ru="Трёх–четырёхзначный код договора без префикса.",
                    call_ru="`p.contract.split('.', 1)[1]`",
                )
            )

        rows.append(
            FilenameBreakdownRow(
                label_ru="Титул + система (как в штампе, ####-XXX)",
                value=parts.title_system,
                pick_hint_ru="Связка номера титула и кода системы/листа.",
                call_ru="`p.title_system` или `ProjectFileName.scan_title_system(s)` (ищет первый ####-XXX в строке).",
            )
        )
        title_num, system_code = cls._split_title_system(parts.title_system)
        if title_num:
            rows.append(
                FilenameBreakdownRow(
                    label_ru="Номер титула (до первого «-»)",
                    value=title_num,
                    pick_hint_ru="Только числовой титул, без кода системы.",
                    call_ru="`p.title_system.split('-', 1)[0]` или `ProjectFileName.scan_title(s)`",
                )
            )
        if system_code:
            rows.append(
                FilenameBreakdownRow(
                    label_ru="Код системы / литера (после первого «-» в титуле)",
                    value=system_code,
                    pick_hint_ru="SKUD, SOT, OD и т.п. вместе с возможными суффиксами в этой части.",
                    call_ru="`p.title_system.split('-', 1)[1]`",
                )
            )

        rows.append(
            FilenameBreakdownRow(
                label_ru=f"Разделитель титул → дисциплина (литерал «{sep_td}»)",
                value=sep_td,
                pick_hint_ru="В шаблоне: `AgccFilenamePatterns.SEP_TITLE_DISCIPLINE_RX` / `SEP_TITLE_DISCIPLINE_STR`.",
                call_ru="`AgccFilenamePatterns.SEP_TITLE_DISCIPLINE_STR`",
            )
        )
        rows.append(
            FilenameBreakdownRow(
                label_ru="Тело дисциплины (код и порядковый номер, без точки слева)",
                value=parts.discipline_block,
                pick_hint_ru="Например OD-0001 или OD; точка перед этим не входит в поле.",
                call_ru="`p.discipline_block`",
            )
        )
        d_code, d_serial = cls._split_discipline(parts.discipline_block)
        if d_code:
            rows.append(
                FilenameBreakdownRow(
                    label_ru="Код дисциплины (до «-» в теле)",
                    value=d_code,
                    pick_hint_ru="OD, MTO и т.д.",
                    call_ru="`p.discipline_block.split('-', 1)[0]`",
                )
            )
        if d_serial:
            rows.append(
                FilenameBreakdownRow(
                    label_ru="Порядковый номер в дисциплине (после «-»)",
                    value=d_serial,
                    pick_hint_ru="Локальный счётчик внутри кода дисциплины.",
                    call_ru="`p.discipline_block.split('-', 1)[1]`",
                )
            )

        if parts.drawing_sheet_page is not None:
            rows.append(
                FilenameBreakdownRow(
                    label_ru="Номер листа/страницы документа (после дисциплины, до _… ревизии)",
                    value=parts.drawing_sheet_page,
                    pick_hint_ru="Необязательный суффикс ``.{1–2 цифры}``: к какой странице комплекта относится лист.",
                    call_ru="`p.drawing_sheet_page`",
                )
            )

        if parts.revision_tail:
            sep_sr = cls.SEP_STEM_REVISION_STR
            sep_rl = cls.SEP_REVISION_LANGUAGE_STR
            sep_lf = cls.SEP_LANGUAGE_FILE_STR
            rows.append(
                FilenameBreakdownRow(
                    label_ru="Хвост: ревизия + язык + тип файла (как в имени)",
                    value=parts.revision_tail,
                    pick_hint_ru="Целиком; структура: `_` + REVISION + `_` + LANG + `.` + FILE_TYPE.",
                    call_ru="`p.revision_tail`",
                )
            )
            rev_parts = cls.split_revision_tail(parts.revision_tail)
            if rev_parts:
                rows.append(
                    FilenameBreakdownRow(
                        label_ru=f"Разделитель ствол → тело ревизии (литерал «{sep_sr}»)",
                        value=sep_sr,
                        pick_hint_ru="Перед `REVISION_BODY` (`01` / `01-AN02`).",
                        call_ru="`AgccFilenamePatterns.SEP_STEM_REVISION_STR`",
                    )
                )
                rows.append(
                    FilenameBreakdownRow(
                        label_ru="Номер листа / ревизии (первая часть REVISION_BODY)",
                        value=rev_parts.rev_sheet,
                        pick_hint_ru="Цифры до опционального `-AN##`.",
                        call_ru=(
                            "`r = AgccFilenamePatterns.split_revision_tail(p.revision_tail); r.rev_sheet` "
                            "(если `r` не `None`)."
                        ),
                    )
                )
                if rev_parts.an:
                    rows.append(
                        FilenameBreakdownRow(
                            label_ru="Номер после AN (вторая часть REVISION_BODY)",
                            value=rev_parts.an,
                            pick_hint_ru="Только если в хвосте есть `-AN##`.",
                            call_ru="`r.an`",
                        )
                    )
                rows.append(
                    FilenameBreakdownRow(
                        label_ru=f"Разделитель ревизия → язык (литерал «{sep_rl}»)",
                        value=sep_rl,
                        pick_hint_ru="Между `REVISION_BODY` и `LANGUAGE`.",
                        call_ru="`AgccFilenamePatterns.SEP_REVISION_LANGUAGE_STR`",
                    )
                )
                rows.append(
                    FilenameBreakdownRow(
                        label_ru="Язык (LANGUAGE: RU / EN / ER)",
                        value=rev_parts.lang,
                        pick_hint_ru="Языковая версия комплекта.",
                        call_ru="`r.lang`",
                    )
                )
                rows.append(
                    FilenameBreakdownRow(
                        label_ru=f"Разделитель язык → тип файла (литерал «{sep_lf}»)",
                        value=sep_lf,
                        pick_hint_ru="Перед расширением без точки в поле FILE_TYPE.",
                        call_ru="`AgccFilenamePatterns.SEP_LANGUAGE_FILE_STR`",
                    )
                )
                rows.append(
                    FilenameBreakdownRow(
                        label_ru="Тип файла (FILE_TYPE: doc, dwg, docx, xls, xlsx, …)",
                        value=rev_parts.file_type,
                        pick_hint_ru="Расширение без точки; см. `AgccFilenamePatterns.FILE_TYPE`.",
                        call_ru="`r.file_type` (алиас: `r.ext`).",
                    )
                )

        if parts.simple_ext:
            rows.append(
                FilenameBreakdownRow(
                    label_ru="Простое расширение после ствола (parse_loose)",
                    value=parts.simple_ext,
                    pick_hint_ru="Например `.doc`, когда нет хвоста `_##_…_RU.ext`.",
                    call_ru="`p.simple_ext`",
                )
            )

        return rows

    @classmethod
    def format_breakdown(cls, text: str, *, width: int = 72) -> str:
        """Человекочитаемый вывод :meth:`explain_breakdown`."""
        lines: list[str] = []
        sep = "=" * min(width, 72)
        lines.append(sep)
        lines.append("Входная строка (как есть):")
        lines.append(repr(text))
        lines.append(sep)
        lines.append(
            "Соглашение: `p = AgccFilenamePatterns.parse_strict(s) or AgccFilenamePatterns.parse_loose(s)`; "
            "ствол: `p.core_stem` = contract + `-` + title + `"
            + cls.SEP_TITLE_DISCIPLINE_STR
            + "` + discipline_body + опционально `.{p.drawing_sheet_page}`; "
            "хвост: `r = AgccFilenamePatterns.split_revision_tail(p.revision_tail)`."
        )
        lines.append(sep)
        for i, row in enumerate(cls.explain_breakdown(text), start=1):
            lines.append(f"{i:2}. {row.label_ru}")
            lines.append(f"    Значение: {row.value!r}")
            lines.append(f"    Когда брать: {row.pick_hint_ru}")
            lines.append(f"    Как получить: {row.call_ru}")
            lines.append("")
        return "\n".join(lines).rstrip() + "\n"


def parse_agcc_mto_xlsx_revision_for_chain(file_name: str) -> tuple[str, str]:
    """Разбор хвоста ревизии AGCC для режима «MTO цепочка ревизий».

    Использует те же токены, что :class:`AgccFilenamePatterns` (``REVISION_BODY``:
    ``01``, ``01-AN02``, ``V`` / ``S`` и т.д.), через :meth:`AgccFilenamePatterns.parse_strict`
    и при необходимости :meth:`AgccFilenamePatterns.parse_loose`, затем
    :meth:`AgccFilenamePatterns.split_revision_tail`.

    Returns:
        ``(revision, appendix)``: для сортировки и уникальности в цепочке —
        ``revision`` — числовая часть ревизии (``01``) или буква ``V``/``S``;
        ``appendix`` — ``""`` или нормализованный токен вида ``AN01`` (как в имени
        ``…_01-AN01_RU.xlsx``). Вместе отображаются как ``01-AN01``.

    Raises:
        ValueError: если имя не разбирается как AGCC с хвостом ``_…_RU.xlsx`` и т.п.
    """
    name = Path(file_name).name
    parts = AgccFilenamePatterns.parse_strict(name) or AgccFilenamePatterns.parse_loose(
        name
    )
    if not parts or not parts.revision_tail:
        raise ValueError(
            "Не удалось разобрать имя как AGCC с хвостом ревизии "
            "(см. ``AgccFilenamePatterns`` / ``_01_RU.xlsx``, ``_01-AN01_RU.xlsx``).\n"
            f"Файл: {name!r}"
        )
    tail = AgccFilenamePatterns.split_revision_tail(parts.revision_tail)
    if not tail:
        raise ValueError(
            "Хвост ревизии не соответствует шаблону ``_01-AN02_RU.xlsx``.\n"
            f"Файл: {name!r}, revision_tail: {parts.revision_tail!r}"
        )
    appendix = f"AN{tail.an}" if tail.an else ""
    if tail.an is not None:
        rev = tail.rev_sheet.split("-", 1)[0]
    else:
        rev = tail.rev_sheet
    return (rev, appendix)


# Project registry: ProjectFileName maps ``project_name`` to a handler class.
FILENAME_PATTERN_BY_PROJECT: dict[str, type[FilenamePatternHandler]] = {
    "AGCC": AgccFilenamePatterns,
}


def _build_dict_doc_name() -> dict[str, list[str]]:
    """Build ``dict_doc_name`` from the handler registry (stem + revision patterns per project)."""
    out: dict[str, list[str]] = {}
    for project_name, handler in FILENAME_PATTERN_BY_PROJECT.items():
        out[project_name] = [
            handler.pattern_core(),
            handler.pattern_revision_with_ext(),
        ]
    return out


# Project key → [stem pattern, revision-tail pattern] for consumers that need raw regex strings.
dict_doc_name = _build_dict_doc_name()


class ProjectFileName:
    """Project-aware facade over project-specific filename handlers.

    Public entry points: ``find_stem_matches`` / ``find_full_matches``,
    ``scan_title_system`` / ``scan_title``, ``parse_strict`` / ``parse_loose``.
    """

    @staticmethod
    def _handler(project_name="AGCC") -> type[FilenamePatternHandler]:
        """Resolve the registered handler for ``project_name``.

        Args:
            project_name: Project key such as ``"AGCC"``.

        Returns:
            Registered handler class implementing ``FilenamePatternHandler``.

        Raises:
            KeyError: If ``project_name`` is not registered.
        """
        try:
            return FILENAME_PATTERN_BY_PROJECT[project_name]
        except KeyError as exc:
            available = ", ".join(sorted(FILENAME_PATTERN_BY_PROJECT))
            raise KeyError(
                f"Unknown project_name: {project_name!r}. Available projects: {available}"
            ) from exc

    # ------------------------------------------------------------------
    # Public API: find_/scan_ => search in arbitrary text;
    # parse_ => parse one file name into structured parts.
    # ------------------------------------------------------------------

    @staticmethod
    def find_stem_matches(file_name, project_name="AGCC"):
        """Find all stem matches for the selected project.

        Args:
            file_name: Arbitrary text that may contain one or more file names.
            project_name: Registered project key.

        Returns:
            List of matched core stems.
        """
        handler = ProjectFileName._handler(project_name)
        return handler.find_stem_matches(file_name)

    @staticmethod
    def find_full_matches(file_name, project_name="AGCC"):
        """Find all full file-name matches for the selected project.

        Args:
            file_name: Arbitrary text that may contain one or more file names.
            project_name: Registered project key.

        Returns:
            List of matched full file names.
        """
        handler = ProjectFileName._handler(project_name)
        return handler.find_full_matches(file_name)

    @staticmethod
    def scan_title_system(file_name, project_name="AGCC"):
        """Scan the first title-system token for the selected project.

        Args:
            file_name: Arbitrary text that may contain a file name.
            project_name: Registered project key.

        Returns:
            First title-system token, or ``None``.
        """
        handler = ProjectFileName._handler(project_name)
        return handler.scan_title_system(file_name)

    @staticmethod
    def scan_title(file_name, project_name="AGCC"):
        """Scan the numeric title part for the selected project.

        Args:
            file_name: Arbitrary text that may contain a file name.
            project_name: Registered project key.

        Returns:
            Numeric title, or ``None``.
        """
        handler = ProjectFileName._handler(project_name)
        return handler.scan_title(file_name)

    @staticmethod
    def parse_strict(file_name, project_name="AGCC"):
        """Parse one file name using the strict rules of the selected project.

        Args:
            file_name: One file name or arbitrary text around it.
            project_name: Registered project key.

        Returns:
            Project-specific parts object, e.g. ``AgccFilenameParts``, or ``None``.
        """
        handler = ProjectFileName._handler(project_name)
        return handler.parse_strict(file_name)

    @staticmethod
    def parse_loose(file_name, project_name="AGCC"):
        """Parse one file name using the relaxed rules of the selected project.

        Args:
            file_name: One file name or arbitrary text around it.
            project_name: Registered project key.

        Returns:
            Project-specific parts object, e.g. ``AgccFilenameParts``, or ``None``.
        """
        handler = ProjectFileName._handler(project_name)
        return handler.parse_loose(file_name)

    # ------------------------------------------------------------------
    # Legacy compatibility API:
    # old get_/check_ names are intentionally kept in one visual block.
    # During future refactoring callers can be moved to find_/scan_/parse_*.
    # ------------------------------------------------------------------

    @staticmethod
    def check_valid_name(file_name, project_name="AGCC"):
        """Legacy alias of :meth:`find_stem_matches`."""
        return ProjectFileName.find_stem_matches(file_name, project_name)

    @staticmethod
    def get_file_name_with_file_extension(file_name, project_name="AGCC"):
        """Legacy alias of :meth:`find_full_matches`."""
        return ProjectFileName.find_full_matches(file_name, project_name)

    @staticmethod
    def get_file_name_without_file_extension(file_name, project_name="AGCC"):
        """Legacy alias of :meth:`find_stem_matches`."""
        return ProjectFileName.find_stem_matches(file_name, project_name)

    @staticmethod
    def check_valid_file_name(file_name, project_name="AGCC"):
        """Legacy alias of :meth:`find_full_matches`."""
        return ProjectFileName.find_full_matches(file_name, project_name)

    @staticmethod
    def get_title_system(file_name, project_name="AGCC"):
        """Legacy alias of :meth:`scan_title_system`."""
        return ProjectFileName.scan_title_system(file_name, project_name)

    @staticmethod
    def get_title(file_name, project_name="AGCC"):
        """Legacy alias of :meth:`scan_title`."""
        return ProjectFileName.scan_title(file_name, project_name)

def print_parse_parts(title: str, parts: AgccFilenameParts | None) -> None:
        """Print parsed parts object in a readable training/example format."""
        print(title)
        print(f"  raw object: {parts!r}")
        if parts is None:
            print("  result: None")
            print()
            return

        print(f"  contract: {parts.contract!r}")
        print(f"  title_system: {parts.title_system!r}")
        print(f"  discipline_block: {parts.discipline_block!r}")
        print(f"  drawing_sheet_page: {parts.drawing_sheet_page!r}")
        print(f"  revision_tail: {parts.revision_tail!r}")
        print(f"  simple_ext: {parts.simple_ext!r}")
        print(f"  core_stem: {parts.core_stem!r}")
        print(f"  full_name: {parts.full_name!r}")

        revision_parts = AgccFilenamePatterns.split_revision_tail(parts.revision_tail)
        print(f"  split_revision_tail(...): {revision_parts!r}")
        if revision_parts is not None:
            print(f"    rev_sheet: {revision_parts.rev_sheet!r}")
            print(f"    an: {revision_parts.an!r}")
            print(f"    lang: {revision_parts.lang!r}")
            print(f"    file_type: {revision_parts.file_type!r}")
            print(f"    ext (alias): {revision_parts.ext!r}")
        print()

def print_project_file_name_guide(input_text: str, *, project_name: str = "AGCC") -> None:
    """Print all facade methods so the module can serve as a quick manual."""
    print("=" * 72)
    print(f"Input text: {input_text!r}")
    print(f"project_name: {project_name!r}")
    print("-" * 72)
    print("New API (recommended):")
    print(
        f"  ProjectFileName.find_stem_matches(...): "
        f"{ProjectFileName.find_stem_matches(input_text, project_name)!r}"
    )
    print(
        f"  ProjectFileName.find_full_matches(...): "
        f"{ProjectFileName.find_full_matches(input_text, project_name)!r}"
    )
    print(
        f"  ProjectFileName.scan_title_system(...): "
        f"{ProjectFileName.scan_title_system(input_text, project_name)!r}"
    )
    print(
        f"  ProjectFileName.scan_title(...): "
        f"{ProjectFileName.scan_title(input_text, project_name)!r}"
    )
    print_parse_parts(
        "  ProjectFileName.parse_strict(...):",
        ProjectFileName.parse_strict(input_text, project_name),
    )
    print_parse_parts(
        "  ProjectFileName.parse_loose(...):",
        ProjectFileName.parse_loose(input_text, project_name),
    )
    print("Legacy compatibility aliases:")
    print(
        f"  ProjectFileName.check_valid_name(...): "
        f"{ProjectFileName.check_valid_name(input_text, project_name)!r}"
    )
    print(
        f"  ProjectFileName.get_file_name_without_file_extension(...): "
        f"{ProjectFileName.get_file_name_without_file_extension(input_text, project_name)!r}"
    )
    print(
        f"  ProjectFileName.check_valid_file_name(...): "
        f"{ProjectFileName.check_valid_file_name(input_text, project_name)!r}"
    )
    print(
        f"  ProjectFileName.get_file_name_with_file_extension(...): "
        f"{ProjectFileName.get_file_name_with_file_extension(input_text, project_name)!r}"
    )
    print(
        f"  ProjectFileName.get_title_system(...): "
        f"{ProjectFileName.get_title_system(input_text, project_name)!r}"
    )
    print(
        f"  ProjectFileName.get_title(...): "
        f"{ProjectFileName.get_title(input_text, project_name)!r}"
    )
    print()


if __name__ in {"__main__"}:
    # Try to print Unicode helper text on Windows consoles without crashing.
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except (AttributeError, ValueError):
        pass

    # ------------------------------------------------------------------
    # Interactive mini-guide:
    # replace any example below with your own string and run the file.
    #
    # New API:
    #   find_* / scan_*  -> search in arbitrary text
    #   parse_*          -> parse one file name into structured parts
    #
    # Legacy compatibility API:
    #   get_* / check_*  -> aliases that forward to the new API
    # ------------------------------------------------------------------

    OD_STYLE_FILE_NAME = "AGCC.287-2879-SOT.OD-0001_01"
    FILE_SYSTEM_STYLE_NAME = "AGCC.287-2869-SOS.CJ-0007_01-AN01_RU.doc "
    STAMP_FILE_NAME_READ_EXAMPLE = (
        "` \nAGCC.287-2879-SOT.OD-0001_01-AN02_RU.doc "
        "AGCC.287-2879-SOT.OD-0001_01-AN02_RU.doc"
    )
    STRICT_PARSE_EXAMPLE = "AGCC.287-2879-SOT.OD-0001_01-AN02_RU.doc"
    LOOSE_PARSE_EXAMPLE = "AGCC.287-2879-SOT.OD.doc"

    # # 1. OD-style string: enough to scan title, but not enough for strict full-name match.
    # print_project_file_name_guide(OD_STYLE_FILE_NAME)

    # 2. File-system full name: good strict parse example.
    print_project_file_name_guide(FILE_SYSTEM_STYLE_NAME)

    # # 3. Typical stamp OCR/readback text: dirty prefix + duplicated full file name.
    # print_project_file_name_guide(STAMP_FILE_NAME_READ_EXAMPLE)

    # 4. Focused examples: one strict and one loose parse, without extra noise.
    print_parse_parts(
        "Focused strict parse example:",
        ProjectFileName.parse_strict(FILE_SYSTEM_STYLE_NAME),
    )
    # print_parse_parts(
    #     "Focused loose parse example:",
    #     ProjectFileName.parse_loose(LOOSE_PARSE_EXAMPLE),
    # )

    # 5. Human-readable decomposition helper from the handler itself.
    print("AgccFilenamePatterns.format_breakdown(STAMP_FILE_NAME_READ_EXAMPLE):")
    print(AgccFilenamePatterns.format_breakdown(FILE_SYSTEM_STYLE_NAME))