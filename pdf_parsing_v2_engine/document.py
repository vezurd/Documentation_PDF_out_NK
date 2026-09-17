"""V2 document and page models — replacements for v1 ``doc_ATTRIBUTES`` / ``PageStampAttributes``.

These dataclasses are v2-native, pickle-safe, and self-contained (no v1 imports).
``V2Document.from_file_path`` uses ``utils.string_parsing`` / ``utils.file_name_converts``
(shared utils, not v1) for file name parsing.

Usage::

    from pdf_parsing_v2_engine.document import V2Document, V2PageData

    doc = V2Document.from_file_path("/path/to/AGCC.287-7417-SOS.MTO-0001_01_RU.pdf")
    # ... run extraction ...
    page_data = V2PageData.from_v2_result(v2_result, doc)
    doc.pages.append(page_data)
"""

from __future__ import annotations

import os
from collections.abc import Iterator
from dataclasses import dataclass, field
from typing import Any

from pdf_parsing_v2_engine.models import V2PageResult
from pdf_parsing_v2_engine.stamp_fields import (
    c_1_DOC_TITLE,
    c_2_Facility_name,
    c_3_Unit_title_name,
    c_4_Document_name,
    c_5_Documentation_type,
    c_6_1_Sheet_number,
    c_6_2_Quantity_of_sheets,
    c_7_Total_number_of_sheets,
    c_8_1,
    c_8_2,
    c_8_3,
    c_8_4,
    c_8_5,
    c_10_1,
    c_10_2,
    c_10_3,
    c_10_4,
    c_10_5,
    c_17_1,
    c_17_2,
    c_17_3,
    c_17_4,
    c_17_5,
    c_18_1,
    c_18_1_1,
    c_18_1_2,
    c_18_1_3,
    c_18_2,
    c_18_2_1,
    c_18_2_2,
    c_18_2_3,
    c_18_3,
    c_18_3_1,
    c_18_3_2,
    c_18_3_3,
    c_26_Document_Revision,
    c_50_File_Name_Stamp,
    c_51_Page_Format,
    c_61_Page_Layers,
    c_62_Page_Bookmarks,
    c_63_Page_Width,
    c_64_Page_Height,
    c_65_Page_Real_Format,
    c_75_File_Full_Name,
)

_SORT_INDEX_BY_DOC_TYPE: dict[str, str] = {
    "MTO": "5001",
    "BOE": "5002",
    "BOM": "5003",
    "BOQ": "5004",
    "VO": "5005",
}

_ALL_ATTRIBUTE_KEYS: tuple[str, ...] = (
    c_1_DOC_TITLE,
    c_2_Facility_name,
    c_3_Unit_title_name,
    c_4_Document_name,
    c_5_Documentation_type,
    c_6_1_Sheet_number,
    c_6_2_Quantity_of_sheets,
    c_7_Total_number_of_sheets,
    c_8_1,
    c_8_2,
    c_8_3,
    c_8_4,
    c_8_5,
    c_10_1,
    c_10_2,
    c_10_3,
    c_10_4,
    c_10_5,
    c_17_1,
    c_17_2,
    c_17_3,
    c_17_4,
    c_17_5,
    c_18_1,
    c_18_1_1,
    c_18_1_2,
    c_18_1_3,
    c_18_2,
    c_18_2_1,
    c_18_2_2,
    c_18_2_3,
    c_18_3,
    c_18_3_1,
    c_18_3_2,
    c_18_3_3,
    c_26_Document_Revision,
    c_50_File_Name_Stamp,
    c_51_Page_Format,
    c_61_Page_Layers,
    c_62_Page_Bookmarks,
    c_63_Page_Width,
    c_64_Page_Height,
    c_65_Page_Real_Format,
    c_75_File_Full_Name,
)

_META_KEYS = (
    c_61_Page_Layers,
    c_62_Page_Bookmarks,
    c_63_Page_Width,
    c_64_Page_Height,
    c_65_Page_Real_Format,
    c_75_File_Full_Name,
)

_INT_LIST_META_KEYS = frozenset({
    c_61_Page_Layers,
    c_62_Page_Bookmarks,
    c_63_Page_Width,
    c_64_Page_Height,
})


def _make_default_attributes() -> dict[str, list]:
    """Create the standard dict_attributes with all known keys mapped to empty lists."""
    return {k: [] for k in _ALL_ATTRIBUTE_KEYS}


def _psa_list_from_metadata(mk: str, mv: object) -> list[object]:
    """Convert one metadata value to a ``dict_attributes`` list (correct element types).

    Ported from ``compat._psa_list_from_metadata``.
    """
    if mv is None:
        return []
    if isinstance(mv, list):
        if not mv:
            return []
        if mk in _INT_LIST_META_KEYS:
            return [int(x) for x in mv]
        if mk == c_65_Page_Real_Format:
            return [str(mv[0])]
        if mk == c_75_File_Full_Name:
            return [str(mv[0])]
        return [str(x) for x in mv]
    if mk in _INT_LIST_META_KEYS:
        return [int(mv)]
    return [str(mv)]


@dataclass
class V2PageData:
    """Page-level stamp data — v2-native replacement for ``PageStampAttributes``.

    ``dict_attributes`` uses the same format as v1 (keys = ``c_*`` string constants,
    values = ``list``). This ensures compatibility with ``checks.py`` where normcontrol
    accesses ``page.dict_attributes[c_18_1][0]`` etc.
    """

    page_num: int
    page_type: str = ""
    page_marka: str = ""
    page_title: str = ""
    page_revision: str = ""
    dict_attributes: dict[str, list] = field(default_factory=_make_default_attributes)

    @classmethod
    def from_v2_result(cls, v2r: V2PageResult, doc: V2Document) -> V2PageData:
        """Convert a ``V2PageResult`` into ``V2PageData``.

        Absorbs the logic previously in ``compat.to_page_stamp_attributes``:
        stamp fields from ``v2r.fields``, metadata (layers/bookmarks/format/size)
        from ``v2r.metadata``.

        Args:
            v2r: extraction result for one page.
            doc: parent document (provides marka, type, title).
        """
        pd = cls(
            page_num=v2r.page_num,
            page_marka=doc.doc_Marka,
            page_type=doc.doc_Type,
            page_title=doc.doc_Title_4d,
        )

        for attr_key in pd.dict_attributes:
            if attr_key in _META_KEYS:
                continue
            fr = v2r.fields.get(attr_key)
            if fr is None:
                pd.dict_attributes[attr_key] = [""]
                continue
            raw = fr.cleaned_value if fr.cleaned_value is not None else (fr.raw_value or "")
            if isinstance(raw, list):
                raw = " ".join(str(x) for x in raw) if raw else ""
            elif not isinstance(raw, str):
                raw = str(raw) if raw else ""
            text = raw.strip()
            pd.dict_attributes[attr_key] = [text]

        md = v2r.metadata
        for mk in _META_KEYS:
            if mk not in md:
                if mk in (c_61_Page_Layers, c_62_Page_Bookmarks):
                    pd.dict_attributes[mk] = [-1]
                else:
                    pd.dict_attributes[mk] = [""]
                continue
            mv = md[mk]
            pd.dict_attributes[mk] = _psa_list_from_metadata(mk, mv) or [""]

        return pd


@dataclass
class V2FieldDiagnostic:
    """Public field-level extraction diagnostic for downstream consumers."""

    doc_name: str
    page_num: int
    field_id: str
    raw_value: str
    cleaned_value: str
    warning_code: str
    warning_message: str
    warning_tier: str | None
    clean_tier: str | None


@dataclass
class V2Document:
    """Document-level model — v2-native replacement for ``doc_ATTRIBUTES``.

    Pickle-safe (no file handles or mutable shared state).

    Fields mirror v1 ``doc_ATTRIBUTES`` for full compatibility with the normcontrol
    checks and tag extraction pipeline.
    """

    file_full_path: str
    file_name: str
    doc_Type: str
    doc_Marka: str
    doc_Number: str
    doc_Revision: str
    doc_Title_4d: str
    doc_Short_Title: str | None
    doc_Short_File_Name: str
    doc_OD_style_file_name: str
    doc_Number_for_sort: str
    pages: list[V2PageData] = field(default_factory=list)
    _v2_results: list[V2PageResult] = field(default_factory=list, repr=False)

    def iter_field_diagnostics(self) -> Iterator[V2FieldDiagnostic]:
        """Yield field-level extraction diagnostics as a stable public API.

        This method intentionally hides the internal ``_v2_results`` storage from
        downstream consumers (for example ``pdf_parsing_v2_rules``). Diagnostics are
        emitted only for fields with non-empty ``parse_warnings``.
        """
        for v2_result in self._v2_results:
            for field_id, field_result in v2_result.fields.items():
                if not field_result.parse_warnings:
                    continue
                cleaned = field_result.cleaned_value
                if cleaned is None:
                    cleaned = field_result.raw_value or ""
                elif not isinstance(cleaned, str):
                    cleaned = str(cleaned)
                for warning in field_result.parse_warnings:
                    yield V2FieldDiagnostic(
                        doc_name=self.doc_OD_style_file_name,
                        page_num=v2_result.page_num,
                        field_id=field_id,
                        raw_value=field_result.raw_value or "",
                        cleaned_value=cleaned,
                        warning_code=warning.code,
                        warning_message=warning.message,
                        warning_tier=warning.tier,
                        clean_tier=field_result.clean_tier,
                    )

    def get_v2_page_result(self, page_num: int) -> V2PageResult | None:
        """Return the extraction result for *page_num* (1-based), if already stored.

        Populated after a successful extraction pass (see ``V2Document._v2_results``).
        Prefer this over re-running ``extract_page`` when geometry and fields are
        still valid for the same file.

        Args:
            page_num: 1-based PDF page index.

        Returns:
            Matching ``V2PageResult`` or ``None`` if that page was not extracted.
        """
        for v2_result in self._v2_results:
            if int(v2_result.page_num) == int(page_num):
                return v2_result
        return None

    @classmethod
    def from_file_path(cls, file_path: str) -> V2Document:
        """Parse file name and create a document descriptor.

        Uses ``utils.string_parsing`` and ``utils.file_name_converts`` (shared
        project utilities, not v1-specific) to extract document metadata from
        the file name following the project naming convention.

        Args:
            file_path: absolute or relative path to the PDF file.
        """
        from utils import string_parsing
        from utils.file_name_converts import ProjectFileName

        file_name = os.path.basename(file_path)

        doc_type = string_parsing.getDocTypeFromFile(file_name)
        doc_title_4d = string_parsing.getDocTitle(file_name)
        doc_marka = string_parsing.getMarkaFromFileName(file_name)
        doc_number = string_parsing.getDocNumber(file_name)
        doc_revision = string_parsing.getRevisionFromFileName(file_name)
        doc_short_title = ProjectFileName.scan_title_system(file_name)
        doc_short_file_name = f"{doc_title_4d}-{doc_marka}.{doc_type}"
        doc_od_style_file_name = string_parsing.getOdStyleFileName(file_name)

        sort_index = _SORT_INDEX_BY_DOC_TYPE.get(doc_type, doc_number)

        return cls(
            file_full_path=file_path,
            file_name=file_name,
            doc_Type=doc_type,
            doc_Marka=doc_marka,
            doc_Number=doc_number,
            doc_Revision=doc_revision,
            doc_Title_4d=doc_title_4d,
            doc_Short_Title=doc_short_title,
            doc_Short_File_Name=doc_short_file_name,
            doc_OD_style_file_name=doc_od_style_file_name,
            doc_Number_for_sort=sort_index,
        )

    def print_debug(self) -> None:
        """Print document-level fields to stdout (DS/MTO debug when ``info_flag`` is on).

        Uses a simple key/value layout so callers do not depend on v1 PrettyTable dumps.
        """
        rows = (
            ("file_full_path", self.file_full_path),
            ("file_name", self.file_name),
            ("doc_Type", self.doc_Type),
            ("doc_Marka", self.doc_Marka),
            ("doc_Number", self.doc_Number),
            ("doc_Revision", self.doc_Revision),
            ("doc_Title_4d", self.doc_Title_4d),
            ("doc_Short_Title", self.doc_Short_Title),
            ("doc_Short_File_Name", self.doc_Short_File_Name),
            ("doc_OD_style_file_name", self.doc_OD_style_file_name),
            ("doc_Number_for_sort", self.doc_Number_for_sort),
            ("pages", len(self.pages)),
        )
        width = max(len(k) for k, _ in rows)
        for key, val in rows:
            print(f"  {key:<{width}}  {val!s}")
