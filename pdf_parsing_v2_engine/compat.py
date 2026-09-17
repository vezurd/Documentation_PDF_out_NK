"""Compatibility bridge: V2PageResult → V2PageData.

Historically converted V2PageResult to v1 ``PageStampAttributes``.
Now delegates to ``V2PageData.from_v2_result`` (v2-native).

The ``to_v2_page_data`` function is the primary public API.
``to_page_stamp_attributes`` is kept as a deprecated alias.
"""

from __future__ import annotations

from pdf_parsing_v2_engine.document import V2Document, V2PageData
from pdf_parsing_v2_engine.models import V2PageResult


def to_v2_page_data(v2_result: V2PageResult, doc: V2Document) -> V2PageData:
    """Convert extraction result to ``V2PageData`` (v2-native page model).

    Thin wrapper around ``V2PageData.from_v2_result``.

    Args:
        v2_result: extraction result for one page.
        doc: parent ``V2Document`` (provides marka, type, title).
    """
    return V2PageData.from_v2_result(v2_result, doc)


def to_page_stamp_attributes(v2_result: V2PageResult) -> V2PageData:
    """Deprecated: use ``to_v2_page_data`` instead.

    Kept for backward compatibility. Creates a V2PageData without doc context
    (page_marka/page_title will be empty — caller must set them manually,
    which is what v2_pipeline and parallel already do).
    """
    pd = V2PageData(
        page_num=v2_result.page_num,
        page_type=v2_result.doc_type,
    )
    from pdf_parsing_v2_engine.document import _META_KEYS, _psa_list_from_metadata
    from pdf_parsing_v2_engine.stamp_fields import c_61_Page_Layers, c_62_Page_Bookmarks

    for attr_key in pd.dict_attributes:
        if attr_key in _META_KEYS:
            continue
        fr = v2_result.fields.get(attr_key)
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

    md = v2_result.metadata
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
