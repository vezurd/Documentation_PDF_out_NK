"""Registry of PDF document-level properties (legacy v1 attribute names).

Values feed ``V2PageResult.metadata`` and ``PageStampAttributes.dict_attributes``
for ``rules_check`` and related tooling.
"""

from __future__ import annotations

# Legacy keys — aligned with historical v1 meta keys / ``compat._META_KEYS``.
KEY_PAGE_LAYERS = "61_Page_Layers"
KEY_PAGE_ANNOTATIONS = "62_Page_Bookmarks"
KEY_PAGE_WIDTH_MM = "63_Page_Width"
KEY_PAGE_HEIGHT_MM = "64_Page_Height"
KEY_PAGE_REAL_FORMAT = "65_Page_Real_Format"
KEY_FILE_NAME = "file_name"

MANDATORY_PROPERTY_KEYS: tuple[str, ...] = (
    KEY_PAGE_LAYERS,
    KEY_PAGE_ANNOTATIONS,
    KEY_PAGE_WIDTH_MM,
    KEY_PAGE_HEIGHT_MM,
    KEY_PAGE_REAL_FORMAT,
    KEY_FILE_NAME,
)

RESERVED_FIELD_IDS: frozenset[str] = frozenset(MANDATORY_PROPERTY_KEYS)

# Human-readable labels (English — engine package).
PROPERTY_LABELS: dict[str, str] = {
    KEY_PAGE_LAYERS: "Optional content / layers (document)",
    KEY_PAGE_ANNOTATIONS: "Text annotations on page",
    KEY_PAGE_WIDTH_MM: "Page width (mm, displayed)",
    KEY_PAGE_HEIGHT_MM: "Page height (mm, displayed)",
    KEY_PAGE_REAL_FORMAT: "Inferred paper format (A4, A3, …)",
    KEY_FILE_NAME: "PDF file basename",
}

# Points per mm — same as v1 ``elements_coordinates.scale_number``.
SCALE_POINTS_PER_MM: float = 2.83444
