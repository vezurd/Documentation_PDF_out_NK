"""Tests for PDF document properties (61–65, file_name) and v2 compat."""

from __future__ import annotations

import os
import tempfile
import unittest

import fitz

from pdf_parsing_v2_engine.compat import to_page_stamp_attributes
from pdf_parsing_v2_engine.document_properties import (
    RESERVED_FIELD_IDS,
    build_document_properties_metadata,
)
from pdf_parsing_v2_engine.models import FieldDef, FrameInfo, StampTemplate, V2PageResult
from pdf_parsing_v2_engine.stamp_extractor import _validate_template_reserved_field_ids


class DocumentPropertiesTest(unittest.TestCase):
    def test_metadata_layers_and_bookmarks_are_int_lists(self) -> None:
        path = tempfile.NamedTemporaryFile(suffix=".pdf", delete=False).name
        try:
            doc = fitz.open()
            doc.new_page(width=595, height=842)
            doc.save(path)
            doc.close()
            doc = fitz.open(path)
            page = doc[0]
            meta = build_document_properties_metadata(page, file_basename="x.pdf")
            doc.close()
            self.assertEqual(meta["61_Page_Layers"], [-1])
            self.assertEqual(meta["62_Page_Bookmarks"], [-1])
            self.assertIsInstance(meta["63_Page_Width"][0], int)
            self.assertIsInstance(meta["64_Page_Height"][0], int)
            self.assertEqual(meta["file_name"], ["x.pdf"])
        finally:
            if os.path.isfile(path):
                os.unlink(path)

    def test_compat_psa_int_lists_for_normcontrol_keys(self) -> None:
        v2 = V2PageResult(
            page_num=1,
            doc_type="MTO",
            frame=FrameInfo(
                x0=0.0,
                y0=0.0,
                x1=100.0,
                y1=100.0,
                page_width=200.0,
                page_height=200.0,
                rotation=0,
                border_left_mm=0.0,
                border_bottom_mm=0.0,
                border_top_mm=0.0,
            ),
            template_name="t",
            template_score=1.0,
            fields={},
            metadata={
                "61_Page_Layers": [-1],
                "62_Page_Bookmarks": [1],
                "63_Page_Width": [210],
                "64_Page_Height": [297],
                "65_Page_Real_Format": ["A4"],
                "file_name": ["a.pdf"],
            },
        )
        psa = to_page_stamp_attributes(v2)
        self.assertEqual(psa.dict_attributes["61_Page_Layers"], [-1])
        self.assertEqual(psa.dict_attributes["62_Page_Bookmarks"], [1])
        self.assertEqual(psa.dict_attributes["63_Page_Width"], [210])
        self.assertEqual(psa.dict_attributes["64_Page_Height"], [297])

    def test_reserved_id_without_document_property_raises(self) -> None:
        bad = FieldDef(
            id="61_Page_Layers",
            label="oops",
            bbox_mm=(1.0, 1.0, 2.0, 2.0),
            clean=None,
            validate_regex=None,
            expected="optional",
            padding_mm=None,
            field_type="data",
            expected_text=None,
            is_anchor=False,
            outside_stamp=False,
            origin="frame_bottom_right",
            stretch_to_page=(),
            bound_top=None,
            bound_bottom=None,
            bound_left=None,
            bound_right=None,
            document_property=None,
        )
        tmpl = StampTemplate(
            schema_version=2,
            name="bad",
            doc_types=["MTO"],
            page_selector="all",
            priority=1,
            origin="frame_bottom_right",
            padding_mm=0.5,
            grid_adapt=False,
            grid_tolerance_template_mm=2.0,
            grid_tolerance_detected_mm=0.5,
            snap_max_distance_mm=5.0,
            max_shape_change_ratio=2.5,
            cascade_score_threshold=0.4,
            frame_mode="gost",
            fields=[bad],
            grid_lines=[],
        )
        with self.assertRaises(ValueError):
            _validate_template_reserved_field_ids(tmpl)

    def test_reserved_field_ids_cover_mandatory_keys(self) -> None:
        from pdf_parsing_v2_engine.document_properties.registry import MANDATORY_PROPERTY_KEYS

        for k in MANDATORY_PROPERTY_KEYS:
            self.assertIn(k, RESERVED_FIELD_IDS)


if __name__ == "__main__":
    unittest.main()
