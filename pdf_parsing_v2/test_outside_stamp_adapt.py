"""Regression: outside_stamp bboxes skip global stamp transform; inner edge snaps to frame."""

from __future__ import annotations

from dataclasses import replace
from unittest.mock import MagicMock, patch

from pdf_parsing_v2_engine.coord_transform import snap_outside_stamp_vertical_to_frame
from pdf_parsing_v2_engine.grid_matcher import (
    CellBbox,
    _template_field_bbox,
    adapt_by_cell_assignment,
)
from pdf_parsing_v2_engine.models import FieldDef, FrameInfo, StampTemplate

_FRAME = FrameInfo(
    x0=72.0,
    y0=72.0,
    x1=520.0,
    y1=770.0,
    page_width=595.0,
    page_height=842.0,
    rotation=0,
    border_left_mm=5.0,
    border_bottom_mm=5.0,
    border_top_mm=5.0,
)


def _template() -> StampTemplate:
    # Two inside fields: placed with vertical slack so a small global shift of
    # detected cells does not distort d_all via clipping (avoids bogus sx/sy).
    inner = FieldDef(
        id="inner_a",
        label="inner_a",
        bbox_mm=(10.0, 80.0, 30.0, 95.0),
        clean=None,
        outside_stamp=False,
        origin="frame_bottom_right",
    )
    inner_b = FieldDef(
        id="inner_b",
        label="inner_b",
        bbox_mm=(85.0, 80.0, 105.0, 95.0),
        clean=None,
        outside_stamp=False,
        origin="frame_bottom_right",
    )
    out_tr = FieldDef(
        id="out_tr",
        label="out_tr",
        bbox_mm=(0.0, 5.0, 40.0, 15.0),
        clean=None,
        outside_stamp=True,
        origin="frame_top_right",
        stretch_to_page=("left",),
    )
    out_br = FieldDef(
        id="out_br",
        label="out_br",
        bbox_mm=(0.0, 5.0, 40.0, 15.0),
        clean=None,
        outside_stamp=True,
        origin="frame_bottom_right",
    )
    return StampTemplate(
        schema_version=2,
        name="t",
        doc_types=["DWG"],
        page_selector="first",
        fields=[inner, inner_b, out_tr, out_br],
    )


def test_outside_stamp_bboxes_skip_global_transform() -> None:
    tpl = _template()
    inner_a = tpl.fields[0]
    inner_b = tpl.fields[1]
    tb_a = _template_field_bbox(inner_a, _FRAME)
    tb_b = _template_field_bbox(inner_b, _FRAME)
    dx, dy = 15.0, -10.0
    det_a = CellBbox(tb_a.x0 + dx, tb_a.y0 + dy, tb_a.x1 + dx, tb_a.y1 + dy)
    det_b = CellBbox(tb_b.x0 + dx, tb_b.y0 + dy, tb_b.x1 + dx, tb_b.y1 + dy)

    def _fake_cells(**_kwargs):
        return [det_a, det_b]

    with patch("pdf_parsing_v2_engine.grid_matcher.get_detected_stamp_cells", _fake_cells):
        results, _info = adapt_by_cell_assignment(tpl, _FRAME, MagicMock(), cfg={})

    by_id = {k: ar for k, ar in results}
    assert by_id["inner_a"].status == "matched"
    assert by_id["inner_b"].status == "matched"

    tdx, tdy = _info.transform_dx, _info.transform_dy
    assert abs(tdx) + abs(tdy) > 1.0

    for fid in ("out_tr", "out_br"):
        fd = next(f for f in tpl.fields if f.id == fid)
        snapped_mm = snap_outside_stamp_vertical_to_frame(fd, _FRAME)
        fd_eff = replace(fd, bbox_mm=snapped_mm) if snapped_mm is not None else fd
        want = _template_field_bbox(fd_eff, _FRAME)
        got = by_id[fid].bbox
        assert by_id[fid].status == "excluded"
        assert abs(got.x0 - want.x0) < 1.0
        assert abs(got.y0 - want.y0) < 1.0
        assert abs(got.x1 - want.x1) < 1.0
        assert abs(got.y1 - want.y1) < 1.0


def test_pipeline_snap_aligns_inside_edge_to_frame() -> None:
    """Excluded outside_stamp rects have inner horizontal edge on frame line after snap."""
    tpl = _template()
    inner_a = tpl.fields[0]
    inner_b = tpl.fields[1]
    tb_a = _template_field_bbox(inner_a, _FRAME)
    tb_b = _template_field_bbox(inner_b, _FRAME)
    det_a = CellBbox(tb_a.x0, tb_a.y0, tb_a.x1, tb_a.y1)
    det_b = CellBbox(tb_b.x0, tb_b.y0, tb_b.x1, tb_b.y1)

    def _fake_cells(**_kwargs):
        return [det_a, det_b]

    with patch("pdf_parsing_v2_engine.grid_matcher.get_detected_stamp_cells", _fake_cells):
        results, _ = adapt_by_cell_assignment(tpl, _FRAME, MagicMock(), cfg={})

    by_id = {k: ar for k, ar in results}
    out_br = by_id["out_br"].bbox
    out_tr = by_id["out_tr"].bbox
    ph = _FRAME.page_height
    # fitz y-down; pdfminer inner edge y=frame.y0 -> fitz y = ph - frame.y0 (top of footer rect)
    assert abs(min(out_br.y0, out_br.y1) - (ph - _FRAME.y0)) < 1.0
    # pdfminer inner edge y=frame.y1 -> fitz y = ph - frame.y1 (bottom of header rect)
    assert abs(max(out_tr.y0, out_tr.y1) - (ph - _FRAME.y1)) < 1.0


def test_expand_fields_by_bindings_skips_outside_stamp() -> None:
    from pdf_parsing_v2_engine.grid_matcher import AdaptResult, expand_fields_by_bindings

    tpl = _template()
    keys = [f.id for f in tpl.fields]
    inner_tb = _template_field_bbox(tpl.fields[0], _FRAME)
    out_tb = _template_field_bbox(tpl.fields[2], _FRAME)
    # Stale bindings on outside field must not move it in expand_fields_by_bindings
    fd0 = tpl.fields[0]
    fd1 = tpl.fields[1]
    fd_out = FieldDef(
        id=tpl.fields[2].id,
        label=tpl.fields[2].label,
        bbox_mm=tpl.fields[2].bbox_mm,
        clean=tpl.fields[2].clean,
        outside_stamp=True,
        origin=tpl.fields[2].origin,
        stretch_to_page=tpl.fields[2].stretch_to_page,
        bound_top="h_fake",
        bound_bottom="h_fake2",
        bound_left="v_fake",
        bound_right="v_fake2",
    )
    fields = [fd0, fd1, fd_out, tpl.fields[3]]
    inner_b_tb = _template_field_bbox(tpl.fields[1], _FRAME)
    out_br_tb = _template_field_bbox(tpl.fields[3], _FRAME)

    results = [
        (keys[0], AdaptResult(
            keys[0], inner_tb, "matched", inner_tb, [inner_tb], 0.0,
        )),
        (keys[1], AdaptResult(
            keys[1], inner_b_tb, "matched", inner_b_tb, [inner_b_tb], 0.0,
        )),
        (keys[2], AdaptResult(
            keys[2], out_tb, "excluded", out_tb, [], 0.0,
        )),
        (keys[3], AdaptResult(
            keys[3], out_br_tb, "excluded", out_br_tb, [], 0.0,
        )),
    ]
    snapped = {"h_fake": 10.0, "h_fake2": 500.0, "v_fake": 20.0, "v_fake2": 400.0}
    expand_fields_by_bindings(results, fields, keys, snapped)
    assert results[2][1].bbox == out_tb
