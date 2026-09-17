"""snap_outside_stamp_vertical_to_frame: внутренняя горизонталь поля вне штампа к линии рамки."""

from __future__ import annotations

from pdf_parsing_v2_engine.coord_transform import (
    absolute_to_field_bbox_mm,
    resolve_field_bbox_core,
    snap_outside_stamp_vertical_to_frame,
)
from pdf_parsing_v2_engine.models import FieldDef, FrameInfo

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


def _fd_outside(
    bbox_mm: tuple[float, float, float, float],
    *,
    origin: str = "frame_bottom_right",
    stretch: tuple[str, ...] = ("bottom",),
) -> FieldDef:
    return FieldDef(
        id="o",
        label="o",
        bbox_mm=bbox_mm,
        clean=None,
        validate_regex=None,
        outside_stamp=True,
        origin=origin,
        stretch_to_page=stretch,
        expected="optional",
    )


def test_snap_bottom_origin_aligns_core_top_to_frame_y0() -> None:
    bb = absolute_to_field_bbox_mm("frame_bottom_right", _FRAME, 480.0, 10.0, 500.0, 50.0)
    fd = _fd_outside(bb)
    core0 = resolve_field_bbox_core(fd, _FRAME)
    assert abs(max(core0[1], core0[3]) - 50.0) < 0.02

    new_bb = snap_outside_stamp_vertical_to_frame(fd, _FRAME)
    assert new_bb is not None
    fd2 = _fd_outside(new_bb)
    core1 = resolve_field_bbox_core(fd2, _FRAME)
    assert abs(max(core1[1], core1[3]) - _FRAME.y0) < 0.02
    assert min(core1[1], core1[3]) < _FRAME.y0


def test_snap_top_origin_aligns_core_bottom_to_frame_y1() -> None:
    bb = absolute_to_field_bbox_mm("frame_top_right", _FRAME, 480.0, 780.0, 500.0, 820.0)
    fd = _fd_outside(bb, origin="frame_top_right", stretch=("top",))
    new_bb = snap_outside_stamp_vertical_to_frame(fd, _FRAME)
    assert new_bb is not None
    fd2 = _fd_outside(new_bb, origin="frame_top_right", stretch=("top",))
    core = resolve_field_bbox_core(fd2, _FRAME)
    assert abs(min(core[1], core[3]) - _FRAME.y1) < 0.02
    assert max(core[1], core[3]) > _FRAME.y1


def test_snap_when_field_core_inside_frame_still_moves_inner_edge() -> None:
    """Previously rejected by guards; inner edge must snap to frame.y0 (footer)."""
    bb = absolute_to_field_bbox_mm("frame_bottom_right", _FRAME, 480.0, 100.0, 500.0, 200.0)
    fd = _fd_outside(bb)
    new_bb = snap_outside_stamp_vertical_to_frame(fd, _FRAME)
    assert new_bb is not None
    fd2 = _fd_outside(new_bb)
    core = resolve_field_bbox_core(fd2, _FRAME)
    assert abs(max(core[1], core[3]) - _FRAME.y0) < 0.02


def test_snap_preserves_height() -> None:
    bb = absolute_to_field_bbox_mm("frame_bottom_right", _FRAME, 480.0, 10.0, 500.0, 50.0)
    fd = _fd_outside(bb)
    h0 = resolve_field_bbox_core(fd, _FRAME)[3] - resolve_field_bbox_core(fd, _FRAME)[1]
    new_bb = snap_outside_stamp_vertical_to_frame(fd, _FRAME)
    assert new_bb is not None
    fd2 = _fd_outside(new_bb)
    core = resolve_field_bbox_core(fd2, _FRAME)
    h1 = core[3] - core[1]
    assert abs(h0 - h1) < 0.02


def test_snap_returns_none_without_outside_stamp() -> None:
    bb = absolute_to_field_bbox_mm("frame_bottom_right", _FRAME, 480.0, 10.0, 500.0, 50.0)
    fd = FieldDef(
        id="i",
        label="i",
        bbox_mm=bb,
        clean=None,
        validate_regex=None,
        outside_stamp=False,
        origin="frame_bottom_right",
        stretch_to_page=("bottom",),
        expected="optional",
    )
    assert snap_outside_stamp_vertical_to_frame(fd, _FRAME) is None


def test_snap_returns_none_when_both_vertical_stretches() -> None:
    bb = absolute_to_field_bbox_mm("frame_bottom_right", _FRAME, 480.0, 10.0, 500.0, 50.0)
    fd = _fd_outside(bb, stretch=("bottom", "top"))
    assert snap_outside_stamp_vertical_to_frame(fd, _FRAME) is None
