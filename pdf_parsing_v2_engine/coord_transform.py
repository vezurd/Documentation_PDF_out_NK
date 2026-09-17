"""
Перевод bbox шаблона (мм от origin-угла рамки) → pdfminer pts → fitz.Rect.

Координатные системы
--------------------
1. **Шаблон (bbox_mm)**: ``[h1, v1, h2, v2]`` — мм от origin-угла рамки.
   Для ``frame_bottom_right``: h = влево от правого края, v = вверх от низа.
   Для ``frame_top_right``: h = влево от правого края, v = вниз от верха.
   Для ``frame_bottom_left``: h = вправо от левого края, v = вверх от низа.
   Для ``frame_top_left``: h = вправо от левого края, v = вниз от верха.
   Отрицательные значения = за пределами рамки.

2. **pdfminer / pdfplumber (displayed)**: x → вправо, y → вверх, начало — левый нижний
   угол *отображаемой* страницы (pts).  ``FrameInfo.x0..y1`` хранятся здесь.

3. **fitz displayed**: x → вправо, y → вниз, начало — левый верхний угол
   *отображаемой* страницы (pts).  ``fitz_page.rect`` задан в этой системе.

4. **fitz unrotated** (для rotation=90/270 относительно displayed): исходное
   пространство PDF до `/Rotate`. ``get_textbox()`` / rawdict bboxes — здесь.
   Переход displayed↔unrotated: ``fitz_displayed_to_unrotated`` /
   ``fitz_unrotated_to_displayed`` через ``page.derotation_matrix`` /
   ``page.rotation_matrix`` (углы rect × матрица → AABB).
"""

from __future__ import annotations

import fitz

from pdf_parsing_v2_engine.models import FieldDef, FrameInfo, TemplateGridLine

SCALE: float = 2.83444  # pts / mm


# ---------------------------------------------------------------------------
# Origin helpers
# ---------------------------------------------------------------------------

def get_origin_points(frame: FrameInfo) -> dict[str, tuple[float, float]]:
    """4 угла рамки в абсолютных pdfminer-координатах (x вправо, y вверх).

    Returns dict: origin_name → (x_pts, y_pts).
    """
    return {
        "frame_bottom_right": (frame.x1, frame.y0),
        "frame_top_right":    (frame.x1, frame.y1),
        "frame_bottom_left":  (frame.x0, frame.y0),
        "frame_top_left":     (frame.x0, frame.y1),
    }


def _origin_xy(origin: str, frame: FrameInfo) -> tuple[float, float]:
    """Origin corner in pdfminer absolute coords."""
    pts = get_origin_points(frame)
    return pts.get(origin, pts["frame_bottom_right"])


# ---------------------------------------------------------------------------
# bbox_mm → absolute pdfminer coords
# ---------------------------------------------------------------------------

def resolve_field_bbox_core(
    field: FieldDef,
    frame: FrameInfo,
) -> tuple[float, float, float, float]:
    """origin + bbox_mm → pdfminer rect **before** ``stretch_to_page`` clamping.

    Returns ``(x0, y0, x1, y1)`` with x0 < x1, y0 < y1.
    """
    h1, v1, h2, v2 = field.bbox_mm
    origin = field.origin

    if origin == "frame_bottom_right":
        ax0 = frame.x1 - h1 * SCALE
        ay0 = frame.y0 + v1 * SCALE
        ax1 = frame.x1 - h2 * SCALE
        ay1 = frame.y0 + v2 * SCALE
    elif origin == "frame_top_right":
        ax0 = frame.x1 - h1 * SCALE
        ay0 = frame.y1 - v1 * SCALE
        ax1 = frame.x1 - h2 * SCALE
        ay1 = frame.y1 - v2 * SCALE
    elif origin == "frame_bottom_left":
        ax0 = frame.x0 + h1 * SCALE
        ay0 = frame.y0 + v1 * SCALE
        ax1 = frame.x0 + h2 * SCALE
        ay1 = frame.y0 + v2 * SCALE
    elif origin == "frame_top_left":
        ax0 = frame.x0 + h1 * SCALE
        ay0 = frame.y1 - v1 * SCALE
        ax1 = frame.x0 + h2 * SCALE
        ay1 = frame.y1 - v2 * SCALE
    else:
        ax0 = frame.x1 - h1 * SCALE
        ay0 = frame.y0 + v1 * SCALE
        ax1 = frame.x1 - h2 * SCALE
        ay1 = frame.y0 + v2 * SCALE

    if ax0 > ax1:
        ax0, ax1 = ax1, ax0
    if ay0 > ay1:
        ay0, ay1 = ay1, ay0

    return (ax0, ay0, ax1, ay1)


def pdfminer_rect_restore_before_stretch(
    pm_after_stretch: tuple[float, float, float, float],
    stretch_to_page: tuple[str, ...],
    core_pm: tuple[float, float, float, float],
) -> tuple[float, float, float, float]:
    """Replace page-clamped edges with *core* coordinates for bbox_mm inversion.

    ``resolve_field_bbox`` overwrites stretched edges with page bounds; this
    reverses that using the pre-stretch rect from :func:`resolve_field_bbox_core`
    for the **same** field (same bbox_mm + origin as the current template state).

    Args:
        pm_after_stretch: Observed pdfminer rect (e.g. from the editor scene).
        stretch_to_page: Same tuple as ``FieldDef.stretch_to_page``.
        core_pm: ``resolve_field_bbox_core(field, frame)`` for the field whose
            stretch produced ``pm_after_stretch`` (before user drag, use current
            ``field_def``; after drag, core may drift — caller may fall back).
    """
    x0, y0, x1, y1 = pm_after_stretch
    cx0, cy0, cx1, cy1 = core_pm
    for edge in stretch_to_page:
        if edge == "bottom":
            y0 = cy0
        elif edge == "top":
            y1 = cy1
        elif edge == "left":
            x0 = cx0
        elif edge == "right":
            x1 = cx1
    return (x0, y0, x1, y1)


def resolve_field_bbox(
    field: FieldDef,
    frame: FrameInfo,
) -> tuple[float, float, float, float]:
    """Раскрыть параметризацию: origin + stretch_to_page → абсолютные pdfminer коорд.

    Returns ``(x0, y0, x1, y1)`` with x0 < x1, y0 < y1.
    """
    ax0, ay0, ax1, ay1 = resolve_field_bbox_core(field, frame)

    # stretch_to_page: заменить соответствующий край на край страницы
    for edge in field.stretch_to_page:
        if edge == "bottom":
            ay0 = 0.0
        elif edge == "top":
            ay1 = frame.page_height
        elif edge == "left":
            ax0 = 0.0
        elif edge == "right":
            ax1 = frame.page_width

    return (ax0, ay0, ax1, ay1)


def field_bbox_to_absolute(
    field: FieldDef,
    frame: FrameInfo,
) -> tuple[float, float, float, float]:
    """Convert template bbox_mm → absolute pdfminer coords (pts).

    Delegates to :func:`resolve_field_bbox` which handles all origins + stretch.
    """
    return resolve_field_bbox(field, frame)


# ---------------------------------------------------------------------------
# Reverse: absolute pdfminer coords → bbox_mm for a given origin
# ---------------------------------------------------------------------------

def absolute_to_field_bbox_mm(
    origin: str,
    frame: FrameInfo,
    pm_x0: float,
    pm_y0: float,
    pm_x1: float,
    pm_y1: float,
) -> tuple[float, float, float, float]:
    """Reverse of resolve_field_bbox (without stretch): absolute pdfminer → bbox_mm.

    Returns (h1, v1, h2, v2) in mm, matching the convention of *origin*.
    """
    if origin == "frame_bottom_right":
        h1 = (frame.x1 - pm_x0) / SCALE
        v1 = (pm_y0 - frame.y0) / SCALE
        h2 = (frame.x1 - pm_x1) / SCALE
        v2 = (pm_y1 - frame.y0) / SCALE
    elif origin == "frame_top_right":
        h1 = (frame.x1 - pm_x0) / SCALE
        v1 = (frame.y1 - pm_y1) / SCALE
        h2 = (frame.x1 - pm_x1) / SCALE
        v2 = (frame.y1 - pm_y0) / SCALE
    elif origin == "frame_bottom_left":
        h1 = (pm_x0 - frame.x0) / SCALE
        v1 = (pm_y0 - frame.y0) / SCALE
        h2 = (pm_x1 - frame.x0) / SCALE
        v2 = (pm_y1 - frame.y0) / SCALE
    elif origin == "frame_top_left":
        h1 = (pm_x0 - frame.x0) / SCALE
        v1 = (frame.y1 - pm_y1) / SCALE
        h2 = (pm_x1 - frame.x0) / SCALE
        v2 = (frame.y1 - pm_y0) / SCALE
    else:
        h1 = (frame.x1 - pm_x0) / SCALE
        v1 = (pm_y0 - frame.y0) / SCALE
        h2 = (frame.x1 - pm_x1) / SCALE
        v2 = (pm_y1 - frame.y0) / SCALE
    return (h1, v1, h2, v2)


def snap_outside_stamp_vertical_to_frame(
    field: FieldDef,
    frame: FrameInfo,
    *,
    eps_pts: float = 0.05,
) -> tuple[float, float, float, float] | None:
    """Adjust *bbox_mm* so the inner horizontal edge meets the frame line.

    For ``outside_stamp`` fields: **stretch** clamps the outer edge to the page;
    **origin** picks which frame corner defines ``bbox_mm``. Vertically, the inner
    edge (footer: ``y1`` → ``frame.y0``; header: ``y0`` → ``frame.y1``) is snapped to
    the frame; horizontal extent and width are unchanged (height preserved).

    Returns ``None`` if not ``outside_stamp``, both vertical stretches are set,
    origin is not a frame corner variant, or the inner edge is already on the
    frame line (within *eps_pts*).

    Args:
        field: Field definition; uses ``outside_stamp``, ``stretch_to_page``,
            ``origin``, ``bbox_mm``.
        frame: Detected drawing frame (pdfminer y-up).
        eps_pts: Tolerance for "already snapped" short-circuit.

    Returns:
        New ``(h1, v1, h2, v2)`` in mm, or ``None`` if no change is applied.
    """
    if not getattr(field, "outside_stamp", False):
        return None
    st = field.stretch_to_page
    if "bottom" in st and "top" in st:
        return None

    core = resolve_field_bbox_core(field, frame)
    x0, y0, x1, y1 = core
    o = field.origin
    h = y1 - y0

    if o in ("frame_bottom_right", "frame_bottom_left"):
        target = frame.y0
        if abs(y1 - target) < eps_pts:
            return None
        new_core = (x0, target - h, x1, target)
    elif o in ("frame_top_right", "frame_top_left"):
        target = frame.y1
        if abs(y0 - target) < eps_pts:
            return None
        new_core = (x0, target, x1, target + h)
    else:
        return None

    if new_core[1] >= new_core[3] - eps_pts:
        return None
    return absolute_to_field_bbox_mm(field.origin, frame, *new_core)


# ---------------------------------------------------------------------------
# pdfminer ↔ fitz conversions
# ---------------------------------------------------------------------------

def pdfminer_to_fitz(
    rect_pm: tuple[float, float, float, float],
    page_height: float,
    rotation: int = 0,
) -> fitz.Rect:
    """Convert a pdfminer-displayed rect (y-up) → fitz **displayed** Rect (y-down)."""
    pm_x0, pm_y0, pm_x1, pm_y1 = rect_pm
    return fitz.Rect(pm_x0, page_height - pm_y1, pm_x1, page_height - pm_y0)


def fitz_rect_transform_by_matrix(rect: fitz.Rect, matrix: fitz.Matrix) -> fitz.Rect:
    """Apply *matrix* to the four corners of *rect*; return axis-aligned bbox (fitz.Rect).

    Used for rotation 90°/270° where an axis-aligned rect in one space maps to a
    rotated box whose tight AABB must be taken in the target space.
    """
    corners = (
        (rect.x0, rect.y0),
        (rect.x1, rect.y0),
        (rect.x0, rect.y1),
        (rect.x1, rect.y1),
    )
    xs: list[float] = []
    ys: list[float] = []
    for x, y in corners:
        pt = fitz.Point(x, y) * matrix
        xs.append(pt.x)
        ys.append(pt.y)
    return fitz.Rect(min(xs), min(ys), max(xs), max(ys))


def unrotated_fitz_rect_to_pdfminer_bbox(
    rect: fitz.Rect,
    fitz_page: fitz.Page,
) -> tuple[float, float, float, float]:
    """Map an **unrotated** fitz rect (e.g. from ``get_drawings()``) → pdfminer displayed bbox.

    For rotation 0°/180° *rect* is already in displayed fitz space (y-down); for 90°/270° it is
    first mapped with ``fitz_page.rotation_matrix`` to displayed fitz, then y-flipped to pdfminer.
    Returns ``(pm_x0, pm_y0, pm_x1, pm_y1)`` with y0 < y1 in pdfminer (y-up).
    """
    ph = fitz_page.rect.height
    rot = fitz_page.rotation % 360
    if rot in (90, 270):
        rd = fitz_rect_transform_by_matrix(rect, fitz_page.rotation_matrix)
    else:
        rd = rect
    return (rd.x0, ph - rd.y1, rd.x1, ph - rd.y0)


def fitz_displayed_to_unrotated(rect: fitz.Rect, fitz_page: fitz.Page) -> fitz.Rect:
    """Convert a *displayed* fitz rect → *unrotated* fitz rect (rotation=90/270).

    Uses ``fitz_page.derotation_matrix`` (same convention as PyMuPDF for ``get_textbox``).
    For rotation 0°/180° returns *rect* unchanged.
    """
    if fitz_page.rotation % 360 not in (90, 270):
        return rect
    return fitz_rect_transform_by_matrix(rect, fitz_page.derotation_matrix)


def fitz_unrotated_to_displayed(rect: fitz.Rect, fitz_page: fitz.Page) -> fitz.Rect:
    """Convert an *unrotated* fitz rect → *displayed* fitz rect (rotation=90/270).

    Uses ``fitz_page.rotation_matrix``. Inverse of :func:`fitz_displayed_to_unrotated`.
    For rotation 0°/180° returns *rect* unchanged.
    """
    if fitz_page.rotation % 360 not in (90, 270):
        return rect
    return fitz_rect_transform_by_matrix(rect, fitz_page.rotation_matrix)


# ---------------------------------------------------------------------------
# Grid line coordinate conversion
# ---------------------------------------------------------------------------

def grid_line_to_fitz_pts(
    line: TemplateGridLine,
    frame: FrameInfo,
    origin: str = "frame_bottom_right",
) -> tuple[str, float, float, float, str | None, str]:
    """Convert a TemplateGridLine to fitz displayed-coordinate values.

    Returns ``(orientation, pos_pts, start_pts, end_pts, boundary, id)``.

    For a H-line: ``pos_pts`` is the fitz y-coordinate; ``start_pts``/``end_pts``
    are x-coordinates (start_pts <= end_pts after normalisation).
    For a V-line: ``pos_pts`` is the fitz x-coordinate; ``start_pts``/``end_pts``
    are y-coordinates (start_pts <= end_pts).

    Uses the same axis conventions as :func:`resolve_field_bbox`.
    """
    ph = frame.page_height

    if origin == "frame_bottom_right":
        if line.orientation == "h":
            pos_pm_y = frame.y0 + line.pos_mm * SCALE
            pos_pts = ph - pos_pm_y
            sp = frame.x1 - line.start_mm * SCALE
            ep = frame.x1 - line.end_mm * SCALE
            start_pts, end_pts = min(sp, ep), max(sp, ep)
        else:  # "v"
            pos_pts = frame.x1 - line.pos_mm * SCALE
            sp_y = ph - (frame.y0 + line.start_mm * SCALE)
            ep_y = ph - (frame.y0 + line.end_mm * SCALE)
            start_pts, end_pts = min(sp_y, ep_y), max(sp_y, ep_y)

    elif origin == "frame_top_right":
        if line.orientation == "h":
            pos_pm_y = frame.y1 - line.pos_mm * SCALE
            pos_pts = ph - pos_pm_y
            sp = frame.x1 - line.start_mm * SCALE
            ep = frame.x1 - line.end_mm * SCALE
            start_pts, end_pts = min(sp, ep), max(sp, ep)
        else:
            pos_pts = frame.x1 - line.pos_mm * SCALE
            sp_y = ph - (frame.y1 - line.start_mm * SCALE)
            ep_y = ph - (frame.y1 - line.end_mm * SCALE)
            start_pts, end_pts = min(sp_y, ep_y), max(sp_y, ep_y)

    elif origin == "frame_bottom_left":
        if line.orientation == "h":
            pos_pm_y = frame.y0 + line.pos_mm * SCALE
            pos_pts = ph - pos_pm_y
            sp = frame.x0 + line.start_mm * SCALE
            ep = frame.x0 + line.end_mm * SCALE
            start_pts, end_pts = min(sp, ep), max(sp, ep)
        else:
            pos_pts = frame.x0 + line.pos_mm * SCALE
            sp_y = ph - (frame.y0 + line.start_mm * SCALE)
            ep_y = ph - (frame.y0 + line.end_mm * SCALE)
            start_pts, end_pts = min(sp_y, ep_y), max(sp_y, ep_y)

    else:  # "frame_top_left"
        if line.orientation == "h":
            pos_pm_y = frame.y1 - line.pos_mm * SCALE
            pos_pts = ph - pos_pm_y
            sp = frame.x0 + line.start_mm * SCALE
            ep = frame.x0 + line.end_mm * SCALE
            start_pts, end_pts = min(sp, ep), max(sp, ep)
        else:
            pos_pts = frame.x0 + line.pos_mm * SCALE
            sp_y = ph - (frame.y1 - line.start_mm * SCALE)
            ep_y = ph - (frame.y1 - line.end_mm * SCALE)
            start_pts, end_pts = min(sp_y, ep_y), max(sp_y, ep_y)

    return (line.orientation, pos_pts, start_pts, end_pts, line.boundary, line.id)


def fitz_pts_to_grid_line_mm(
    orientation: str,
    pos_fitz: float,
    start_fitz: float,
    end_fitz: float,
    frame: FrameInfo,
    origin: str = "frame_bottom_right",
) -> tuple[float, float, float]:
    """Reverse of grid_line_to_fitz_pts: fitz displayed coords → (pos_mm, start_mm, end_mm).

    Only ``frame_bottom_right`` is currently implemented (the default template origin).
    """
    ph = frame.page_height

    if origin == "frame_bottom_right":
        if orientation == "h":
            pos_pm_y = ph - pos_fitz
            pos_mm = (pos_pm_y - frame.y0) / SCALE
            start_mm = (frame.x1 - max(start_fitz, end_fitz)) / SCALE
            end_mm = (frame.x1 - min(start_fitz, end_fitz)) / SCALE
        else:
            pos_mm = (frame.x1 - pos_fitz) / SCALE
            s_pm = ph - max(start_fitz, end_fitz)
            e_pm = ph - min(start_fitz, end_fitz)
            start_mm = (s_pm - frame.y0) / SCALE
            end_mm = (e_pm - frame.y0) / SCALE
    else:
        raise NotImplementedError(f"fitz_pts_to_grid_line_mm: origin {origin!r} not implemented")

    return (round(pos_mm, 3), round(start_mm, 3), round(end_mm, 3))


# ---------------------------------------------------------------------------
# High-level: field → fitz.Rect (displayed)
# ---------------------------------------------------------------------------

def field_to_fitz_rect(
    field: FieldDef,
    frame: FrameInfo,
    padding_mm: float = 0.0,
) -> fitz.Rect:
    """Template field → fitz Rect in **displayed** coordinates.

    Steps:
    1. ``resolve_field_bbox`` → pdfminer abs rect (handles all origins + stretch)
    2. Expand by *padding_mm* on each side
    3. ``pdfminer_to_fitz`` → fitz.Rect in displayed coords
    4. Clamp to page bounds
    """
    x0, y0, x1, y1 = resolve_field_bbox(field, frame)

    if padding_mm:
        pad_pts = padding_mm * SCALE
        x0 -= pad_pts
        y0 -= pad_pts
        x1 += pad_pts
        y1 += pad_pts

    r = pdfminer_to_fitz((x0, y0, x1, y1), frame.page_height, frame.rotation)

    page_rect = fitz.Rect(0, 0, frame.page_width, frame.page_height)
    return r & page_rect
