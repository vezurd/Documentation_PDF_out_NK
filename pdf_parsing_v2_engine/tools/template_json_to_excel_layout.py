"""
Построить приблизительный макет Excel по JSON-шаблону v2: bbox_mm → merge-ячейки.

Сетка строится только по полям внутри штампа (без outside_stamp), чтобы крайние поля
не «мельчили» общую сетку. Поля outside_stamp выносятся в полосы над/под блоком штампа
по положению v (выше/ниже рамки) с горизонтальным слотом по h.

Подписи: полный field id; для «маленьких» полей — только ведущее число из id.

Запуск (из корня репозитория):
  python pdf_parsing_v2_engine/tools/template_json_to_excel_layout.py path/to/template.json
  python ... --scale-h 1.2 --scale-v 0.9
  # scale-h / scale-x → ось h bbox (колонки Excel); scale-v / scale-y → ось v (строки)
"""

from __future__ import annotations

import argparse
import bisect
import json
import re
from pathlib import Path
from typing import Any

import xlsxwriter

# мм: границы с ближним шагом схлопываем (float из JSON)
_EDGE_QUANT_MM = 0.02
# поле считаем «маленьким» для короткой подписи
_SMALL_AREA_MM2 = 140.0
_SMALL_MAX_SIDE_MM = 11.0
# умолчания масштаба сетки (синхронно с argparse в main)
_DEFAULT_SCALE_H = 1.0
_DEFAULT_SCALE_V = 1.0
# визуальный размер ячеек Excel (мельче было из‑за 2.3 / 11 pt)
_DEFAULT_COL_WIDTH = 4.8
_DEFAULT_ROW_HEIGHT = 19.0


def _quantize(mm: float) -> float:
    return round(mm / _EDGE_QUANT_MM) * _EDGE_QUANT_MM


def _short_id(field_id: str) -> str:
    """Только ведущие цифры id (до первого нецифрового символа)."""
    m = re.match(r"^(\d+)", field_id)
    if m:
        return m.group(1)
    return field_id[:12]


def _display_label(field_id: str, small: bool) -> str:
    return _short_id(field_id) if small else field_id


def _field_label(field: dict[str, Any]) -> str:
    return str(field.get("id", ""))


def _is_outside_stamp(f: dict[str, Any]) -> bool:
    return bool(f.get("outside_stamp"))


def _collect_edges(
    fields: list[dict[str, Any]],
    *,
    scale_h: float = 1.0,
    scale_v: float = 1.0,
) -> tuple[list[float], list[float]]:
    hs: set[float] = set()
    vs: set[float] = set()
    for f in fields:
        bb = f.get("bbox_mm")
        if not bb or len(bb) != 4:
            continue
        h1, v1, h2, v2 = (float(bb[0]), float(bb[1]), float(bb[2]), float(bb[3]))
        hs.update((_quantize(h1 * scale_h), _quantize(h2 * scale_h)))
        vs.update((_quantize(v1 * scale_v), _quantize(v2 * scale_v)))
    return sorted(hs), sorted(vs)


def _span_on_axis(
    lo: float, hi: float, edges: list[float]
) -> tuple[int, int] | None:
    """Индексы интервалов [edges[i], edges[i+1]], пересекающих [lo, hi] (после quant)."""
    a, b = min(lo, hi), max(lo, hi)
    if len(edges) < 2:
        return None
    if b - a < _EDGE_QUANT_MM * 0.25:
        x = (a + b) * 0.5
        j = bisect.bisect_right(edges, x) - 1
        j = max(0, min(j, len(edges) - 2))
        return j, j
    i0 = None
    i1 = None
    for i in range(len(edges) - 1):
        seg_lo, seg_hi = edges[i], edges[i + 1]
        if seg_hi <= a or seg_lo >= b:
            continue
        if i0 is None:
            i0 = i
        i1 = i
    if i0 is None or i1 is None:
        return None
    return i0, i1


def _excel_cols(g0: int, g1: int, n_intervals: int) -> tuple[int, int]:
    """h: малый индекс интервала = правый край штампа → больший excel col справа."""
    ex0 = (n_intervals - 1) - g1
    ex1 = (n_intervals - 1) - g0
    return ex0, ex1


def _excel_rows(r0: int, r1: int, n_intervals: int) -> tuple[int, int]:
    """v: малый индекс = низ штампа → низ листа Excel."""
    ex0 = (n_intervals - 1) - r1
    ex1 = (n_intervals - 1) - r0
    return ex0, ex1


def _inside_mm_bounds(
    inside_fields: list[dict[str, Any]], scale_h: float, scale_v: float
) -> tuple[float, float, float, float]:
    """h_lo, h_hi, v_lo, v_hi в масштабированных мм (для полос вне штампа)."""
    hs: list[float] = []
    vs: list[float] = []
    for f in inside_fields:
        bb = f.get("bbox_mm")
        if not bb or len(bb) != 4:
            continue
        h1, v1, h2, v2 = (float(bb[0]), float(bb[1]), float(bb[2]), float(bb[3]))
        hs.extend((h1 * scale_h, h2 * scale_h))
        vs.extend((v1 * scale_v, v2 * scale_v))
    if not hs or not vs:
        return 0.0, 1.0, 0.0, 1.0
    return min(hs), max(hs), min(vs), max(vs)


def _field_small(bb: list[Any], small_area_mm2: float, small_max_side_mm: float) -> bool:
    h1, v1, h2, v2 = (float(bb[0]), float(bb[1]), float(bb[2]), float(bb[3]))
    w_mm = abs(h1 - h2)
    h_mm = abs(v1 - v2)
    area = w_mm * h_mm
    return area < small_area_mm2 or (w_mm < small_max_side_mm and h_mm < small_max_side_mm)


def template_to_excel(
    template_path: Path,
    out_path: Path,
    *,
    small_area_mm2: float = _SMALL_AREA_MM2,
    small_max_side_mm: float = _SMALL_MAX_SIDE_MM,
    scale_h: float = _DEFAULT_SCALE_H,
    scale_v: float = _DEFAULT_SCALE_V,
    col_width: float = _DEFAULT_COL_WIDTH,
    row_height: float = _DEFAULT_ROW_HEIGHT,
) -> None:
    if scale_h <= 0 or scale_v <= 0:
        raise SystemExit("scale_h и scale_v должны быть > 0")

    data = json.loads(template_path.read_text(encoding="utf-8"))
    fields: list[dict[str, Any]] = list(data.get("fields") or [])
    if not fields:
        raise SystemExit("В JSON нет fields")

    inside_fields = [f for f in fields if not _is_outside_stamp(f)]
    outside_fields = [f for f in fields if _is_outside_stamp(f)]
    if not inside_fields:
        inside_fields = list(fields)
        outside_fields = []

    h_edges, v_edges = _collect_edges(inside_fields, scale_h=scale_h, scale_v=scale_v)
    if len(h_edges) < 2 or len(v_edges) < 2:
        raise SystemExit("Недостаточно границ для сетки (внутри штампа)")

    n_h = len(h_edges) - 1
    n_v = len(v_edges) - 1

    h_lo_in, h_hi_in, v_lo_in, v_hi_in = _inside_mm_bounds(inside_fields, scale_h, scale_v)
    h_span_in = h_hi_in - h_lo_in + 1e-9
    v_span_in = v_hi_in - v_lo_in + 1e-9

    items: list[dict[str, Any]] = []
    for f in inside_fields:
        bb = f.get("bbox_mm")
        if not bb or len(bb) != 4:
            continue
        h1, v1, h2, v2 = (float(bb[0]), float(bb[1]), float(bb[2]), float(bb[3]))
        h1q = _quantize(h1 * scale_h)
        v1q = _quantize(v1 * scale_v)
        h2q = _quantize(h2 * scale_h)
        v2q = _quantize(v2 * scale_v)
        small = _field_small(bb, small_area_mm2, small_max_side_mm)
        hs = _span_on_axis(h1q, h2q, h_edges)
        vs = _span_on_axis(v1q, v2q, v_edges)
        if hs is None or vs is None:
            continue
        g0, g1 = hs
        r0, r1 = vs
        c0, c1 = _excel_cols(g0, g1, n_h)
        er0, er1 = _excel_rows(r0, r1, n_v)
        w_mm = abs(h1 - h2)
        h_mm = abs(v1 - v2)
        items.append(
            {
                "id": _field_label(f),
                "small": small,
                "c0": min(c0, c1),
                "c1": max(c0, c1),
                "r0": min(er0, er1),
                "r1": max(er0, er1),
                "area": w_mm * h_mm,
            }
        )

    items.sort(key=lambda x: (-x["area"], x["id"]))
    if not items and not outside_fields:
        raise SystemExit("Не удалось сопоставить ни одного поля сетке")

    # --- outside: классификация над / под рамкой по центру v ---
    above_out: list[dict[str, Any]] = []
    below_out: list[dict[str, Any]] = []
    for f in outside_fields:
        bb = f.get("bbox_mm")
        if not bb or len(bb) != 4:
            continue
        h1, v1, h2, v2 = (float(bb[0]), float(bb[1]), float(bb[2]), float(bb[3]))
        vc = (v1 + v2) * 0.5 * scale_v
        small = _field_small(bb, small_area_mm2, small_max_side_mm)
        if vc > v_hi_in + _EDGE_QUANT_MM:
            above_out.append({"f": f, "small": small})
        else:
            below_out.append({"f": f, "small": small})

    meta_top = 3
    above_start = meta_top
    n_above = len(above_out)
    data_row0 = above_start + n_above + (1 if n_above else 0)
    below_start = data_row0 + n_v + 1

    workbook = xlsxwriter.Workbook(str(out_path))
    sheet = workbook.add_worksheet("layout")
    fmt_small = workbook.add_format({"font_size": 8, "text_wrap": True, "valign": "vcenter"})
    fmt_norm = workbook.add_format({"font_size": 9, "text_wrap": True, "valign": "vcenter"})
    fmt_annex = workbook.add_format({"font_size": 9, "text_wrap": True, "valign": "vcenter", "bg_color": "#F2F2F2"})
    fmt_header = workbook.add_format({"bold": True, "font_size": 10})

    sheet.write(0, 0, f"template: {template_path.name}", fmt_header)
    grid_note = (
        f"grid: {n_h}×{n_v} (inside stamp only), quant {_EDGE_QUANT_MM} mm, "
        f"scale_h={scale_h}, scale_v={scale_v}; outside: {len(above_out)} above / {len(below_out)} below"
    )
    sheet.write(1, 0, grid_note, fmt_header)

    max_col = n_h - 1

    def _annex_h_slot(h1: float, h2: float) -> tuple[int, int]:
        hc = (h1 + h2) * 0.5 * scale_h
        t = (hc - h_lo_in) / h_span_in
        t = max(0.0, min(1.0, t))
        c0 = int(t * (n_h - 1))
        w_mm = abs(h1 - h2) * scale_h
        frac = w_mm / h_span_in
        wcols = max(1, min(n_h - c0, int(round(frac * n_h)) + 1))
        c1 = min(n_h - 1, c0 + wcols - 1)
        return c0, c1

    # Полоса «над штампом» (в шаблоне v выше рамки)
    for i, pack in enumerate(above_out):
        f = pack["f"]
        bb = f["bbox_mm"]
        h1, v1, h2, v2 = (float(bb[0]), float(bb[1]), float(bb[2]), float(bb[3]))
        c0, c1 = _annex_h_slot(h1, h2)
        row = above_start + i
        label = _display_label(_field_label(f), pack["small"])
        max_col = max(max_col, c1)
        if c0 == c1:
            sheet.write(row, c0, label, fmt_annex)
        else:
            sheet.merge_range(row, c0, row, c1, label, fmt_annex)

    used: set[tuple[int, int]] = set()

    def try_occupy(r0: int, r1: int, c0: int, c1: int) -> bool:
        for r in range(r0, r1 + 1):
            for c in range(c0, c1 + 1):
                if (r, c) in used:
                    return False
        for r in range(r0, r1 + 1):
            for c in range(c0, c1 + 1):
                used.add((r, c))
        return True

    skipped: list[str] = []
    fallback: list[str] = []

    def place_inside(it: dict[str, Any]) -> None:
        nonlocal max_col
        r0, r1 = it["r0"] + data_row0, it["r1"] + data_row0
        c0, c1 = it["c0"], it["c1"]
        label = _display_label(it["id"], it["small"])
        fmt = fmt_small if it["small"] else fmt_norm
        max_col = max(max_col, c1)
        if r0 == r1 and c0 == c1:
            if (r0, c0) not in used:
                used.add((r0, c0))
                sheet.write(r0, c0, label, fmt)
                return
            skipped.append(it["id"])
            return
        if try_occupy(r0, r1, c0, c1):
            sheet.merge_range(r0, c0, r1, c1, label, fmt)
            return
        # fallback: одна свободная ячейка внутри прямоугольника
        for r in range(r0, r1 + 1):
            for c in range(c0, c1 + 1):
                if (r, c) not in used:
                    used.add((r, c))
                    sheet.write(r, c, label, fmt)
                    fallback.append(it["id"])
                    return
        skipped.append(it["id"])

    for it in items:
        place_inside(it)

    # Полоса «под штампом»
    for i, pack in enumerate(below_out):
        f = pack["f"]
        bb = f["bbox_mm"]
        h1, v1, h2, v2 = (float(bb[0]), float(bb[1]), float(bb[2]), float(bb[3]))
        c0, c1 = _annex_h_slot(h1, h2)
        row = below_start + i
        label = _display_label(_field_label(f), pack["small"])
        max_col = max(max_col, c1)
        if c0 == c1:
            sheet.write(row, c0, label, fmt_annex)
        else:
            sheet.merge_range(row, c0, row, c1, label, fmt_annex)

    max_row_used = data_row0 + n_v - 1
    if items:
        max_row_used = max(max_row_used, max(it["r1"] + data_row0 for it in items))
    if below_out:
        max_row_used = max(max_row_used, below_start + len(below_out) - 1)

    for c in range(max_col + 1):
        sheet.set_column(c, c, col_width)
    for r in range(meta_top, max_row_used + 1):
        sheet.set_row(r, row_height)

    notes: list[str] = []
    if fallback:
        notes.append(f"merge→1cell: {', '.join(fallback)}")
    if skipped:
        notes.append(f"still skipped: {', '.join(skipped)}")
    if notes:
        sheet.write(2, 0, " | ".join(notes), fmt_header)

    workbook.close()
    print(f"[template_json_to_excel_layout] written: {out_path} (inside {len(items)}, "
          f"above {len(above_out)}, below {len(below_out)}, fallback {len(fallback)}, skipped {len(skipped)})")
    if skipped:
        print("  skipped:", ", ".join(skipped))


def main() -> None:
    ap = argparse.ArgumentParser(description="Excel layout from v2 template JSON (bbox_mm grid)")
    ap.add_argument("template_json", type=Path, help="Path to template .json")
    ap.add_argument(
        "--out",
        type=Path,
        default=None,
        help="Output .xlsx (default: same dir as json, <stem>_layout.xlsx)",
    )
    ap.add_argument(
        "--scale-h",
        type=float,
        default=_DEFAULT_SCALE_H,
        metavar="K",
        help=f"Множитель оси h bbox (→ колонки Excel), по умолчанию {_DEFAULT_SCALE_H}",
    )
    ap.add_argument(
        "--scale-v",
        type=float,
        default=_DEFAULT_SCALE_V,
        metavar="K",
        help=f"Множитель оси v bbox (→ строки Excel), по умолчанию {_DEFAULT_SCALE_V}",
    )
    ap.add_argument(
        "--scale-x",
        type=float,
        default=None,
        metavar="K",
        dest="scale_x",
        help="Синоним --scale-h",
    )
    ap.add_argument(
        "--scale-y",
        type=float,
        default=None,
        metavar="K",
        dest="scale_y",
        help="Синоним --scale-v",
    )
    ap.add_argument(
        "--col-width",
        type=float,
        default=_DEFAULT_COL_WIDTH,
        metavar="W",
        help=f"Ширина колонки Excel, по умолчанию {_DEFAULT_COL_WIDTH}",
    )
    ap.add_argument(
        "--row-height",
        type=float,
        default=_DEFAULT_ROW_HEIGHT,
        metavar="H",
        help=f"Высота строки (pt), по умолчанию {_DEFAULT_ROW_HEIGHT}",
    )
    args = ap.parse_args()
    tpl = args.template_json.resolve()
    if not tpl.is_file():
        raise SystemExit(f"Not found: {tpl}")
    out = args.out.resolve() if args.out else tpl.parent / f"{tpl.stem}_layout.xlsx"
    sh = args.scale_h if args.scale_x is None else args.scale_x
    sv = args.scale_v if args.scale_y is None else args.scale_y
    template_to_excel(
        tpl,
        out,
        scale_h=sh,
        scale_v=sv,
        col_width=args.col_width,
        row_height=args.row_height,
    )


if __name__ == "__main__":
    main()
