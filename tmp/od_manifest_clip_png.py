"""One-off: render OD manifest clip / frame / stamp bbox over PDF pages (PNG).

Run from repo root:
  python tmp/od_manifest_clip_png.py <pdf_path> [out_png_prefix]

Requires network path to be reachable if PDF is on UNC share.
"""
from __future__ import annotations

import os
import sys
from pathlib import Path

_REPO = Path(__file__).resolve().parents[1]
if str(_REPO) not in sys.path:
    sys.path.insert(0, str(_REPO))

import fitz

from pdf_parsing_v2 import load_project_templates
from pdf_parsing_v2.v2_config import load_v2_config, resolve_templates_dir
from pdf_parsing_v2_engine.stamp_extractor import extract_page
from pdf_parsing_v2_od.od_manifest_clip import build_manifest_clip_frame_minus_stamp

try:
    from PIL import Image, ImageDraw
except ImportError as e:
    raise SystemExit("Pillow required: pip install Pillow") from e

_ZOOM = 2.0
_MAX_PAGES = 4
_COL_FRAME = (0, 120, 255, 180)
_COL_STAMP = (255, 80, 0, 160)
_COL_CLIP = (0, 220, 80, 200)
_COL_LABEL_BG = (255, 255, 255, 220)


def _pdf_rect_to_px(rect: fitz.Rect, zm: float) -> tuple[int, int, int, int]:
    x0 = int(rect.x0 * zm)
    y0 = int(rect.y0 * zm)
    x1 = int(rect.x1 * zm)
    y1 = int(rect.y1 * zm)
    return (x0, y0, x1, y1)


def _stamp_bbox(meta: dict) -> tuple[float, float, float, float] | None:
    seb = meta.get("stamp_effective_bbox")
    if not isinstance(seb, (list, tuple)) or len(seb) != 4:
        return None
    try:
        return (float(seb[0]), float(seb[1]), float(seb[2]), float(seb[3]))
    except (TypeError, ValueError):
        return None


def main() -> int:
    if len(sys.argv) < 2:
        print("usage: python tmp/od_manifest_clip_png.py <pdf> [out_prefix]")
        return 2
    pdf_path = os.path.abspath(sys.argv[1].strip())
    out_prefix = (
        sys.argv[2].strip()
        if len(sys.argv) > 2
        else os.path.join(os.path.dirname(__file__), "od_manifest_clip_overlay")
    )

    if not os.path.isfile(pdf_path):
        print(f"PDF not found or not a file: {pdf_path}")
        return 1

    cfg = load_v2_config()
    td = resolve_templates_dir(cfg)
    project = (cfg.get("project") or "").strip() or "agcc_287"
    templates = load_project_templates(td, project)
    if not templates:
        print(f"No templates for project={project!r} in {td}")
        return 1

    doc = fitz.open(pdf_path)
    try:
        n = min(doc.page_count, _MAX_PAGES)
        mat = fitz.Matrix(_ZOOM, _ZOOM)
        for i in range(n):
            page = doc[i]
            pn = i + 1
            result = extract_page(page, "OD", pn, templates, cfg)
            fr = result.frame
            frame_rect = fitz.Rect(fr.x0, fr.y0, fr.x1, fr.y1)
            meta = result.metadata or {}
            seb = _stamp_bbox(meta)
            tmpl = next((t for t in templates if t.name == result.template_name), None)
            origin = tmpl.origin if tmpl is not None else "frame_bottom_right"

            page_rect = fitz.Rect(0.0, 0.0, page.rect.width, page.rect.height)
            stamp_rect = (
                fitz.Rect(seb[0], seb[1], seb[2], seb[3]) if seb else fitz.Rect()
            )
            clip = (
                build_manifest_clip_frame_minus_stamp(
                    frame_rect=frame_rect,
                    stamp_rect=stamp_rect,
                    page_rect=page_rect,
                    origin=origin,
                )
                if seb
                else None
            )

            pix = page.get_pixmap(matrix=mat, alpha=False)
            img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
            overlay = Image.new("RGBA", img.size, (0, 0, 0, 0))
            draw = ImageDraw.Draw(overlay, "RGBA")

            zm = _ZOOM
            draw.rectangle(_pdf_rect_to_px(frame_rect, zm), outline=_COL_FRAME[:3], width=3)
            if seb and not stamp_rect.is_empty:
                draw.rectangle(_pdf_rect_to_px(stamp_rect, zm), outline=_COL_STAMP[:3], width=3)
                draw.rectangle(
                    _pdf_rect_to_px(stamp_rect, zm),
                    fill=_COL_STAMP,
                )
            if clip is not None and not clip.is_empty:
                draw.rectangle(_pdf_rect_to_px(clip, zm), outline=_COL_CLIP[:3], width=4)
                draw.rectangle(_pdf_rect_to_px(clip, zm), fill=_COL_CLIP)

            composed = Image.alpha_composite(img.convert("RGBA"), overlay).convert("RGB")

            legend = Image.new("RGBA", (composed.width, composed.height + 72), (255, 255, 255, 255))
            legend.paste(composed, (0, 0))
            ld = ImageDraw.Draw(legend)
            y = composed.height + 6
            lh = 18
            lines = [
                f"p{pn} template={result.template_name!r} origin={origin!r}",
                f"frame=[{frame_rect.x0:.1f},{frame_rect.y0:.1f},{frame_rect.x1:.1f},{frame_rect.y1:.1f}]",
                (
                    f"stamp_effective_bbox={[round(x,2) for x in seb]}"
                    if seb
                    else "stamp_effective_bbox=MISSING"
                ),
                (
                    f"manifest_clip(find_tables)=[{clip.x0:.1f},{clip.y0:.1f},{clip.x1:.1f},{clip.y1:.1f}]"
                    if clip and not clip.is_empty
                    else "manifest_clip=None or empty"
                ),
            ]
            for k, text in enumerate(lines):
                ld.text((8, y + k * lh), text, fill=(20, 20, 20))

            outp = f"{out_prefix}_p{pn}.png"
            legend.convert("RGB").save(outp, "PNG")
            print("wrote", outp)
    finally:
        doc.close()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
