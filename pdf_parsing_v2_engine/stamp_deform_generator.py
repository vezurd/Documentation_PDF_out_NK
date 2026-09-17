"""Generate paired deformed-stamp PDF fixtures for v2 benchmarks.

The tool requires two aligned inputs:
1. A source PDF containing the original stamp.
2. A background PDF where the same page exists without the stamp.

The output page is composed from the background page plus a clipped stamp region
inserted with anisotropic scaling via PyMuPDF ``show_pdf_page()``.
"""

from __future__ import annotations

import csv
import json
from dataclasses import dataclass
from pathlib import Path
from typing import Any

import fitz

from pdf_parsing_v2_engine.coord_transform import field_to_fitz_rect
from pdf_parsing_v2_engine.frame_detector import find_frame
from pdf_parsing_v2_engine.models import StampTemplate

DEFAULT_DATASET_ROOT = Path(__file__).resolve().parent / "templates" / "stamp_deform_bench"
DEFAULT_STAMPED_INPUT_DIR = DEFAULT_DATASET_ROOT / "input" / "stamped"
DEFAULT_BACKGROUND_INPUT_DIR = DEFAULT_DATASET_ROOT / "input" / "background"
DEFAULT_TEMPLATE_PATH = DEFAULT_DATASET_ROOT / "template" / "stamp_template.json"
DEFAULT_OUTPUT_DIR = DEFAULT_DATASET_ROOT / "output" / "generated"
DEFAULT_MANIFEST_DIR = DEFAULT_DATASET_ROOT / "output" / "manifests"
DEFAULT_MANIFEST_JSONL = DEFAULT_MANIFEST_DIR / "manifest.jsonl"
DEFAULT_MANIFEST_CSV = DEFAULT_MANIFEST_DIR / "manifest.csv"

DEFAULT_PAGE_NUM = 1
DEFAULT_RANGE_MIN = -25
DEFAULT_RANGE_MAX = 25
DEFAULT_STEP = 5
DEFAULT_MODE = "grid"
DEFAULT_ANCHOR = "bottom_right"

_MODES = frozenset({"grid", "diagonal", "axis-x", "axis-y", "cross"})
_ANCHORS = frozenset({"bottom_right", "bottom_left", "top_right", "top_left", "center"})


@dataclass(frozen=True)
class PairJob:
    """One source/background PDF pair to process."""

    source_pdf: Path
    background_pdf: Path
    relative_key: str


@dataclass(frozen=True)
class DeformationSpec:
    """One deformation point in percent and scale form."""

    dx_percent: int
    dy_percent: int

    @property
    def sx_percent(self) -> int:
        return 100 + self.dx_percent

    @property
    def sy_percent(self) -> int:
        return 100 + self.dy_percent

    @property
    def sx_scale(self) -> float:
        return self.sx_percent / 100.0

    @property
    def sy_scale(self) -> float:
        return self.sy_percent / 100.0

    @property
    def suffix(self) -> str:
        return f"sx{self.sx_percent:03d}__sy{self.sy_percent:03d}"


@dataclass(frozen=True)
class PagePairInfo:
    """Validated page-level information for one source/background pair."""

    source_page_index: int
    background_page_index: int
    source_page_count: int
    source_page_rect: fitz.Rect
    stamp_bbox_source: fitz.Rect


@dataclass(frozen=True)
class ManifestRow:
    """One manifest record for a generated or dry-run output."""

    source_pdf: str
    background_pdf: str
    output_pdf: str
    page_num: int
    dx_percent: int
    dy_percent: int
    sx_percent: int
    sy_percent: int
    sx_scale: float
    sy_scale: float
    stamp_bbox_source: list[float]
    stamp_bbox_target: list[float]
    template_name: str
    mode: str
    anchor: str
    status: str
    warning: str

    def to_dict(self) -> dict[str, Any]:
        """Convert to a JSON/CSV-friendly dict."""
        return {
            "source_pdf": self.source_pdf,
            "background_pdf": self.background_pdf,
            "output_pdf": self.output_pdf,
            "page_num": self.page_num,
            "dx_percent": self.dx_percent,
            "dy_percent": self.dy_percent,
            "sx_percent": self.sx_percent,
            "sy_percent": self.sy_percent,
            "sx_scale": round(self.sx_scale, 6),
            "sy_scale": round(self.sy_scale, 6),
            "stamp_bbox_source": list(self.stamp_bbox_source),
            "stamp_bbox_target": list(self.stamp_bbox_target),
            "template_name": self.template_name,
            "mode": self.mode,
            "anchor": self.anchor,
            "status": self.status,
            "warning": self.warning,
        }


@dataclass(frozen=True)
class GenerationSummary:
    """High-level result returned by the generator."""

    jobs_count: int
    variants_per_job: int
    total_variants: int
    dry_run: bool
    manifest_rows: list[ManifestRow]


@dataclass(frozen=True)
class GenerationConfig:
    """Runtime configuration for paired stamp deformation generation."""

    input_path: Path
    background_input_path: Path
    template_path: Path
    output_dir: Path = DEFAULT_OUTPUT_DIR
    manifest_jsonl_path: Path = DEFAULT_MANIFEST_JSONL
    manifest_csv_path: Path = DEFAULT_MANIFEST_CSV
    page_num: int = DEFAULT_PAGE_NUM
    x_range: tuple[int, int] = (DEFAULT_RANGE_MIN, DEFAULT_RANGE_MAX)
    y_range: tuple[int, int] = (DEFAULT_RANGE_MIN, DEFAULT_RANGE_MAX)
    step: int = DEFAULT_STEP
    mode: str = DEFAULT_MODE
    anchor: str = DEFAULT_ANCHOR
    dry_run: bool = False
    cfg: dict[str, Any] | None = None


def generate_deformed_stamp_dataset(config: GenerationConfig) -> GenerationSummary:
    """Generate or preview paired deformed-stamp fixtures.

    Args:
        config: Generation parameters and filesystem paths.

    Returns:
        Summary with all manifest rows.

    Raises:
        ValueError: If the inputs or deformation grid are invalid.
    """
    _validate_config(config)
    template = StampTemplate.from_json(str(config.template_path))
    jobs = resolve_pair_jobs(config.input_path, config.background_input_path)
    specs = build_deformation_specs(
        x_range=config.x_range,
        y_range=config.y_range,
        step=config.step,
        mode=config.mode,
    )

    manifest_rows: list[ManifestRow] = []
    if not config.dry_run:
        config.output_dir.mkdir(parents=True, exist_ok=True)
        config.manifest_jsonl_path.parent.mkdir(parents=True, exist_ok=True)
        config.manifest_csv_path.parent.mkdir(parents=True, exist_ok=True)

    for job in jobs:
        pair_info = inspect_pair_job(
            job=job,
            template=template,
            page_num=config.page_num,
            cfg=config.cfg,
        )
        for spec in specs:
            target_bbox = compute_target_bbox(
                source_bbox=pair_info.stamp_bbox_source,
                sx_scale=spec.sx_scale,
                sy_scale=spec.sy_scale,
                anchor=config.anchor,
            )
            warning = ""
            if not _rect_inside_page(target_bbox, pair_info.source_page_rect):
                warning = "target stamp bbox exceeds page bounds"

            relative_path = Path(job.relative_key)
            output_name = f"{relative_path.stem}__{spec.suffix}.pdf"
            output_pdf = config.output_dir / relative_path.parent / output_name

            row = ManifestRow(
                source_pdf=str(job.source_pdf),
                background_pdf=str(job.background_pdf),
                output_pdf=str(output_pdf),
                page_num=config.page_num,
                dx_percent=spec.dx_percent,
                dy_percent=spec.dy_percent,
                sx_percent=spec.sx_percent,
                sy_percent=spec.sy_percent,
                sx_scale=spec.sx_scale,
                sy_scale=spec.sy_scale,
                stamp_bbox_source=rect_to_list(pair_info.stamp_bbox_source),
                stamp_bbox_target=rect_to_list(target_bbox),
                template_name=template.name,
                mode=config.mode,
                anchor=config.anchor,
                status="dry_run" if config.dry_run else "generated",
                warning=warning,
            )
            manifest_rows.append(row)

            if config.dry_run:
                continue

            if warning:
                raise ValueError(
                    f"Target bbox leaves page bounds for {job.source_pdf} "
                    f"({spec.suffix}): {target_bbox}"
                )

            output_pdf.parent.mkdir(parents=True, exist_ok=True)
            render_output_pdf(
                source_pdf=job.source_pdf,
                background_pdf=job.background_pdf,
                output_pdf=output_pdf,
                pair_info=pair_info,
                stamp_bbox_target=target_bbox,
            )

    if not config.dry_run:
        write_manifest_jsonl(config.manifest_jsonl_path, manifest_rows)
        write_manifest_csv(config.manifest_csv_path, manifest_rows)

    return GenerationSummary(
        jobs_count=len(jobs),
        variants_per_job=len(specs),
        total_variants=len(jobs) * len(specs),
        dry_run=config.dry_run,
        manifest_rows=manifest_rows,
    )


def resolve_pair_jobs(input_path: Path, background_input_path: Path) -> list[PairJob]:
    """Resolve file-file or dir-dir paired inputs."""
    input_path = input_path.resolve()
    background_input_path = background_input_path.resolve()

    if input_path.is_file() and background_input_path.is_file():
        return [
            PairJob(
                source_pdf=input_path,
                background_pdf=background_input_path,
                relative_key=input_path.name,
            )
        ]

    if input_path.is_dir() and background_input_path.is_dir():
        source_map = _scan_pdf_tree(input_path)
        background_map = _scan_pdf_tree(background_input_path)
        missing = sorted(item["relative_key"] for key, item in source_map.items() if key not in background_map)
        if missing:
            preview = ", ".join(missing[:5])
            raise ValueError(
                "Missing background PDFs for relative paths: "
                f"{preview}{' ...' if len(missing) > 5 else ''}"
            )

        jobs = [
            PairJob(
                source_pdf=item["path"],
                background_pdf=background_map[key]["path"],
                relative_key=item["relative_key"],
            )
            for key, item in sorted(source_map.items(), key=lambda pair: pair[1]["relative_key"])
        ]
        return jobs

    raise ValueError(
        "Input pairing must be file-file or dir-dir: "
        f"{input_path} vs {background_input_path}"
    )


def build_deformation_specs(
    *,
    x_range: tuple[int, int],
    y_range: tuple[int, int],
    step: int,
    mode: str,
) -> list[DeformationSpec]:
    """Expand the requested mode into a list of deformation points."""
    if step <= 0:
        raise ValueError("step must be > 0")
    if mode not in _MODES:
        raise ValueError(f"Unsupported mode: {mode!r}")

    x_values = _inclusive_range(x_range[0], x_range[1], step)
    y_values = _inclusive_range(y_range[0], y_range[1], step)

    if mode == "grid":
        pairs = [(dx, dy) for dx in x_values for dy in y_values]
    elif mode == "diagonal":
        diagonal = sorted(set(x_values).intersection(y_values))
        if not diagonal:
            raise ValueError("diagonal mode requires overlapping X/Y percent values")
        pairs = [(d, d) for d in diagonal]
    elif mode == "axis-x":
        pairs = [(dx, 0) for dx in x_values]
    elif mode == "axis-y":
        pairs = [(0, dy) for dy in y_values]
    else:
        unique_pairs: list[tuple[int, int]] = []
        seen: set[tuple[int, int]] = set()
        for pair in [(dx, 0) for dx in x_values] + [(0, dy) for dy in y_values]:
            if pair in seen:
                continue
            seen.add(pair)
            unique_pairs.append(pair)
        pairs = unique_pairs

    specs = [DeformationSpec(dx_percent=dx, dy_percent=dy) for dx, dy in pairs]
    for spec in specs:
        if spec.sx_scale <= 0 or spec.sy_scale <= 0:
            raise ValueError(
                "Scale percent must stay above zero; "
                f"got {spec.sx_percent}% / {spec.sy_percent}%"
            )
    return specs


def inspect_pair_job(
    *,
    job: PairJob,
    template: StampTemplate,
    page_num: int,
    cfg: dict[str, Any] | None,
) -> PagePairInfo:
    """Inspect one pair and compute the source stamp bbox."""
    with fitz.open(job.source_pdf) as source_doc, fitz.open(job.background_pdf) as background_doc:
        source_page_index = _resolve_source_page_index(source_doc, page_num)
        background_page_index = _resolve_background_page_index(
            source_doc=source_doc,
            background_doc=background_doc,
            source_page_index=source_page_index,
        )
        source_page = source_doc[source_page_index]
        background_page = background_doc[background_page_index]
        _validate_page_compatibility(source_page, background_page, job)

        stamp_bbox_source = compute_stamp_bbox_source(
            fitz_page=source_page,
            template=template,
            cfg=cfg,
        )

        return PagePairInfo(
            source_page_index=source_page_index,
            background_page_index=background_page_index,
            source_page_count=source_doc.page_count,
            source_page_rect=fitz.Rect(source_page.rect),
            stamp_bbox_source=stamp_bbox_source,
        )


def compute_stamp_bbox_source(
    *,
    fitz_page: fitz.Page,
    template: StampTemplate,
    cfg: dict[str, Any] | None,
) -> fitz.Rect:
    """Compute the stamp bbox using v2 template union semantics."""
    frame_template = template if template.frame_mode == "drawing_union" else None
    frame, _ = find_frame(
        fitz_page,
        frame_mode=template.frame_mode,
        template=frame_template,
        cfg=cfg,
    )
    rects = [
        field_to_fitz_rect(field, frame, padding_mm=0.0)
        for field in template.fields
        if not field.outside_stamp and not field.document_property
    ]
    rects = [rect for rect in rects if not rect.is_empty and not rect.is_infinite]
    if not rects:
        raise ValueError(
            f"Template {template.name!r} has no inside-stamp geometric fields to union"
        )
    x0 = min(rect.x0 for rect in rects)
    y0 = min(rect.y0 for rect in rects)
    x1 = max(rect.x1 for rect in rects)
    y1 = max(rect.y1 for rect in rects)
    bbox = fitz.Rect(x0, y0, x1, y1) & fitz_page.rect
    if bbox.is_empty:
        raise ValueError(f"Computed empty stamp bbox for template {template.name!r}")
    return bbox


def compute_target_bbox(
    *,
    source_bbox: fitz.Rect,
    sx_scale: float,
    sy_scale: float,
    anchor: str,
) -> fitz.Rect:
    """Scale a rect anisotropically around the selected anchor point."""
    if anchor not in _ANCHORS:
        raise ValueError(f"Unsupported anchor: {anchor!r}")

    width = source_bbox.width * sx_scale
    height = source_bbox.height * sy_scale
    if width <= 0 or height <= 0:
        raise ValueError("Scaled bbox width/height must stay positive")

    if anchor == "bottom_right":
        return fitz.Rect(source_bbox.x1 - width, source_bbox.y1 - height, source_bbox.x1, source_bbox.y1)
    if anchor == "bottom_left":
        return fitz.Rect(source_bbox.x0, source_bbox.y1 - height, source_bbox.x0 + width, source_bbox.y1)
    if anchor == "top_right":
        return fitz.Rect(source_bbox.x1 - width, source_bbox.y0, source_bbox.x1, source_bbox.y0 + height)
    if anchor == "top_left":
        return fitz.Rect(source_bbox.x0, source_bbox.y0, source_bbox.x0 + width, source_bbox.y0 + height)

    cx = (source_bbox.x0 + source_bbox.x1) * 0.5
    cy = (source_bbox.y0 + source_bbox.y1) * 0.5
    return fitz.Rect(cx - width * 0.5, cy - height * 0.5, cx + width * 0.5, cy + height * 0.5)


def render_output_pdf(
    *,
    source_pdf: Path,
    background_pdf: Path,
    output_pdf: Path,
    pair_info: PagePairInfo,
    stamp_bbox_target: fitz.Rect,
) -> None:
    """Render one output PDF using the paired background mode."""
    with fitz.open(source_pdf) as source_doc, fitz.open(background_pdf) as background_doc:
        out_doc = fitz.open()
        try:
            _seed_output_doc_from_background_or_source(
                out_doc=out_doc,
                source_doc=source_doc,
                background_doc=background_doc,
            )
            dest_page = out_doc[pair_info.source_page_index]
            dest_page.show_pdf_page(
                stamp_bbox_target,
                source_doc,
                pair_info.source_page_index,
                keep_proportion=False,
                overlay=True,
                clip=pair_info.stamp_bbox_source,
            )
            out_doc.save(output_pdf, garbage=3, deflate=True)
        finally:
            out_doc.close()


def _seed_output_doc_from_background_or_source(
    *,
    out_doc: fitz.Document,
    source_doc: fitz.Document,
    background_doc: fitz.Document,
) -> None:
    """Seed the output document while preserving original page geometry."""
    if background_doc.page_count == source_doc.page_count:
        out_doc.insert_pdf(background_doc)
        return
    if background_doc.page_count == 1 and source_doc.page_count == 1:
        out_doc.insert_pdf(background_doc)
        return
    raise ValueError(
        "Paired background PDF must have the same page count as source "
        "for geometry-preserving output assembly"
    )


def rect_to_list(rect: fitz.Rect) -> list[float]:
    """Serialize a rect as rounded float values."""
    return [round(float(value), 3) for value in (rect.x0, rect.y0, rect.x1, rect.y1)]


def write_manifest_jsonl(path: Path, rows: list[ManifestRow]) -> None:
    """Write manifest rows to JSONL."""
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8") as fh:
        for row in rows:
            fh.write(json.dumps(row.to_dict(), ensure_ascii=False) + "\n")


def write_manifest_csv(path: Path, rows: list[ManifestRow]) -> None:
    """Write manifest rows to CSV."""
    path.parent.mkdir(parents=True, exist_ok=True)
    fieldnames = [
        "source_pdf",
        "background_pdf",
        "output_pdf",
        "page_num",
        "dx_percent",
        "dy_percent",
        "sx_percent",
        "sy_percent",
        "sx_scale",
        "sy_scale",
        "stamp_bbox_source",
        "stamp_bbox_target",
        "template_name",
        "mode",
        "anchor",
        "status",
        "warning",
    ]
    with path.open("w", encoding="utf-8", newline="") as fh:
        writer = csv.DictWriter(fh, fieldnames=fieldnames)
        writer.writeheader()
        for row in rows:
            payload = row.to_dict()
            payload["stamp_bbox_source"] = json.dumps(payload["stamp_bbox_source"], ensure_ascii=False)
            payload["stamp_bbox_target"] = json.dumps(payload["stamp_bbox_target"], ensure_ascii=False)
            writer.writerow(payload)


def _inclusive_range(start: int, stop: int, step: int) -> list[int]:
    """Return an inclusive integer range for positive *step*."""
    if step <= 0:
        raise ValueError("step must be > 0")
    if start > stop:
        raise ValueError(f"Invalid range: start {start} > stop {stop}")
    return list(range(start, stop + 1, step))


def _scan_pdf_tree(root: Path) -> dict[str, dict[str, Any]]:
    """Scan a directory recursively and map relative PDF paths to absolute ones."""
    mapping: dict[str, dict[str, Any]] = {}
    for pdf_path in sorted(root.rglob("*.pdf")):
        if not pdf_path.is_file():
            continue
        relative_key = pdf_path.relative_to(root).as_posix()
        mapping[relative_key.lower()] = {
            "relative_key": relative_key,
            "path": pdf_path.resolve(),
        }
    return mapping


def _resolve_source_page_index(source_doc: fitz.Document, page_num: int) -> int:
    """Convert a 1-based page number to a valid source page index."""
    if page_num < 1 or page_num > source_doc.page_count:
        raise ValueError(
            f"Requested page_num={page_num} is outside source PDF page range 1..{source_doc.page_count}"
        )
    return page_num - 1


def _resolve_background_page_index(
    *,
    source_doc: fitz.Document,
    background_doc: fitz.Document,
    source_page_index: int,
) -> int:
    """Resolve which background page should be used for the selected source page."""
    if background_doc.page_count == source_doc.page_count:
        return source_page_index
    if background_doc.page_count == 1:
        return 0
    raise ValueError(
        "Background PDF must have either the same page count as the source "
        f"({source_doc.page_count}) or exactly one page, got {background_doc.page_count}"
    )


def _validate_page_compatibility(
    source_page: fitz.Page,
    background_page: fitz.Page,
    job: PairJob,
) -> None:
    """Validate that paired pages are visually aligned for composition."""
    if source_page.rotation != background_page.rotation:
        raise ValueError(
            f"Rotation mismatch for {job.relative_key}: "
            f"source={source_page.rotation}, background={background_page.rotation}"
        )
    if not _same_rect(source_page.rect, background_page.rect):
        raise ValueError(
            f"Page size mismatch for {job.relative_key}: "
            f"source={source_page.rect}, background={background_page.rect}"
        )


def _same_rect(a: fitz.Rect, b: fitz.Rect, eps: float = 0.01) -> bool:
    """Approximate rect equality for page size checks."""
    return (
        abs(a.x0 - b.x0) <= eps
        and abs(a.y0 - b.y0) <= eps
        and abs(a.x1 - b.x1) <= eps
        and abs(a.y1 - b.y1) <= eps
    )


def _rect_inside_page(rect: fitz.Rect, page_rect: fitz.Rect, eps: float = 0.01) -> bool:
    """Check whether *rect* stays inside the visible page bounds."""
    return (
        rect.x0 >= page_rect.x0 - eps
        and rect.y0 >= page_rect.y0 - eps
        and rect.x1 <= page_rect.x1 + eps
        and rect.y1 <= page_rect.y1 + eps
    )


def _validate_config(config: GenerationConfig) -> None:
    """Validate high-level generator configuration."""
    if config.page_num < 1:
        raise ValueError("page_num must be >= 1")
    if config.mode not in _MODES:
        raise ValueError(f"Unsupported mode: {config.mode!r}")
    if config.anchor not in _ANCHORS:
        raise ValueError(f"Unsupported anchor: {config.anchor!r}")
    if not config.template_path.is_file():
        raise ValueError(
            "Template JSON not found: "
            f"{config.template_path}. Put a template at the default path or pass --template."
        )
    if not config.input_path.exists():
        raise ValueError(f"Input path not found: {config.input_path}")
    if not config.background_input_path.exists():
        raise ValueError(f"Background input path not found: {config.background_input_path}")

