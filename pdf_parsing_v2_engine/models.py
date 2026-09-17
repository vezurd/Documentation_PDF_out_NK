"""
Модели данных v2: шаблоны полей, результаты извлечения, каталоги и наборы шаблонов.

Схема JSON шаблона — см. `.cursor/rules/AI_v2_engine.mdc` (раздел «Формат JSON-шаблона»).
"""

from __future__ import annotations

import json
import os
from dataclasses import dataclass, field, replace
from typing import Any

from pdf_parsing_v2_engine.stamp_text.outcome import ParseWarning


_EXPECTED_VALUES = frozenset({"required", "optional", "absent"})
_PAGE_SELECTOR_VALUES = frozenset({"first", "rest", "all"})
_ORIGIN_VALUES = frozenset({
    "frame_bottom_right",
    "frame_top_right",
    "frame_bottom_left",
    "frame_top_left",
})
_STRETCH_EDGES = frozenset({"bottom", "top", "left", "right"})
_FIELD_TYPE_VALUES = frozenset({"data", "label", "empty"})
_LINE_ORIENTATIONS = frozenset({"h", "v"})
_LINE_BOUNDARIES = frozenset({"top", "bottom", "left", "right"})
_FRAME_MODE_VALUES = frozenset({"gost", "drawing_union"})

_DEFAULT_DEBUG_REPORT_ORDER = 999


def normalize_field_type(raw: str) -> str:
    """Миграция: static→label, static_empty→empty; неизвестное → data."""
    if raw == "static":
        return "label"
    if raw == "static_empty":
        return "empty"
    if raw in _FIELD_TYPE_VALUES:
        return raw
    return "data"


def normalize_frame_mode(raw: str | None) -> str:
    """StampTemplate.frame_mode: ``gost`` (default) or ``drawing_union``."""
    s = (raw or "").strip().lower()
    if s in _FRAME_MODE_VALUES:
        return s
    return "gost"


def normalize_expected(raw: str) -> str:
    """Допустимое значение FieldDef.expected (required|optional|absent)."""
    e = (raw or "").strip().lower()
    if e in _EXPECTED_VALUES:
        return e
    return "optional"


def _expect(cond: bool, msg: str) -> None:
    if not cond:
        raise ValueError(msg)


@dataclass
class TemplateGridLine:
    """One grid/wireframe line of the stamp template.

    Coordinates use the same origin as the template's ``origin`` field
    (default ``frame_bottom_right``).

    For a H-line (horizontal): ``pos_mm`` is on the V-axis (mm from origin),
    ``start_mm``/``end_mm`` are on the H-axis.
    For a V-line (vertical): ``pos_mm`` is on the H-axis, ``start_mm``/``end_mm``
    are on the V-axis.
    """

    id: str
    orientation: str          # "h" or "v"
    pos_mm: float             # position on perpendicular axis (mm from origin)
    start_mm: float           # start on parallel axis (mm from origin)
    end_mm: float             # end on parallel axis (mm from origin)
    boundary: str | None = None  # "top"/"bottom"/"left"/"right" or None
    start_line_id: str | None = None  # ID of perpendicular line this line meets at start
    end_line_id: str | None = None    # ID of perpendicular line this line meets at end

    def __post_init__(self) -> None:
        _expect(
            self.orientation in _LINE_ORIENTATIONS,
            f"TemplateGridLine.orientation must be 'h' or 'v', got {self.orientation!r}",
        )
        _expect(
            self.boundary is None or self.boundary in _LINE_BOUNDARIES,
            f"TemplateGridLine.boundary must be one of {_LINE_BOUNDARIES} or None, "
            f"got {self.boundary!r}",
        )

    @classmethod
    def from_dict(cls, d: dict[str, Any]) -> "TemplateGridLine":
        return cls(
            id=str(d["id"]),
            orientation=str(d.get("orientation", "h")),
            pos_mm=float(d["pos_mm"]),
            start_mm=float(d.get("start_mm", 0.0)),
            end_mm=float(d.get("end_mm", 0.0)),
            boundary=(str(d["boundary"]) if d.get("boundary") else None),
            start_line_id=(str(d["start_line_id"]) if d.get("start_line_id") else None),
            end_line_id=(str(d["end_line_id"]) if d.get("end_line_id") else None),
        )

    def to_dict(self) -> dict[str, Any]:
        d: dict[str, Any] = {
            "id": self.id,
            "orientation": self.orientation,
            "pos_mm": self.pos_mm,
            "start_mm": self.start_mm,
            "end_mm": self.end_mm,
        }
        if self.boundary:
            d["boundary"] = self.boundary
        if self.start_line_id:
            d["start_line_id"] = self.start_line_id
        if self.end_line_id:
            d["end_line_id"] = self.end_line_id
        return d


@dataclass
class FieldDef:
    id: str
    label: str
    bbox_mm: tuple[float, float, float, float]
    clean: str | None
    validate_regex: str | None = None
    expected: str = "optional"
    padding_mm: float | None = None
    field_type: str = "data"
    expected_text: str | None = None
    is_anchor: bool = False
    outside_stamp: bool = False
    origin: str = "frame_bottom_right"
    stretch_to_page: tuple[str, ...] = ()
    bound_top: str | None = None     # ID of H-line bounding from above (sets fitz y0)
    bound_bottom: str | None = None  # ID of H-line bounding from below (sets fitz y1)
    bound_left: str | None = None    # ID of V-line bounding from left (sets fitz x0)
    bound_right: str | None = None   # ID of V-line bounding from right (sets fitz x1)
    # If set, value comes from ``pdf_parsing_v2.document_properties`` (no bbox text extraction).
    document_property: str | None = None

    def __post_init__(self) -> None:
        _expect(self.expected in _EXPECTED_VALUES, f"FieldDef.expected must be one of {_EXPECTED_VALUES}, got {self.expected!r}")
        _expect(len(self.bbox_mm) == 4, "bbox_mm must be a 4-tuple (mm from origin corner)")
        _expect(self.field_type in _FIELD_TYPE_VALUES, f"FieldDef.field_type must be one of {_FIELD_TYPE_VALUES}, got {self.field_type!r}")
        _expect(self.origin in _ORIGIN_VALUES, f"FieldDef.origin must be one of {_ORIGIN_VALUES}, got {self.origin!r}")
        _expect(
            all(e in _STRETCH_EDGES for e in self.stretch_to_page),
            f"FieldDef.stretch_to_page entries must be from {_STRETCH_EDGES}, got {self.stretch_to_page!r}",
        )
        if self.document_property:
            from pdf_parsing_v2_engine.document_properties.registry import MANDATORY_PROPERTY_KEYS

            _expect(
                self.document_property in MANDATORY_PROPERTY_KEYS,
                f"FieldDef.document_property must be one of {MANDATORY_PROPERTY_KEYS}, "
                f"got {self.document_property!r}",
            )
            _expect(
                self.document_property == self.id,
                "FieldDef.document_property must equal FieldDef.id for reserved document keys",
            )


@dataclass
class StampTemplate:
    schema_version: int
    name: str
    doc_types: list[str]
    page_selector: str
    priority: int = 10
    origin: str = "frame_bottom_right"
    padding_mm: float = 0.5
    grid_adapt: bool = False
    grid_tolerance_template_mm: float = 2.0
    grid_tolerance_detected_mm: float = 0.5
    snap_max_distance_mm: float = 5.0
    max_shape_change_ratio: float = 2.5
    cascade_score_threshold: float = 0.4
    find_tables_snap_x_tolerance: float | None = None
    find_tables_snap_y_tolerance: float | None = None
    frame_mode: str = "gost"
    fields: list[FieldDef] = field(default_factory=list)
    grid_lines: list[TemplateGridLine] = field(default_factory=list)
    source_path: str = ""   # абсолютный путь к JSON-файлу шаблона (заполняется при from_json)

    def __post_init__(self) -> None:
        self.frame_mode = normalize_frame_mode(self.frame_mode)
        _expect(self.page_selector in _PAGE_SELECTOR_VALUES, f"page_selector must be one of {_PAGE_SELECTOR_VALUES}")
        _expect(self.origin in _ORIGIN_VALUES, f"origin must be one of {_ORIGIN_VALUES}")
        _expect(self.schema_version >= 1, "schema_version must be >= 1")

    @classmethod
    def from_json(cls, path: str) -> StampTemplate:
        with open(path, encoding="utf-8") as f:
            data = json.load(f)
        return cls.from_dict(data, source_path=path)

    @classmethod
    def from_dict(cls, data: dict[str, Any], source_path: str = "") -> StampTemplate:
        _expect(isinstance(data, dict), "template root must be a JSON object")
        fields_raw = data.get("fields") or []
        _expect(isinstance(fields_raw, list), "fields must be a list")
        fields: list[FieldDef] = []
        for i, fd in enumerate(fields_raw):
            _expect(isinstance(fd, dict), f"fields[{i}] must be an object")
            bbox = fd.get("bbox_mm")
            _expect(isinstance(bbox, list) and len(bbox) == 4, f"fields[{i}].bbox_mm must be [4] numbers")
            raw_origin = str(fd.get("origin", ""))
            field_origin = raw_origin if raw_origin in _ORIGIN_VALUES else "frame_bottom_right"
            raw_stretch = fd.get("stretch_to_page") or []
            stretch = tuple(s for s in raw_stretch if s in _STRETCH_EDGES)
            fields.append(
                FieldDef(
                    id=str(fd["id"]),
                    label=str(fd.get("label", "")),
                    bbox_mm=tuple(float(x) for x in bbox),
                    clean=fd.get("clean"),
                    validate_regex=fd.get("validate_regex"),
                    expected=str(fd.get("expected", "optional")),
                    padding_mm=(float(fd["padding_mm"]) if fd.get("padding_mm") is not None else None),
                    field_type=normalize_field_type(str(fd.get("field_type", "data"))),
                    expected_text=(str(fd["expected_text"]) if fd.get("expected_text") is not None else None),
                    is_anchor=bool(fd.get("is_anchor", False)),
                    outside_stamp=bool(fd.get("outside_stamp", False)),
                    origin=field_origin,
                    stretch_to_page=stretch,
                    bound_top=(str(fd["bound_top"]) if fd.get("bound_top") else None),
                    bound_bottom=(str(fd["bound_bottom"]) if fd.get("bound_bottom") else None),
                    bound_left=(str(fd["bound_left"]) if fd.get("bound_left") else None),
                    bound_right=(str(fd["bound_right"]) if fd.get("bound_right") else None),
                    document_property=(
                        str(fd["document_property"]) if fd.get("document_property") else None
                    ),
                )
            )
        gl_raw = data.get("grid_lines") or []
        _expect(isinstance(gl_raw, list), "grid_lines must be a list")
        grid_lines: list[TemplateGridLine] = []
        for i, gl in enumerate(gl_raw):
            _expect(isinstance(gl, dict), f"grid_lines[{i}] must be an object")
            try:
                grid_lines.append(TemplateGridLine.from_dict(gl))
            except (KeyError, ValueError) as e:
                raise ValueError(f"grid_lines[{i}]: {e}") from e

        def _opt_float(key: str) -> float | None:
            if key not in data or data[key] is None:
                return None
            return float(data[key])

        try:
            return cls(
                schema_version=int(data["schema_version"]),
                name=str(data.get("name", os.path.basename(source_path) or "template")),
                doc_types=[str(x) for x in data.get("doc_types", [])],
                page_selector=str(data.get("page_selector", "first")),
                priority=int(data.get("priority", 10)),
                origin=str(data.get("origin", "frame_bottom_right")),
                padding_mm=float(data.get("padding_mm", 0.5)),
                grid_adapt=bool(data.get("grid_adapt", False)),
                grid_tolerance_template_mm=float(data.get("grid_tolerance_template_mm", 2.0)),
                grid_tolerance_detected_mm=float(data.get("grid_tolerance_detected_mm", 0.5)),
                snap_max_distance_mm=float(data.get("snap_max_distance_mm", 5.0)),
                max_shape_change_ratio=float(data.get("max_shape_change_ratio", 2.5)),
                cascade_score_threshold=float(data.get("cascade_score_threshold", 0.4)),
                find_tables_snap_x_tolerance=_opt_float("find_tables_snap_x_tolerance"),
                find_tables_snap_y_tolerance=_opt_float("find_tables_snap_y_tolerance"),
                frame_mode=normalize_frame_mode(data.get("frame_mode")),
                fields=fields,
                grid_lines=grid_lines,
                source_path=str(os.path.abspath(source_path)) if source_path else "",
            )
        except KeyError as e:
            raise ValueError(f"StampTemplate missing required key: {e}") from e

    def _field_def_to_dict(self, f: FieldDef, catalog: FieldCatalog | None) -> dict[str, Any]:
        """Serialize one field. If *catalog* has an entry for f.id, omit label/field_type (catalog is source)."""
        entry = catalog.get_entry(f.id) if catalog else None
        d: dict[str, Any] = {
            "id": f.id,
            "bbox_mm": list(f.bbox_mm),
            "clean": f.clean,
            **({"validate_regex": f.validate_regex} if f.validate_regex else {}),
            "expected": f.expected,
            **({"padding_mm": f.padding_mm} if f.padding_mm is not None else {}),
            **({"expected_text": f.expected_text} if f.expected_text is not None else {}),
            **({"is_anchor": True} if f.is_anchor else {}),
            **({"outside_stamp": True} if f.outside_stamp else {}),
            **({"origin": f.origin} if f.origin != "frame_bottom_right" else {}),
            **({"stretch_to_page": list(f.stretch_to_page)} if f.stretch_to_page else {}),
            **({"bound_top": f.bound_top} if f.bound_top else {}),
            **({"bound_bottom": f.bound_bottom} if f.bound_bottom else {}),
            **({"bound_left": f.bound_left} if f.bound_left else {}),
            **({"bound_right": f.bound_right} if f.bound_right else {}),
            **({"document_property": f.document_property} if f.document_property else {}),
        }
        if entry is None:
            d["label"] = f.label
            d["field_type"] = f.field_type
        return d

    def to_dict(self, catalog: FieldCatalog | None = None) -> dict[str, Any]:
        """Optional *catalog*: when set, fields listed in the catalog omit label/field_type in JSON."""
        d: dict[str, Any] = {
            "schema_version": self.schema_version,
            "name": self.name,
            "doc_types": list(self.doc_types),
            "page_selector": self.page_selector,
            "priority": self.priority,
            "origin": self.origin,
            "padding_mm": self.padding_mm,
            "grid_adapt": self.grid_adapt,
            "grid_tolerance_template_mm": self.grid_tolerance_template_mm,
            "grid_tolerance_detected_mm": self.grid_tolerance_detected_mm,
            "snap_max_distance_mm": self.snap_max_distance_mm,
            "max_shape_change_ratio": self.max_shape_change_ratio,
            "cascade_score_threshold": self.cascade_score_threshold,
            "fields": [self._field_def_to_dict(f, catalog) for f in self.fields],
        }
        if self.frame_mode != "gost":
            d["frame_mode"] = self.frame_mode
        if self.find_tables_snap_x_tolerance is not None:
            d["find_tables_snap_x_tolerance"] = self.find_tables_snap_x_tolerance
        if self.find_tables_snap_y_tolerance is not None:
            d["find_tables_snap_y_tolerance"] = self.find_tables_snap_y_tolerance
        if self.grid_lines:
            d["grid_lines"] = [gl.to_dict() for gl in self.grid_lines]
        return d

    def to_json(self, path: str, catalog: FieldCatalog | None = None) -> None:
        os.makedirs(os.path.dirname(os.path.abspath(path)) or ".", exist_ok=True)
        with open(path, "w", encoding="utf-8") as f:
            json.dump(self.to_dict(catalog=catalog), f, ensure_ascii=False, indent=2)


@dataclass
class FrameInfo:
    """Рамка чертежа / штампа в координатах pdfminer (x вправо, y вверх от низа страницы), пункты."""

    x0: float
    y0: float
    x1: float
    y1: float
    page_width: float
    page_height: float
    rotation: int
    border_left_mm: float
    border_bottom_mm: float
    border_top_mm: float


@dataclass
class FieldResult:
    field_id: str
    raw_value: str
    cleaned_value: str | None
    bbox_pts: tuple[float, float, float, float]
    is_valid: bool | None
    parse_warnings: list[ParseWarning] = field(default_factory=list)
    clean_tier: str | None = None


@dataclass
class V2PageResult:
    page_num: int
    doc_type: str
    frame: FrameInfo
    template_name: str
    template_score: float
    fields: dict[str, FieldResult]
    metadata: dict[str, Any] = field(default_factory=dict)
    warnings: list[str] = field(default_factory=list)


@dataclass
class TemplateScore:
    template: StampTemplate
    score: float
    required_found: int
    required_total: int
    absent_violations: int
    details: list[str] = field(default_factory=list)


@dataclass
class FieldCatalogEntry:
    id: str
    label: str
    short_num: int
    description: str = ""
    group: str = ""
    default_field_type: str = "data"
    # v2 debug Excel report (see v2_report.save_v2_debug_report)
    debug_report_order: int = _DEFAULT_DEBUG_REPORT_ORDER
    include_in_debug_report: bool = True
    debug_report_header_note: str = ""
    # Catalog: this id is a PDF document-property field (hidden in template editor canvas).
    is_document_property: bool = False

    def __post_init__(self) -> None:
        _expect(
            self.default_field_type in _FIELD_TYPE_VALUES,
            f"default_field_type must be one of {_FIELD_TYPE_VALUES}, got {self.default_field_type!r}",
        )


@dataclass
class FieldCatalog:
    name: str
    projects: list[str] = field(default_factory=list)
    description: str = ""
    entries: list[FieldCatalogEntry] = field(default_factory=list)

    def __post_init__(self) -> None:
        ids = [e.id for e in self.entries]
        if len(ids) != len(set(ids)):
            dup = [i for i in ids if ids.count(i) > 1]
            raise ValueError(f"FieldCatalog: duplicate entry ids: {sorted(set(dup))!r}")

    @classmethod
    def from_json(cls, path: str) -> FieldCatalog:
        with open(path, encoding="utf-8") as f:
            data = json.load(f)
        _expect(isinstance(data, dict), "catalog root must be object")
        entries_raw = data.get("entries") or []
        entries: list[FieldCatalogEntry] = []
        for i, e in enumerate(entries_raw):
            _expect(isinstance(e, dict), f"entries[{i}] must be object")
            dro_raw = e.get("debug_report_order", _DEFAULT_DEBUG_REPORT_ORDER)
            try:
                debug_report_order = int(dro_raw)
            except (TypeError, ValueError):
                debug_report_order = _DEFAULT_DEBUG_REPORT_ORDER
            entries.append(
                FieldCatalogEntry(
                    id=str(e["id"]),
                    label=str(e.get("label", "")),
                    short_num=int(e.get("short_num", 0)),
                    description=str(e.get("description", "")),
                    group=str(e.get("group", "")),
                    default_field_type=normalize_field_type(str(e.get("default_field_type", "data"))),
                    debug_report_order=debug_report_order,
                    include_in_debug_report=bool(e.get("include_in_debug_report", True)),
                    debug_report_header_note=str(e.get("debug_report_header_note", "")),
                    is_document_property=bool(e.get("is_document_property", False)),
                )
            )
        return cls(
            name=str(data.get("name", os.path.basename(path))),
            projects=[str(x) for x in data.get("projects", [])],
            description=str(data.get("description", "")),
            entries=entries,
        )

    def to_dict(self) -> dict[str, Any]:
        d: dict[str, Any] = {"name": self.name}
        if self.projects:
            d["projects"] = list(self.projects)
        if self.description:
            d["description"] = self.description
        d["entries"] = [
            {
                "id": e.id,
                "label": e.label,
                "short_num": e.short_num,
                "description": e.description,
                "group": e.group,
                "default_field_type": e.default_field_type,
                "debug_report_order": e.debug_report_order,
                "include_in_debug_report": e.include_in_debug_report,
                "debug_report_header_note": e.debug_report_header_note,
                **({"is_document_property": True} if e.is_document_property else {}),
            }
            for e in self.entries
        ]
        return d

    def to_json(self, path: str) -> None:
        os.makedirs(os.path.dirname(os.path.abspath(path)) or ".", exist_ok=True)
        with open(path, "w", encoding="utf-8") as f:
            json.dump(self.to_dict(), f, ensure_ascii=False, indent=2)

    def get_entry(self, field_id: str) -> FieldCatalogEntry | None:
        for e in self.entries:
            if e.id == field_id:
                return e
        return None

    def all_ids(self) -> list[str]:
        return [e.id for e in self.entries]


def merge_field_defs_with_catalog(
    fields: list[FieldDef],
    catalog: FieldCatalog | None,
) -> list[FieldDef]:
    """Подставить label и field_type из каталога для совпадающих id; clean остаётся из *fields*."""
    if not catalog or not catalog.entries:
        return fields
    by_id = {e.id: e for e in catalog.entries}
    out: list[FieldDef] = []
    for fd in fields:
        e = by_id.get(fd.id)
        if e is None:
            out.append(fd)
            continue
        doc_prop = fd.document_property
        if e.is_document_property:
            doc_prop = e.id
        out.append(
            replace(
                fd,
                label=e.label,
                field_type=e.default_field_type,
                document_property=doc_prop,
            )
        )
    return out


def merge_stamp_template_with_catalog(
    tmpl: StampTemplate,
    catalog: FieldCatalog | None,
) -> StampTemplate:
    if not catalog:
        return tmpl
    return replace(
        tmpl,
        fields=merge_field_defs_with_catalog(tmpl.fields, catalog),
    )


# ---------------------------------------------------------------------------
# TemplateSet — deprecated, use project folder structure (catalog.json + templates/*.json).
# Kept for backward compatibility; will be removed in a future version.
# ---------------------------------------------------------------------------

@dataclass
class TemplateSet:
    """DEPRECATED: use project folder structure instead (catalog.json + templates in same folder)."""

    name: str
    projects: list[str]
    field_catalog: str
    template_files: list[str]

    @classmethod
    def from_json(cls, path: str) -> TemplateSet:
        with open(path, encoding="utf-8") as f:
            data = json.load(f)
        _expect(isinstance(data, dict), "TemplateSet root must be object")
        return cls(
            name=str(data.get("name", os.path.basename(path))),
            projects=[str(x) for x in data.get("projects", [])],
            field_catalog=str(data.get("field_catalog", "")),
            template_files=[str(x) for x in data.get("template_files", [])],
        )

    def to_dict(self) -> dict[str, Any]:
        return {
            "name": self.name,
            "projects": list(self.projects),
            "field_catalog": self.field_catalog,
            "template_files": list(self.template_files),
        }

    def to_json(self, path: str) -> None:
        os.makedirs(os.path.dirname(os.path.abspath(path)) or ".", exist_ok=True)
        with open(path, "w", encoding="utf-8") as f:
            json.dump(self.to_dict(), f, ensure_ascii=False, indent=2)
