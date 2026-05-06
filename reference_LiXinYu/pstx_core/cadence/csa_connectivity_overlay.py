# -*- coding: utf-8 -*-
"""Overlay CSA geometry findings with Cadence page connectivity semantics."""

from __future__ import annotations

from collections import Counter
import re
from pathlib import Path
from typing import Dict, Iterable, List, Mapping, Optional, Sequence, Tuple

from pstx_core.cadence.page_model import (
    CadenceObject,
    CadencePageModel,
    ConnectivityComponent,
    SEMANTIC_OBJECT_TYPES,
    load_cadence_page_model,
)


Point = Tuple[int, int]
BBox = Tuple[float, float, float, float]
CSA_CONNECTIVITY_OVERLAY_SCHEMA_VERSION = "pstx-csa-connectivity-overlay.v1"

PAGE_LABEL_RE = re.compile(r"PAGE\s*(\d+)", re.IGNORECASE)
COORD_RE = re.compile(r"\((-?\d+(?:\.\d+)?)\s*,\s*(-?\d+(?:\.\d+)?)\)")


def _page_no_from_row(row: Mapping[str, object]) -> int:
    raw = str(row.get("页面") or row.get("page_label") or row.get("page") or "").strip()
    match = PAGE_LABEL_RE.search(raw)
    if match:
        return int(match.group(1))
    value = row.get("page_no")
    try:
        return int(value or 0)
    except (TypeError, ValueError):
        return 0


def _wanted_page(page: Optional[int]) -> int:
    try:
        return int(page or 0)
    except (TypeError, ValueError):
        return 0


def _filter_rows_by_page(rows: Sequence[dict], page: Optional[int]) -> List[dict]:
    wanted = _wanted_page(page)
    if wanted <= 0:
        return list(rows)
    return [row for row in rows if _page_no_from_row(row) == wanted]


def _infer_project_root_from_path(path: Path) -> Optional[Path]:
    path = path.expanduser()
    probe = path.parent if path.is_file() else path
    if probe.name.lower() == "sch_1":
        return probe.parent
    if (probe / "sch_1").is_dir():
        return probe
    return None


def _row_file_path(geometry_root: str, row: Mapping[str, object]) -> Optional[Path]:
    raw_file = str(row.get("文件") or row.get("file") or "").strip()
    if not raw_file:
        return None
    path = Path(raw_file).expanduser()
    if path.is_absolute():
        return path
    root = Path(str(geometry_root or "")).expanduser()
    if not str(root):
        return path
    return root / path


def _project_root_for_row(source_root: str,
                          geometry_root: str,
                          row: Mapping[str, object]) -> Optional[Path]:
    candidates: List[Optional[Path]] = []
    row_path = _row_file_path(geometry_root, row)
    if row_path is not None:
        candidates.append(_infer_project_root_from_path(row_path))
    for raw in (source_root, geometry_root):
        if str(raw or "").strip():
            candidates.append(_infer_project_root_from_path(Path(str(raw)).expanduser()))
    for candidate in candidates:
        if candidate is not None:
            return candidate
    return None


def _model_key(root: Path, page_no: int) -> str:
    return f"{root.resolve()}::{page_no}"


def _load_model(root: Path,
                page_no: int,
                cache: Dict[str, CadencePageModel],
                warnings: List[str]) -> Optional[CadencePageModel]:
    key = _model_key(root, page_no)
    if key in cache:
        return cache[key]
    try:
        model = load_cadence_page_model(
            root,
            "project",
            page_no,
            include_raw_unknown=False,
            collect_junctions=False,
        )
    except Exception as exc:
        warnings.append(f"PAGE{page_no} 连接语义加载失败：{exc}")
        return None
    cache[key] = model
    if not model.csa_exists and not model.csv_exists:
        warnings.append(f"PAGE{page_no} 未找到可叠加的 page{page_no}.csa/csv：{root / 'sch_1'}")
    elif model.csa_error or model.csv_error:
        detail = "; ".join(item for item in [model.csa_error, model.csv_error] if item)
        warnings.append(f"PAGE{page_no} 连接语义解析受限：{detail}")
    return model


def _point_from_row(row: Mapping[str, object]) -> Optional[Point]:
    try:
        return int(row.get("X")), int(row.get("Y"))
    except (TypeError, ValueError):
        raw = str(row.get("坐标") or "").strip()
        match = COORD_RE.search(raw)
        if not match:
            return None
        return int(float(match.group(1))), int(float(match.group(2)))


def _wire_contains_point(wire: CadenceObject, point: Point) -> bool:
    if len(wire.coords) < 2:
        return False
    (x1, y1), (x2, y2) = wire.coords[0], wire.coords[1]
    px, py = point
    if y1 == y2:
        return py == y1 and min(x1, x2) <= px <= max(x1, x2)
    if x1 == x2:
        return px == x1 and min(y1, y2) <= py <= max(y1, y2)
    return False


def _component_wires(model: CadencePageModel,
                     component: ConnectivityComponent) -> List[CadenceObject]:
    ids = set(component.object_ids)
    return [
        item for item in model.objects
        if item.object_type == "WIRE" and item.object_id in ids
    ]


def _component_match_method(model: CadencePageModel,
                            component: ConnectivityComponent,
                            point: Point) -> str:
    if point in component.dot_points:
        return "dot_point"
    if point in component.junctions:
        return "junction"
    if any(_wire_contains_point(wire, point) for wire in _component_wires(model, component)):
        return "wire_contains"
    return ""


def _dedupe(values: Iterable[object]) -> List[object]:
    result: List[object] = []
    seen = set()
    for value in values:
        marker = repr(value)
        if marker in seen:
            continue
        seen.add(marker)
        result.append(value)
    return result


def _component_summary(component: ConnectivityComponent) -> dict:
    return {
        "component_id": component.component_id,
        "signal_names": list(component.signal_names),
        "labels": list(component.labels),
        "ports": list(component.ports),
        "offpage_connectors": list(component.offpage_connectors),
        "bus_names": list(component.bus_names),
        "no_connect_points": [list(point) for point in component.no_connect_points],
        "semantic_object_ids": list(component.semantic_object_ids),
        "object_ids": list(component.object_ids),
        "bbox": list(component.bbox) if component.bbox else None,
    }


def _merge_component_field(components: Sequence[ConnectivityComponent], field: str) -> List[object]:
    values: List[object] = []
    for component in components:
        values.extend(getattr(component, field, []) or [])
    return _dedupe(values)


def _dot_overlay_row(row: Mapping[str, object],
                     model: Optional[CadencePageModel],
                     point: Optional[Point],
                     root: Optional[Path]) -> dict:
    base = {
        "page": _page_no_from_row(row),
        "page_label": str(row.get("页面") or ""),
        "file": str(row.get("文件") or ""),
        "index": row.get("序号", ""),
        "coordinate": list(point) if point else None,
        "binding_status": "unmatched",
        "binding_method": "none",
        "component_id": "",
        "component_ids": [],
        "components": [],
        "signal_names": [],
        "labels": [],
        "ports": [],
        "offpage_connectors": [],
        "bus_names": [],
        "no_connect_points": [],
        "semantic_object_ids": [],
        "object_ids": [],
        "project_root": str(root) if root else "",
        "note": "未命中连接组；未做距离猜测。",
    }
    if point is None:
        base["binding_status"] = "invalid_coordinate"
        base["note"] = "几何行缺少可解析坐标。"
        return base
    if model is None:
        base["binding_status"] = "missing_page_model"
        base["note"] = "未能加载同页 Cadence 连接语义模型。"
        return base

    matches: List[Tuple[ConnectivityComponent, str]] = []
    for component in model.connectivity:
        method = _component_match_method(model, component, point)
        if method:
            matches.append((component, method))
    components = [item[0] for item in matches]
    methods = _dedupe(item[1] for item in matches)
    if not components:
        return base
    if len(components) == 1:
        base["binding_status"] = "matched"
        base["component_id"] = components[0].component_id
        base["note"] = "坐标命中同页连接组；仅作为页级 evidence。"
    else:
        base["binding_status"] = "ambiguous"
        base["note"] = "坐标命中多个连接组；需要人工复核。"
    base["binding_method"] = ",".join(str(item) for item in methods) if methods else "wire_contains"
    base["component_ids"] = [component.component_id for component in components]
    base["components"] = [_component_summary(component) for component in components]
    base["signal_names"] = _merge_component_field(components, "signal_names")
    base["labels"] = _merge_component_field(components, "labels")
    base["ports"] = _merge_component_field(components, "ports")
    base["offpage_connectors"] = _merge_component_field(components, "offpage_connectors")
    base["bus_names"] = _merge_component_field(components, "bus_names")
    base["no_connect_points"] = _dedupe(
        list(point) for component in components for point in component.no_connect_points
    )
    base["semantic_object_ids"] = _merge_component_field(components, "semantic_object_ids")
    base["object_ids"] = _merge_component_field(components, "object_ids")
    return base


def _float_value(value: object) -> Optional[float]:
    try:
        return float(str(value).strip())
    except (TypeError, ValueError):
        return None


def _bbox_from_circle_row(row: Mapping[str, object]) -> Optional[BBox]:
    raw = str(row.get("外接框") or "").strip()
    matches = COORD_RE.findall(raw)
    if len(matches) >= 2:
        x1, y1 = float(matches[0][0]), float(matches[0][1])
        x2, y2 = float(matches[1][0]), float(matches[1][1])
        return min(x1, x2), min(y1, y2), max(x1, x2), max(y1, y2)
    center = str(row.get("圆心") or "").strip()
    center_match = COORD_RE.search(center)
    radius = _float_value(row.get("半径"))
    if center_match and radius is not None:
        cx, cy = float(center_match.group(1)), float(center_match.group(2))
        return cx - radius, cy - radius, cx + radius, cy + radius
    return None


def _point_in_bbox(point: Point, bbox: BBox) -> bool:
    x, y = point
    xmin, ymin, xmax, ymax = bbox
    return xmin <= x <= xmax and ymin <= y <= ymax


def _bbox_intersects(left: Optional[Sequence[float]],
                     right: Optional[Sequence[float]]) -> bool:
    if not left or not right or len(left) < 4 or len(right) < 4:
        return False
    lx1, ly1, lx2, ly2 = [float(item) for item in left[:4]]
    rx1, ry1, rx2, ry2 = [float(item) for item in right[:4]]
    return not (lx2 < rx1 or rx2 < lx1 or ly2 < ry1 or ry2 < ly1)


def _semantic_name(item: CadenceObject) -> str:
    return str(item.attributes.get("name") or item.attributes.get("value") or "").strip()


def _semantic_in_bbox(item: CadenceObject, bbox: BBox) -> bool:
    if any(_point_in_bbox(point, bbox) for point in item.coords):
        return True
    return _bbox_intersects(item.bbox, bbox)


def _semantic_summary(item: CadenceObject) -> dict:
    return {
        "object_id": item.object_id,
        "type": item.object_type,
        "name": _semantic_name(item),
        "line_no": item.line_no,
        "coords": [list(point) for point in item.coords],
        "bbox": list(item.bbox) if item.bbox else None,
    }


def _circle_overlay_row(row: Mapping[str, object],
                        model: Optional[CadencePageModel],
                        bbox: Optional[BBox],
                        root: Optional[Path]) -> dict:
    base = {
        "page": _page_no_from_row(row),
        "page_label": str(row.get("页面") or ""),
        "file": str(row.get("文件") or ""),
        "index": row.get("序号", ""),
        "object_type": str(row.get("对象类型") or ""),
        "line_no": row.get("行号", ""),
        "center": str(row.get("圆心") or ""),
        "radius": str(row.get("半径") or ""),
        "bbox": list(bbox) if bbox else None,
        "contained_semantic_count": 0,
        "contained_semantic_types": {},
        "contained_semantic_objects": [],
        "intersecting_component_count": 0,
        "intersecting_components": [],
        "project_root": str(root) if root else "",
        "connection_claim": False,
        "note": "画圈对象仅提供范围内 evidence，不声明电气连接结论。",
    }
    if bbox is None:
        base["note"] = "画圈对象缺少可解析外接框。"
        return base
    if model is None:
        base["note"] = "未能加载同页 Cadence 连接语义模型。"
        return base

    semantics = [
        item for item in model.objects
        if item.object_type in SEMANTIC_OBJECT_TYPES and _semantic_in_bbox(item, bbox)
    ]
    components = [
        component for component in model.connectivity
        if _bbox_intersects(component.bbox, bbox)
    ]
    type_counts = Counter(item.object_type for item in semantics)
    base["contained_semantic_count"] = len(semantics)
    base["contained_semantic_types"] = dict(sorted(type_counts.items()))
    base["contained_semantic_objects"] = [_semantic_summary(item) for item in semantics]
    base["intersecting_component_count"] = len(components)
    base["intersecting_components"] = [_component_summary(component) for component in components]
    return base


def _limited_rows(rows: Sequence[dict], limit: int) -> Tuple[List[dict], bool]:
    limit = max(1, min(5000, int(limit or 200)))
    return list(rows[:limit]), len(rows) > limit


def _unique_warnings(warnings: Sequence[str]) -> List[str]:
    return [str(item) for item in _dedupe(str(warning) for warning in warnings if warning)]


def build_csa_connectivity_overlay(
    csa_geometry: Mapping[str, object],
    *,
    source_root: str = "",
    page: Optional[int] = None,
    stdout: str = "summary",
    limit: int = 200,
) -> dict:
    """Build an opt-in semantic overlay for CSA geometry rows."""

    mode = str(stdout or "summary").strip().lower()
    if mode not in {"summary", "hits", "details", "full"}:
        mode = "summary"
    geometry_root = str(csa_geometry.get("root") or "")
    dot_rows = _filter_rows_by_page(list(csa_geometry.get("dot_cross_rows", []) or []), page)
    circle_rows = _filter_rows_by_page(list(csa_geometry.get("circle_rows", []) or []), page)
    warnings: List[str] = []
    model_cache: Dict[str, CadencePageModel] = {}

    def model_for(row: Mapping[str, object]) -> Tuple[Optional[CadencePageModel], Optional[Path]]:
        page_no = _page_no_from_row(row)
        if page_no <= 0:
            warnings.append("CSA 几何行缺少可解析页码，无法叠加连接语义。")
            return None, None
        root = _project_root_for_row(source_root, geometry_root, row)
        if root is None:
            warnings.append(f"PAGE{page_no} 无法从输入路径推断 project root，已保留几何结果。")
            return None, None
        return _load_model(root, page_no, model_cache, warnings), root

    dot_overlay_rows: List[dict] = []
    for row in dot_rows:
        model, root = model_for(row)
        dot_overlay_rows.append(_dot_overlay_row(row, model, _point_from_row(row), root))

    circle_overlay_rows: List[dict] = []
    for row in circle_rows:
        model, root = model_for(row)
        circle_overlay_rows.append(_circle_overlay_row(row, model, _bbox_from_circle_row(row), root))

    status_counts = Counter(row["binding_status"] for row in dot_overlay_rows)
    digest = {
        "schema_version": CSA_CONNECTIVITY_OVERLAY_SCHEMA_VERSION,
        "enabled": True,
        "page_filter": _wanted_page(page),
        "dot_cross_count": len(dot_overlay_rows),
        "dot_cross_matched_count": int(status_counts.get("matched", 0) or 0),
        "dot_cross_ambiguous_count": int(status_counts.get("ambiguous", 0) or 0),
        "dot_cross_unmatched_count": int(status_counts.get("unmatched", 0) or 0),
        "circle_count": len(circle_overlay_rows),
        "circle_with_semantics_count": sum(
            1 for row in circle_overlay_rows
            if int(row.get("contained_semantic_count", 0) or 0) > 0
        ),
        "circle_with_components_count": sum(
            1 for row in circle_overlay_rows
            if int(row.get("intersecting_component_count", 0) or 0) > 0
        ),
        "warning_count": len(_unique_warnings(warnings)),
    }
    payload = {
        "schema_version": CSA_CONNECTIVITY_OVERLAY_SCHEMA_VERSION,
        "digest": digest,
        "warnings": _unique_warnings(warnings),
        "dot_cross_overlay_rows": [],
        "circle_overlay_rows": [],
        "truncated": False,
        "truncation": {
            "dot_cross_overlay_rows": False,
            "circle_overlay_rows": False,
        },
    }
    if mode in {"hits", "details", "full"}:
        payload["dot_cross_overlay_rows"], payload["truncation"]["dot_cross_overlay_rows"] = _limited_rows(
            dot_overlay_rows,
            limit,
        )
    if mode in {"hits", "full"}:
        payload["circle_overlay_rows"], payload["truncation"]["circle_overlay_rows"] = _limited_rows(
            circle_overlay_rows,
            limit,
        )
    payload["truncated"] = any(payload["truncation"].values())
    return payload
