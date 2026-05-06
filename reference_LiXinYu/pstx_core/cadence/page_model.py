# -*- coding: utf-8 -*-
"""Cadence DE HDL page-level semantic model for compare harness tools."""

from __future__ import annotations

import hashlib
import json
import re
from bisect import bisect_left, bisect_right
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Sequence, Set, Tuple

from pstx_core.cadence.csa_geometry import (
    ARC_LINE_RE,
    CIRCLE_LINE_RE,
    DOT_RE,
    PAGE_NUMBER_RE,
    SIG_NAME_RE,
    WIRE_RE,
    parse_arc_line_as_circle,
    parse_circle_line,
)


Point = Tuple[int, int]
CADENCE_PAGE_SCHEMA_VERSION = "pstx-cadence-page.v1"

PROPERTY_RE = re.compile(
    r"^\s*['\"]?([A-Za-z_][A-Za-z0-9_ ]{0,80})['\"]?\s*=\s*(.+?)\s*;?\s*$"
)
TEXT_COMMANDS = {"TEXT", "NOTE", "ANNOTATE", "DISPLAY", "COMMENT"}
SEMANTIC_OBJECT_TYPES = {"NET_LABEL", "PORT", "OFFPAGE", "BUS", "NO_CONNECT"}
NET_LABEL_COMMANDS = {"LABEL", "NET_LABEL", "NETLABEL", "NET_NAME", "NETNAME", "SIGNAL_LABEL"}
PORT_COMMANDS = {"PORT", "INPORT", "OUTPORT", "IOPORT", "INPUT_PORT", "OUTPUT_PORT", "BIDIR_PORT"}
OFFPAGE_COMMANDS = {"OFFPAGE", "OFF_PAGE", "OFFPAGE_CONNECTOR", "OFFPAGECONN", "OFFSHEET", "OFF_SHEET"}
BUS_COMMANDS = {"BUS", "BUS_LABEL", "BUSLABEL", "BUS_NAME", "BUSNAME"}
NO_CONNECT_COMMANDS = {"NO_CONNECT", "NOCONNECT", "NC"}
SEMANTIC_KEY_TO_TYPE = {
    "NET_NAME": "NET_LABEL",
    "NET_LABEL": "NET_LABEL",
    "SIGNAL_NAME": "NET_LABEL",
    "SIG_NAME": "NET_LABEL",
    "PORT_NAME": "PORT",
    "OFFPAGE_NAME": "OFFPAGE",
    "OFF_PAGE_NAME": "OFFPAGE",
    "BUS_NAME": "BUS",
    "BUS_LABEL": "BUS",
    "NO_CONNECT": "NO_CONNECT",
    "NOCONNECT": "NO_CONNECT",
}
COORD_PAIR_RE = re.compile(r"\((-?\d+)\s*,?\s*(-?\d+)\)")
DIRECTION_TOKENS = {"IN", "OUT", "INPUT", "OUTPUT", "INOUT", "BIDIR", "BI"}


@dataclass
class CadenceObject:
    object_id: str
    object_type: str
    line_no: int
    raw: str
    coords: List[Point] = field(default_factory=list)
    bbox: Optional[Tuple[int, int, int, int]] = None
    attributes: Dict[str, object] = field(default_factory=dict)
    fingerprint: str = ""

    def to_dict(self, *, include_raw: bool = True) -> dict:
        payload = {
            "object_id": self.object_id,
            "type": self.object_type,
            "line_no": self.line_no,
            "coords": [list(point) for point in self.coords],
            "bbox": list(self.bbox) if self.bbox else None,
            "attributes": dict(self.attributes),
            "fingerprint": self.fingerprint,
        }
        if include_raw:
            payload["raw"] = self.raw
        return payload


@dataclass
class ConnectivityComponent:
    component_id: str
    object_ids: List[str]
    signal_names: List[str]
    bbox: Optional[Tuple[int, int, int, int]]
    junctions: List[Point]
    dot_points: List[Point]
    wire_count: int
    dot_count: int
    fingerprint: str
    semantic_object_ids: List[str] = field(default_factory=list)
    labels: List[str] = field(default_factory=list)
    ports: List[str] = field(default_factory=list)
    offpage_connectors: List[str] = field(default_factory=list)
    bus_names: List[str] = field(default_factory=list)
    no_connect_points: List[Point] = field(default_factory=list)
    unbound_semantics: List[dict] = field(default_factory=list)

    def to_dict(self) -> dict:
        return {
            "component_id": self.component_id,
            "object_ids": list(self.object_ids),
            "signal_names": list(self.signal_names),
            "semantic_object_ids": list(self.semantic_object_ids),
            "labels": list(self.labels),
            "ports": list(self.ports),
            "offpage_connectors": list(self.offpage_connectors),
            "bus_names": list(self.bus_names),
            "bbox": list(self.bbox) if self.bbox else None,
            "junctions": [list(point) for point in self.junctions],
            "dot_points": [list(point) for point in self.dot_points],
            "no_connect_points": [list(point) for point in self.no_connect_points],
            "unbound_semantics": list(self.unbound_semantics),
            "wire_count": self.wire_count,
            "dot_count": self.dot_count,
            "fingerprint": self.fingerprint,
        }


@dataclass
class CadencePageModel:
    side: str
    page_no: int
    csv_path: str
    csa_path: str
    csv_exists: bool
    csa_exists: bool
    page_number: str = ""
    csv_encoding: str = ""
    csa_encoding: str = ""
    csv_error: str = ""
    csa_error: str = ""
    objects: List[CadenceObject] = field(default_factory=list)
    connectivity: List[ConnectivityComponent] = field(default_factory=list)
    unbound_semantics: List[CadenceObject] = field(default_factory=list)
    csv_properties: Dict[str, str] = field(default_factory=dict)
    csv_rows: List[str] = field(default_factory=list)

    def object_by_id(self, object_id: str) -> Optional[CadenceObject]:
        for item in self.objects:
            if item.object_id == object_id:
                return item
        return None

    def connectivity_by_id(self, component_id: str) -> Optional[ConnectivityComponent]:
        for item in self.connectivity:
            if item.component_id == component_id:
                return item
        return None

    def counts(self) -> dict:
        result: Dict[str, int] = {}
        for item in self.objects:
            result[item.object_type] = result.get(item.object_type, 0) + 1
        result["CONNECTIVITY"] = len(self.connectivity)
        result["CSV_PROPERTY"] = len(self.csv_properties)
        return result

    def digest(self) -> dict:
        return {
            "side": self.side,
            "page": self.page_no,
            "page_label": f"PAGE{self.page_no}",
            "page_number": self.page_number,
            "csv_exists": self.csv_exists,
            "csa_exists": self.csa_exists,
            "csv_path": self.csv_path,
            "csa_path": self.csa_path,
            "csv_error": self.csv_error,
            "csa_error": self.csa_error,
            "counts": self.counts(),
            "object_count": len(self.objects),
            "connectivity_count": len(self.connectivity),
            "unbound_semantic_count": len(self.unbound_semantics),
            "csv_property_count": len(self.csv_properties),
        }

    def object_preview(self, limit: int = 12) -> List[dict]:
        return [item.to_dict(include_raw=False) for item in self.objects[:max(0, limit)]]

    def connectivity_summary(self) -> dict:
        counts = self.counts()
        semantic_counts = {key: int(counts.get(key, 0) or 0) for key in sorted(SEMANTIC_OBJECT_TYPES)}
        bound_semantic_ids: Set[str] = set()
        for item in self.connectivity:
            bound_semantic_ids.update(item.semantic_object_ids)
        return {
            "schema_version": CADENCE_PAGE_SCHEMA_VERSION,
            "page": self.page_no,
            "page_label": f"PAGE{self.page_no}",
            "page_number": self.page_number,
            "wire_count": int(counts.get("WIRE", 0) or 0),
            "dot_count": int(counts.get("DOT", 0) or 0),
            "connectivity_count": len(self.connectivity),
            "semantic_counts": semantic_counts,
            "bound_semantic_count": len(bound_semantic_ids),
            "unbound_semantic_count": len(self.unbound_semantics),
            "unknown_count": int(counts.get("UNKNOWN", 0) or 0),
            "csv_property_count": len(self.csv_properties),
            "status": _model_status(self),
        }

    def to_trace_dict(self) -> dict:
        return {
            **self.digest(),
            "csv_encoding": self.csv_encoding,
            "csa_encoding": self.csa_encoding,
            "csv_properties": dict(self.csv_properties),
            "csv_rows": list(self.csv_rows[:200]),
            "objects": [item.to_dict(include_raw=True) for item in self.objects],
            "connectivity": [item.to_dict() for item in self.connectivity],
            "unbound_semantics": [item.to_dict(include_raw=True) for item in self.unbound_semantics],
        }


def _model_status(model: CadencePageModel) -> str:
    if model.csv_error or model.csa_error:
        return "parse_limited"
    if not model.csv_exists and not model.csa_exists:
        return "missing"
    if model.csa_exists and not model.connectivity and not model.objects:
        return "empty"
    return "ok"


class _DSU:
    def __init__(self, ids: Iterable[str]) -> None:
        self.parent = {item: item for item in ids}

    def find(self, item: str) -> str:
        while self.parent[item] != item:
            self.parent[item] = self.parent[self.parent[item]]
            item = self.parent[item]
        return item

    def union(self, left: str, right: str) -> None:
        left_root = self.find(left)
        right_root = self.find(right)
        if left_root != right_root:
            self.parent[right_root] = left_root


@dataclass(frozen=True)
class _IndexedWire:
    item: CadenceObject
    object_id: str
    x1: int
    y1: int
    x2: int
    y2: int
    xmin: int
    xmax: int
    ymin: int
    ymax: int
    is_h: bool
    is_v: bool


def _indexed_wire(item: CadenceObject) -> Optional[_IndexedWire]:
    if len(item.coords) < 2:
        return None
    (x1, y1), (x2, y2) = item.coords[0], item.coords[1]
    return _IndexedWire(
        item=item,
        object_id=item.object_id,
        x1=x1,
        y1=y1,
        x2=x2,
        y2=y2,
        xmin=min(x1, x2),
        xmax=max(x1, x2),
        ymin=min(y1, y2),
        ymax=max(y1, y2),
        is_h=y1 == y2 and x1 != x2,
        is_v=x1 == x2 and y1 != y2,
    )


def _build_wire_indexes(wires: Sequence[CadenceObject]) -> Tuple[
    List[_IndexedWire],
    Dict[int, List[_IndexedWire]],
    Dict[int, List[_IndexedWire]],
    List[int],
]:
    indexed: List[_IndexedWire] = []
    horizontals_by_y: Dict[int, List[_IndexedWire]] = {}
    verticals_by_x: Dict[int, List[_IndexedWire]] = {}
    for wire in wires:
        entry = _indexed_wire(wire)
        if entry is None:
            continue
        indexed.append(entry)
        if entry.is_h:
            horizontals_by_y.setdefault(entry.y1, []).append(entry)
        elif entry.is_v:
            verticals_by_x.setdefault(entry.x1, []).append(entry)
    vertical_xs = sorted(verticals_by_x)
    return indexed, horizontals_by_y, verticals_by_x, vertical_xs


def _indexed_wires_touching_point(
    point: Point,
    horizontals_by_y: Dict[int, List[_IndexedWire]],
    verticals_by_x: Dict[int, List[_IndexedWire]],
) -> List[_IndexedWire]:
    px, py = point
    touching: List[_IndexedWire] = []
    for wire in horizontals_by_y.get(py, []):
        if wire.xmin <= px <= wire.xmax:
            touching.append(wire)
    for wire in verticals_by_x.get(px, []):
        if wire.ymin <= py <= wire.ymax:
            touching.append(wire)
    return touching


def _decode_bytes(data: bytes) -> Tuple[str, str]:
    for encoding in ("utf-8-sig", "utf-16", "gb18030", "latin-1"):
        try:
            return data.decode(encoding), encoding
        except UnicodeDecodeError:
            continue
    return data.decode("utf-8", errors="replace"), "utf-8-replace"


def _stable_hash(value) -> str:
    payload = json.dumps(value, ensure_ascii=False, sort_keys=True, default=str)
    return hashlib.sha1(payload.encode("utf-8")).hexdigest()[:16]


def _bbox(points: Sequence[Point]) -> Optional[Tuple[int, int, int, int]]:
    if not points:
        return None
    xs = [point[0] for point in points]
    ys = [point[1] for point in points]
    return min(xs), min(ys), max(xs), max(ys)


def _norm_coord(value: int, tolerance: int) -> int:
    tolerance = max(0, int(tolerance or 0))
    if tolerance <= 0:
        return int(value)
    return int(round(int(value) / tolerance) * tolerance)


def _norm_points(points: Sequence[Point], tolerance: int) -> List[Point]:
    return [(_norm_coord(x, tolerance), _norm_coord(y, tolerance)) for x, y in points]


def _line_command(raw: str) -> str:
    match = re.match(r"^\s*([A-Za-z_][A-Za-z0-9_]*)", raw or "")
    return match.group(1).upper() if match else ""


def _clean_value(value: object) -> str:
    text = str(value or "").strip().strip(";").strip()
    if len(text) >= 2 and text[0] == text[-1] and text[0] in {"'", '"'}:
        text = text[1:-1].strip()
    return text


def _extract_points(raw: str) -> List[Point]:
    return [(int(x), int(y)) for x, y in COORD_PAIR_RE.findall(str(raw or ""))]


def _strip_coords(raw: str) -> str:
    return COORD_PAIR_RE.sub(" ", str(raw or ""))


def _semantic_type_for(raw: str, prop_match: Optional[re.Match] = None) -> str:
    command = _line_command(raw)
    if command in NET_LABEL_COMMANDS:
        return "NET_LABEL"
    if command in PORT_COMMANDS:
        return "PORT"
    if command in OFFPAGE_COMMANDS:
        return "OFFPAGE"
    if command in BUS_COMMANDS:
        return "BUS"
    if command in NO_CONNECT_COMMANDS:
        return "NO_CONNECT"
    if prop_match:
        key = prop_match.group(1).strip().upper().replace(" ", "_")
        return SEMANTIC_KEY_TO_TYPE.get(key, "")
    return ""


def _semantic_name(raw: str, object_type: str, prop_match: Optional[re.Match] = None) -> Tuple[str, str]:
    if prop_match:
        return _clean_value(prop_match.group(2)), ""

    command = _line_command(raw)
    quoted = [item.strip() for item in re.findall(r"['\"]([^'\"]+)['\"]", str(raw or "")) if item.strip()]
    if quoted:
        return quoted[-1].rstrip(";").strip(), ""

    text = _strip_coords(raw)
    text = re.sub(r";.*$", "", text).replace(":", " ")
    tokens = [token.strip().strip(",").strip("'\"") for token in text.split() if token.strip()]
    if tokens and tokens[0].upper() == command:
        tokens = tokens[1:]
    filtered: List[str] = []
    direction = ""
    for token in tokens:
        upper = token.upper()
        if re.fullmatch(r"[-+]?\d+(?:\.\d+)?", token):
            continue
        if upper in {"J", "LAST"}:
            continue
        if upper in DIRECTION_TOKENS:
            direction = upper
            continue
        filtered.append(token)
    if object_type == "NO_CONNECT" and not filtered:
        return "NO_CONNECT", direction
    return (filtered[-1] if filtered else ""), direction


def _semantic_attrs(raw: str, object_type: str, prop_match: Optional[re.Match] = None) -> Dict[str, object]:
    name, direction = _semantic_name(raw, object_type, prop_match)
    attrs: Dict[str, object] = {
        "semantic_kind": object_type.lower(),
        "command": _line_command(raw),
        "name": name,
    }
    if direction:
        attrs["direction"] = direction
    if prop_match:
        attrs["key"] = prop_match.group(1).strip()
        attrs["value"] = _clean_value(prop_match.group(2))
    return attrs


def _parse_csv(path: Path) -> Tuple[Dict[str, str], List[str], str, str, str]:
    if not path.is_file():
        return {}, [], "", "", ""
    try:
        text, encoding = _decode_bytes(path.read_bytes())
    except OSError as exc:
        return {}, [], "", "", str(exc)
    properties: Dict[str, str] = {}
    rows: List[str] = []
    page_number = ""
    for raw_line in text.splitlines():
        line = raw_line.strip()
        if not line:
            continue
        rows.append(line)
        page_match = PAGE_NUMBER_RE.search(line)
        if page_match and not page_number:
            page_number = page_match.group(1).upper()
        prop_match = PROPERTY_RE.match(line)
        if prop_match:
            key = prop_match.group(1).strip().strip('"').strip("'")
            value = prop_match.group(2).strip().strip(";").strip().strip('"').strip("'")
            properties[key] = value
    return properties, rows, page_number, encoding, ""


def _object_fingerprint(item: CadenceObject, coordinate_tolerance: int) -> str:
    coords = _norm_points(item.coords, coordinate_tolerance)
    attrs = dict(item.attributes)
    attrs.pop("raw", None)
    return _stable_hash({
        "type": item.object_type,
        "coords": sorted(coords),
        "attrs": attrs,
        "raw": item.raw.strip() if item.object_type == "UNKNOWN" else "",
    })


def _wire_contains_point(wire: CadenceObject, point: Point) -> bool:
    coords = wire.coords
    if len(coords) < 2:
        return False
    (x1, y1), (x2, y2) = coords[0], coords[1]
    px, py = point
    if y1 == y2:
        return py == y1 and min(x1, x2) <= px <= max(x1, x2)
    if x1 == x2:
        return px == x1 and min(y1, y2) <= py <= max(y1, y2)
    return False


def _wire_intersection(left: CadenceObject, right: CadenceObject) -> Optional[Point]:
    if left.object_type != "WIRE" or right.object_type != "WIRE":
        return None
    if len(left.coords) < 2 or len(right.coords) < 2:
        return None
    (lx1, ly1), (lx2, ly2) = left.coords[0], left.coords[1]
    (rx1, ry1), (rx2, ry2) = right.coords[0], right.coords[1]
    left_h = ly1 == ly2 and lx1 != lx2
    left_v = lx1 == lx2 and ly1 != ly2
    right_h = ry1 == ry2 and rx1 != rx2
    right_v = rx1 == rx2 and ry1 != ry2
    if left_h and right_v:
        point = (rx1, ly1)
    elif left_v and right_h:
        point = (lx1, ry1)
    else:
        return None
    if _wire_contains_point(left, point) and _wire_contains_point(right, point):
        return point
    return None


def _semantic_name_for(item: CadenceObject) -> str:
    return str(item.attributes.get("name") or item.attributes.get("value") or "").strip()


def _semantic_points_for(items: Sequence[CadenceObject]) -> List[Point]:
    points: List[Point] = []
    for item in items:
        points.extend(item.coords)
    return points


def _semantic_objects(objects: Sequence[CadenceObject]) -> List[CadenceObject]:
    return [item for item in objects if item.object_type in SEMANTIC_OBJECT_TYPES]


def _touching_semantics(group: Sequence[CadenceObject],
                        semantics: Sequence[CadenceObject]) -> List[CadenceObject]:
    result: List[CadenceObject] = []
    for semantic in semantics:
        if not semantic.coords:
            continue
        if any(_wire_contains_point(wire, point) for wire in group for point in semantic.coords):
            result.append(semantic)
    return result


def _semantic_names(items: Sequence[CadenceObject], object_type: str) -> List[str]:
    return sorted({
        _semantic_name_for(item)
        for item in items
        if item.object_type == object_type and _semantic_name_for(item)
    })


def _build_connectivity(
    objects: Sequence[CadenceObject],
    *,
    collect_junctions: bool = True,
) -> List[ConnectivityComponent]:
    wires = [item for item in objects if item.object_type == "WIRE"]
    dots = [item for item in objects if item.object_type == "DOT" and item.coords]
    semantics = _semantic_objects(objects)
    if not wires:
        return []
    dsu = _DSU(item.object_id for item in wires)
    _, horizontals_by_y, verticals_by_x, vertical_xs = _build_wire_indexes(wires)
    endpoint_map: Dict[Point, List[str]] = {}
    for wire in wires:
        for point in wire.coords[:2]:
            endpoint_map.setdefault(point, []).append(wire.object_id)
    for ids in endpoint_map.values():
        for other in ids[1:]:
            dsu.union(ids[0], other)

    dot_touching_ids: Dict[int, List[str]] = {}
    for dot in dots:
        touching = [
            wire.object_id
            for wire in _indexed_wires_touching_point(dot.coords[0], horizontals_by_y, verticals_by_x)
        ]
        dot_touching_ids[id(dot)] = touching
        for other in touching[1:]:
            dsu.union(touching[0], other)

    intersection_pairs: List[Tuple[Point, str, str]] = []
    for horizontals in horizontals_by_y.values():
        for horizontal in horizontals:
            left = bisect_left(vertical_xs, horizontal.xmin)
            right = bisect_right(vertical_xs, horizontal.xmax)
            for x in vertical_xs[left:right]:
                for vertical in verticals_by_x.get(x, []):
                    if vertical.ymin <= horizontal.y1 <= vertical.ymax:
                        dsu.union(horizontal.object_id, vertical.object_id)
                        if collect_junctions:
                            intersection_pairs.append((
                                (vertical.x1, horizontal.y1),
                                horizontal.object_id,
                                vertical.object_id,
                            ))

    semantic_touching_ids: Dict[str, Set[str]] = {}
    for semantic in semantics:
        touching: Set[str] = set()
        for point in semantic.coords:
            touching.update(
                wire.object_id
                for wire in _indexed_wires_touching_point(point, horizontals_by_y, verticals_by_x)
            )
        semantic_touching_ids[semantic.object_id] = touching

    junctions_by_root: Dict[str, Set[Point]] = {}
    if collect_junctions:
        for point, left_id, _right_id in intersection_pairs:
            junctions_by_root.setdefault(dsu.find(left_id), set()).add(point)

    groups: Dict[str, List[CadenceObject]] = {}
    for wire in wires:
        groups.setdefault(dsu.find(wire.object_id), []).append(wire)
    components: List[ConnectivityComponent] = []
    for group_index, group in enumerate(sorted(groups.values(), key=lambda values: values[0].object_id), start=1):
        root = dsu.find(group[0].object_id)
        group_semantics = [
            semantic for semantic in semantics
            if any(dsu.find(wire_id) == root for wire_id in semantic_touching_ids.get(semantic.object_id, set()))
        ]
        object_ids = sorted(item.object_id for item in group)
        semantic_object_ids = sorted(item.object_id for item in group_semantics)
        all_points: List[Point] = []
        for wire in group:
            all_points.extend(wire.coords[:2])
        dot_points = [
            dot.coords[0] for dot in dots
            if any(dsu.find(wire_id) == root for wire_id in dot_touching_ids.get(id(dot), []))
        ]
        junctions = sorted(junctions_by_root.get(root, set())) if collect_junctions else []
        signal_names = sorted({
            str(wire.attributes.get("sig_name") or "").strip()
            for wire in group
            if str(wire.attributes.get("sig_name") or "").strip()
        })
        labels = _semantic_names(group_semantics, "NET_LABEL")
        ports = _semantic_names(group_semantics, "PORT")
        offpage_connectors = _semantic_names(group_semantics, "OFFPAGE")
        bus_names = _semantic_names(group_semantics, "BUS")
        no_connect_points = sorted({
            point
            for item in group_semantics
            if item.object_type == "NO_CONNECT"
            for point in item.coords
        })
        fingerprint = _stable_hash({
            "signals": signal_names,
            "labels": labels,
            "ports": ports,
            "offpage_connectors": offpage_connectors,
            "bus_names": bus_names,
            "no_connect_points": no_connect_points,
            "wire_segments": sorted([
                sorted(_norm_points(wire.coords[:2], 0))
                for wire in group
            ]),
            "dots": sorted(dot_points),
        })
        components.append(ConnectivityComponent(
            component_id=f"conn-{group_index}",
            object_ids=object_ids + semantic_object_ids,
            signal_names=signal_names,
            bbox=_bbox(all_points + dot_points + _semantic_points_for(group_semantics)),
            junctions=junctions,
            dot_points=sorted(dot_points),
            wire_count=len(group),
            dot_count=len(dot_points),
            fingerprint=fingerprint,
            semantic_object_ids=semantic_object_ids,
            labels=labels,
            ports=ports,
            offpage_connectors=offpage_connectors,
            bus_names=bus_names,
            no_connect_points=no_connect_points,
        ))
    return components


def _unbound_semantics(objects: Sequence[CadenceObject],
                       connectivity: Sequence[ConnectivityComponent]) -> List[CadenceObject]:
    bound_ids: Set[str] = set()
    for component in connectivity:
        bound_ids.update(component.semantic_object_ids)
    return [item for item in _semantic_objects(objects) if item.object_id not in bound_ids]


def _parse_csa_objects(text: str,
                       page_no: int,
                       *,
                       coordinate_tolerance: int = 0,
                       include_raw_unknown: bool = True) -> Tuple[List[CadenceObject], str]:
    objects: List[CadenceObject] = []
    page_number = ""
    last_wire: Optional[CadenceObject] = None
    counters: Dict[str, int] = {}

    def add_object(object_type: str,
                   line_no: int,
                   raw: str,
                   coords: Optional[List[Point]] = None,
                   attributes: Optional[Dict[str, object]] = None) -> CadenceObject:
        counters[object_type] = counters.get(object_type, 0) + 1
        object_id = f"p{page_no}-{object_type.lower()}-{counters[object_type]}"
        item = CadenceObject(
            object_id=object_id,
            object_type=object_type,
            line_no=line_no,
            raw=raw,
            coords=list(coords or []),
            bbox=_bbox(coords or []),
            attributes=dict(attributes or {}),
        )
        item.fingerprint = _object_fingerprint(item, coordinate_tolerance)
        objects.append(item)
        return item

    for line_no, raw_line in enumerate(str(text or "").splitlines(), start=1):
        raw = raw_line.strip()
        if not raw:
            continue
        page_match = PAGE_NUMBER_RE.search(raw)
        if page_match and not page_number:
            page_number = page_match.group(1).upper()

        matched = False
        events = []
        for match in WIRE_RE.finditer(raw):
            events.append((match.start(), "WIRE", match))
        for match in DOT_RE.finditer(raw):
            events.append((match.start(), "DOT", match))
        for match in SIG_NAME_RE.finditer(raw):
            events.append((match.start(), "SIG", match))
        events.sort(key=lambda item: item[0])

        for _, kind, match in events:
            matched = True
            if kind == "WIRE":
                x1, y1, x2, y2 = map(int, match.groups())
                last_wire = add_object("WIRE", line_no, match.group(0).strip(), [(x1, y1), (x2, y2)])
            elif kind == "DOT":
                x, y = map(int, match.groups()[-2:])
                add_object("DOT", line_no, match.group(0).strip(), [(x, y)])
            elif kind == "SIG" and last_wire is not None:
                sig_name = match.group(1).strip().rstrip(";").strip()
                last_wire.attributes["sig_name"] = sig_name
                last_wire.fingerprint = _object_fingerprint(last_wire, coordinate_tolerance)

        if CIRCLE_LINE_RE.match(raw):
            circle = parse_circle_line(raw, line_no)
            if circle:
                matched = True
                add_object(
                    "CIRCLE",
                    line_no,
                    raw,
                    [
                        (int(round(circle.bbox_xmin)), int(round(circle.bbox_ymin))),
                        (int(round(circle.bbox_xmax)), int(round(circle.bbox_ymax))),
                    ],
                    {
                        "center": [circle.center_x, circle.center_y],
                        "radius": circle.radius,
                        "parse_note": circle.parse_note,
                    },
                )
        elif ARC_LINE_RE.match(raw):
            circle = parse_arc_line_as_circle(raw, line_no)
            if circle:
                matched = True
                add_object(
                    "ARC",
                    line_no,
                    raw,
                    [
                        (int(round(circle.bbox_xmin)), int(round(circle.bbox_ymin))),
                        (int(round(circle.bbox_xmax)), int(round(circle.bbox_ymax))),
                    ],
                    {
                        "center": [circle.center_x, circle.center_y],
                        "radius": circle.radius,
                        "parse_note": circle.parse_note,
                    },
                )

        prop_match = PROPERTY_RE.match(raw)
        semantic_type = _semantic_type_for(raw, prop_match)
        if semantic_type:
            matched = True
            add_object(
                semantic_type,
                line_no,
                raw,
                _extract_points(raw),
                _semantic_attrs(raw, semantic_type, prop_match),
            )

        if prop_match:
            matched = True
            add_object(
                "PROPERTY",
                line_no,
                raw,
                [],
                {"key": prop_match.group(1).strip(), "value": prop_match.group(2).strip().rstrip(";").strip()},
            )
        command = _line_command(raw)
        if command in TEXT_COMMANDS:
            matched = True
            add_object("TEXT", line_no, raw, [], {"command": command})
        if not matched and include_raw_unknown:
            add_object("UNKNOWN", line_no, raw, [], {"command": command})
    return objects, page_number


def load_cadence_page_model(project_root: str | Path,
                            side: str,
                            page_no: int,
                            *,
                            coordinate_tolerance: int = 0,
                            include_raw_unknown: bool = True,
                            collect_junctions: bool = True) -> CadencePageModel:
    root = Path(project_root).expanduser().resolve()
    sch_dir = root / "sch_1"
    csv_path = sch_dir / f"page{int(page_no)}.csv"
    csa_path = sch_dir / f"page{int(page_no)}.csa"
    csv_properties, csv_rows, csv_page_number, csv_encoding, csv_error = _parse_csv(csv_path)
    csa_encoding = ""
    csa_error = ""
    csa_page_number = ""
    objects: List[CadenceObject] = []
    if csa_path.is_file():
        try:
            csa_text, csa_encoding = _decode_bytes(csa_path.read_bytes())
            objects, csa_page_number = _parse_csa_objects(
                csa_text,
                int(page_no),
                coordinate_tolerance=coordinate_tolerance,
                include_raw_unknown=include_raw_unknown,
            )
        except Exception as exc:
            csa_error = str(exc)
    connectivity = _build_connectivity(objects, collect_junctions=collect_junctions)
    unbound_semantics = _unbound_semantics(objects, connectivity)
    return CadencePageModel(
        side=side,
        page_no=int(page_no),
        csv_path=str(csv_path),
        csa_path=str(csa_path),
        csv_exists=csv_path.is_file(),
        csa_exists=csa_path.is_file(),
        page_number=csa_page_number or csv_page_number,
        csv_encoding=csv_encoding,
        csa_encoding=csa_encoding,
        csv_error=csv_error,
        csa_error=csa_error,
        objects=objects,
        connectivity=connectivity,
        unbound_semantics=unbound_semantics,
        csv_properties=csv_properties,
        csv_rows=csv_rows,
    )


def _page_file_number(path: Path) -> int:
    match = re.match(r"^page(\d+)\.(?:csv|csa)$", path.name, re.IGNORECASE)
    return int(match.group(1)) if match else 0


def iter_cadence_page_numbers(project_root: str | Path) -> List[int]:
    root = Path(project_root).expanduser()
    sch_dir = root / "sch_1"
    if not sch_dir.is_dir():
        return []
    pages = {
        number
        for path in sch_dir.iterdir()
        for number in [_page_file_number(path)]
        if path.is_file() and number > 0
    }
    return sorted(pages)


def cadence_connectivity_row(model: CadencePageModel) -> Dict[str, object]:
    summary = model.connectivity_summary()
    semantic_counts = summary.get("semantic_counts") or {}
    return {
        "页码": f"PAGE{model.page_no}",
        "PAGE_NUMBER": model.page_number,
        "WIRE": summary.get("wire_count", 0),
        "DOT": summary.get("dot_count", 0),
        "连接组": summary.get("connectivity_count", 0),
        "网络标签": semantic_counts.get("NET_LABEL", 0),
        "端口": semantic_counts.get("PORT", 0),
        "跨页连接": semantic_counts.get("OFFPAGE", 0),
        "Bus": semantic_counts.get("BUS", 0),
        "No Connect": semantic_counts.get("NO_CONNECT", 0),
        "未绑定语义": summary.get("unbound_semantic_count", 0),
        "未知行": summary.get("unknown_count", 0),
        "解析状态": summary.get("status", ""),
    }


def build_cadence_connectivity_summary(project_root: str | Path) -> dict:
    root = Path(project_root).expanduser()
    if not str(project_root or "").strip():
        return {
            "enabled": False,
            "schema_version": CADENCE_PAGE_SCHEMA_VERSION,
            "root": "",
            "page_count": 0,
            "rows": [],
            "warnings": [],
        }
    warnings: List[str] = []
    pages = iter_cadence_page_numbers(root)
    if not pages:
        warnings.append(f"未在项目根路径下找到 sch_1/page*.csv 或 page*.csa：{root}")
    rows = [
        cadence_connectivity_row(
            load_cadence_page_model(
                root,
                "project",
                page,
                include_raw_unknown=True,
                collect_junctions=False,
            )
        )
        for page in pages
    ]
    return {
        "enabled": bool(pages),
        "schema_version": CADENCE_PAGE_SCHEMA_VERSION,
        "root": str(root),
        "page_count": len(pages),
        "rows": rows,
        "warnings": warnings,
    }


def _limited_items(items: Sequence[dict], limit: int) -> Tuple[List[dict], bool]:
    limit = max(0, int(limit or 0))
    return list(items[:limit]), len(items) > limit


def build_cadence_page_payload(project_root: str | Path,
                               page: int,
                               *,
                               stdout: str = "summary",
                               object_id: str = "",
                               limit: int = 200,
                               include_raw_unknown: bool = True) -> dict:
    page_no = int(page or 0)
    if page_no <= 0:
        raise ValueError("page 必须是正整数。")
    root = Path(project_root).expanduser()
    if not str(project_root or "").strip():
        raise ValueError("cadence-page 需要 project_root 或 bundle 中的 project_root。")
    stdout = str(stdout or "summary").strip().lower()
    if stdout not in {"summary", "objects", "full"}:
        raise ValueError("stdout 必须是 summary、objects 或 full。")
    limit = max(1, min(5000, int(limit or 200)))
    model = load_cadence_page_model(
        root,
        "project",
        page_no,
        include_raw_unknown=include_raw_unknown,
        collect_junctions=bool(stdout == "full" or str(object_id or "").startswith("conn-")),
    )

    object_payload = None
    object_kind = ""
    if object_id:
        obj = model.object_by_id(object_id)
        if obj:
            object_payload = obj.to_dict(include_raw=True)
            object_kind = "object"
        else:
            conn = model.connectivity_by_id(object_id)
            if not conn:
                raise ValueError(f"PAGE{page_no} 不存在 Cadence 对象或连接组：{object_id}")
            object_payload = conn.to_dict()
            object_kind = "connectivity"

    object_dicts: List[dict] = []
    connectivity_dicts: List[dict] = []
    unbound_dicts: List[dict] = []
    truncated = False
    if stdout in {"objects", "full"}:
        object_dicts = [item.to_dict(include_raw=stdout == "full") for item in model.objects]
        connectivity_dicts = [item.to_dict() for item in model.connectivity]
        unbound_dicts = [item.to_dict(include_raw=stdout == "full") for item in model.unbound_semantics]
        object_dicts, objects_truncated = _limited_items(object_dicts, limit)
        connectivity_dicts, connectivity_truncated = _limited_items(connectivity_dicts, limit)
        unbound_dicts, unbound_truncated = _limited_items(unbound_dicts, limit)
        truncated = objects_truncated or connectivity_truncated or unbound_truncated

    warnings = []
    if model.csv_error:
        warnings.append(f"CSV 解析受限：{model.csv_error}")
    if model.csa_error:
        warnings.append(f"CSA 解析受限：{model.csa_error}")
    if not model.csv_exists and not model.csa_exists:
        warnings.append(f"PAGE{page_no} 未找到 page{page_no}.csv 或 page{page_no}.csa。")

    return {
        "schema_version": CADENCE_PAGE_SCHEMA_VERSION,
        "project_root": str(root),
        "page": page_no,
        "page_label": f"PAGE{page_no}",
        "stdout": stdout,
        "digest": model.digest(),
        "connectivity_summary": model.connectivity_summary(),
        "objects": object_dicts,
        "connectivity": connectivity_dicts,
        "unbound_semantics": unbound_dicts,
        "object": object_payload,
        "object_kind": object_kind,
        "warnings": warnings,
        "truncated": truncated,
        "readonly": True,
    }


def _index_by_fingerprint(items: Sequence[dict]) -> Dict[str, List[dict]]:
    index: Dict[str, List[dict]] = {}
    for item in items:
        index.setdefault(str(item.get("fingerprint") or ""), []).append(item)
    return index


def _pop_match(index: Dict[str, List[dict]], fingerprint: str) -> Optional[dict]:
    bucket = index.get(fingerprint)
    if not bucket:
        return None
    item = bucket.pop(0)
    if not bucket:
        index.pop(fingerprint, None)
    return item


def _diff_items(left_items: Sequence[dict],
                right_items: Sequence[dict],
                *,
                item_type: str,
                max_items: int) -> Tuple[List[dict], int]:
    right_index = _index_by_fingerprint(right_items)
    removed: List[dict] = []
    for left in left_items:
        match = _pop_match(right_index, str(left.get("fingerprint") or ""))
        if match:
            continue
        removed.append(left)
    added: List[dict] = []
    for bucket in right_index.values():
        for right in bucket:
            added.append(right)

    # When object counts are balanced, pair unmatched objects as semantic changes.
    # This captures coordinate moves or property edits without falling back to
    # noisy added+removed pairs. Unknown objects stay as raw add/remove evidence.
    diffs: List[dict] = []
    if item_type != "UNKNOWN":
        while removed and added:
            diffs.append({"type": "changed", "item_type": item_type, "left": removed.pop(0), "right": added.pop(0)})
    diffs.extend({"type": "removed", "item_type": item_type, "left": item, "right": None} for item in removed)
    diffs.extend({"type": "added", "item_type": item_type, "left": None, "right": item} for item in added)
    return diffs[:max(0, max_items)], max(0, len(diffs) - max(0, max_items))


def compare_page_models(left: CadencePageModel,
                        right: CadencePageModel,
                        *,
                        max_diff_items: int = 40) -> dict:
    diffs: List[dict] = []
    omitted = 0
    if left.csv_exists != right.csv_exists:
        diffs.append({"type": "file_presence", "item_type": "CSV", "left": left.csv_exists, "right": right.csv_exists})
    if left.csa_exists != right.csa_exists:
        diffs.append({"type": "file_presence", "item_type": "CSA", "left": left.csa_exists, "right": right.csa_exists})
    if left.page_number != right.page_number:
        diffs.append({"type": "changed", "item_type": "PAGE_NUMBER", "left": left.page_number, "right": right.page_number})

    left_props = {key: str(value) for key, value in left.csv_properties.items()}
    right_props = {key: str(value) for key, value in right.csv_properties.items()}
    for key in sorted(set(left_props) | set(right_props)):
        if left_props.get(key) != right_props.get(key):
            diffs.append({
                "type": "changed" if key in left_props and key in right_props else ("removed" if key in left_props else "added"),
                "item_type": "CSV_PROPERTY",
                "key": key,
                "left": left_props.get(key, ""),
                "right": right_props.get(key, ""),
            })

    for object_type in sorted({item.object_type for item in left.objects + right.objects}):
        left_items = [item.to_dict(include_raw=False) for item in left.objects if item.object_type == object_type]
        right_items = [item.to_dict(include_raw=False) for item in right.objects if item.object_type == object_type]
        item_diffs, item_omitted = _diff_items(
            left_items,
            right_items,
            item_type=object_type,
            max_items=max_diff_items,
        )
        diffs.extend(item_diffs)
        omitted += item_omitted

    left_conn = [item.to_dict() for item in left.connectivity]
    right_conn = [item.to_dict() for item in right.connectivity]
    conn_diffs, conn_omitted = _diff_items(
        left_conn,
        right_conn,
        item_type="CONNECTIVITY",
        max_items=max_diff_items,
    )
    diffs.extend(conn_diffs)
    omitted += conn_omitted

    full_diff_count = len(diffs) + omitted
    returned = diffs[:max(0, max_diff_items)]
    status = "same" if full_diff_count == 0 else "changed"
    if not left.csv_exists and not left.csa_exists and not right.csv_exists and not right.csa_exists:
        status = "both_missing"
    elif not left.csv_exists and not left.csa_exists:
        status = "missing_left"
    elif not right.csv_exists and not right.csa_exists:
        status = "missing_right"
    elif left.csv_error or left.csa_error or right.csv_error or right.csa_error:
        status = "parse_limited"
    return {
        "page": left.page_no,
        "status": status,
        "left_digest": left.digest(),
        "right_digest": right.digest(),
        "diff_count": full_diff_count,
        "returned_diff_count": len(returned),
        "omitted_diff_count": max(0, full_diff_count - len(returned)),
        "diffs": returned,
    }
