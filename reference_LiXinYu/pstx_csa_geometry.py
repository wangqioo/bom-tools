# -*- coding: utf-8 -*-
"""CSA geometry checks for Cadence DE HDL page*.csa files."""

from __future__ import annotations

import math
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Sequence, Set, Tuple

Point = Tuple[int, int]

PAGE_FILE_RE = re.compile(r"^page(\d+)\.csa$", re.IGNORECASE)
WIRE_RE = re.compile(
    r"\bWIRE\s+\S+\s+\S+\s+"
    r"\((-?\d+)\s*,?\s*(-?\d+)\)\s*"
    r"\((-?\d+)\s*,?\s*(-?\d+)\)",
    re.IGNORECASE,
)
DOT_RE = re.compile(
    r"(?<![A-Za-z0-9_])DOT\s+\S+(?:\s+\S+)?\s+"
    r"\((-?\d+)\s*,?\s*(-?\d+)\)",
    re.IGNORECASE,
)
SIG_NAME_RE = re.compile(
    r"\bFORCEPROP\s+\S+\s+LAST\s+SIG_NAME\s+(.+?)(?=\s+J\s+\d+\b|;|$)",
    re.IGNORECASE,
)
PAGE_NUMBER_RE = re.compile(
    r"(?:\bSET\s+)?['\"]?\bPAGE_NUMBER\b['\"]?\s*(?:=\s*)?['\"]?([A-Z]*\d+)['\"]?",
    re.IGNORECASE,
)
CIRCLE_LINE_RE = re.compile(r"^\s*CIRCLE\b", re.IGNORECASE)
ARC_LINE_RE = re.compile(r"^\s*ARC\b", re.IGNORECASE)
COORD_RE = re.compile(r"\((-?\d+)\s*,?\s*(-?\d+)\)")
NUMBER_RE = re.compile(r"[-+]?\d+(?:\.\d+)?")


@dataclass
class Wire:
    wid: int
    line_no: int
    raw: str
    x1: int
    y1: int
    x2: int
    y2: int
    sig_name: str = ""

    @property
    def is_h(self) -> bool:
        return self.y1 == self.y2 and self.x1 != self.x2

    @property
    def is_v(self) -> bool:
        return self.x1 == self.x2 and self.y1 != self.y2

    @property
    def endpoints(self) -> Tuple[Point, Point]:
        return (self.x1, self.y1), (self.x2, self.y2)

    def x_range(self) -> Tuple[int, int]:
        return min(self.x1, self.x2), max(self.x1, self.x2)

    def y_range(self) -> Tuple[int, int]:
        return min(self.y1, self.y2), max(self.y1, self.y2)

    def contains_point(self, point: Point) -> bool:
        px, py = point
        if self.is_h:
            xmin, xmax = self.x_range()
            return self.y1 == py and xmin <= px <= xmax
        if self.is_v:
            ymin, ymax = self.y_range()
            return self.x1 == px and ymin <= py <= ymax
        return False

    def directions_from_point(self, point: Point) -> Set[str]:
        if not self.contains_point(point):
            return set()
        px, py = point
        dirs: Set[str] = set()
        if self.is_h:
            xmin, xmax = self.x_range()
            if xmin < px:
                dirs.add("left")
            if px < xmax:
                dirs.add("right")
        elif self.is_v:
            ymin, ymax = self.y_range()
            if ymin < py:
                dirs.add("down")
            if py < ymax:
                dirs.add("up")
        return dirs


@dataclass
class Dot:
    line_no: int
    x: int
    y: int
    raw: str

    @property
    def point(self) -> Point:
        return self.x, self.y


@dataclass
class CircleMark:
    object_type: str
    line_no: int
    center_x: float
    center_y: float
    radius: float
    diameter: float
    bbox_xmin: float
    bbox_ymin: float
    bbox_xmax: float
    bbox_ymax: float
    width: float
    height: float
    raw: str
    parse_note: str


@dataclass
class CrossFinding:
    x: int
    y: int
    dot_line: int
    h_wire_lines: str
    v_wire_lines: str
    all_wire_lines: str
    labels: str
    detail: str


@dataclass
class PageResult:
    page_no: int
    page_label: str
    page_name: str
    file: str
    relative_file: str
    cross_count: int
    circle_count: int
    wire_count: int
    dot_count: int
    findings: List[CrossFinding]
    circles: List[CircleMark]
    error: str = ""


class DSU:
    def __init__(self, ids: Iterable[int]) -> None:
        self.parent: Dict[int, int] = {i: i for i in ids}

    def find(self, value: int) -> int:
        while self.parent[value] != value:
            self.parent[value] = self.parent[self.parent[value]]
            value = self.parent[value]
        return value

    def union(self, left: int, right: int) -> None:
        left_root = self.find(left)
        right_root = self.find(right)
        if left_root != right_root:
            self.parent[right_root] = left_root


def _fmt_num(value: float) -> str:
    if abs(value - round(value)) < 1e-9:
        return str(int(round(value)))
    return f"{value:.2f}"


def _page_no(path: Path) -> int:
    match = PAGE_FILE_RE.match(path.name)
    return int(match.group(1)) if match else 10**12


def _page_label(page_no: int) -> str:
    return f"PAGE{page_no}" if page_no < 10**12 else "UNKNOWN"


def _normalize_sig(sig: str) -> str:
    return sig.strip().rstrip(";").strip()


def _extract_coords(raw: str) -> List[Tuple[float, float]]:
    return [(float(x), float(y)) for x, y in COORD_RE.findall(raw)]


def _circle_from_center_radius(
    object_type: str,
    line_no: int,
    raw: str,
    center: Tuple[float, float],
    radius: float,
    note: str,
) -> CircleMark:
    cx, cy = center
    r = abs(radius)
    return CircleMark(
        object_type=object_type,
        line_no=line_no,
        center_x=cx,
        center_y=cy,
        radius=r,
        diameter=2 * r,
        bbox_xmin=cx - r,
        bbox_ymin=cy - r,
        bbox_xmax=cx + r,
        bbox_ymax=cy + r,
        width=2 * r,
        height=2 * r,
        raw=raw,
        parse_note=note,
    )


def _circle_from_bbox(
    object_type: str,
    line_no: int,
    raw: str,
    p1: Tuple[float, float],
    p2: Tuple[float, float],
    note: str,
) -> CircleMark:
    x1, y1 = p1
    x2, y2 = p2
    xmin, xmax = sorted([x1, x2])
    ymin, ymax = sorted([y1, y2])
    width = xmax - xmin
    height = ymax - ymin
    center = ((xmin + xmax) / 2, (ymin + ymax) / 2)
    radius = max(width, height) / 2
    return CircleMark(
        object_type=object_type,
        line_no=line_no,
        center_x=center[0],
        center_y=center[1],
        radius=radius,
        diameter=2 * radius,
        bbox_xmin=xmin,
        bbox_ymin=ymin,
        bbox_xmax=xmax,
        bbox_ymax=ymax,
        width=width,
        height=height,
        raw=raw,
        parse_note=note,
    )


def _fit_circle_three_points(
    p1: Tuple[float, float],
    p2: Tuple[float, float],
    p3: Tuple[float, float],
) -> Optional[Tuple[float, float, float]]:
    x1, y1 = p1
    x2, y2 = p2
    x3, y3 = p3
    temp = x2 * x2 + y2 * y2
    bc = (x1 * x1 + y1 * y1 - temp) / 2.0
    cd = (temp - x3 * x3 - y3 * y3) / 2.0
    det = (x1 - x2) * (y2 - y3) - (x2 - x3) * (y1 - y2)
    if abs(det) < 1e-9:
        return None
    cx = (bc * (y2 - y3) - cd * (y1 - y2)) / det
    cy = ((x1 - x2) * cd - (x2 - x3) * bc) / det
    return cx, cy, math.hypot(cx - x1, cy - y1)


def parse_circle_line(raw: str, line_no: int, circle_two_point_mode: str = "center_radius") -> Optional[CircleMark]:
    coords = _extract_coords(raw)
    if len(coords) >= 2:
        if circle_two_point_mode == "bbox":
            return _circle_from_bbox(
                "CIRCLE", line_no, raw, coords[0], coords[1],
                "CIRCLE two-point mode: bbox diagonal points.",
            )
        radius = math.hypot(coords[1][0] - coords[0][0], coords[1][1] - coords[0][1])
        return _circle_from_center_radius(
            "CIRCLE", line_no, raw, coords[0], radius,
            "CIRCLE two-point mode: center + radius point.",
        )

    if len(coords) == 1:
        nums = [float(n) for n in NUMBER_RE.findall(COORD_RE.sub(" ", raw))]
        if nums:
            return _circle_from_center_radius(
                "CIRCLE", line_no, raw, coords[0], nums[-1],
                "CIRCLE parsed as center + numeric radius.",
            )
    return None


def parse_arc_line_as_circle(raw: str, line_no: int) -> Optional[CircleMark]:
    coords = _extract_coords(raw)
    if len(coords) >= 3:
        fit = _fit_circle_three_points(coords[0], coords[1], coords[2])
        if fit is None:
            return None
        cx, cy, radius = fit
        return _circle_from_center_radius(
            "ARC_FIT", line_no, raw, (cx, cy), radius,
            "ARC parsed by fitting a circle through three points; manually confirm it is a real mark circle.",
        )
    if len(coords) == 2:
        return _circle_from_bbox(
            "ARC_DIAMETER_GUESS", line_no, raw, coords[0], coords[1],
            "ARC with two points parsed as a weak diameter/bbox guess; manually confirm.",
        )
    return None


def parse_csa_text(
    text: str,
    page_no: int,
    *,
    circle_two_point_mode: str = "center_radius",
    include_arcs: bool = True,
) -> Tuple[List[Wire], List[Dot], List[CircleMark], str]:
    wires: List[Wire] = []
    dots: List[Dot] = []
    circles: List[CircleMark] = []
    page_name = _page_label(page_no)
    last_wire: Optional[Wire] = None

    for line_no, raw_line in enumerate(str(text or "").splitlines(), start=1):
        raw = raw_line.strip()
        if not raw:
            continue

        page_match = PAGE_NUMBER_RE.search(raw)
        if page_match:
            page_name = page_match.group(1).upper()

        events = []
        for match in WIRE_RE.finditer(raw):
            events.append((match.start(), "WIRE", match))
        for match in DOT_RE.finditer(raw):
            events.append((match.start(), "DOT", match))
        for match in SIG_NAME_RE.finditer(raw):
            events.append((match.start(), "SIG", match))
        events.sort(key=lambda item: item[0])

        for _, kind, match in events:
            if kind == "WIRE":
                x1, y1, x2, y2 = map(int, match.groups())
                wire = Wire(
                    wid=len(wires),
                    line_no=line_no,
                    raw=match.group(0).strip(),
                    x1=x1,
                    y1=y1,
                    x2=x2,
                    y2=y2,
                )
                wires.append(wire)
                last_wire = wire
            elif kind == "DOT":
                x, y = map(int, match.groups()[-2:])
                dots.append(Dot(line_no=line_no, x=x, y=y, raw=match.group(0).strip()))
            elif kind == "SIG" and last_wire is not None:
                last_wire.sig_name = _normalize_sig(match.group(1))

        if CIRCLE_LINE_RE.match(raw):
            circle = parse_circle_line(raw, line_no, circle_two_point_mode)
            if circle:
                circles.append(circle)
        elif include_arcs and ARC_LINE_RE.match(raw):
            circle = parse_arc_line_as_circle(raw, line_no)
            if circle:
                circles.append(circle)

    return wires, dots, circles, page_name


def _wire_component_labels(wires: Sequence[Wire]) -> Dict[int, Set[str]]:
    dsu = DSU(wire.wid for wire in wires)
    endpoint_map: Dict[Point, List[int]] = {}
    for wire in wires:
        if not (wire.is_h or wire.is_v):
            continue
        for endpoint in wire.endpoints:
            endpoint_map.setdefault(endpoint, []).append(wire.wid)

    for ids in endpoint_map.values():
        if len(ids) < 2:
            continue
        for other in ids[1:]:
            dsu.union(ids[0], other)

    root_labels: Dict[int, Set[str]] = {}
    for wire in wires:
        root = dsu.find(wire.wid)
        root_labels.setdefault(root, set())
        if wire.sig_name:
            root_labels[root].add(wire.sig_name)
    return {wire.wid: root_labels.get(dsu.find(wire.wid), set()) for wire in wires}


def find_dot_four_way_crosses(wires: Sequence[Wire], dots: Sequence[Dot]) -> List[CrossFinding]:
    findings: List[CrossFinding] = []
    labels_by_wire = _wire_component_labels(wires)
    required = {"left", "right", "down", "up"}

    for dot in dots:
        touching = [
            wire for wire in wires
            if (wire.is_h or wire.is_v) and wire.contains_point(dot.point)
        ]
        if not touching:
            continue

        directions: Set[str] = set()
        for wire in touching:
            directions.update(wire.directions_from_point(dot.point))
        if not required.issubset(directions):
            continue

        h_wires = [wire for wire in touching if wire.is_h]
        v_wires = [wire for wire in touching if wire.is_v]
        labels: Set[str] = set()
        for wire in touching:
            labels.update(labels_by_wire.get(wire.wid, set()))
            if wire.sig_name:
                labels.add(wire.sig_name)
        findings.append(CrossFinding(
            x=dot.x,
            y=dot.y,
            dot_line=dot.line_no,
            h_wire_lines=",".join(str(wire.line_no) for wire in h_wires),
            v_wire_lines=",".join(str(wire.line_no) for wire in v_wires),
            all_wire_lines=",".join(str(wire.line_no) for wire in touching),
            labels=",".join(sorted(labels)) if labels else "",
            detail="DOT point has WIREs extending left/right/up/down. T junctions and dotless crosses are ignored.",
        ))

    return findings


def collect_page_files(project_root: str | Path) -> List[Path]:
    root = Path(project_root).expanduser()
    csa_root = root if root.name.lower() == "sch_1" else root / "sch_1"
    if not csa_root.is_dir():
        return []
    files = [path for path in csa_root.glob("*.csa") if PAGE_FILE_RE.match(path.name)]
    return sorted(files, key=lambda path: (_page_no(path), str(path).lower()))


def _decode_csa_bytes(data: bytes) -> str:
    for encoding in ("utf-8-sig", "utf-16", "gb18030", "latin-1"):
        try:
            return data.decode(encoding)
        except UnicodeDecodeError:
            continue
    return data.decode("utf-8", errors="replace")


def analyze_one_page(
    file_path: str | Path,
    *,
    root: str | Path = "",
    circle_two_point_mode: str = "center_radius",
    include_arcs: bool = True,
) -> PageResult:
    path = Path(file_path)
    page_no = _page_no(path)
    page_label = _page_label(page_no)
    csa_root = Path(root).expanduser() if root else path.parent
    try:
        relative_file = str(path.relative_to(csa_root))
    except ValueError:
        relative_file = path.name

    try:
        text = _decode_csa_bytes(path.read_bytes())
        wires, dots, circles, page_name = parse_csa_text(
            text,
            page_no,
            circle_two_point_mode=circle_two_point_mode,
            include_arcs=include_arcs,
        )
        findings = find_dot_four_way_crosses(wires, dots)
        return PageResult(
            page_no=page_no,
            page_label=page_label,
            page_name=page_name,
            file=str(path),
            relative_file=relative_file,
            cross_count=len(findings),
            circle_count=len(circles),
            wire_count=len(wires),
            dot_count=len(dots),
            findings=findings,
            circles=circles,
        )
    except Exception as exc:
        return PageResult(
            page_no=page_no,
            page_label=page_label,
            page_name=page_label,
            file=str(path),
            relative_file=relative_file,
            cross_count=0,
            circle_count=0,
            wire_count=0,
            dot_count=0,
            findings=[],
            circles=[],
            error=str(exc),
        )


def _summary_row(result: PageResult) -> dict:
    return {
        "页面": result.page_label,
        "CSA页名": result.page_name,
        "文件": result.relative_file,
        "DOT四向十字数": result.cross_count,
        "画圈对象数": result.circle_count,
        "WIRE数": result.wire_count,
        "DOT数": result.dot_count,
        "错误": result.error,
    }


def _cross_rows(result: PageResult) -> List[dict]:
    rows = []
    for index, item in enumerate(result.findings, start=1):
        rows.append({
            "页面": result.page_label,
            "CSA页名": result.page_name,
            "文件": result.relative_file,
            "序号": index,
            "坐标": f"({item.x},{item.y})",
            "X": item.x,
            "Y": item.y,
            "DOT行号": item.dot_line,
            "水平WIRE行号": item.h_wire_lines,
            "垂直WIRE行号": item.v_wire_lines,
            "全部WIRE行号": item.all_wire_lines,
            "关联信号": item.labels,
            "说明": "带 DOT 的四向十字交叉，需人工复核是否符合绘图规范。",
        })
    return rows


def _circle_rows(result: PageResult) -> List[dict]:
    rows = []
    for index, item in enumerate(result.circles, start=1):
        rows.append({
            "页面": result.page_label,
            "CSA页名": result.page_name,
            "文件": result.relative_file,
            "序号": index,
            "对象类型": item.object_type,
            "行号": item.line_no,
            "圆心": f"({_fmt_num(item.center_x)},{_fmt_num(item.center_y)})",
            "半径": _fmt_num(item.radius),
            "直径": _fmt_num(item.diameter),
            "外接框": (
                f"({_fmt_num(item.bbox_xmin)},{_fmt_num(item.bbox_ymin)})"
                f"-({_fmt_num(item.bbox_xmax)},{_fmt_num(item.bbox_ymax)})"
            ),
            "宽": _fmt_num(item.width),
            "高": _fmt_num(item.height),
            "解析说明": item.parse_note,
            "原始行": item.raw,
        })
    return rows


def analyze_csa_geometry(
    project_root: str | Path,
    *,
    include_arcs: bool = True,
    circle_two_point_mode: str = "center_radius",
) -> dict:
    files = collect_page_files(project_root)
    if not files:
        root = Path(project_root).expanduser()
        csa_root = root if root.name.lower() == "sch_1" else root / "sch_1"
        return {
            "enabled": False,
            "root": str(csa_root),
            "page_count": 0,
            "cross_count": 0,
            "circle_count": 0,
            "error_count": 0,
            "summary_rows": [],
            "dot_cross_rows": [],
            "circle_rows": [],
            "warnings": [],
        }

    root = files[0].parent
    results = [
        analyze_one_page(
            path,
            root=root,
            circle_two_point_mode=circle_two_point_mode,
            include_arcs=include_arcs,
        )
        for path in files
    ]
    summary_rows = [_summary_row(result) for result in results]
    dot_cross_rows: List[dict] = []
    circle_rows: List[dict] = []
    for result in results:
        dot_cross_rows.extend(_cross_rows(result))
        circle_rows.extend(_circle_rows(result))
    warnings = [
        f"CSA 页面解析异常：{result.relative_file}: {result.error}"
        for result in results
        if result.error
    ]
    return {
        "enabled": True,
        "root": str(root),
        "page_count": len(results),
        "cross_count": sum(result.cross_count for result in results),
        "circle_count": sum(result.circle_count for result in results),
        "error_count": sum(1 for result in results if result.error),
        "summary_rows": summary_rows,
        "dot_cross_rows": dot_cross_rows,
        "circle_rows": circle_rows,
        "warnings": warnings,
    }
