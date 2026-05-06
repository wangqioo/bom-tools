# -*- coding: utf-8 -*-
"""CSA geometry checks for Cadence DE HDL page*.csa files."""

from __future__ import annotations

import csv
import html as html_lib
import json
import math
import os
import re
from concurrent.futures import ProcessPoolExecutor, ThreadPoolExecutor, as_completed
from dataclasses import asdict, dataclass, field
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Sequence, Set, Tuple

Point = Tuple[int, int]
CSA_GEOMETRY_SCHEMA_VERSION = "pstx-csa-geometry.v1"

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
    source_lines: List[dict] = field(default_factory=list)

    @property
    def position_text(self) -> str:
        return f"({_fmt_num(self.center_x)},{_fmt_num(self.center_y)})"

    @property
    def size_text(self) -> str:
        return f"r={_fmt_num(self.radius)},d={_fmt_num(self.diameter)}"


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
    dot_raw: str = ""
    wire_raws: str = ""
    source_lines: List[dict] = field(default_factory=list)


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

    @property
    def cross_positions(self) -> str:
        return ";".join(f"({item.x},{item.y})" for item in self.findings)

    @property
    def circle_positions(self) -> str:
        return ";".join(item.position_text for item in self.circles)

    @property
    def circle_sizes(self) -> str:
        return ";".join(item.size_text for item in self.circles)


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


def _source_context(source_lines: Sequence[str], line_numbers: Iterable[int], radius: int = 1) -> List[dict]:
    if not source_lines:
        return []
    selected: Set[int] = set()
    for raw_line_no in line_numbers:
        try:
            line_no = int(raw_line_no)
        except (TypeError, ValueError):
            continue
        for candidate in range(max(1, line_no - radius), min(len(source_lines), line_no + radius) + 1):
            selected.add(candidate)
    return [
        {"line_no": line_no, "text": source_lines[line_no - 1].strip()}
        for line_no in sorted(selected)
    ]


def _source_line_text(source_lines: Sequence[str], line_no: int, fallback: str = "") -> str:
    try:
        index = int(line_no) - 1
    except (TypeError, ValueError):
        return fallback
    if 0 <= index < len(source_lines):
        return source_lines[index].strip()
    return fallback


def _format_source_lines(source_lines: Sequence[dict]) -> str:
    return " | ".join(
        f"{item.get('line_no', '')}: {item.get('text', '')}"
        for item in source_lines
        if item.get("line_no") or item.get("text")
    )


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
    source_lines = str(text or "").splitlines()

    for line_no, raw_line in enumerate(source_lines, start=1):
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
                circle.source_lines = _source_context(source_lines, [line_no])
                circles.append(circle)
        elif include_arcs and ARC_LINE_RE.match(raw):
            circle = parse_arc_line_as_circle(raw, line_no)
            if circle:
                circle.source_lines = _source_context(source_lines, [line_no])
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


def wires_touching_point(wires: Sequence[Wire], point: Point) -> List[Wire]:
    return [
        wire for wire in wires
        if (wire.is_h or wire.is_v) and wire.contains_point(point)
    ]


def find_dot_four_way_crosses(
    wires: Sequence[Wire],
    dots: Sequence[Dot],
    source_lines: Optional[Sequence[str]] = None,
) -> List[CrossFinding]:
    findings: List[CrossFinding] = []
    labels_by_wire = _wire_component_labels(wires)
    required = {"left", "right", "down", "up"}

    for dot in dots:
        touching = wires_touching_point(wires, dot.point)
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
            dot_raw=_source_line_text(source_lines or [], dot.line_no, dot.raw),
            wire_raws=" | ".join(
                f"{wire.line_no}: {_source_line_text(source_lines or [], wire.line_no, wire.raw)}"
                for wire in touching
            ),
            source_lines=_source_context(source_lines or [], [dot.line_no, *(wire.line_no for wire in touching)]),
        ))

    return findings


def _page_file_sort_key(path: Path) -> tuple[int, str]:
    return _page_no(path), str(path).lower()


def _candidate_scan_dir(root: Path, recursive: bool) -> Path:
    if recursive:
        return root
    if root.name.lower() == "sch_1":
        return root
    sch_dir = root / "sch_1"
    return sch_dir if sch_dir.is_dir() else root


def collect_page_files(
    project_root: str | Path,
    *,
    recursive: bool = False,
    strict: bool = False,
) -> List[Path]:
    root = Path(project_root).expanduser()
    if root.is_file():
        if PAGE_FILE_RE.match(root.name):
            return [root]
        if strict:
            raise ValueError(f"Input file is not page数字.csa: {root}")
        return []
    if not root.exists():
        if strict:
            raise FileNotFoundError(f"Path does not exist: {root}")
        return []
    if not root.is_dir():
        if strict:
            raise ValueError(f"Input is not a file or folder: {root}")
        return []

    csa_root = _candidate_scan_dir(root, recursive)
    if not csa_root.is_dir():
        if strict:
            raise FileNotFoundError(f"CSA scan directory does not exist: {csa_root}")
        return []
    iterator = csa_root.rglob("*.csa") if recursive else csa_root.glob("*.csa")
    files = [path for path in iterator if PAGE_FILE_RE.match(path.name)]
    if strict and not files:
        raise FileNotFoundError(f"No page数字.csa files found: {root}")
    return sorted(files, key=_page_file_sort_key)


def find_missing_pages(files: Sequence[Path]) -> List[int]:
    nums = sorted({_page_no(Path(path)) for path in files if _page_no(Path(path)) < 10**12})
    if not nums:
        return []
    seen = set(nums)
    return [page_no for page_no in range(nums[0], nums[-1] + 1) if page_no not in seen]


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
        findings = find_dot_four_way_crosses(wires, dots, text.splitlines())
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


def _analyze_one_page_worker(task: tuple[str, str, str, bool]) -> PageResult:
    file_path, root, circle_two_point_mode, include_arcs = task
    return analyze_one_page(
        file_path,
        root=root,
        circle_two_point_mode=circle_two_point_mode,
        include_arcs=include_arcs,
    )


def scan_pages(
    files: Sequence[Path],
    *,
    workers: Optional[int] = None,
    executor_kind: str = "thread",
    circle_two_point_mode: str = "center_radius",
    include_arcs: bool = False,
    root: str | Path = "",
) -> List[PageResult]:
    if not files:
        return []
    if executor_kind not in {"thread", "process", "serial"}:
        raise ValueError(f"unsupported CSA executor: {executor_kind}")
    scan_root = str(Path(root).expanduser()) if root else str(Path(files[0]).parent)
    max_workers = workers or min(os.cpu_count() or 1, len(files))
    max_workers = max(1, min(max_workers, len(files)))
    tasks = [
        (str(Path(path)), scan_root, circle_two_point_mode, include_arcs)
        for path in files
    ]
    if executor_kind == "serial" or max_workers == 1 or len(tasks) == 1:
        results = [_analyze_one_page_worker(task) for task in tasks]
    else:
        executor_cls = ThreadPoolExecutor if executor_kind == "thread" else ProcessPoolExecutor
        results = []
        with executor_cls(max_workers=max_workers) as executor:
            futures = [executor.submit(_analyze_one_page_worker, task) for task in tasks]
            for future in as_completed(futures):
                results.append(future.result())
    return sorted(results, key=lambda result: (result.page_no, result.file.lower()))


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
            "DOT原始行": item.dot_raw,
            "相关WIRE原始行": item.wire_raws,
            "证据上下文": _format_source_lines(item.source_lines),
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
            "证据上下文": _format_source_lines(item.source_lines),
        })
    return rows


def _package_summary_row(result: PageResult) -> dict:
    return {
        "page_no": result.page_no if result.page_no < 10**12 else "",
        "page_name": result.page_name,
        "file": result.file,
        "cross_count": result.cross_count,
        "cross_positions": result.cross_positions,
        "circle_count": result.circle_count,
        "circle_positions": result.circle_positions,
        "circle_sizes": result.circle_sizes,
        "wire_count": result.wire_count,
        "dot_count": result.dot_count,
        "error": result.error,
    }


def _package_cross_rows(result: PageResult) -> List[dict]:
    rows = []
    for index, item in enumerate(result.findings, start=1):
        rows.append({
            "page_no": result.page_no if result.page_no < 10**12 else "",
            "page_name": result.page_name,
            "file": result.file,
            "index_in_page": index,
            "x": item.x,
            "y": item.y,
            "dot_line": item.dot_line,
            "h_wire_lines": item.h_wire_lines,
            "v_wire_lines": item.v_wire_lines,
            "all_wire_lines": item.all_wire_lines,
            "labels": item.labels,
            "detail": item.detail,
            "dot_raw": item.dot_raw,
            "wire_raws": item.wire_raws,
            "source_context": _format_source_lines(item.source_lines),
        })
    return rows


def _package_circle_rows(result: PageResult) -> List[dict]:
    rows = []
    for index, item in enumerate(result.circles, start=1):
        rows.append({
            "page_no": result.page_no if result.page_no < 10**12 else "",
            "page_name": result.page_name,
            "file": result.file,
            "index_in_page": index,
            "object_type": item.object_type,
            "line_no": item.line_no,
            "center_x": _fmt_num(item.center_x),
            "center_y": _fmt_num(item.center_y),
            "radius": _fmt_num(item.radius),
            "diameter": _fmt_num(item.diameter),
            "bbox_xmin": _fmt_num(item.bbox_xmin),
            "bbox_ymin": _fmt_num(item.bbox_ymin),
            "bbox_xmax": _fmt_num(item.bbox_xmax),
            "bbox_ymax": _fmt_num(item.bbox_ymax),
            "width": _fmt_num(item.width),
            "height": _fmt_num(item.height),
            "raw": item.raw,
            "parse_note": item.parse_note,
            "source_context": _format_source_lines(item.source_lines),
        })
    return rows


def _write_csv_rows(rows: Sequence[dict], fieldnames: Sequence[str], out_path: str | Path) -> None:
    target = Path(out_path).expanduser()
    target.parent.mkdir(parents=True, exist_ok=True)
    with target.open("w", newline="", encoding="utf-8-sig") as handle:
        writer = csv.DictWriter(handle, fieldnames=fieldnames)
        writer.writeheader()
        for row in rows:
            writer.writerow({field: row.get(field, "") for field in fieldnames})


def _html_escape(value: object) -> str:
    return html_lib.escape(str(value if value is not None else ""), quote=True)


def _html_table(title: str, rows: Sequence[dict], fieldnames: Sequence[str]) -> str:
    if not rows:
        return (
            f"<section><h2>{_html_escape(title)}</h2>"
            "<p class=\"empty\">No rows.</p></section>"
        )
    head = "".join(f"<th>{_html_escape(field)}</th>" for field in fieldnames)
    body_parts = []
    for row in rows:
        cells = "".join(
            f"<td><pre>{_html_escape(row.get(field, ''))}</pre></td>"
            for field in fieldnames
        )
        body_parts.append(f"<tr>{cells}</tr>")
    return (
        f"<section><h2>{_html_escape(title)}</h2>"
        "<div class=\"table-wrap\"><table>"
        f"<thead><tr>{head}</tr></thead>"
        f"<tbody>{''.join(body_parts)}</tbody>"
        "</table></div></section>"
    )


def write_summary_csv(results: Sequence[PageResult], out_path: str | Path) -> None:
    _write_csv_rows(
        [_package_summary_row(result) for result in results],
        [
            "page_no", "page_name", "file", "cross_count", "cross_positions",
            "circle_count", "circle_positions", "circle_sizes", "wire_count", "dot_count", "error",
        ],
        out_path,
    )


def write_cross_detail_csv(results: Sequence[PageResult], out_path: str | Path) -> None:
    rows: List[dict] = []
    for result in results:
        rows.extend(_package_cross_rows(result))
    _write_csv_rows(
        rows,
        [
            "page_no", "page_name", "file", "index_in_page", "x", "y", "dot_line",
            "h_wire_lines", "v_wire_lines", "all_wire_lines", "labels", "detail",
            "dot_raw", "wire_raws", "source_context",
        ],
        out_path,
    )


def write_circle_detail_csv(results: Sequence[PageResult], out_path: str | Path) -> None:
    rows: List[dict] = []
    for result in results:
        rows.extend(_package_circle_rows(result))
    _write_csv_rows(
        rows,
        [
            "page_no", "page_name", "file", "index_in_page", "object_type", "line_no",
            "center_x", "center_y", "radius", "diameter", "bbox_xmin", "bbox_ymin",
            "bbox_xmax", "bbox_ymax", "width", "height", "raw", "parse_note",
            "source_context",
        ],
        out_path,
    )


def page_result_to_dict(result: PageResult) -> dict:
    payload = asdict(result)
    payload.update({
        "cross_positions": result.cross_positions,
        "circle_positions": result.circle_positions,
        "circle_sizes": result.circle_sizes,
    })
    return payload


def write_json(results: Sequence[PageResult], out_path: str | Path) -> None:
    target = Path(out_path).expanduser()
    target.parent.mkdir(parents=True, exist_ok=True)
    target.write_text(
        json.dumps([page_result_to_dict(result) for result in results], ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


def write_html_report(results: Sequence[PageResult], out_path: str | Path) -> None:
    summary_rows = [_package_summary_row(result) for result in results]
    cross_rows: List[dict] = []
    circle_rows: List[dict] = []
    for result in results:
        cross_rows.extend(_package_cross_rows(result))
        circle_rows.extend(_package_circle_rows(result))

    cross_count = sum(result.cross_count for result in results)
    circle_count = sum(result.circle_count for result in results)
    error_count = sum(1 for result in results if result.error)
    sections = [
        _html_table(
            "Page Summary",
            summary_rows,
            [
                "page_no", "page_name", "file", "cross_count", "cross_positions",
                "circle_count", "circle_positions", "circle_sizes", "wire_count", "dot_count", "error",
            ],
        ),
        _html_table(
            "DOT Four-Way Cross Findings",
            cross_rows,
            [
                "page_no", "page_name", "file", "index_in_page", "x", "y", "dot_line",
                "h_wire_lines", "v_wire_lines", "all_wire_lines", "labels", "detail",
                "dot_raw", "wire_raws", "source_context",
            ],
        ),
        _html_table(
            "Circle And ARC Candidates",
            circle_rows,
            [
                "page_no", "page_name", "file", "index_in_page", "object_type", "line_no",
                "center_x", "center_y", "radius", "diameter", "bbox_xmin", "bbox_ymin",
                "bbox_xmax", "bbox_ymax", "width", "height", "raw", "parse_note",
                "source_context",
            ],
        ),
    ]
    document = f"""<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <title>CSA Geometry Report</title>
  <style>
    body {{ font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif; margin: 24px; color: #172033; background: #f7f8fb; }}
    header, section {{ background: #fff; border: 1px solid #d9dfeb; border-radius: 8px; margin-bottom: 18px; padding: 18px; }}
    h1 {{ font-size: 24px; margin: 0 0 8px; }}
    h2 {{ font-size: 18px; margin: 0 0 12px; }}
    .digest {{ display: flex; flex-wrap: wrap; gap: 10px; margin-top: 12px; }}
    .digest span {{ border: 1px solid #cfd7e6; border-radius: 6px; padding: 6px 10px; background: #fbfcff; }}
    .table-wrap {{ overflow-x: auto; }}
    table {{ border-collapse: collapse; width: 100%; min-width: 860px; }}
    th, td {{ border: 1px solid #d9dfeb; text-align: left; vertical-align: top; padding: 8px; }}
    th {{ background: #edf2fa; font-weight: 600; }}
    pre {{ margin: 0; white-space: pre-wrap; word-break: break-word; font-family: ui-monospace, SFMono-Regular, Menlo, monospace; font-size: 12px; line-height: 1.35; }}
    .empty {{ color: #667085; margin: 0; }}
  </style>
</head>
<body>
  <header>
    <h1>CSA Geometry Report</h1>
    <p>Geometry evidence only. DOT crosses and circle candidates require engineering review; this report does not assert electrical shorts.</p>
    <div class="digest">
      <span>Pages: {_html_escape(len(results))}</span>
      <span>DOT crosses: {_html_escape(cross_count)}</span>
      <span>Circle candidates: {_html_escape(circle_count)}</span>
      <span>Parse errors: {_html_escape(error_count)}</span>
    </div>
  </header>
  {''.join(sections)}
</body>
</html>
"""
    target = Path(out_path).expanduser()
    target.parent.mkdir(parents=True, exist_ok=True)
    target.write_text(document, encoding="utf-8")


def write_csa_geometry_reports(
    results: Sequence[PageResult],
    out_dir: str | Path,
    *,
    summary_name: str = "cross_circle_summary.csv",
    cross_detail_name: str = "dot_cross_detail.csv",
    circle_detail_name: str = "circle_detail.csv",
    json_report: bool = False,
    json_name: str = "cross_circle_report.json",
    html_report: bool = False,
    html_name: str = "cross_circle_report.html",
) -> Dict[str, Optional[str]]:
    target_dir = Path(out_dir).expanduser()
    target_dir.mkdir(parents=True, exist_ok=True)
    summary_path = target_dir / summary_name
    cross_path = target_dir / cross_detail_name
    circle_path = target_dir / circle_detail_name
    write_summary_csv(results, summary_path)
    write_cross_detail_csv(results, cross_path)
    write_circle_detail_csv(results, circle_path)
    json_path = target_dir / json_name if json_report else None
    if json_path:
        write_json(results, json_path)
    html_path = target_dir / html_name if html_report else None
    if html_path:
        write_html_report(results, html_path)
    return {
        "summary_csv": str(summary_path),
        "cross_detail_csv": str(cross_path),
        "circle_detail_csv": str(circle_path),
        "json_report": str(json_path) if json_path else None,
        "html_report": str(html_path) if html_path else None,
    }


def build_csa_geometry_result(
    results: Sequence[PageResult],
    *,
    root: str | Path,
    missing_pages: Optional[Sequence[int]] = None,
) -> dict:
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
    missing = list(missing_pages or [])
    if missing:
        preview = ",".join(str(item) for item in missing[:50])
        suffix = " ..." if len(missing) > 50 else ""
        warnings.append(f"CSA page number gap: missing page numbers {preview}{suffix}")
    return {
        "enabled": bool(results),
        "root": str(Path(root).expanduser()),
        "page_count": len(results),
        "cross_count": sum(result.cross_count for result in results),
        "circle_count": sum(result.circle_count for result in results),
        "error_count": sum(1 for result in results if result.error),
        "missing_pages": missing,
        "summary_rows": summary_rows,
        "dot_cross_rows": dot_cross_rows,
        "circle_rows": circle_rows,
        "warnings": warnings,
    }


def scan_csa_geometry(
    project_root: str | Path,
    *,
    recursive: bool = False,
    workers: Optional[int] = None,
    executor_kind: str = "thread",
    circle_two_point_mode: str = "center_radius",
    include_arcs: bool = False,
    check_missing: bool = False,
    strict: bool = False,
    page: Optional[int] = None,
) -> tuple[List[PageResult], dict]:
    files = collect_page_files(project_root, recursive=recursive, strict=strict)
    page_no = int(page or 0)
    if page_no > 0:
        files = [path for path in files if _page_no(path) == page_no]
        if strict and not files:
            raise FileNotFoundError(f"No page{page_no}.csa files found: {project_root}")
    root = Path(project_root).expanduser()
    scan_root = root if (root.is_file() or recursive) else _candidate_scan_dir(root, recursive=False)
    missing = find_missing_pages(files) if check_missing else []
    results = scan_pages(
        files,
        workers=workers,
        executor_kind=executor_kind,
        circle_two_point_mode=circle_two_point_mode,
        include_arcs=include_arcs,
        root=scan_root if scan_root.exists() else (files[0].parent if files else scan_root),
    )
    return results, build_csa_geometry_result(results, root=scan_root, missing_pages=missing)


def _limited_rows(rows: Sequence[dict], limit: int) -> tuple[List[dict], bool]:
    if limit < 1:
        limit = 1
    return list(rows[:limit]), len(rows) > limit


def build_csa_geometry_payload(
    csa_geometry: dict,
    *,
    stdout: str = "summary",
    limit: int = 200,
    page: Optional[int] = None,
    semantic_overlay: Optional[dict] = None,
) -> dict:
    mode = stdout if stdout in {"summary", "hits", "details", "full"} else "summary"
    limit = max(1, min(5000, int(limit or 200)))
    summary_rows = list(csa_geometry.get("summary_rows", []) or [])
    dot_rows = list(csa_geometry.get("dot_cross_rows", []) or [])
    circle_rows = list(csa_geometry.get("circle_rows", []) or [])
    page_no = int(page or 0)
    if page_no > 0:
        page_label = f"PAGE{page_no}"
        summary_rows = [row for row in summary_rows if str(row.get("页面") or "") == page_label]
        dot_rows = [row for row in dot_rows if str(row.get("页面") or "") == page_label]
        circle_rows = [row for row in circle_rows if str(row.get("页面") or "") == page_label]
    digest_summary_rows = list(summary_rows)
    if mode == "hits":
        summary_rows = [
            row for row in summary_rows
            if int(row.get("DOT四向十字数", 0) or 0) or int(row.get("画圈对象数", 0) or 0) or row.get("错误")
        ]

    missing_pages = list(csa_geometry.get("missing_pages", []) or [])
    if page_no > 0:
        missing_pages = [item for item in missing_pages if int(item or 0) == page_no]

    digest = {
        "schema_version": CSA_GEOMETRY_SCHEMA_VERSION,
        "enabled": bool(csa_geometry.get("enabled")),
        "root": str(csa_geometry.get("root", "") or ""),
        "page_count": len(digest_summary_rows) if page_no > 0 else int(csa_geometry.get("page_count", 0) or 0),
        "cross_count": len(dot_rows) if page_no > 0 else int(csa_geometry.get("cross_count", 0) or 0),
        "circle_count": len(circle_rows) if page_no > 0 else int(csa_geometry.get("circle_count", 0) or 0),
        "error_count": int(csa_geometry.get("error_count", 0) or 0),
        "missing_page_count": len(missing_pages),
        "warning_count": len(csa_geometry.get("warnings", []) or []),
        "page_filter": page_no,
    }
    payload = {
        "schema_version": CSA_GEOMETRY_SCHEMA_VERSION,
        "digest": digest,
        "missing_pages": missing_pages,
        "warnings": list(csa_geometry.get("warnings", []) or []),
        "summary_rows": [],
        "dot_cross_rows": [],
        "circle_rows": [],
        "written": {},
        "truncated": False,
        "truncation": {
            "summary_rows": False,
            "dot_cross_rows": False,
            "circle_rows": False,
        },
    }
    if mode in {"hits", "details", "full"}:
        payload["summary_rows"], payload["truncation"]["summary_rows"] = _limited_rows(summary_rows, limit)
        payload["dot_cross_rows"], payload["truncation"]["dot_cross_rows"] = _limited_rows(dot_rows, limit)
    if mode in {"hits", "full"}:
        payload["circle_rows"], payload["truncation"]["circle_rows"] = _limited_rows(circle_rows, limit)
    if semantic_overlay is not None:
        payload["semantic_overlay"] = semantic_overlay
    payload["truncated"] = any(payload["truncation"].values())
    return payload


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
            "missing_pages": [],
            "summary_rows": [],
            "dot_cross_rows": [],
            "circle_rows": [],
            "warnings": [],
        }

    root = files[0].parent
    results = scan_pages(
        files,
        workers=None,
        executor_kind="thread",
        circle_two_point_mode=circle_two_point_mode,
        include_arcs=include_arcs,
        root=root,
    )
    return build_csa_geometry_result(results, root=root)
