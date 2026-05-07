# -*- coding: utf-8 -*-
"""
CSA 几何规范检查工具 v1.0
检查 Cadence DE HDL 导出的 sch_1/page*.csa 文件中的几何规范问题。

功能：
  - DOT 四向十字交叉检测（同坐标存在四向正交 WIRE + DOT）
  - CIRCLE/ARC 画圈对象检测
  - 多页面批量分析
  - 结果树形表展示 + 导出 CSV

依赖：无（纯 Python 标准库）
运行：python csa_checker.py
"""

from __future__ import annotations

import csv
import math
import os
import re
from dataclasses import dataclass, field
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from typing import Dict, Iterable, List, Optional, Set, Tuple
import tkinter as tk

# ══════════════════════════════════════════════════════════
# 零、CSA 解析引擎（内联自 pstx_csa_geometry.py）
# ══════════════════════════════════════════════════════════

Point = Tuple[int, int]

PAGE_FILE_RE = re.compile(r"^page(\d+)\.csa$", re.IGNORECASE)
WIRE_RE = re.compile(
    r"\bWIRE\s+\S+(?:\s+\S+)?\s+"
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


def _natural_sort_key(value: str):
    parts = re.split(r'(\d+)', str(value or '').upper())
    return [int(p) if p.isdigit() else p for p in parts]


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
    findings: List[CrossFinding] = field(default_factory=list)
    circles: List[CircleMark] = field(default_factory=list)
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
    object_type: str, line_no: int, raw: str,
    center: Tuple[float, float], radius: float, note: str,
) -> CircleMark:
    cx, cy = center
    r = abs(radius)
    return CircleMark(
        object_type=object_type, line_no=line_no,
        center_x=cx, center_y=cy, radius=r, diameter=2 * r,
        bbox_xmin=cx - r, bbox_ymin=cy - r, bbox_xmax=cx + r, bbox_ymax=cy + r,
        width=2 * r, height=2 * r, raw=raw, parse_note=note,
    )


def _circle_from_bbox(
    object_type: str, line_no: int, raw: str,
    p1: Tuple[float, float], p2: Tuple[float, float], note: str,
) -> CircleMark:
    x1, y1 = p1; x2, y2 = p2
    xmin, xmax = sorted([x1, x2])
    ymin, ymax = sorted([y1, y2])
    center = ((xmin + xmax) / 2, (ymin + ymax) / 2)
    radius = max(xmax - xmin, ymax - ymin) / 2
    return CircleMark(
        object_type=object_type, line_no=line_no,
        center_x=center[0], center_y=center[1], radius=radius, diameter=2 * radius,
        bbox_xmin=xmin, bbox_ymin=ymin, bbox_xmax=xmax, bbox_ymax=ymax,
        width=xmax - xmin, height=ymax - ymin, raw=raw, parse_note=note,
    )


def _fit_circle_three_points(
    p1: Tuple[float, float], p2: Tuple[float, float], p3: Tuple[float, float],
) -> Optional[Tuple[float, float, float]]:
    x1, y1 = p1; x2, y2 = p2; x3, y3 = p3
    temp = x2 * x2 + y2 * y2
    bc = (x1 * x1 + y1 * y1 - temp) / 2.0
    cd = (temp - x3 * x3 - y3 * y3) / 2.0
    det = (x1 - x2) * (y2 - y3) - (x2 - x3) * (y1 - y2)
    if abs(det) < 1e-9:
        return None
    cx = (bc * (y2 - y3) - cd * (y1 - y2)) / det
    cy = ((x1 - x2) * cd - (x2 - x3) * bc) / det
    return cx, cy, math.hypot(cx - x1, cy - y1)


def parse_circle_line(raw: str, line_no: int, mode: str = "center_radius") -> Optional[CircleMark]:
    coords = _extract_coords(raw)
    if len(coords) >= 2:
        if mode == "bbox":
            return _circle_from_bbox("CIRCLE", line_no, raw, coords[0], coords[1],
                                     "CIRCLE two-point mode: bbox diagonal points.")
        radius = math.hypot(coords[1][0] - coords[0][0], coords[1][1] - coords[0][1])
        return _circle_from_center_radius("CIRCLE", line_no, raw, coords[0], radius,
                                          "CIRCLE two-point mode: center + radius point.")
    if len(coords) == 1:
        nums = [float(n) for n in NUMBER_RE.findall(COORD_RE.sub(" ", raw))]
        if nums:
            return _circle_from_center_radius("CIRCLE", line_no, raw, coords[0], nums[-1],
                                              "CIRCLE parsed as center + numeric radius.")
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
    text: str, page_no: int, *,
    circle_mode: str = "center_radius",
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
                wire = Wire(wid=len(wires), line_no=line_no, raw=match.group(0).strip(),
                            x1=x1, y1=y1, x2=x2, y2=y2)
                wires.append(wire)
                last_wire = wire
            elif kind == "DOT":
                x, y = map(int, match.groups()[-2:])
                dots.append(Dot(line_no=line_no, x=x, y=y, raw=match.group(0).strip()))
            elif kind == "SIG" and last_wire is not None:
                last_wire.sig_name = _normalize_sig(match.group(1))

        if CIRCLE_LINE_RE.match(raw):
            circle = parse_circle_line(raw, line_no, circle_mode)
            if circle:
                circles.append(circle)
        elif include_arcs and ARC_LINE_RE.match(raw):
            circle = parse_arc_line_as_circle(raw, line_no)
            if circle:
                circles.append(circle)

    return wires, dots, circles, page_name


def _wire_component_labels(wires: List[Wire]) -> Dict[int, Set[str]]:
    dsu = DSU(w.wid for w in wires)
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


def find_dot_four_way_crosses(wires: List[Wire], dots: List[Dot]) -> List[CrossFinding]:
    findings: List[CrossFinding] = []
    labels_by_wire = _wire_component_labels(wires)
    required = {"left", "right", "down", "up"}

    for dot in dots:
        touching = [w for w in wires if (w.is_h or w.is_v) and w.contains_point(dot.point)]
        if not touching:
            continue
        directions: Set[str] = set()
        for w in touching:
            directions.update(w.directions_from_point(dot.point))
        if not required.issubset(directions):
            continue
        h_wires = [w for w in touching if w.is_h]
        v_wires = [w for w in touching if w.is_v]
        labels: Set[str] = set()
        for w in touching:
            labels.update(labels_by_wire.get(w.wid, set()))
            if w.sig_name:
                labels.add(w.sig_name)
        findings.append(CrossFinding(
            x=dot.x, y=dot.y, dot_line=dot.line_no,
            h_wire_lines=",".join(str(w.line_no) for w in h_wires),
            v_wire_lines=",".join(str(w.line_no) for w in v_wires),
            all_wire_lines=",".join(str(w.line_no) for w in touching),
            labels=",".join(sorted(labels)) if labels else "",
            detail="DOT point has WIREs extending left/right/up/down. T junctions and dotless crosses are ignored.",
        ))
    return findings


def collect_page_files(project_root: str) -> List[Path]:
    root = Path(project_root).expanduser()
    csa_root = root if root.name.lower() == "sch_1" else root / "sch_1"
    if not csa_root.is_dir():
        return []
    files = [p for p in csa_root.glob("*.csa") if PAGE_FILE_RE.match(p.name)]
    return sorted(files, key=lambda p: (_page_no(p), str(p).lower()))


def _decode_csa_bytes(data: bytes) -> str:
    for enc in ("utf-8-sig", "utf-16", "gb18030", "latin-1"):
        try:
            return data.decode(enc)
        except UnicodeDecodeError:
            continue
    return data.decode("utf-8", errors="replace")


def analyze_one_page(
    file_path: str, *, root: str = "",
    circle_mode: str = "center_radius",
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
            text, page_no, circle_mode=circle_mode, include_arcs=include_arcs)
        findings = find_dot_four_way_crosses(wires, dots)
        return PageResult(
            page_no=page_no, page_label=page_label, page_name=page_name,
            file=str(path), relative_file=relative_file,
            cross_count=len(findings), circle_count=len(circles),
            wire_count=len(wires), dot_count=len(dots),
            findings=findings, circles=circles,
        )
    except Exception as exc:
        return PageResult(
            page_no=page_no, page_label=page_label, page_name=page_label,
            file=str(path), relative_file=relative_file,
            cross_count=0, circle_count=0, wire_count=0, dot_count=0,
            findings=[], circles=[], error=str(exc),
        )


# ══════════════════════════════════════════════════════════
# 一、GUI 辅助函数
# ══════════════════════════════════════════════════════════

def _make_tree(parent, columns, height=12):
    outer = tk.Frame(parent)
    tree = ttk.Treeview(outer, columns=columns, show='headings', height=height)
    vsb = ttk.Scrollbar(outer, orient='vertical', command=tree.yview)
    hsb = ttk.Scrollbar(outer, orient='horizontal', command=tree.xview)
    tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
    tree.grid(row=0, column=0, sticky='nsew')
    vsb.grid(row=0, column=1, sticky='ns')
    hsb.grid(row=1, column=0, sticky='ew')
    outer.grid_rowconfigure(0, weight=1)
    outer.grid_columnconfigure(0, weight=1)
    return outer, tree


def _sort_tree(tree, col, reverse: bool):
    items = [(tree.set(iid, col), iid) for iid in tree.get_children('')]
    try:
        items.sort(key=lambda t: (float(t[0]) if t[0] else float('-inf')), reverse=reverse)
    except ValueError:
        items.sort(key=lambda t: _natural_sort_key(t[0]), reverse=reverse)
    for idx, (_, iid) in enumerate(items):
        tree.move(iid, '', idx)
    arrow = ' ▲' if not reverse else ' ▼'
    for c in tree['columns']:
        base = tree.heading(c, 'text').rstrip(' ▲▼')
        tree.heading(c, text=base + arrow if c == col else base,
                     command=lambda _c=c: _sort_tree(tree, _c, c != col or not reverse))


def _fill_tree(tree, rows: list, columns: list = None):
    tree.delete(*tree.get_children())
    if not rows:
        return
    cols = columns or list(rows[0].keys())
    tree['columns'] = cols
    for c in cols:
        tree.heading(c, text=c, anchor='w',
                     command=lambda _c=c: _sort_tree(tree, _c, False))
        tree.column(c, width=min(max(len(c) * 9, 80), 200), anchor='w', stretch=True)
    for row in rows:
        tree.insert('', 'end', values=[str(row.get(c, '')) for c in cols])


# ══════════════════════════════════════════════════════════
# 二、主 GUI 类
# ══════════════════════════════════════════════════════════

class CsaCheckerApp(tk.Tk):

    def __init__(self):
        super().__init__()
        self.title('CSA 几何规范检查工具 v1.0')
        self.geometry('1100x700')
        self.minsize(860, 500)

        self.proj_root = tk.StringVar()
        self.include_arcs_var = tk.BooleanVar(value=True)
        self.circle_mode_var = tk.StringVar(value='center_radius')

        self._results: List[PageResult] = []
        self._build_ui()

    def _section(self, parent, title):
        f = ttk.LabelFrame(parent, text=title, padding=8)
        f.pack(fill='x', padx=10, pady=4)
        return f

    def _build_ui(self):
        # ── 顶部控制区 ──
        ctrl = self._section(self, '项目选择')
        row = tk.Frame(ctrl); row.pack(fill='x')
        tk.Label(row, text='项目根目录（含 sch_1/ 的那一层）：').pack(side='left')
        ttk.Entry(row, textvariable=self.proj_root, width=52).pack(side='left', padx=6)
        ttk.Button(row, text='浏览…', command=self._browse).pack(side='left')

        row2 = tk.Frame(ctrl); row2.pack(fill='x', pady=(4, 0))
        ttk.Checkbutton(row2, text='包含 ARC 拟合圆检测（需人工确认）',
                        variable=self.include_arcs_var).pack(side='left')
        tk.Label(row2, text='   CIRCLE 解析模式：').pack(side='left', padx=(16, 0))
        mode_cb = ttk.Combobox(row2, textvariable=self.circle_mode_var,
                               values=['center_radius', 'bbox'], width=14, state='readonly')
        mode_cb.pack(side='left', padx=4)

        ttk.Button(ctrl, text='开始分析', command=self._run_analysis).pack(pady=(8, 0))

        # ── 结果表格区 ──
        nb = ttk.Notebook(self)
        nb.pack(fill='both', expand=True, padx=10, pady=6)

        for title, cols, attr in [
            ('概要', ['页面', 'CSA页名', 'DOT四向十字数', '画圈对象数', 'WIRE数', 'DOT数', '错误'], '_tree_sum'),
            ('DOT四向十字', ['页面', 'CSA页名', '序号', '坐标', 'X', 'Y', 'DOT行号',
                             '水平WIRE行号', '垂直WIRE行号', '关联信号'], '_tree_cross'),
            ('画圈对象', ['页面', 'CSA页名', '序号', '对象类型', '行号', '圆心', '半径', '直径',
                         '外接框', '宽', '高', '解析说明'], '_tree_circle'),
        ]:
            f = ttk.Frame(nb); nb.add(f, text=f'  {title}  ')
            outer, tree = _make_tree(f, cols, height=16)
            outer.pack(fill='both', expand=True)
            setattr(self, attr, tree)

        # ── 状态栏 ──
        bot = tk.Frame(self); bot.pack(fill='x', padx=10, pady=4)
        self.status = tk.Label(bot, text='请选择项目根目录开始分析', fg='#888')
        self.status.pack(side='left')
        ttk.Button(bot, text='导出 CSV', command=self._export_csv).pack(side='right')

    def _browse(self):
        folder = filedialog.askdirectory(title='选择项目根目录（包含 sch_1/ 子目录）')
        if folder:
            self.proj_root.set(folder)

    def _run_analysis(self):
        root = self.proj_root.get().strip()
        if not root:
            messagebox.showwarning('提示', '请先选择项目根目录')
            return
        if not os.path.isdir(root):
            messagebox.showerror('错误', f'目录不存在：{root}')
            return

        self.status.configure(text='分析中…', fg='#2d6cdf')
        self.update()

        try:
            files = collect_page_files(root)
            if not files:
                messagebox.showinfo('结果', f'在 {root}/sch_1/ 下未找到 page*.csa 文件')
                self.status.configure(text='未找到 CSA 文件', fg='#888')
                return

            csa_root = files[0].parent
            self._results = [
                analyze_one_page(str(f), root=str(csa_root),
                                 circle_mode=self.circle_mode_var.get(),
                                 include_arcs=self.include_arcs_var.get())
                for f in files
            ]

            total_crosses = sum(r.cross_count for r in self._results)
            total_circles = sum(r.circle_count for r in self._results)
            errors = sum(1 for r in self._results if r.error)
            msg = (f'完成：{len(self._results)} 页, '
                   f'{total_crosses} 个四向十字, '
                   f'{total_circles} 个画圈对象')
            if errors:
                msg += f', {errors} 页解析异常'
            self.status.configure(text=msg, fg='#2a8a2a' if not errors else '#b06000')

            self._refresh_tables()
        except Exception as e:
            import traceback
            messagebox.showerror('错误', f'{e}\n\n{traceback.format_exc()}')
            self.status.configure(text='分析失败', fg='red')

    def _refresh_tables(self):
        # 概要表
        sum_rows = []
        for r in self._results:
            sum_rows.append({
                '页面': r.page_label, 'CSA页名': r.page_name,
                '文件': r.relative_file, 'DOT四向十字数': r.cross_count,
                '画圈对象数': r.circle_count, 'WIRE数': r.wire_count,
                'DOT数': r.dot_count, '错误': r.error,
            })
        _fill_tree(self._tree_sum, sum_rows,
                   ['页面', 'CSA页名', '文件', 'DOT四向十字数', '画圈对象数', 'WIRE数', 'DOT数', '错误'])

        # DOT 四向十字表
        cross_rows = []
        for r in self._results:
            for idx, item in enumerate(r.findings, start=1):
                cross_rows.append({
                    '页面': r.page_label, 'CSA页名': r.page_name, '文件': r.relative_file,
                    '序号': idx, '坐标': f'({item.x},{item.y})',
                    'X': item.x, 'Y': item.y, 'DOT行号': item.dot_line,
                    '水平WIRE行号': item.h_wire_lines, '垂直WIRE行号': item.v_wire_lines,
                    '全部WIRE行号': item.all_wire_lines, '关联信号': item.labels,
                    '说明': item.detail,
                })
        _fill_tree(self._tree_cross, cross_rows,
                   ['页面', 'CSA页名', '文件', '序号', '坐标', 'X', 'Y', 'DOT行号',
                    '水平WIRE行号', '垂直WIRE行号', '关联信号', '说明'])

        # 画圈对象表
        circle_rows = []
        for r in self._results:
            for idx, item in enumerate(r.circles, start=1):
                circle_rows.append({
                    '页面': r.page_label, 'CSA页名': r.page_name, '文件': r.relative_file,
                    '序号': idx, '对象类型': item.object_type, '行号': item.line_no,
                    '圆心': f'({_fmt_num(item.center_x)},{_fmt_num(item.center_y)})',
                    '半径': _fmt_num(item.radius), '直径': _fmt_num(item.diameter),
                    '外接框': f'({_fmt_num(item.bbox_xmin)},{_fmt_num(item.bbox_ymin)})-'
                             f'({_fmt_num(item.bbox_xmax)},{_fmt_num(item.bbox_ymax)})',
                    '宽': _fmt_num(item.width), '高': _fmt_num(item.height),
                    '解析说明': item.parse_note, '原始行': item.raw,
                })
        _fill_tree(self._tree_circle, circle_rows,
                   ['页面', 'CSA页名', '文件', '序号', '对象类型', '行号', '圆心', '半径', '直径',
                    '外接框', '宽', '高', '解析说明'])

    def _export_csv(self):
        if not self._results:
            messagebox.showwarning('提示', '请先执行分析')
            return
        out = filedialog.asksaveasfilename(
            title='导出 CSA 分析结果', initialfile='csa_分析结果.csv',
            defaultextension='.csv', filetypes=[('CSV 文件', '*.csv')])
        if not out:
            return

        try:
            with open(out, 'w', newline='', encoding='utf-8-sig') as f:
                writer = csv.writer(f)
                # 概要
                writer.writerow(['=== CSA 几何规范检查 概要 ==='])
                writer.writerow(['页面', 'CSA页名', '文件', 'DOT四向十字数', '画圈对象数', 'WIRE数', 'DOT数', '错误'])
                for r in self._results:
                    writer.writerow([r.page_label, r.page_name, r.relative_file,
                                     r.cross_count, r.circle_count, r.wire_count, r.dot_count, r.error])

                # DOT 四向十字
                writer.writerow([])
                writer.writerow(['=== DOT 四向十字详细 ==='])
                writer.writerow(['页面', '序号', '坐标', 'X', 'Y', 'DOT行号',
                                 '水平WIRE行号', '垂直WIRE行号', '关联信号', '说明'])
                for r in self._results:
                    for idx, item in enumerate(r.findings, start=1):
                        writer.writerow([r.page_label, idx, f'({item.x},{item.y})',
                                         item.x, item.y, item.dot_line,
                                         item.h_wire_lines, item.v_wire_lines, item.labels, item.detail])

                # 画圈对象
                writer.writerow([])
                writer.writerow(['=== 画圈对象详细 ==='])
                writer.writerow(['页面', '序号', '对象类型', '行号', '圆心', '半径', '直径',
                                 '外接框', '宽', '高', '解析说明'])
                for r in self._results:
                    for idx, item in enumerate(r.circles, start=1):
                        writer.writerow([r.page_label, idx, item.object_type, item.line_no,
                                         f'({_fmt_num(item.center_x)},{_fmt_num(item.center_y)})',
                                         _fmt_num(item.radius), _fmt_num(item.diameter),
                                         f'({_fmt_num(item.bbox_xmin)},{_fmt_num(item.bbox_ymin)})-'
                                         f'({_fmt_num(item.bbox_xmax)},{_fmt_num(item.bbox_ymax)})',
                                         _fmt_num(item.width), _fmt_num(item.height), item.parse_note])
            self.status.configure(text=f'导出成功：{out}', fg='#2a8a2a')
        except Exception as e:
            messagebox.showerror('错误', f'导出失败：{e}')


if __name__ == '__main__':
    app = CsaCheckerApp()
    app.mainloop()
