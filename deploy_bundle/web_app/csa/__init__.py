# -*- coding: utf-8 -*-
"""CSA 几何检查工具 — Blueprint"""

import os, re
from pathlib import Path

from shared import (
    render_template, request, jsonify,
    UPLOAD_DIR,
    FEISHU_PRESET_TABLES,
)
from flask import Blueprint

csa_bp = Blueprint('csa_tool', __name__)

Point = tuple
PAGE_FILE_RE = re.compile(r"^page(\d+)\.csa$", re.IGNORECASE)


def _csa_page_no(path):
    m = PAGE_FILE_RE.match(os.path.basename(str(path)))
    return int(m.group(1)) if m else 10 ** 12


def _csa_page_label(no):
    return f"PAGE{no}" if no < 10 ** 12 else "UNKNOWN"


def _csa_parse_text(text, page_no):
    wires = []
    dots = []
    wire_re = re.compile(
        r"\bWIRE\s+\S+\s+\S+\s+\((-?\d+)\s*,?\s*(-?\d+)\)\s*\((-?\d+)\s*,?\s*(-?\d+)\)",
        re.IGNORECASE)
    dot_re = re.compile(
        r"(?<![A-Za-z0-9_])DOT\s+\S+(?:\s+\S+)?\s+\((-?\d+)\s*,?\s*(-?\d+)\)",
        re.IGNORECASE)

    for line_no, raw_line in enumerate(text.splitlines(), start=1):
        raw = raw_line.strip()
        if not raw:
            continue
        for m in wire_re.finditer(raw):
            x1, y1, x2, y2 = map(int, m.groups())
            wires.append({
                'wid': len(wires), 'line_no': line_no,
                'x1': x1, 'y1': y1, 'x2': x2, 'y2': y2,
                'sig': '', 'raw': m.group(0).strip(),
            })
        for m in dot_re.finditer(raw):
            x, y = map(int, m.groups()[-2:])
            dots.append({'line_no': line_no, 'x': x, 'y': y, 'raw': m.group(0).strip()})

    findings = []
    required = {'left', 'right', 'down', 'up'}
    for dot in dots:
        dx, dy = dot['x'], dot['y']
        touching = [
            w for w in wires
            if ((w['y1'] == w['y2'] == dy and min(w['x1'], w['x2']) <= dx <= max(w['x1'], w['x2']))
                or (w['x1'] == w['x2'] == dx and min(w['y1'], w['y2']) <= dy <= max(w['y1'], w['y2'])))
        ]
        if not touching:
            continue
        dirs = set()
        for w in touching:
            if w['y1'] == w['y2']:  # horizontal
                if min(w['x1'], w['x2']) < dx:
                    dirs.add('left')
                if dx < max(w['x1'], w['x2']):
                    dirs.add('right')
            else:  # vertical
                if min(w['y1'], w['y2']) < dy:
                    dirs.add('down')
                if dy < max(w['y1'], w['y2']):
                    dirs.add('up')
        if required.issubset(dirs):
            findings.append({
                'x': dx, 'y': dy, 'dot_line': dot['line_no'], 'wires': len(touching),
            })
    return len(wires), len(dots), findings


def _analyze_csa_dir(project_dir):
    root = Path(project_dir)
    csa_dir = root if root.name.lower() == 'sch_1' else root / 'sch_1'
    if not csa_dir.is_dir():
        raise FileNotFoundError(f"目录不存在：{csa_dir}")
    files = sorted(
        [f for f in csa_dir.glob('*.csa') if PAGE_FILE_RE.match(f.name)],
        key=lambda f: (_csa_page_no(f), str(f).lower()))
    if not files:
        raise FileNotFoundError(f"在 {csa_dir} 中未找到 page*.csa 文件")
    results = []
    for fp in files:
        try:
            data = fp.read_bytes()
            for enc in ['utf-8-sig', 'utf-16', 'gb18030', 'latin-1']:
                try:
                    text = data.decode(enc)
                    break
                except Exception:
                    continue
            else:
                text = data.decode('utf-8', errors='replace')
            wc, dc, findings = _csa_parse_text(text, _csa_page_no(fp))
            results.append({
                'page': _csa_page_label(_csa_page_no(fp)),
                'file': fp.name, 'wires': wc, 'dots': dc,
                'crosses': len(findings), 'findings': findings,
            })
        except Exception as e:
            results.append({
                'page': _csa_page_label(_csa_page_no(fp)),
                'file': fp.name, 'wires': 0, 'dots': 0,
                'crosses': 0, 'findings': [], 'error': str(e),
            })
    total_crosses = sum(r['crosses'] for r in results)
    total_errors = sum(1 for r in results if r.get('error'))
    return results, total_crosses, total_errors


# ── 路由 ─────────────────────────────────────────────────────

@csa_bp.route('/csa', methods=['GET', 'POST'])
def tool_csa():
    if request.method == 'POST':
        files_upload = request.files.getlist('files')
        if files_upload and any(f.filename for f in files_upload):
            uid = str(uuid.uuid4())[:8]
            proj_dir = os.path.join(UPLOAD_DIR, f"csa_proj_{uid}")
            os.makedirs(os.path.join(proj_dir, 'sch_1'), exist_ok=True)
            for f in files_upload:
                if f.filename:
                    f.save(os.path.join(proj_dir, 'sch_1', os.path.basename(f.filename)))
            try:
                results, total_crosses, total_errors = _analyze_csa_dir(proj_dir)
                return jsonify({
                    'success': True, 'results': results,
                    'total_crosses': total_crosses,
                    'total_errors': total_errors, 'pages': len(results),
                })
            except Exception as e:
                return jsonify({'success': False, 'error': str(e)})
        else:
            return jsonify({'success': False, 'error': '请上传 page*.csa 文件'})
    return render_template('index.html', tables=FEISHU_PRESET_TABLES)
