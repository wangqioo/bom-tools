# -*- coding: utf-8 -*-
"""PLM 上传工具 — Blueprint"""

import os, uuid, re, json, threading, time
from zipfile import ZipFile, ZIP_DEFLATED
from flask import Blueprint
from activity import track_tool_activity
from shared import (
    openpyxl, Workbook, Font, PatternFill, Alignment, Border, Side,
    get_column_letter,
    request, jsonify,
    UPLOAD_DIR, OUTPUT_DIR, _col_int,
    _open_workbook, _request_int, _save_uploaded_excel,
)

plm_bp = Blueprint('plm', __name__)

_PLM_ATTACHMENT_JOBS = {}
_PLM_ATTACHMENT_JOBS_LOCK = threading.Lock()


def _new_attachment_job(hqpn):
    job_id = uuid.uuid4().hex
    now = time.time()
    job = {
        'id': job_id,
        'status': 'queued',
        'stage': '\u4efb\u52a1\u5df2\u521b\u5efa',
        'progress': 3,
        'hqpn': hqpn,
        'logs': [],
        'download': '',
        'filename': '',
        'source_path': '',
        'error': '',
        'created_at': now,
        'updated_at': now,
    }
    with _PLM_ATTACHMENT_JOBS_LOCK:
        _PLM_ATTACHMENT_JOBS[job_id] = job
    return job_id


def _attachment_progress_from_message(message, current):
    text = str(message or '')
    rules = [
        ('\u542f\u52a8\u6d4f\u89c8\u5668', 8, '\u542f\u52a8\u6d4f\u89c8\u5668'),
        ('\u6253\u5f00 EIP', 15, '\u6253\u5f00 EIP'),
        ('\u8fdb\u5165 PLM', 25, '\u8fdb\u5165 PLM'),
        ('Open PLM search page', 35, '\u6253\u5f00 PLM \u641c\u7d22\u9875'),
        ('\u76f4\u63a5\u8fdb\u5165 PLM \u641c\u7d22\u9875', 35, '\u6253\u5f00 PLM \u641c\u7d22\u9875'),
        ('\u641c\u7d22\u6599\u53f7', 45, '\u641c\u7d22 HQ \u6599\u53f7'),
        ('\u6253\u5f00\u7b2c\u4e00\u6761\u641c\u7d22\u7ed3\u679c', 58, '\u6253\u5f00\u7269\u6599\u8be6\u60c5'),
        ('\u8fdb\u5165\u5185\u5bb9\u9875', 68, '\u8fdb\u5165\u5185\u5bb9\u9875'),
        ('\u52fe\u9009\u9644\u4ef6\u5e76\u4e0b\u8f7d', 76, '\u52fe\u9009\u9644\u4ef6'),
        ('\u8bc6\u522b PDF \u9644\u4ef6', 80, '\u8bc6\u522b\u9644\u4ef6'),
        ('Selected all attachment rows', 84, '\u52fe\u9009\u9644\u4ef6\u884c'),
        ('No immediate download event', 88, '\u7b49\u5f85 PDF \u9884\u89c8\u9875'),
        ('Downloaded PDF response', 94, '\u4fdd\u5b58 PDF \u9644\u4ef6'),
        ('Downloaded PDF viewer', 94, '\u4fdd\u5b58 PDF \u9644\u4ef6'),
        ('Downloaded selected attachments', 94, '\u4fdd\u5b58\u9644\u4ef6\u538b\u7f29\u5305'),
        ('\u4e0b\u8f7d\u5b8c\u6210', 98, '\u6574\u7406\u4e0b\u8f7d\u6587\u4ef6'),
    ]
    for needle, progress, stage in rules:
        if needle in text:
            return max(current, progress), stage
    return min(max(current, 5) + 1, 90), text[:80] or '\u5904\u7406\u4e2d'


def _update_attachment_job(job_id, **updates):
    with _PLM_ATTACHMENT_JOBS_LOCK:
        job = _PLM_ATTACHMENT_JOBS.get(job_id)
        if not job:
            return
        job.update(updates)
        job['updated_at'] = time.time()


def _append_attachment_log(job_id, message):
    message = str(message)
    with _PLM_ATTACHMENT_JOBS_LOCK:
        job = _PLM_ATTACHMENT_JOBS.get(job_id)
        if not job:
            return
        job['logs'].append(message)
        progress, stage = _attachment_progress_from_message(message, int(job.get('progress') or 0))
        job['progress'] = progress
        job['stage'] = stage
        job['updated_at'] = time.time()


def _snapshot_attachment_job(job_id):
    with _PLM_ATTACHMENT_JOBS_LOCK:
        job = _PLM_ATTACHMENT_JOBS.get(job_id)
        if not job:
            return None
        return dict(job, logs=list(job.get('logs') or []))


def _cleanup_attachment_jobs():
    cutoff = time.time() - 3600
    with _PLM_ATTACHMENT_JOBS_LOCK:
        for job_id, job in list(_PLM_ATTACHMENT_JOBS.items()):
            if job.get('updated_at', 0) < cutoff:
                _PLM_ATTACHMENT_JOBS.pop(job_id, None)
PLM_HEADERS = [
    "序号", "料号", "型号", "物料描述", "单耗",
    "替代关系\n(A:完全替代/N:独供/X:不完全替代)",
    "位号", "生产厂家", "是否环保", "温敏属性", "备注",
    "主辅BOM标记\n(仅允许填写二供/三供/四供/五供/六供/七供/八供)",
    "MBG优选属性", "CBG优选属性", "DBG优选属性", "首制程", "次制程", "次制程单耗",
    "是否可量产下单", "次制程位号", "ABG优选属性", "IFM_PART", "PCD_PART",
    "是否受EAR管控", "ECCN",
]

PLM_IDX_SEQ  = 0
PLM_IDX_HQPN = 1
PLM_IDX_QTY  = 4
PLM_IDX_MARK = 11  # 主辅BOM标记

def _detect_columns(ws, header_row):
    result = {}
    found_headers = []
    scan_cols = max((ws.max_column or 0) + 5, 30)
    for ci in range(1, scan_cols + 1):
        raw = ws.cell(row=header_row, column=ci).value
        if raw is None:
            continue
        h = str(raw).replace('\n', '').replace('\r', '').strip()
        hl = h.lower().replace(' ', '')
        if h:
            found_headers.append(f"{get_column_letter(ci)}:{h}")
        if '序号' in h:
            result.setdefault('seq', ci)
        if 'hq' in hl and 'pn' in hl:
            result.setdefault('hq_pn', ci)
        if '主二供' in h or '主供' in h:
            result.setdefault('supply_type', ci)
        if '用量' in h or '单耗' in h:
            result.setdefault('qty', ci)
    return result, found_headers


def _safe_qty(v):
    if v is None:
        return None
    s = str(v).strip()
    if s == "":
        return None
    try:
        return float(s)
    except ValueError:
        return None


def _split_col_refs(raw):
    refs = []
    for part in re.split(r'[\s,\uFF0C;\uFF1B]+', str(raw or '').strip()):
        part = part.strip()
        if part:
            refs.append(part)
    return refs


def _safe_filename_part(value):
    text = str(value or '').strip() or '\u672a\u547d\u540d'
    return re.sub(r'[\\/*?:"<>|]', '_', text)


def _do_convert(in_file, sheet_name, header_row,
                col_seq, col_hqpn, col_stype, col_qty, project_name, out_file):
    wb_in = _open_workbook(in_file, data_only=True)
    ws_in = wb_in[sheet_name]
    max_col = ws_in.max_column

    data_rows = []
    for ri in range(header_row + 1, ws_in.max_row + 1):
        rv = {ci: ws_in.cell(row=ri, column=ci).value for ci in range(1, max_col + 1)}
        if any(v is not None and str(v).strip() for v in rv.values()):
            data_rows.append(rv)

    wb_out = Workbook()
    ws_out = wb_out.active
    ws_out.title = "PLM导入"

    bdr = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"), bottom=Side(style="thin"),
    )
    meta_font = Font(bold=True, size=10)

    # Rows 1-2: metadata
    ws_out.cell(row=1, column=1, value="料号:").font = meta_font
    ws_out.cell(row=1, column=2, value=project_name or "").font = Font(size=10)
    ws_out.cell(row=1, column=3, value="描述:").font = meta_font
    ws_out.cell(row=1, column=5, value="项目配置名:").font = meta_font
    ws_out.cell(row=1, column=7, value="工程师:").font = meta_font
    ws_out.cell(row=2, column=1, value="版本:").font = meta_font
    ws_out.cell(row=2, column=3, value="替代项").font = meta_font
    ws_out.cell(row=2, column=5, value="BOM名称:").font = meta_font
    ws_out.cell(row=2, column=7, value="归档部门:").font = meta_font

    # Row 3: headers
    for offset, hdr_txt in enumerate(PLM_HEADERS):
        c = ws_out.cell(row=3, column=offset + 1, value=hdr_txt)
        c.font = Font(bold=True, color="FF0000", size=9)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = bdr
        ws_out.column_dimensions[get_column_letter(offset + 1)].width = 14
    ws_out.column_dimensions[get_column_letter(PLM_IDX_HQPN + 1)].width = 22
    ws_out.row_dimensions[3].height = 60

    # Data rows from row 4
    dr = 4
    total = 0
    skipped = 0
    skip_logs = []
    for rv in data_rows:
        source_seq = rv.get(col_seq)
        if not source_seq or str(source_seq).strip() == "":
            skipped += 1
            skip_logs.append("  跳过（序号为空）")
            continue

        qty_raw = rv.get(col_qty)
        if qty_raw is None or str(qty_raw).strip() == "":
            skipped += 1
            skip_logs.append(f"  跳过（用量为空）: 序号 {str(source_seq).strip()}")
            continue

        qty = _safe_qty(qty_raw)
        if qty is None:
            skipped += 1
            skip_logs.append(f"  跳过（用量非数字）: 序号 {str(source_seq).strip()}")
            continue

        hqpn = rv.get(col_hqpn)
        hqpn_str = str(hqpn).strip() if hqpn is not None else ""
        stype_str = str(rv.get(col_stype) or "").strip()


        def wc(idx, val, row=dr):
            cc = ws_out.cell(row=row, column=idx + 1, value=val)
            cc.alignment = Alignment(horizontal="left", vertical="center")
            cc.border = bdr

        wc(PLM_IDX_SEQ, source_seq)
        wc(PLM_IDX_HQPN, hqpn_str)

        if qty != 0:
            wc(PLM_IDX_QTY, qty)

        if stype_str and stype_str != "主供":
            wc(PLM_IDX_MARK, stype_str)

        dr += 1
        total += 1

    wb_out.save(out_file)
    wb_in.close()
    return total, skipped, skip_logs


# ── 路由 ─────────────────────────────────────────────────────

@plm_bp.route('/api/plm/detect', methods=['POST'])
def api_plm_detect():
    file = request.files.get('file')
    if not file:
        return jsonify({'success': False, 'error': '请上传文件'})
    uid = str(uuid.uuid4())[:8]
    try:
        in_path = _save_uploaded_excel(file, "plm_pre", uid)
        wb = _open_workbook(in_path, read_only=True, data_only=True)
    except ValueError as e:
        return jsonify({'success': False, 'error': str(e)})
    sheets = wb.sheetnames
    wb.close()

    sheet_name = request.form.get('sheet_name', '')
    if not sheet_name or sheet_name not in sheets:
        sheet_name = sheets[0] if sheets else ''
    header_row = _request_int('header_row', 4)
    if header_row is None:
        return jsonify({'success': False, 'error': '表头行必须是大于等于 1 的数字'})

    try:
        wb2 = _open_workbook(in_path, data_only=True)
    except ValueError as e:
        return jsonify({'success': False, 'error': str(e)})
    ws = wb2[sheet_name] if sheet_name else wb2[wb2.sheetnames[0]]
    detected, raw_headers = _detect_columns(ws, header_row)

    # Preview
    preview_headers = [ws.cell(row=header_row, column=ci).value for ci in range(1, ws.max_column + 1)]
    preview = []
    for ri in range(header_row + 1, min(header_row + 51, ws.max_row + 1)):
        row = [ws.cell(row=ri, column=ci).value for ci in range(1, ws.max_column + 1)]
        if any(v is not None and str(v).strip() for v in row):
            preview.append([str(v) if v is not None else "" for v in row])
    wb2.close()

    result = {k: get_column_letter(v) for k, v in detected.items() if v}
    return jsonify({
        'success': True,
        'uid': uid,
        'sheets': sheets,
        'current_sheet': sheet_name,
        'headers': raw_headers,
        'preview_headers': [str(h) if h is not None else "" for h in preview_headers],
        'preview': preview,
        'detected': result,
    })


@plm_bp.route('/api/plm/convert', methods=['POST'])
@track_tool_activity('PLM格式转换')
def api_plm_convert():
    file = request.files.get('file')
    if not file:
        return jsonify({'success': False, 'error': '请上传文件'})

    uid = str(uuid.uuid4())[:8]
    try:
        in_path = _save_uploaded_excel(file, "plm_in", uid)
    except ValueError as e:
        return jsonify({'success': False, 'error': str(e)})

    sheet_name = request.form.get('sheet', '')
    header_row = _request_int('header_row', 4)
    if header_row is None:
        return jsonify({'success': False, 'error': '表头行必须是大于等于 1 的数字'})
    col_seq_str = request.form.get('col_seq', '')
    col_hqpn_str = request.form.get('col_hqpn', '')
    col_stype_str = request.form.get('col_stype', '')
    col_qty_str = request.form.get('col_qty', '')
    qty_configs_str = request.form.get('qty_configs', '')
    project_name = request.form.get('project_name', '')

    try:
        wb = _open_workbook(in_path, read_only=True, data_only=True)
    except ValueError as e:
        return jsonify({'success': False, 'error': str(e)})
    sheets = wb.sheetnames
    wb.close()
    if not sheet_name or sheet_name not in sheets:
        sheet_name = sheets[0]

    # Auto-detect if columns not specified
    if not all([col_seq_str, col_hqpn_str, col_stype_str]) or (not col_qty_str and not qty_configs_str):
        wb2 = _open_workbook(in_path, data_only=True)
        ws = wb2[sheet_name]
        detected, raw_headers = _detect_columns(ws, header_row)
        wb2.close()
        if not col_seq_str and 'seq' in detected:
            col_seq_str = get_column_letter(detected['seq'])
        if not col_hqpn_str and 'hq_pn' in detected:
            col_hqpn_str = get_column_letter(detected['hq_pn'])
        if not col_stype_str and 'supply_type' in detected:
            col_stype_str = get_column_letter(detected['supply_type'])
        if not col_qty_str and 'qty' in detected:
            col_qty_str = get_column_letter(detected['qty'])

    col_seq = _col_int(col_seq_str)
    col_hqpn = _col_int(col_hqpn_str)
    col_stype = _col_int(col_stype_str)

    qty_jobs = []
    if qty_configs_str.strip():
        try:
            qty_configs = json.loads(qty_configs_str)
        except Exception:
            return jsonify({'success': False, 'error': '\u7528\u91cf\u914d\u7f6e\u683c\u5f0f\u9519\u8bef'})
        for cfg in qty_configs if isinstance(qty_configs, list) else []:
            col_qty = _col_int((cfg or {}).get('qty_col', ''))
            if not col_qty:
                continue
            qty_project_name = str((cfg or {}).get('name') or '').strip()
            if not qty_project_name:
                qty_project_name = f"\u7528\u91cf{get_column_letter(col_qty)}"
            qty_jobs.append((col_qty, qty_project_name))
    else:
        col_qty_refs = _split_col_refs(col_qty_str)
        if not col_qty_refs and col_qty_str.strip():
            col_qty_refs = [col_qty_str.strip()]
        col_qty_list = [_col_int(ref) for ref in col_qty_refs]
        col_qty_list = [ci for ci in col_qty_list if ci]

        wb_hdr = _open_workbook(in_path, read_only=True, data_only=True)
        ws_hdr = wb_hdr[sheet_name]
        for col_qty in col_qty_list:
            header_val = ws_hdr.cell(row=header_row, column=col_qty).value
            qty_project_name = str(header_val or '').strip() or f"\u7528\u91cf{get_column_letter(col_qty)}"
            if project_name.strip() and len(col_qty_list) == 1:
                qty_project_name = project_name.strip()
            qty_jobs.append((col_qty, qty_project_name))
        wb_hdr.close()

    if not all([col_seq, col_hqpn, col_stype]) or not qty_jobs:
        return jsonify({
            'success': False,
            'error': '\u8bf7\u6307\u5b9a\u6709\u6548\u7684\u5e8f\u53f7\u5217\u3001HQ PN \u5217\u3001\u4e3b\u4e8c\u4f9b\u5217\u3001\u7528\u91cf\u5217\uff08\u53ef\u6dfb\u52a0\u591a\u4e2a\u7528\u91cf\u914d\u7f6e\uff09',
        })

    if len(qty_jobs) == 1:
        col_qty, qty_project_name = qty_jobs[0]
        safe_proj = _safe_filename_part(qty_project_name)
        out_name = f"PLM\u5bfc\u5165_{safe_proj}_{uid}.xlsx"
        out_path = os.path.join(OUTPUT_DIR, out_name)
        total, skipped, skip_logs = _do_convert(
            in_path, sheet_name, header_row,
            col_seq, col_hqpn, col_stype, col_qty, qty_project_name, out_path,
        )
        return jsonify({
            'success': True,
            'download': f'/download/{out_name}',
            'total': total,
            'skipped': skipped,
            'skip_logs': skip_logs,
            'files': [{'name': out_name, 'project_name': qty_project_name,
                       'qty_col': get_column_letter(col_qty),
                       'total': total, 'skipped': skipped}],
        })

    results = []
    all_skip_logs = []
    zip_name = f"PLM\u5bfc\u5165\u6279\u91cf_{uid}.zip"
    zip_path = os.path.join(OUTPUT_DIR, zip_name)
    used_names = set()
    with ZipFile(zip_path, 'w', ZIP_DEFLATED) as zf:
        for col_qty, qty_project_name in qty_jobs:
            safe_proj = _safe_filename_part(qty_project_name)
            out_name = f"PLM\u5bfc\u5165_{safe_proj}_{uid}.xlsx"
            n = 2
            while out_name in used_names:
                out_name = f"PLM\u5bfc\u5165_{safe_proj}_{uid}_{n}.xlsx"
                n += 1
            used_names.add(out_name)

            out_path = os.path.join(OUTPUT_DIR, out_name)
            total, skipped, skip_logs = _do_convert(
                in_path, sheet_name, header_row,
                col_seq, col_hqpn, col_stype, col_qty, qty_project_name, out_path,
            )
            zf.write(out_path, arcname=out_name)
            results.append({
                'name': out_name,
                'project_name': qty_project_name,
                'qty_col': get_column_letter(col_qty),
                'total': total,
                'skipped': skipped,
            })
            all_skip_logs.extend([f"[{qty_project_name}] {msg}" for msg in skip_logs])

    return jsonify({
        'success': True,
        'download': f'/download/{zip_name}',
        'total': sum(r['total'] for r in results),
        'skipped': sum(r['skipped'] for r in results),
        'skip_logs': all_skip_logs,
        'files': results,
        'is_zip': True,
    })

@plm_bp.route('/api/plm/spec_extract', methods=['POST'])
@track_tool_activity('规格型号提取')
def api_spec_extract():
    """提取单列规格型号，去除空格，输出单列 Excel"""
    import json as _json
    f = request.files.get('file')
    if not f:
        return jsonify({'success': False, 'error': '未上传文件'})
    cfg_str = request.form.get('config', '{}')
    try:
        cfg = _json.loads(cfg_str)
    except Exception:
        return jsonify({'success': False, 'error': 'config 格式错误'})

    try:
        header_row = int(cfg.get('header_row', 1))
    except (TypeError, ValueError):
        return jsonify({'success': False, 'error': '表头行必须是大于等于 1 的数字'})
    if header_row < 1:
        return jsonify({'success': False, 'error': '表头行必须是大于等于 1 的数字'})
    sheet_name = cfg.get('sheet_name', '')
    col_name = (cfg.get('col_name') or '').strip()
    exclude_col_name = (cfg.get('exclude_col_name') or cfg.get('hq_col_name') or '').strip()
    if not col_name:
        return jsonify({'success': False, 'error': '未指定提取列'})

    uid = str(uuid.uuid4())[:8]
    try:
        path = _save_uploaded_excel(f, 'se', uid)
    except ValueError as e:
        return jsonify({'success': False, 'error': str(e)})

    try:
        wb = _open_workbook(path, data_only=True)
        sheets = wb.sheetnames
        if not sheet_name or sheet_name not in sheets:
            sheet_name = sheets[0]
        ws = wb[sheet_name]

        headers = [_cell_str(ws.cell(row=header_row, column=ci).value)
                   for ci in range(1, ws.max_column + 1)]
        if col_name not in headers:
            return jsonify({'success': False,
                            'error': f'列 "{col_name}" 不存在，请检查表头行设置'})
        col_idx = headers.index(col_name) + 1  # 1-based
        exclude_col_idx = None
        if exclude_col_name:
            if exclude_col_name not in headers:
                return jsonify({'success': False,
                                'error': f'剔除列 "{exclude_col_name}" 不存在，请检查表头行设置'})
            exclude_col_idx = headers.index(exclude_col_name) + 1

        values = []
        skipped_excluded = 0
        for ri in range(header_row + 1, ws.max_row + 1):
            if exclude_col_idx is not None and _cell_str(ws.cell(row=ri, column=exclude_col_idx).value):
                skipped_excluded += 1
                continue
            v = _cell_str(ws.cell(row=ri, column=col_idx).value)
            if v:
                values.append(v.replace(' ', '').replace('\u3000', ''))
        wb.close()

        # Write output
        wb_out = Workbook()
        ws_out = wb_out.active
        ws_out.title = '规格型号'
        ws_out.cell(row=1, column=1, value='规格型号').font = Font(bold=True)
        for i, v in enumerate(values, 2):
            ws_out.cell(row=i, column=1, value=v)
        ws_out.column_dimensions['A'].width = 40

        out_name = f'spec_{uid}.xlsx'
        wb_out.save(os.path.join(OUTPUT_DIR, out_name))
        return jsonify({'success': True, 'download': f'/download/{out_name}',
                        'count': len(values), 'skipped_excluded': skipped_excluded})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)})



@plm_bp.route('/api/plm/auto_spec_reverse', methods=['POST'])
@track_tool_activity('PLM规格反查')
def api_auto_spec_reverse():
    """Run an integrated Playwright PLM automation feature."""
    username = (request.form.get('username') or '').strip()
    password = request.form.get('password') or ''
    f = request.files.get('file')
    if not username:
        return jsonify({'success': False, 'error': '请输入账号'})
    if not password:
        return jsonify({'success': False, 'error': '请输入密码'})
    if not f:
        return jsonify({'success': False, 'error': '请选择需要上传的 Excel 文件'})

    uid = str(uuid.uuid4())[:8]
    try:
        in_path = _save_uploaded_excel(f, 'plm_auto', uid)
    except ValueError as e:
        return jsonify({'success': False, 'error': str(e)})

    logs = []

    def add_log(message):
        logs.append(str(message))

    try:
        from pathlib import Path as _Path
        from playwright.sync_api import sync_playwright
        from .automation import require_feature, run_plm_feature

        feature = require_feature('spec_reverse_material')
        with sync_playwright() as playwright:
            output_path = run_plm_feature(
                playwright,
                username=username,
                password=password,
                feature=feature,
                upload_file=_Path(in_path),
                output_dir=_Path(OUTPUT_DIR),
                headless=False,
                log=add_log,
            )
    except ImportError as e:
        return jsonify({
            'success': False,
            'error': '缺少 Playwright 依赖，请在 BOM 工具环境安装 requirements.txt 并执行 playwright install chromium',
            'log': chr(10).join(logs + [str(e)]),
        })
    except Exception as e:
        return jsonify({'success': False, 'error': str(e), 'log': chr(10).join(logs)})

    output_path = str(output_path)
    if not os.path.exists(output_path):
        return jsonify({'success': False, 'error': '自动化完成但未找到导出文件', 'log': chr(10).join(logs)})

    out_name = os.path.basename(output_path)
    return jsonify({
        'success': True,
        'download': f'/download/{out_name}',
        'filename': out_name,
        'source_path': output_path,
        'log': chr(10).join(logs),
    })

@plm_bp.route('/api/plm/auto_hq_attachments', methods=['POST'])
@track_tool_activity('PLM\u9644\u4ef6\u4e0b\u8f7d')
def api_auto_hq_attachments():
    """Start a background PLM attachment download job."""
    username = (request.form.get('username') or '').strip()
    password = request.form.get('password') or ''
    hqpn = (request.form.get('hqpn') or '').strip()
    if not username:
        return jsonify({'success': False, 'error': '\u8bf7\u8f93\u5165\u8d26\u53f7'})
    if not password:
        return jsonify({'success': False, 'error': '\u8bf7\u8f93\u5165\u5bc6\u7801'})
    if not hqpn:
        return jsonify({'success': False, 'error': '\u8bf7\u8f93\u5165 HQ \u6599\u53f7'})

    _cleanup_attachment_jobs()
    job_id = _new_attachment_job(hqpn)

    def run_job():
        _update_attachment_job(job_id, status='running', stage='\u51c6\u5907\u542f\u52a8\u6d4f\u89c8\u5668', progress=5)
        try:
            from pathlib import Path as _Path
            from playwright.sync_api import sync_playwright
            from .automation import run_hq_attachment_download

            with sync_playwright() as playwright:
                output_path = run_hq_attachment_download(
                    playwright,
                    username=username,
                    password=password,
                    hqpn=hqpn,
                    output_dir=_Path(OUTPUT_DIR),
                    headless=False,
                    log=lambda message: _append_attachment_log(job_id, message),
                )
            output_path = str(output_path)
            if not os.path.exists(output_path):
                raise RuntimeError('\u81ea\u52a8\u5316\u5b8c\u6210\u4f46\u672a\u627e\u5230\u4e0b\u8f7d\u6587\u4ef6')
            out_name = os.path.basename(output_path)
            _update_attachment_job(
                job_id,
                status='done',
                stage='\u4e0b\u8f7d\u5b8c\u6210',
                progress=100,
                download=f'/download/{out_name}',
                filename=out_name,
                source_path=output_path,
            )
        except ImportError as e:
            _append_attachment_log(job_id, str(e))
            _update_attachment_job(job_id, status='error', stage='\u7f3a\u5c11 Playwright \u4f9d\u8d56', progress=100, error='\u7f3a\u5c11 Playwright \u4f9d\u8d56\uff0c\u8bf7\u5b89\u88c5 requirements.txt \u5e76\u6267\u884c playwright install chromium')
        except Exception as e:
            _append_attachment_log(job_id, str(e))
            _update_attachment_job(job_id, status='error', stage='\u6267\u884c\u5931\u8d25', progress=100, error=str(e))

    threading.Thread(target=run_job, daemon=True).start()
    return jsonify({'success': True, 'job_id': job_id, 'status_url': f'/api/plm/auto_hq_attachments/status/{job_id}'})


@plm_bp.route('/api/plm/auto_hq_attachments/status/<job_id>', methods=['GET'])
def api_auto_hq_attachments_status(job_id):
    job = _snapshot_attachment_job(job_id)
    if not job:
        return jsonify({'success': False, 'error': '\u4efb\u52a1\u4e0d\u5b58\u5728\u6216\u5df2\u8fc7\u671f'}), 404
    return jsonify({
        'success': True,
        'job_id': job['id'],
        'status': job.get('status'),
        'stage': job.get('stage'),
        'progress': job.get('progress'),
        'hqpn': job.get('hqpn'),
        'download': job.get('download'),
        'filename': job.get('filename'),
        'source_path': job.get('source_path'),
        'error': job.get('error'),
        'log': chr(10).join(job.get('logs') or []),
    })


