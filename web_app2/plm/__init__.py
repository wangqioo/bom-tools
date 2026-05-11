# -*- coding: utf-8 -*-
"""PLM 上传工具 — Blueprint"""

import os, uuid, re, json
from zipfile import BadZipFile, ZipFile, ZIP_DEFLATED
from flask import Blueprint
from shared import (
    openpyxl, Workbook, Font, PatternFill, Alignment, Border, Side,
    get_column_letter,
    request, jsonify,
    UPLOAD_DIR, OUTPUT_DIR, _col_int,
)

plm_bp = Blueprint('plm', __name__)

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

BAD_EXCEL_ERROR = '无法读取文件，可能原因：① 文件是 .xls 旧格式（请另存为 .xlsx）；② 公司加解密软件未启动导致文件被加密，请检查后重试'


def _request_int(name, default=1, min_value=1):
    try:
        value = int(request.form.get(name, default))
    except (TypeError, ValueError):
        return None
    if min_value is not None and value < min_value:
        return None
    return value




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
                col_hqpn, col_stype, col_qty, project_name, out_file):
    wb_in = openpyxl.load_workbook(in_file, data_only=True)
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
    ws_out.cell(row=1, column=3, value="描述:").font = meta_font
    ws_out.cell(row=1, column=5, value="项目配置名:").font = meta_font
    ws_out.cell(row=1, column=6, value=project_name or "").font = Font(size=10)
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
    seq = 0

    for rv in data_rows:
        hqpn = rv.get(col_hqpn)
        if not hqpn or str(hqpn).strip() == "":
            continue

        qty = _safe_qty(rv.get(col_qty))
        if qty is None:
            skipped += 1
            skip_logs.append(f"  跳过（用量为空）: {str(hqpn).strip()}")
            continue

        stype_str = str(rv.get(col_stype) or "").strip()
        is_primary = (stype_str == "主供" or stype_str == "")

        if is_primary:
            seq += 1

        def wc(idx, val, row=dr):
            cc = ws_out.cell(row=row, column=idx + 1, value=val)
            cc.alignment = Alignment(horizontal="left", vertical="center")
            cc.border = bdr

        wc(PLM_IDX_SEQ, seq)
        wc(PLM_IDX_HQPN, str(hqpn).strip())

        if is_primary and qty > 0:
            wc(PLM_IDX_QTY, qty)

        # For secondary/tertiary supply, fill 主辅BOM标记
        if not is_primary and stype_str:
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
    if file.filename and file.filename.lower().endswith('.xls') and not file.filename.lower().endswith('.xlsx'):
        return jsonify({'success': False, 'error': '不支持 .xls 格式，请在 Excel 中另存为 .xlsx 后重试'})
    uid = str(uuid.uuid4())[:8]
    in_path = os.path.join(UPLOAD_DIR, f"plm_pre_{uid}.xlsx")
    file.save(in_path)

    try:
        wb = openpyxl.load_workbook(in_path, read_only=True, data_only=True)
    except BadZipFile:
        return jsonify({'success': False, 'error': '无法读取文件，可能原因：① 文件是 .xls 旧格式（请另存为 .xlsx）；② 公司加解密软件未启动导致文件被加密，请检查后重试'})
    sheets = wb.sheetnames
    wb.close()

    sheet_name = request.form.get('sheet_name', '')
    if not sheet_name or sheet_name not in sheets:
        sheet_name = sheets[0] if sheets else ''
    header_row = _request_int('header_row', 4)
    if header_row is None:
        return jsonify({'success': False, 'error': '表头行必须是大于等于 1 的数字'})

    wb2 = openpyxl.load_workbook(in_path, data_only=True)
    ws = wb2[sheet_name] if sheet_name else wb2[wb2.sheetnames[0]]
    detected, raw_headers = _detect_columns(ws, header_row)

    # Preview
    preview_headers = [ws.cell(row=header_row, column=ci).value for ci in range(1, ws.max_column + 1)]
    preview = []
    for ri in range(header_row + 1, min(header_row + 4, ws.max_row + 1)):
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
def api_plm_convert():
    file = request.files.get('file')
    if not file:
        return jsonify({'success': False, 'error': '请上传文件'})

    if file.filename and file.filename.lower().endswith('.xls') and not file.filename.lower().endswith('.xlsx'):
        return jsonify({'success': False, 'error': '不支持 .xls 格式，请在 Excel 中另存为 .xlsx 后重试'})
    uid = str(uuid.uuid4())[:8]
    in_path = os.path.join(UPLOAD_DIR, f"plm_in_{uid}.xlsx")
    file.save(in_path)

    sheet_name = request.form.get('sheet', '')
    header_row = _request_int('header_row', 4)
    if header_row is None:
        return jsonify({'success': False, 'error': '表头行必须是大于等于 1 的数字'})
    col_hqpn_str = request.form.get('col_hqpn', '')
    col_stype_str = request.form.get('col_stype', '')
    col_qty_str = request.form.get('col_qty', '')
    qty_configs_str = request.form.get('qty_configs', '')
    project_name = request.form.get('project_name', '')

    try:
        wb = openpyxl.load_workbook(in_path, read_only=True, data_only=True)
    except BadZipFile:
        return jsonify({'success': False, 'error': '无法读取文件，可能原因：① 文件是 .xls 旧格式（请另存为 .xlsx）；② 公司加解密软件未启动导致文件被加密，请检查后重试'})
    sheets = wb.sheetnames
    wb.close()
    if not sheet_name or sheet_name not in sheets:
        sheet_name = sheets[0]

    # Auto-detect if columns not specified
    if not all([col_hqpn_str, col_stype_str]) or (not col_qty_str and not qty_configs_str):
        wb2 = openpyxl.load_workbook(in_path, data_only=True)
        ws = wb2[sheet_name]
        detected, raw_headers = _detect_columns(ws, header_row)
        wb2.close()
        if not col_hqpn_str and 'hq_pn' in detected:
            col_hqpn_str = get_column_letter(detected['hq_pn'])
        if not col_stype_str and 'supply_type' in detected:
            col_stype_str = get_column_letter(detected['supply_type'])
        if not col_qty_str and 'qty' in detected:
            col_qty_str = get_column_letter(detected['qty'])

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

        wb_hdr = openpyxl.load_workbook(in_path, read_only=True, data_only=True)
        ws_hdr = wb_hdr[sheet_name]
        for col_qty in col_qty_list:
            header_val = ws_hdr.cell(row=header_row, column=col_qty).value
            qty_project_name = str(header_val or '').strip() or f"\u7528\u91cf{get_column_letter(col_qty)}"
            if project_name.strip() and len(col_qty_list) == 1:
                qty_project_name = project_name.strip()
            qty_jobs.append((col_qty, qty_project_name))
        wb_hdr.close()

    if not all([col_hqpn, col_stype]) or not qty_jobs:
        return jsonify({
            'success': False,
            'error': '\u8bf7\u6307\u5b9a\u6709\u6548\u7684 HQ PN \u5217\u3001\u4e3b\u4e8c\u4f9b\u5217\u3001\u7528\u91cf\u5217\uff08\u53ef\u6dfb\u52a0\u591a\u4e2a\u7528\u91cf\u914d\u7f6e\uff09',
        })

    if len(qty_jobs) == 1:
        col_qty, qty_project_name = qty_jobs[0]
        safe_proj = _safe_filename_part(qty_project_name)
        out_name = f"PLM\u5bfc\u5165_{safe_proj}_{uid}.xlsx"
        out_path = os.path.join(OUTPUT_DIR, out_name)
        total, skipped, skip_logs = _do_convert(
            in_path, sheet_name, header_row,
            col_hqpn, col_stype, col_qty, qty_project_name, out_path,
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
                col_hqpn, col_stype, col_qty, qty_project_name, out_path,
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
    col_name   = (cfg.get('col_name') or '').strip()
    if not col_name:
        return jsonify({'success': False, 'error': '未指定提取列'})

    uid = str(uuid.uuid4())[:8]
    path = os.path.join(UPLOAD_DIR, f'se_{uid}.xlsx')
    f.save(path)

    try:
        from shared import _cell_str
        wb = openpyxl.load_workbook(path, data_only=True)
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

        values = []
        for ri in range(header_row + 1, ws.max_row + 1):
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
                        'count': len(values)})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)})
