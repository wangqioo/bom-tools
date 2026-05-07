# -*- coding: utf-8 -*-
"""PLM 上传转换工具 — Blueprint"""

import os, uuid, json

from shared import (
    openpyxl, Workbook, Font, PatternFill, Alignment, Border, Side,
    get_column_letter,
    render_template, request, jsonify,
    UPLOAD_DIR, OUTPUT_DIR, _col_int,
    FEISHU_PRESET_TABLES,
)
from flask import Blueprint

plm_bp = Blueprint('plm_tool', __name__)

PLM_HEADERS = [
    "序号", "料号", "型号", "物料描述", "单耗",
    "替代关系\n(A:完全替代/N:独供/X:不完全替代)", "位号", "生产厂家", "是否环保",
    "温敏属性", "备注",
    "主辅BOM标记\n(仅允许填写二供/三供/四供/五供/六供/七供/八供)",
    "MBG优选属性", "CBG优选属性", "DBG优选属性", "首制程", "次制程", "次制程单耗",
    "是否可量产下单", "次制程位号", "ABG优选属性", "IFM_PART", "PCD_PART",
    "是否受EAR管控", "ECCN",
]


def _detect_columns_plm(ws, header_row):
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


def _do_plm_convert(in_file, sheet_name, header_row, col_hqpn, col_stype, col_qty, project_name, out_file):
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
    bdr = Border(left=Side('thin'), right=Side('thin'), top=Side('thin'), bottom=Side('thin'))
    meta_font = Font(bold=True, size=10)

    ws_out.cell(row=1, column=1, value="料号:").font = meta_font
    ws_out.cell(row=1, column=3, value="描述:").font = meta_font
    ws_out.cell(row=1, column=5, value="项目配置名:").font = meta_font
    ws_out.cell(row=1, column=6, value=project_name or "").font = Font(size=10)
    ws_out.cell(row=1, column=7, value="工程师:").font = meta_font
    ws_out.cell(row=2, column=1, value="版本:").font = meta_font
    ws_out.cell(row=2, column=3, value="替代项").font = meta_font
    ws_out.cell(row=2, column=5, value="BOM名称:").font = meta_font
    ws_out.cell(row=2, column=7, value="归档部门:").font = meta_font

    for offset, hdr_txt in enumerate(PLM_HEADERS):
        c = ws_out.cell(row=3, column=offset + 1, value=hdr_txt)
        c.font = Font(bold=True, color='FF0000', size=9)
        c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        c.border = bdr
        ws_out.column_dimensions[get_column_letter(offset + 1)].width = 14
    ws_out.column_dimensions['B'].width = 22
    ws_out.row_dimensions[3].height = 60

    dr = 4
    seq = 0
    total = 0
    skipped = 0
    skip_logs = []
    for rv in data_rows:
        hqpn = str(rv.get(col_hqpn) or '').strip()
        if not hqpn:
            continue

        qty_raw = rv.get(col_qty)
        if qty_raw is None or str(qty_raw).strip() == '':
            skipped += 1
            skip_logs.append(f"  跳过（用量为空）: {hqpn}")
            continue
        try:
            qty = float(qty_raw)
        except ValueError:
            qty = None
        if qty is None:
            skipped += 1
            skip_logs.append(f"  跳过（用量为空）: {hqpn}")
            continue

        stype = str(rv.get(col_stype) or '').strip()
        is_primary = (stype == '主供' or stype == '')

        if is_primary:
            seq += 1

        def w(idx, val):
            cell = ws_out.cell(row=dr, column=idx + 1, value=val)
            cell.border = bdr

        w(0, seq)
        w(1, hqpn)
        if is_primary and qty > 0:
            w(4, qty)
        dr += 1
        total += 1

    wb_out.save(out_file)
    return total, skipped, skip_logs


# ── 路由 ─────────────────────────────────────────────────────

@plm_bp.route('/api/plm/sheets', methods=['POST'])
def api_plm_sheets():
    file = request.files.get('file')
    if not file:
        return jsonify({'success': False, 'error': '请上传文件'})
    uid = str(uuid.uuid4())[:8]
    in_path = os.path.join(UPLOAD_DIR, f"plm_pre_{uid}.xlsx")
    file.save(in_path)
    wb = openpyxl.load_workbook(in_path, read_only=True, data_only=True)
    sheets = wb.sheetnames
    wb.close()
    sheet_name = request.form.get('sheet_name', '')
    if not sheet_name or sheet_name not in sheets:
        sheet_name = sheets[0] if sheets else ''
    wb2 = openpyxl.load_workbook(in_path, data_only=True)
    ws = wb2[sheet_name] if sheet_name else wb2[wb2.sheetnames[0]]
    header_row = int(request.form.get('header_row', 4))
    detected, raw_headers = _detect_columns_plm(ws, header_row)
    wb2.close()
    result = {}
    for k, ci in detected.items():
        col_letter = get_column_letter(ci) if ci else ''
        if col_letter:
            result[k] = col_letter
    return jsonify({
        'success': True, 'sheets': sheets, 'current_sheet': sheet_name,
        'detected': result, 'headers': raw_headers,
    })


@plm_bp.route('/plm', methods=['GET', 'POST'])
def tool_plm():
    if request.method == 'POST':
        file = request.files.get('file')
        if not file:
            return "请上传文件", 400
        sheet_name = request.form.get('sheet', '')
        header_row = int(request.form.get('header_row', 4))
        project_name = request.form.get('project_name', '')
        col_hqpn_str = request.form.get('col_hqpn', '')
        col_stype_str = request.form.get('col_stype', '')
        col_qty_str = request.form.get('col_qty', '')

        uid = str(uuid.uuid4())[:8]
        in_path = os.path.join(UPLOAD_DIR, f"plm_in_{uid}.xlsx")
        out_path = os.path.join(OUTPUT_DIR, f"PLM导入BOM_{uid}.xlsx")
        file.save(in_path)

        wb = openpyxl.load_workbook(in_path, read_only=True, data_only=True)
        sheets = wb.sheetnames
        wb.close()
        if not sheet_name or sheet_name not in sheets:
            sheet_name = sheets[0]

        wb2 = openpyxl.load_workbook(in_path, data_only=True)
        ws = wb2[sheet_name]
        detected, raw_headers = _detect_columns_plm(ws, header_row)
        wb2.close()

        if not col_hqpn_str and 'hq_pn' in detected:
            col_hqpn_str = str(detected['hq_pn'])
        if not col_stype_str and 'supply_type' in detected:
            col_stype_str = str(detected['supply_type'])
        if not col_qty_str and 'qty' in detected:
            col_qty_str = str(detected['qty'])

        col_hqpn = _col_int(col_hqpn_str)
        col_stype = _col_int(col_stype_str)
        col_qty = _col_int(col_qty_str)
        if not all([col_hqpn, col_stype, col_qty]):
            return jsonify({
                'success': False, 'error': '请指定有效的 HQ PN列、主二供列、用量列',
                'detected': detected, 'headers': raw_headers,
            })
        total, skipped, skip_logs = _do_plm_convert(
            in_path, sheet_name, header_row, col_hqpn, col_stype, col_qty, project_name, out_path)
        return jsonify({
            'success': True, 'total': total, 'skipped': skipped, 'skip_logs': skip_logs,
            'download': f'/download/PLM导入BOM_{uid}.xlsx',
            'sheets': sheets, 'detected': detected, 'headers': raw_headers,
        })
    return render_template('index.html', tables=FEISHU_PRESET_TABLES)
