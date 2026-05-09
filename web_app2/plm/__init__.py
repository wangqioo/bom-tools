# -*- coding: utf-8 -*-
"""PLM 上传工具 — Blueprint"""

import os, uuid, re
from zipfile import BadZipFile
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
    header_row = int(request.form.get('header_row', 4))

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
    header_row = int(request.form.get('header_row', 4))
    col_hqpn_str = request.form.get('col_hqpn', '')
    col_stype_str = request.form.get('col_stype', '')
    col_qty_str = request.form.get('col_qty', '')
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
    if not all([col_hqpn_str, col_stype_str, col_qty_str]):
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
    col_qty = _col_int(col_qty_str)

    if not all([col_hqpn, col_stype, col_qty]):
        return jsonify({
            'success': False,
            'error': '请指定有效的 HQ PN 列、主二供列、用量列',
        })

    safe_proj = re.sub(r'[\\/*?:"<>|]', '_', project_name or '未命名')
    out_name = f"PLM导入_{safe_proj}_{uid}.xlsx"
    out_path = os.path.join(OUTPUT_DIR, out_name)

    total, skipped, skip_logs = _do_convert(
        in_path, sheet_name, header_row,
        col_hqpn, col_stype, col_qty, project_name, out_path,
    )

    return jsonify({
        'success': True,
        'download': f'/download/{out_name}',
        'total': total,
        'skipped': skipped,
        'skip_logs': skip_logs,
    })
