# -*- coding: utf-8 -*-
"""BOM 转换工具 — Blueprint"""

import os, uuid, re
from flask import Blueprint, render_template
from activity import track_tool_activity
from shared import (
    openpyxl, Workbook, Font, PatternFill, Alignment, Border, Side,
    get_column_letter, column_index_from_string,
    request, jsonify,
    UPLOAD_DIR, OUTPUT_DIR, _col_int, FEISHU_PRESET_TABLES,
    _open_workbook, _request_int, _save_uploaded_excel,
)

bom_bp = Blueprint('bom', __name__)

SUPPLIER_LABELS = ["主供","二供","三供","四供","五供","六供","七供","八供","九供","十供"]

# ── 供应商解析 ─────────────────────────────────────────────────

def _parse_combined(raw):
    if not raw or str(raw).strip() == "":
        return []
    s = str(raw).strip().replace("：", ":").replace("∥", "||").replace("‖", "||")
    if "||" in s:
        entries = [e.strip() for e in re.split(r"\|\|", s) if e.strip()]
    elif re.search(r'[^\s]+:[^\s]+\s{2,}[^\s]+:[^\s]+', s):
        entries = [e.strip() for e in re.split(r'\s{2,}', s) if e.strip()]
    elif re.search(r'[^\s:]+:[^\s]+(?:\s+[^\s:]+:[^\s]+)+', s):
        entries = [e.strip() for e in re.split(r'\s+(?=[^\s:]+:[^\s])', s) if e.strip()]
    else:
        entries = [s.strip()]
    result = []
    for entry in entries:
        if ":" in entry:
            b, m = entry.split(":", 1)
            result.append((b.strip(), m.strip()))
        elif "/" in entry and len(entry.split("/")) == 2:
            b, m = entry.split("/", 1)
            result.append((b.strip(), m.strip()))
        elif entry:
            result.append(("", entry.strip()))
    return result


def _parse_split(brand_raw, model_raw):
    brands = [b.strip() for b in str(brand_raw or "").split(";") if b.strip()] if brand_raw else []
    models = [m.strip() for m in str(model_raw or "").split(";") if m.strip()] if model_raw else []
    result = []
    for i in range(max(len(brands), len(models), 1)):
        b = brands[i] if i < len(brands) else ""
        m = models[i] if i < len(models) else ""
        if b or m:
            result.append((b, m))
    return result


def _parse_format_c(brand_raw, model_raw):
    brand_names = []
    if brand_raw:
        s = str(brand_raw).strip()
        matches = re.findall(r'\d{4}-([^\[:\]]+)\[', s)
        if matches:
            brand_names = [m.strip() for m in matches]
        else:
            brand_names = [b.strip() for b in s.split(":") if b.strip()]
    models = [m.strip() for m in str(model_raw or "").split(":") if m.strip()] if model_raw else []
    result = []
    for i in range(max(len(brand_names), len(models), 1)):
        b = brand_names[i] if i < len(brand_names) else ""
        m = models[i] if i < len(models) else ""
        if b or m:
            result.append((b, m))
    return result


def _parse_suppliers(bv, mv, fmt):
    if fmt == "C":
        return _parse_format_c(bv, mv)
    if fmt == "B":
        return _parse_split(bv, mv)
    return _parse_combined(bv)


def _safe_qty(qv):
    try:
        q = float(qv)
        return int(q) if q == int(q) else q
    except Exception:
        return qv if qv not in (None, "") else ""


# ── 列自动检测 ─────────────────────────────────────────────────

def _detect_columns(ws, header_row):
    data_rows = list(range(header_row + 1, min(header_row + 11, ws.max_row + 1)))
    all_cols = {}
    for ci in range(1, ws.max_column + 1):
        hv = ws.cell(row=header_row, column=ci).value
        hs = str(hv).strip() if hv else ""
        letter = get_column_letter(ci)
        samples = [ws.cell(row=r, column=ci).value for r in data_rows]
        strs = [str(v).strip() for v in samples if v is not None]
        role = "other"
        score = 0

        b_code = sum(1 for v in strs if re.search(r'\d{4}-[^\[]+\[', v))
        if b_code >= 2 or (any(k in hs for k in ["制造商", "Manufacturer"]) and "型号" not in hs and b_code >= 1):
            role = "brand_code"
            score = b_code * 25 + (50 if "制造商" in hs else 0)

        m_code = sum(1 for v in strs if ":" in v and not re.search(r'\d{4}-[^\[]+\[', v) and "||" not in v)
        if "制造商型号" in hs or "Manufacturer P/N" in hs:
            if role == "other":
                role = "model_code"
                score = 85
        elif m_code >= 3 and role == "other":
            role = "model_code"
            score = m_code * 12

        b_comb = sum(1 for v in strs if "||" in v or re.search(r"[A-Za-z0-9]+:[A-Za-z0-9]", v))
        if role == "other" and (b_comb >= 2 or "品牌型号" in hs):
            role = "brand_combined"
            score = b_comb * 20 + (40 if "品牌型号" in hs else 0)

        b_split = sum(1 for v in strs if ";" in v and not re.search(r"[A-Za-z0-9]+:[A-Za-z0-9]", v))
        if any(k in hs for k in ["厂家", "厂商", "Manufacturer", "Brand"]) and role == "other":
            role = "brand_split"
            score = 80
        elif b_split >= 3 and role == "other":
            role = "brand_split"
            score = b_split * 15

        m_split = sum(1 for v in strs if ";" in v)
        if "型号" in hs and "品牌" not in hs and "制造商" not in hs and role == "other":
            role = "model_split"
            score = 80
        elif m_split >= 3 and role == "other":
            role = "model_split"
            score = m_split * 12

        numeric = sum(1 for v in samples if v is not None and str(v).replace(".", "").isdigit())
        if any(k in hs for k in ["用量", "数量", "qty", "quantity", "Quantity"]):
            role = "qty"
            score = 85
        elif numeric >= len(data_rows) * 0.6 and role == "other":
            role = "qty"
            score = numeric * 10

        avg_len = sum(len(v) for v in strs) / max(len(strs), 1)
        if any(k in hs for k in ["名称", "品名", "物料名", "描述", "项目描述", "description", "Description"]):
            if role == "other":
                role = "name"
                score = 75
        elif avg_len > 8 and role == "other":
            role = "name"
            score = int(avg_len * 2)

        all_cols[ci] = {"letter": letter, "header": hs, "role": role, "score": score, "sample": strs[:3]}

    best = {}
    for ci, info in all_cols.items():
        r = info["role"]
        if r != "other" and (r not in best or info["score"] > best[r]["score"]):
            best[r] = {"ci": ci, **info}
    return all_cols, best


# ── 输出：HQ 内部评审 BOM ─────────────────────────────────────

def _write_review_bom(rows, output_file, project_name):
    wb = Workbook()
    ws = wb.active
    ws.title = "SW节点整机BOM配置"
    thin = Side(style="thin")
    bdr = Border(left=thin, right=thin, top=thin, bottom=thin)

    def S(cell, bold=False, bg=None, color="000000", h="center", v="center", size=11):
        cell.font = Font(bold=bold, color=color, size=size)
        if bg:
            cell.fill = PatternFill("solid", start_color=bg)
        cell.alignment = Alignment(horizontal=h, vertical=v)

    ws.merge_cells("A1:A2"); ws["A1"] = "项目名称"; S(ws["A1"], bold=True, bg="92D050", size=14)
    ws.merge_cells("B1:B2"); ws["B1"] = project_name; S(ws["B1"], bold=True, bg="92D050", size=14)
    ws.merge_cells("E1:I2"); ws["E1"] = "整机BOM配置表"; S(ws["E1"], bold=True, bg="92D050", size=16)
    ws["J1"] = "配置说明"; S(ws["J1"], bold=True, bg="92D050")
    ws["K1"] = "TBD"; S(ws["K1"], bg="BDD7EE")
    ws.row_dimensions[1].height = 30
    ws.merge_cells("A3:I3"); ws["A3"] = "SW节点HQ SN"
    S(ws["A3"], bold=True, bg="FFFF00", color="FF0000", size=12)
    ws["K3"] = ""; S(ws["K3"], bg="FFC000", color="FF0000")
    ws.row_dimensions[3].height = 20

    headers = ["序号", "组件子类", "虚拟层/物料", "物料类型", "HQ PN", "物料名称", "厂商型号", "厂商", "主二供", "", "用量"]
    for ci, h in enumerate(headers, 1):
        c = ws.cell(row=4, column=ci, value=h)
        S(c, bold=True, bg="D9D9D9")
        c.border = bdr
    ws.row_dimensions[4].height = 22

    dr = 5
    for item in rows:
        for si, (brand, model, qty) in enumerate(item["suppliers"]):
            label = SUPPLIER_LABELS[si] if si < len(SUPPLIER_LABELS) else f"{si+1}供"
            for ci, val in enumerate([item["seq"], "", "", "", "", item["name"], model, brand, label, "", qty], 1):
                c = ws.cell(row=dr, column=ci, value=val)
                c.border = bdr
                c.alignment = Alignment(horizontal="center", vertical="center")
            dr += 1

    for i, w in enumerate([6, 10, 12, 10, 18, 35, 30, 20, 8, 6, 8], 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    wb.save(output_file)
    return dr - 5


# ── 输出：原格式展开 ───────────────────────────────────────────

def _write_expanded_bom(ws_in, header_row, col_brand, col_model, col_qty, fmt, out_file):
    wb_out = Workbook()
    ws_out = wb_out.active
    ws_out.title = ws_in.title
    thin = Side(style="thin")
    bdr = Border(left=thin, right=thin, top=thin, bottom=thin)
    hdr_fill = PatternFill("solid", start_color="D9D9D9")
    max_col = ws_in.max_column

    out_map = [("seq", None, "序号")]
    for ci in range(1, max_col + 1):
        h = ws_in.cell(row=header_row, column=ci).value or ""
        if fmt == "A":
            if ci == col_brand:
                out_map.append(("brand", ci, "厂商"))
                out_map.append(("model", None, "型号"))
            else:
                out_map.append(("orig", ci, str(h)))
        else:
            if ci == col_brand:
                out_map.append(("brand", ci, "厂商"))
            elif col_model and ci == col_model:
                out_map.append(("model", ci, "型号"))
            else:
                out_map.append(("orig", ci, str(h)))
    out_map.append(("sole", None, "是否独供"))

    for out_ci, (typ, _, h) in enumerate(out_map, 1):
        c = ws_out.cell(row=1, column=out_ci, value=h)
        c.font = Font(bold=True)
        c.fill = hdr_fill
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = bdr
        ws_out.column_dimensions[get_column_letter(out_ci)].width = 6 if typ == "seq" else (10 if typ == "sole" else 18)

    dr = 2
    total = 0
    skipped = 0
    seq = 0
    for ri in range(header_row + 1, ws_in.max_row + 1):
        row_vals = {ci: ws_in.cell(row=ri, column=ci).value for ci in range(1, max_col + 1)}
        if not any(v is not None and str(v).strip() for v in row_vals.values()):
            skipped += 1
            continue

        bv = row_vals.get(col_brand)
        mv = row_vals.get(col_model) if col_model else None
        qv = row_vals.get(col_qty)
        suppliers = _parse_suppliers(bv, mv, fmt)
        if not suppliers:
            suppliers = [("", "")]
        mq = _safe_qty(qv)
        seq += 1
        sole_val = "是" if len(suppliers) == 1 else "否"

        for si, (brand, model) in enumerate(suppliers):
            for out_ci, (typ, src_ci, _) in enumerate(out_map, 1):
                if typ == "seq":
                    val = seq
                elif typ == "sole":
                    val = sole_val
                elif si == 0:
                    if typ == "brand":
                        val = brand
                    elif typ == "model":
                        val = model
                    else:
                        val = mq if src_ci == col_qty else row_vals.get(src_ci)
                else:
                    if typ == "brand":
                        val = brand
                    elif typ == "model":
                        val = model
                    else:
                        val = None
                c = ws_out.cell(row=dr, column=out_ci, value=val)
                c.alignment = Alignment(horizontal="left", vertical="center")
                c.border = bdr
            dr += 1
            total += 1

    wb_out.save(out_file)
    return total, skipped


# ── 路由 ─────────────────────────────────────────────────────

@bom_bp.route('/api/bom/detect', methods=['POST'])
def api_bom_detect():
    file = request.files.get('file')
    if not file:
        return jsonify({'success': False, 'error': '请上传文件'})
    uid = str(uuid.uuid4())[:8]
    try:
        in_path = _save_uploaded_excel(file, "bom_pre", uid)
        wb = _open_workbook(in_path, read_only=True, data_only=True)
    except ValueError as e:
        return jsonify({'success': False, 'error': str(e)})
    sheets = wb.sheetnames
    wb.close()

    sheet_name = request.form.get('sheet_name', '')
    if not sheet_name or sheet_name not in sheets:
        sheet_name = sheets[0] if sheets else ''

    header_row = _request_int('header_row', 1)
    if header_row is None:
        return jsonify({'success': False, 'error': '表头行必须是大于等于 1 的数字'})
    try:
        wb2 = _open_workbook(in_path, data_only=True)
    except ValueError as e:
        return jsonify({'success': False, 'error': str(e)})
    ws = wb2[sheet_name] if sheet_name else wb2[wb2.sheetnames[0]]

    all_cols, best = _detect_columns(ws, header_row)

    # Build preview (first 50 data rows)
    headers = [ws.cell(row=header_row, column=ci).value for ci in range(1, ws.max_column + 1)]
    preview = []
    for ri in range(header_row + 1, min(header_row + 51, ws.max_row + 1)):
        row = [ws.cell(row=ri, column=ci).value for ci in range(1, ws.max_column + 1)]
        if any(v is not None and str(v).strip() for v in row):
            preview.append([str(v) if v is not None else "" for v in row])

    wb2.close()

    # Format detected best columns
    detected = {}
    role_map = {
        "brand_combined": "brand",
        "brand_split": "brand",
        "brand_code": "brand",
        "model_split": "model",
        "model_code": "model",
        "qty": "qty",
        "name": "name",
    }
    fmt_guess = "A"
    for role, info in best.items():
        mapped = role_map.get(role, role)
        if mapped not in detected:
            detected[mapped] = info["letter"]
        if role == "brand_split":
            fmt_guess = "B"
        elif role == "brand_code":
            fmt_guess = "C"

    return jsonify({
        'success': True,
        'uid': uid,
        'sheets': sheets,
        'current_sheet': sheet_name,
        'headers': [str(h) if h is not None else "" for h in headers],
        'preview': preview,
        'detected': detected,
        'fmt_guess': fmt_guess,
    })


@bom_bp.route('/api/bom/convert', methods=['POST'])
@track_tool_activity('BOM格式转换')
def api_bom_convert():
    file = request.files.get('file')
    if not file:
        return jsonify({'success': False, 'error': '请上传文件'})

    uid = str(uuid.uuid4())[:8]
    try:
        in_path = _save_uploaded_excel(file, "bom_in", uid)
    except ValueError as e:
        return jsonify({'success': False, 'error': str(e)})

    sheet_name = request.form.get('sheet', '')
    header_row = _request_int('header_row', 1)
    if header_row is None:
        return jsonify({'success': False, 'error': '表头行必须是大于等于 1 的数字'})
    fmt = request.form.get('fmt', 'A')
    col_brand_str = request.form.get('col_brand', '')
    col_model_str = request.form.get('col_model', '')
    col_qty_str = request.form.get('col_qty', '')
    col_name_str = request.form.get('col_name', '')
    output_mode = request.form.get('output_mode', 'expand')
    project_name = request.form.get('project_name', '')

    try:
        wb = _open_workbook(in_path, read_only=True, data_only=True)
    except ValueError as e:
        return jsonify({'success': False, 'error': str(e)})
    sheets = wb.sheetnames
    wb.close()
    if not sheet_name or sheet_name not in sheets:
        sheet_name = sheets[0]

    col_brand = _col_int(col_brand_str)
    col_model = _col_int(col_model_str) if col_model_str.strip() else None
    col_qty = _col_int(col_qty_str)
    col_name = _col_int(col_name_str) if col_name_str.strip() else None

    if not col_brand:
        return jsonify({'success': False, 'error': '请指定品牌/厂家列'})

    try:
        wb2 = _open_workbook(in_path, data_only=True)
    except ValueError as e:
        return jsonify({'success': False, 'error': str(e)})
    ws = wb2[sheet_name]

    if output_mode == 'expand':
        out_name = f"展开BOM_{uid}.xlsx"
        out_path = os.path.join(OUTPUT_DIR, out_name)
        total, skipped = _write_expanded_bom(ws, header_row, col_brand, col_model, col_qty, fmt, out_path)
        wb2.close()
        return jsonify({
            'success': True,
            'download': f'/download/{out_name}',
            'total': total,
            'skipped': skipped,
        })
    else:
        # HQ format
        if not col_qty:
            wb2.close()
            return jsonify({'success': False, 'error': '输出HQ格式需要指定用量列'})

        rows = []
        seq = 0
        skipped = 0
        for ri in range(header_row + 1, ws.max_row + 1):
            row_vals = {ci: ws.cell(row=ri, column=ci).value for ci in range(1, ws.max_column + 1)}
            bv = row_vals.get(col_brand)
            nv = str(row_vals.get(col_name) or "").strip() if col_name else ""
            if not nv and not bv:
                skipped += 1
                continue
            mv = row_vals.get(col_model) if col_model else None
            qv = row_vals.get(col_qty)
            suppliers = _parse_suppliers(bv, mv, fmt)
            if not suppliers:
                suppliers = [("", "")]
            mq = _safe_qty(qv)
            seq += 1
            rows.append({
                "seq": seq,
                "name": nv,
                "suppliers": [(b, m, mq if si == 0 else 0) for si, (b, m) in enumerate(suppliers)],
            })

        wb2.close()
        out_name = f"整机BOM_{uid}.xlsx"
        out_path = os.path.join(OUTPUT_DIR, out_name)
        total = _write_review_bom(rows, out_path, project_name)
        return jsonify({
            'success': True,
            'download': f'/download/{out_name}',
            'total': total,
            'skipped': skipped,
        })




