# -*- coding: utf-8 -*-
"""BOM 格式转换工具 — Blueprint"""

import os, uuid, re, json, traceback

from shared import (
    openpyxl, Workbook, Font, PatternFill, Alignment, Border, Side,
    get_column_letter, column_index_from_string,
    render_template, request, jsonify,
    UPLOAD_DIR, OUTPUT_DIR, _cell_str, _col_int,
)
from flask import Blueprint

bom_bp = Blueprint('bom_tool', __name__)

SUPPLIER_LABELS = ["主供", "二供", "三供", "四供", "五供", "六供", "七供", "八供", "九供", "十供"]


# ── 解析器 ──────────────────────────────────────────────────

def parse_combined(raw):
    """格式A：品牌型号合并列（|| 或多空格分隔，如 MURATA:GRM188||SAMSUNG:CL10）"""
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


def parse_split(brand_raw, model_raw):
    """格式B：厂家/型号分开列，分号分隔（如 YAGEO;KOA / RC0805;RK73）"""
    brands = [b.strip() for b in str(brand_raw or "").split(";") if b.strip()] if brand_raw else []
    models = [m.strip() for m in str(model_raw or "").split(";") if m.strip()] if model_raw else []
    result = []
    for i in range(max(len(brands), len(models), 1)):
        b = brands[i] if i < len(brands) else ""
        m = models[i] if i < len(models) else ""
        if b or m:
            result.append((b, m))
    return result


def parse_format_c(brand_raw, model_raw):
    """格式C：制造商/型号分开列，冒号分隔，制造商含编号（如 1630-大毅科技[全称]:0362-RALEC[全称]）"""
    brand_names = []
    if brand_raw:
        s = str(brand_raw).strip()
        matches = re.findall(r'\d{4}-([^\[\]:]+)\[', s)
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


def parse_suppliers(bv, mv, fmt):
    if fmt == "C":
        return parse_format_c(bv, mv)
    if fmt == "B":
        return parse_split(bv, mv)
    return parse_combined(bv)


def safe_qty(qv):
    try:
        q = float(qv)
        return int(q) if q == int(q) else q
    except Exception:
        return qv if qv not in (None, "") else ""


# ── 列检测 ──────────────────────────────────────────────────

def detect_columns(ws, header_row):
    """扫描表头行，基于样本评分自动识别各列用途"""
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


# ── 输出器 ──────────────────────────────────────────────────

def write_review_bom(rows, output_file, project_name):
    """输出为 SW节点整机BOM配置表"""
    wb = Workbook()
    ws = wb.active
    ws.title = "SW节点整机BOM配置"
    GREEN = "92D050"
    YELLOW = "FFFF00"
    ORANGE = "FFC000"
    thin = Side(style="thin")
    bdr = Border(left=thin, right=thin, top=thin, bottom=thin)

    def S(cell, bold=False, bg=None, color="000000", h="center", v="center", size=11):
        cell.font = Font(bold=bold, color=color, size=size)
        if bg:
            cell.fill = PatternFill("solid", start_color=bg)
        cell.alignment = Alignment(horizontal=h, vertical=v)

    ws.merge_cells("A1:A2")
    ws["A1"] = "项目名称"
    S(ws["A1"], bold=True, bg=GREEN, size=14)
    ws.merge_cells("B1:B2")
    ws["B1"] = project_name
    S(ws["B1"], bold=True, bg=GREEN, size=14)
    ws.merge_cells("E1:I2")
    ws["E1"] = "整机BOM配置表"
    S(ws["E1"], bold=True, bg=GREEN, size=16)
    ws["J1"] = "配置说明"
    S(ws["J1"], bold=True, bg=GREEN)
    ws["K1"] = "TBD"
    S(ws["K1"], bg="BDD7EE")
    ws.row_dimensions[1].height = 30
    ws.merge_cells("A3:I3")
    ws["A3"] = "SW节点HQ SN"
    S(ws["A3"], bold=True, bg=YELLOW, color="FF0000", size=12)
    ws["K3"] = ""
    S(ws["K3"], bg=ORANGE, color="FF0000")
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
            for ci, val in enumerate([
                item["seq"], "", "", "", "", item["name"], model, brand, label, "", qty
            ], 1):
                c = ws.cell(row=dr, column=ci, value=val)
                c.border = bdr
                c.alignment = Alignment(horizontal="center", vertical="center")
            dr += 1
    for i, w in enumerate([6, 10, 12, 10, 18, 35, 30, 20, 8, 6, 8], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    wb.save(output_file)
    return dr - 5


def write_expanded_bom(ws_in, header_row, col_brand, col_model, col_qty, fmt, out_file):
    """保留客户BOM所有列，将供应商信息拆成多行。"""
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
        ws_out.column_dimensions[get_column_letter(out_ci)].width = (
            6 if typ == "seq" else (10 if typ == "sole" else 18)
        )

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
        suppliers = parse_suppliers(bv, mv, fmt)
        if not suppliers:
            suppliers = [("", "")]
        mq = safe_qty(qv)
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


# ── 扫描助手 ────────────────────────────────────────────────

def _scan_file_result(in_path, sheet_name, header_row):
    """加载已上传的 Excel 文件，扫描列结构，返回完整结果字典"""
    wb = openpyxl.load_workbook(in_path, read_only=True, data_only=True)
    sheets = wb.sheetnames
    wb.close()

    if not sheet_name or sheet_name not in sheets:
        sheet_name = sheets[0] if sheets else ''

    wb2 = openpyxl.load_workbook(in_path, data_only=True)
    ws = wb2[sheet_name] if sheet_name else wb2[wb2.sheetnames[0]]
    all_cols, best = detect_columns(ws, header_row)

    role_label = {
        "brand_combined": "品牌型号(合并A)",
        "brand_split": "厂家(分开B)",
        "model_split": "型号(分开B)",
        "brand_code": "制造商(编号C)",
        "model_code": "制造商型号(C)",
        "qty": "用量",
        "name": "物料名称",
    }
    col_results = {}
    for ci, info in all_cols.items():
        r = info["role"]
        if r in role_label:
            col_results[info["letter"]] = {
                "header": info["header"], "role": role_label[r],
                "sample": info["sample"], "score": info["score"],
            }

    detected_letters = {}
    if "name" in best:
        detected_letters["name"] = best["name"]["letter"]
    if "qty" in best:
        detected_letters["qty"] = best["qty"]["letter"]
    if "brand_code" in best:
        detected_letters["brand"] = best["brand_code"]["letter"]
        if "model_code" in best:
            detected_letters["model"] = best["model_code"]["letter"]
        detected_letters["fmt"] = "C"
    elif "brand_combined" in best:
        detected_letters["brand"] = best["brand_combined"]["letter"]
        detected_letters["fmt"] = "A"
    elif "brand_split" in best:
        detected_letters["brand"] = best["brand_split"]["letter"]
        if "model_split" in best:
            detected_letters["model"] = best["model_split"]["letter"]
        detected_letters["fmt"] = "B"

    preview = []
    max_row = min(ws.max_row, 11)
    max_col = min(ws.max_column, 10)
    for ri in range(1, max_row + 1):
        preview.append([str(ws.cell(row=ri, column=ci).value or '')[:24] for ci in range(1, max_col + 1)])

    headers = [str(ws.cell(row=header_row, column=ci).value or '').strip()
               for ci in range(1, max(ws.max_column, 1) + 1)]
    wb2.close()

    return {
        'sheets': sheets,
        'current_sheet': sheet_name,
        'detected': detected_letters,
        'col_results': col_results,
        'headers': headers,
        'preview': preview,
    }


# ── 路由 ─────────────────────────────────────────────────────

@bom_bp.route('/api/bom/sheets', methods=['POST'])
def api_bom_sheets():
    file = request.files.get('file')
    if not file:
        return jsonify({'success': False, 'error': '请上传文件'})
    uid = str(uuid.uuid4())[:8]
    in_path = os.path.join(UPLOAD_DIR, f"bom_pre_{uid}.xlsx")
    file.save(in_path)

    sheet_name = request.form.get('sheet_name', '')
    header_row = int(request.form.get('header_row', 1))
    result = _scan_file_result(in_path, sheet_name, header_row)

    return jsonify({'success': True, 'file_id': uid, **result})


@bom_bp.route('/api/bom/rescan', methods=['POST'])
def api_bom_rescan():
    file_id = request.form.get('file_id', '')
    if not file_id:
        return jsonify({'success': False, 'error': '缺少文件ID'})
    in_path = os.path.join(UPLOAD_DIR, f"bom_pre_{file_id}.xlsx")
    if not os.path.exists(in_path):
        return jsonify({'success': False, 'error': '文件已过期，请重新上传'})

    sheet_name = request.form.get('sheet_name', '')
    header_row = int(request.form.get('header_row', 1))
    result = _scan_file_result(in_path, sheet_name, header_row)

    return jsonify({'success': True, 'file_id': file_id, **result})


@bom_bp.route('/bom', methods=['GET', 'POST'])
def tool_bom():
    if request.method == 'POST':
        file = request.files.get('file')
        if not file:
            return "请上传文件", 400
        fmt = request.form.get('fmt', 'A')
        mode = request.form.get('mode', 'expand')
        sheet_name = request.form.get('sheet', '')
        header_row = int(request.form.get('header_row', 1))
        project_name = request.form.get('project_name', '')
        col_brand = request.form.get('col_brand', '')
        col_model = request.form.get('col_model', '')
        col_qty = request.form.get('col_qty', '')
        col_name = request.form.get('col_name', '')

        uid = str(uuid.uuid4())[:8]
        in_path = os.path.join(UPLOAD_DIR, f"bom_in_{uid}.xlsx")
        out_path = os.path.join(OUTPUT_DIR, f"BOM转换结果_{uid}.xlsx")
        file.save(in_path)

        wb = openpyxl.load_workbook(in_path, read_only=True, data_only=True)
        sheets = wb.sheetnames
        wb.close()
        if not sheet_name or sheet_name not in sheets:
            sheet_name = sheets[0]

        wb2 = openpyxl.load_workbook(in_path, data_only=True)
        ws = wb2[sheet_name]
        _, best = detect_columns(ws, header_row)

        if not col_brand:
            if "brand_code" in best:
                col_brand = best["brand_code"]["letter"]
            elif "brand_combined" in best:
                col_brand = best["brand_combined"]["letter"]
            elif "brand_split" in best:
                col_brand = best["brand_split"]["letter"]
        if not col_model:
            if "model_code" in best:
                col_model = best["model_code"]["letter"]
            elif "model_split" in best:
                col_model = best["model_split"]["letter"]
        if not col_qty and "qty" in best:
            col_qty = best["qty"]["letter"]
        if not col_name and "name" in best:
            col_name = best["name"]["letter"]

        if fmt == "auto":
            col_brand_int = _col_int(col_brand)
            col_model_int = _col_int(col_model)
            if col_model_int:
                sample = str(ws.cell(row=header_row + 1, column=col_brand_int).value or "")
                fmt = "C" if re.search(r'\d{4}-[^\[]+\[', sample) else "B"
            else:
                fmt = "A"

        col_brand_int = _col_int(col_brand)
        col_model_int = _col_int(col_model)
        col_qty_int = _col_int(col_qty)
        col_name_int = _col_int(col_name)
        wb2.close()

        if not col_brand_int:
            return jsonify({'success': False, 'error': '请指定品牌/厂家列'})

        try:
            if mode == 'expand':
                if not col_qty_int:
                    return jsonify({'success': False, 'error': '原格式展开需要指定用量列'})
                total, skipped = write_expanded_bom(
                    ws, header_row, col_brand_int, col_model_int, col_qty_int, fmt, out_path)
                msg = f"写入 {total} 行（跳过空行 {skipped}）"
            else:
                """内部格式：输出仅含「规格型号」列的单列 Excel"""
                if not col_model_int:
                    return jsonify({'success': False, 'error': '内部格式需要指定型号列'})
                model_vals = []
                skipped = 0
                max_row = ws.max_row
                for ri in range(header_row + 1, max_row + 1):
                    mv = ws.cell(row=ri, column=col_model_int).value
                    if mv is None or str(mv).strip() == "":
                        skipped += 1
                        continue
                    model_vals.append(str(mv).strip())
                wb_simple = Workbook()
                ws_simple = wb_simple.active
                ws_simple.title = "规格型号"
                c = ws_simple.cell(row=1, column=1, value="规格型号")
                c.font = Font(bold=True)
                for i, val in enumerate(model_vals, 2):
                    ws_simple.cell(row=i, column=1, value=val)
                ws_simple.column_dimensions['A'].width = 45
                wb_simple.save(out_path)
                wb_simple.close()
                msg = f"共输出 {len(model_vals)} 行规格型号（跳过空行 {skipped}）"

            return jsonify({
                'success': True, 'message': msg,
                'download': f'/download/BOM转换结果_{uid}.xlsx',
                'sheets': sheets,
            })
        except Exception as e:
            return jsonify({'success': False, 'error': f"{e}\n{traceback.format_exc()}"})

    return render_template('bom.html')
