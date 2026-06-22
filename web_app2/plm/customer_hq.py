# -*- coding: utf-8 -*-
"""Customer BOM to HQ single-board BOM conversion for the PLM tool."""

import os
import uuid

from activity import track_tool_activity
from openpyxl.styles import Color
from bom import _detect_columns as _detect_customer_bom_columns, _safe_qty
from shared import (
    OUTPUT_DIR,
    Alignment,
    Border,
    Font,
    PatternFill,
    Side,
    Workbook,
    get_column_letter,
    jsonify,
    request,
    _cell_str,
    _col_int,
    _open_workbook,
    _request_int,
    _save_uploaded_excel,
)

from . import plm_bp


HQ_SINGLE_BOARD_HEADERS = [
    "序号", "料号", "型号", "物料描述", "单耗", "替代关系", "位号", "生产厂家", "是否环保", "湿敏属性",
    "备注", "主辅BOM标记", "MBG优选属性", "CBG优选属性", "DBG优选属性", "主制控", "子制控", "子制控数量", "ABG优选属性",
]

OPTIONAL_HEADER_TO_FIELD = {
    "是否环保": "environmental",
    "湿敏属性": "msl",
    "备注": "note",
    "主辅BOM标记": "main_aux",
    "MBG优选属性": "mbg_preferred",
    "CBG优选属性": "cbg_preferred",
    "DBG优选属性": "dbg_preferred",
    "主制控": "main_control",
    "子制控": "sub_control",
    "子制控数量": "sub_control_qty",
    "ABG优选属性": "abg_preferred",
}

SAMPLE_COLUMN_WIDTHS = [
    5.88671875, 21.44140625, 31.21875, 43.0, 13.6640625, 13.0, 31.21875, 11.6640625, 21.44140625,
    11.6640625, 13.0, 19.5546875, 13.0, 13.0, 13.0, 13.0, 13.0, 13.0, 13.0,
]


def _detect_seq_column(ws, header_row):
    for ci in range(1, ws.max_column + 1):
        header = _cell_str(ws.cell(row=header_row, column=ci).value).lower()
        if header in ("序号", "seq", "no.", "no"):
            return ci
    return None


def _detect_exact_header_columns(ws, header_row):
    result = {}
    for ci in range(1, ws.max_column + 1):
        field = OPTIONAL_HEADER_TO_FIELD.get(_cell_str(ws.cell(row=header_row, column=ci).value))
        if field:
            result[field] = ci
    return result


def _customer_hq_detect_payload(path, sheet_name, header_row):
    wb = _open_workbook(path, data_only=True)
    try:
        if not sheet_name or sheet_name not in wb.sheetnames:
            sheet_name = wb.sheetnames[0] if wb.sheetnames else ""
        ws = wb[sheet_name]
        all_cols, best = _detect_customer_bom_columns(ws, header_row)
        seq_col = _detect_seq_column(ws, header_row)
        headers = [ws.cell(row=header_row, column=ci).value for ci in range(1, ws.max_column + 1)]
        preview = []
        for ri in range(header_row + 1, min(header_row + 51, ws.max_row + 1)):
            row = [ws.cell(row=ri, column=ci).value for ci in range(1, ws.max_column + 1)]
            if any(_cell_str(v) for v in row):
                preview.append([_cell_str(v) for v in row])

        detected = {}
        if seq_col:
            detected["seq"] = get_column_letter(seq_col)
        role_map = {
            "brand_combined": "brand",
            "brand_split": "brand",
            "brand_code": "brand",
            "model_split": "model",
            "model_code": "model",
            "qty": "qty",
            "name": "name",
        }
        for role, info in best.items():
            mapped = role_map.get(role, role)
            if mapped not in detected:
                detected[mapped] = info["letter"]

        return {
            "success": True,
            "sheets": wb.sheetnames,
            "current_sheet": sheet_name,
            "headers": [f"{info['letter']}:{info['header']}" for _, info in sorted(all_cols.items())],
            "preview_headers": [_cell_str(h) for h in headers],
            "preview": preview,
            "detected": detected,
        }
    finally:
        wb.close()


def _form_text(name, default=""):
    return _cell_str(request.form.get(name)) or default


def _write_hq_single_board_bom(rows, out_path, meta):
    wb = Workbook()
    ws = wb.active
    ws.title = "BOM"
    thin = Side(style="thin", color=Color(auto=1))
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    header_fill = PatternFill(fill_type="solid", fgColor=Color(indexed=23))
    label_font = Font(name="Arial", size=12, bold=True)
    body_font = Font(name="宋体", size=10)

    meta_rows = [
        ["料号", meta["part_no"], "描述", meta["description"], "项目配置名", meta["config_name"], "工程师", meta["engineer"]],
        ["版本", meta["version"], "替代项", meta["alternate_item"], "BOM名称", meta["bom_name"], "归档部门", meta["archive_dept"]],
    ]
    for ri, values in enumerate(meta_rows, 1):
        for ci, value in enumerate(values, 1):
            cell = ws.cell(row=ri, column=ci, value=value)
            cell.border = border
            cell.alignment = Alignment(vertical="center", wrap_text=True)
            cell.font = body_font
            cell.number_format = "@"
            if ci % 2 == 1:
                cell.font = label_font
                cell.fill = header_fill
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    for ci, header in enumerate(HQ_SINGLE_BOARD_HEADERS, 1):
        cell = ws.cell(row=3, column=ci, value=header)
        cell.font = label_font
        cell.border = border
        cell.fill = header_fill
        cell.number_format = "@"
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    for ri, row in enumerate(rows, 4):
        values = [
            row["seq"], row["hqpn"], row["model"], row["description"], row["qty"], row["alternate"], row["refdes"],
            row["manufacturer"], row["environmental"], row["msl"], row["note"], row["main_aux"], row["mbg_preferred"],
            row["cbg_preferred"], row["dbg_preferred"], row["main_control"], row["sub_control"], row["sub_control_qty"],
            row["abg_preferred"],
        ]
        for ci, value in enumerate(values, 1):
            cell = ws.cell(row=ri, column=ci, value=value)
            cell.font = body_font
            cell.border = border
            cell.alignment = Alignment(vertical="center", wrap_text=True)
            if ci != 5:
                cell.number_format = "@"

    for ci, width in enumerate(SAMPLE_COLUMN_WIDTHS, 1):
        ws.column_dimensions[get_column_letter(ci)].width = width
    wb.save(out_path)
    wb.close()


@plm_bp.route("/api/plm/customer_hq_detect", methods=["POST"])
def api_customer_hq_detect():
    file = request.files.get("file")
    if not file:
        return jsonify({"success": False, "error": "请上传客户 BOM 文件"})
    uid = str(uuid.uuid4())[:8]
    try:
        path = _save_uploaded_excel(file, "plm_customer_hq_pre", uid)
        wb = _open_workbook(path, read_only=True, data_only=True)
        sheets = wb.sheetnames
        wb.close()
    except ValueError as exc:
        return jsonify({"success": False, "error": str(exc)})

    sheet_name = request.form.get("sheet_name", "")
    if not sheet_name or sheet_name not in sheets:
        sheet_name = sheets[0] if sheets else ""
    header_row = _request_int("header_row", 1)
    if header_row is None:
        return jsonify({"success": False, "error": "表头行必须是大于等于 1 的数字"})

    try:
        payload = _customer_hq_detect_payload(path, sheet_name, header_row)
        payload["uid"] = uid
        return jsonify(payload)
    except ValueError as exc:
        return jsonify({"success": False, "error": str(exc)})


@plm_bp.route("/api/plm/customer_hq_convert", methods=["POST"])
@track_tool_activity("客户BOM转换成HQ格式单板BOM")
def api_customer_hq_convert():
    file = request.files.get("file")
    if not file:
        return jsonify({"success": False, "error": "请上传客户 BOM 文件"})

    uid = str(uuid.uuid4())[:8]
    try:
        path = _save_uploaded_excel(file, "plm_customer_hq_in", uid)
    except ValueError as exc:
        return jsonify({"success": False, "error": str(exc)})

    sheet_name = request.form.get("sheet", "")
    header_row = _request_int("header_row", 1)
    if header_row is None:
        return jsonify({"success": False, "error": "表头行必须是大于等于 1 的数字"})

    col_seq_str = request.form.get("col_seq", "")
    col_hqpn_str = request.form.get("col_hqpn", "")
    col_brand_str = request.form.get("col_brand", "")
    col_model_str = request.form.get("col_model", "")
    col_qty_str = request.form.get("col_qty", "")
    col_name_str = request.form.get("col_name", "")
    col_refdes_str = request.form.get("col_refdes", "")

    wb_ro = _open_workbook(path, read_only=True, data_only=True)
    try:
        sheets = wb_ro.sheetnames
    finally:
        wb_ro.close()
    if not sheet_name or sheet_name not in sheets:
        sheet_name = sheets[0]

    if not all([col_seq_str, col_brand_str, col_qty_str, col_name_str]):
        detected = _customer_hq_detect_payload(path, sheet_name, header_row).get("detected", {})
        col_seq_str = col_seq_str or detected.get("seq", "")
        col_brand_str = col_brand_str or detected.get("brand", "")
        col_model_str = col_model_str or detected.get("model", "")
        col_qty_str = col_qty_str or detected.get("qty", "")
        col_name_str = col_name_str or detected.get("name", "")

    col_seq = _col_int(col_seq_str)
    col_hqpn = _col_int(col_hqpn_str) if str(col_hqpn_str).strip() else None
    col_brand = _col_int(col_brand_str)
    col_model = _col_int(col_model_str) if str(col_model_str).strip() else None
    col_qty = _col_int(col_qty_str)
    col_name = _col_int(col_name_str)
    col_refdes = _col_int(col_refdes_str) if str(col_refdes_str).strip() else None
    if not col_seq:
        return jsonify({"success": False, "error": "请指定序号列"})
    if not col_brand:
        return jsonify({"success": False, "error": "请指定生产厂家列"})
    if not col_model:
        return jsonify({"success": False, "error": "请指定型号列"})
    if not col_qty:
        return jsonify({"success": False, "error": "请指定单耗/用量列"})
    if not col_name:
        return jsonify({"success": False, "error": "请指定物料描述列"})

    meta = {
        "part_no": _form_text("part_no"),
        "description": _form_text("description", _form_text("project_name", "客户BOM")),
        "config_name": _form_text("config_name", _form_text("project_name", "客户BOM")),
        "engineer": _form_text("engineer"),
        "version": _form_text("version"),
        "alternate_item": _form_text("alternate_item"),
        "bom_name": _form_text("bom_name", _form_text("description", _form_text("project_name", "客户BOM"))),
        "archive_dept": _form_text("archive_dept"),
    }

    wb = _open_workbook(path, data_only=True)
    try:
        ws = wb[sheet_name]
        optional_cols = _detect_exact_header_columns(ws, header_row)
        rows = []
        skipped = 0
        seen_seq = set()
        for ri in range(header_row + 1, ws.max_row + 1):
            row_vals = {ci: ws.cell(row=ri, column=ci).value for ci in range(1, ws.max_column + 1)}
            seq = _cell_str(row_vals.get(col_seq))
            manufacturer = _cell_str(row_vals.get(col_brand))
            model = _cell_str(row_vals.get(col_model))
            description = _cell_str(row_vals.get(col_name))
            if not seq and not any([manufacturer, model, description]):
                skipped += 1
                continue
            is_main = seq not in seen_seq
            seen_seq.add(seq)
            optional_values = {field: _cell_str(row_vals.get(col_idx)) for field, col_idx in optional_cols.items()}
            rows.append({
                "seq": seq,
                "hqpn": _cell_str(row_vals.get(col_hqpn)) if col_hqpn else "",
                "model": model,
                "description": description,
                "qty": _safe_qty(row_vals.get(col_qty)) if is_main else "",
                "alternate": "",
                "refdes": _cell_str(row_vals.get(col_refdes)) if (is_main and col_refdes) else "",
                "manufacturer": manufacturer,
                "environmental": optional_values.get("environmental", ""),
                "msl": optional_values.get("msl", ""),
                "note": optional_values.get("note", ""),
                "main_aux": optional_values.get("main_aux", ""),
                "mbg_preferred": optional_values.get("mbg_preferred", ""),
                "cbg_preferred": optional_values.get("cbg_preferred", ""),
                "dbg_preferred": optional_values.get("dbg_preferred", ""),
                "main_control": optional_values.get("main_control", ""),
                "sub_control": optional_values.get("sub_control", ""),
                "sub_control_qty": optional_values.get("sub_control_qty", ""),
                "abg_preferred": optional_values.get("abg_preferred", ""),
            })
    finally:
        wb.close()

    out_name = f"客户BOM_HQ单板BOM_{uid}.xlsx"
    out_path = os.path.join(OUTPUT_DIR, out_name)
    _write_hq_single_board_bom(rows, out_path, meta)

    return jsonify({
        "success": True,
        "download": f"/download/{out_name}",
        "total": len(rows),
        "skipped": skipped,
        "columns": {
            "seq": get_column_letter(col_seq),
            "hqpn": get_column_letter(col_hqpn) if col_hqpn else "",
            "brand": get_column_letter(col_brand),
            "model": get_column_letter(col_model),
            "qty": get_column_letter(col_qty),
            "name": get_column_letter(col_name),
            "refdes": get_column_letter(col_refdes) if col_refdes else "",
        },
    })
