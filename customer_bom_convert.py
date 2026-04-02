# -*- coding: utf-8 -*-
"""
客户BOM → 内部整机BOM评审格式转换脚本

【使用步骤】
1. 在 CUSTOMERS 字典里找到你的客户（或新增一个）
2. 修改 CURRENT_CUSTOMER 为该客户的key
3. 修改 INPUT_FILE 为实际文件名
4. 运行：python customer_bom_convert.py

【新增客户】
在 CUSTOMERS 中复制一份配置，按实际列填写即可。
列用字母表示：A=1, B=2, C=3, D=4, E=5, F=6, G=7, H=8 ...
"""

import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter, column_index_from_string

# =====================================================================
# 客户配置字典
# 每个客户单独配置，配一次永久生效
# col_name      : 物料名称/品名 所在列（字母）
# col_qty       : 用量 所在列（字母）
# col_brand     : 品牌型号（厂商:型号||格式）所在列（字母）
# header_row    : 表头在第几行
# data_start_row: 数据从第几行开始
# sheet         : 0=第一个Sheet，也可填Sheet名字符串
# =====================================================================
CUSTOMERS = {
    "客户A": {
        "desc":          "客户A - LB-800G系列BOM",
        "col_name":      "D",   # 物料品名
        "col_qty":       "E",   # 用量
        "col_brand":     "G",   # 品牌型号
        "header_row":    1,
        "data_start_row": 2,
        "sheet":         0,
    },

    # ---------- 在下面继续添加新客户 ----------
    # "客户B": {
    #     "desc":          "客户B - XX项目",
    #     "col_name":      "C",
    #     "col_qty":       "F",
    #     "col_brand":     "H",
    #     "header_row":    2,
    #     "data_start_row": 3,
    #     "sheet":         "BOM",
    # },
}

# =====================================================================
# 运行配置 - 每次使用只需改这3行
# =====================================================================
CURRENT_CUSTOMER = "客户A"        # 选择上面配置的客户key
INPUT_FILE       = "客户BOM.xlsx" # 客户BOM文件名
OUTPUT_FILE      = "内部评审BOM.xlsx"
PROJECT_NAME     = ""             # 项目名称，留空则脚本运行时提示输入
# =====================================================================

SUPPLIER_LABELS = ["主供", "二供", "三供", "四供", "五供",
                   "六供", "七供", "八供", "九供", "十供"]


def col_letter_to_num(letter):
    return column_index_from_string(letter.upper())


def parse_brand_model(raw_str):
    """解析 厂商:型号||厂商:型号 格式，返回 [(厂商, 型号), ...]"""
    if not raw_str or str(raw_str).strip() == "":
        return []
    raw = str(raw_str).strip().replace("：", ":")  # 全角冒号兼容
    result = []
    for entry in [e.strip() for e in raw.split("||") if e.strip()]:
        if ":" in entry:
            brand, model = entry.split(":", 1)
            result.append((brand.strip(), model.strip()))
        else:
            result.append(("", entry.strip()))
    return result


def write_review_bom(rows, output_file, project_name):
    wb = Workbook()
    ws = wb.active
    ws.title = "SW节点整机BOM配置"

    GREEN  = "92D050"
    YELLOW = "FFFF00"
    ORANGE = "FFC000"

    def style(cell, bold=False, bg=None, color="000000",
              h="center", v="center", wrap=False, size=11):
        cell.font = Font(bold=bold, color=color, size=size)
        if bg:
            cell.fill = PatternFill("solid", start_color=bg)
        cell.alignment = Alignment(horizontal=h, vertical=v, wrap_text=wrap)

    thin = Side(style="thin")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    # 第1-2行：项目名称
    ws.merge_cells("A1:A2"); ws["A1"] = "项目名称"
    style(ws["A1"], bold=True, bg=GREEN, size=14)
    ws.merge_cells("B1:B2"); ws["B1"] = project_name
    style(ws["B1"], bold=True, bg=GREEN, size=14)
    ws.merge_cells("E1:I2"); ws["E1"] = "整机BOM配置表"
    style(ws["E1"], bold=True, bg=GREEN, size=16)
    ws["J1"] = "配置说明"; style(ws["J1"], bold=True, bg=GREEN)
    ws["K1"] = "TBD";      style(ws["K1"], bg="BDD7EE")
    ws.row_dimensions[1].height = 30

    # 第3行：SW节点
    ws.merge_cells("A3:I3"); ws["A3"] = "SW节点HQ SN"
    style(ws["A3"], bold=True, bg=YELLOW, color="FF0000", size=12)
    ws["K3"] = ""; style(ws["K3"], bg=ORANGE, color="FF0000")
    ws.row_dimensions[3].height = 20

    # 第4行：表头
    headers = ["序号", "组件子类", "虚拟层/物料", "物料类型", "HQ PN",
               "物料名称", "厂商型号", "厂商", "主二供", "", "用量"]
    for ci, h in enumerate(headers, 1):
        c = ws.cell(row=4, column=ci, value=h)
        style(c, bold=True, bg="D9D9D9")
        c.border = border
    ws.row_dimensions[4].height = 22

    # 数据行
    dr = 5
    for item in rows:
        for si, (brand, model, qty) in enumerate(item["suppliers"]):
            label = SUPPLIER_LABELS[si] if si < len(SUPPLIER_LABELS) else f"{si+1}供"
            for ci, val in enumerate(
                [item["seq"], "", "", "", "", item["name"], model, brand, label, "", qty], 1
            ):
                c = ws.cell(row=dr, column=ci, value=val)
                c.border = border
                c.alignment = Alignment(horizontal="center", vertical="center")
            dr += 1

    # 列宽
    for i, w in enumerate([6, 10, 12, 10, 18, 30, 30, 20, 8, 6, 8], 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    wb.save(output_file)
    total_items = len(rows)
    total_rows  = dr - 5
    print(f"完成！共 {total_items} 个物料，展开 {total_rows} 行 → {output_file}")


def convert():
    cfg = CUSTOMERS.get(CURRENT_CUSTOMER)
    if not cfg:
        print(f"[错误] 找不到客户配置：{CURRENT_CUSTOMER}")
        print(f"已有配置：{list(CUSTOMERS.keys())}")
        return

    print(f"客户：{cfg['desc']}")
    print(f"文件：{INPUT_FILE}")

    try:
        wb = openpyxl.load_workbook(INPUT_FILE, data_only=True)
    except Exception as e:
        print(f"[错误] 无法打开文件：{e}")
        return

    sheet = cfg["sheet"]
    ws = wb.worksheets[sheet] if isinstance(sheet, int) else wb[sheet]
    print(f"Sheet：{ws.title}")

    col_name  = col_letter_to_num(cfg["col_name"])
    col_qty   = col_letter_to_num(cfg["col_qty"])
    col_brand = col_letter_to_num(cfg["col_brand"])
    data_start = cfg["data_start_row"]

    # 打印表头确认（方便调试）
    print("表头确认：")
    print(f"  物料名称列 {cfg['col_name']} = {ws.cell(row=cfg['header_row'], column=col_name).value}")
    print(f"  用量列     {cfg['col_qty']} = {ws.cell(row=cfg['header_row'], column=col_qty).value}")
    print(f"  品牌型号列 {cfg['col_brand']} = {ws.cell(row=cfg['header_row'], column=col_brand).value}")

    project = PROJECT_NAME
    if not project:
        project = input("请输入项目名称：").strip()

    rows = []
    seq = 0
    for ri in range(data_start, ws.max_row + 1):
        name_val  = ws.cell(row=ri, column=col_name).value
        qty_val   = ws.cell(row=ri, column=col_qty).value
        brand_val = ws.cell(row=ri, column=col_brand).value

        if not name_val and not brand_val:
            continue

        suppliers_raw = parse_brand_model(brand_val) or [("", "")]

        try:
            main_qty = float(qty_val) if qty_val not in (None, "") else 0
            main_qty = int(main_qty) if main_qty == int(main_qty) else main_qty
        except (ValueError, TypeError):
            main_qty = qty_val

        suppliers = [
            (brand, model, main_qty if i == 0 else 0)
            for i, (brand, model) in enumerate(suppliers_raw)
        ]

        seq += 1
        rows.append({"seq": seq, "name": str(name_val).strip(), "suppliers": suppliers})

    print(f"解析：{len(rows)} 个物料")
    write_review_bom(rows, OUTPUT_FILE, project)


if __name__ == "__main__":
    convert()
