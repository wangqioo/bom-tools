# -*- coding: utf-8 -*-
"""
BOM转换工具 - 图形界面版
运行方式：python bom_gui.py
依赖：pip install openpyxl
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter, column_index_from_string
import os
import re
import threading

# ───────────────────────────── 解析逻辑 ─────────────────────────────

SUPPLIER_LABELS = ["主供", "二供", "三供", "四供", "五供",
                   "六供", "七供", "八供", "九供", "十供"]


def parse_brand_model(raw):
    """
    智能解析品牌型号字段，支持多种格式：
      MURATA:GRM188||SAMSUNG:CL10   → 厂商:型号 || 分隔
      MURATA：GRM188||SAMSUNG：CL10 → 全角冒号
      GRM188||CL10                  → 只有型号没有厂商
      MURATA/GRM188                 → 斜杠分隔
      GRM188                        → 单个型号
    """
    if not raw or str(raw).strip() == "":
        return []
    s = str(raw).strip()
    s = s.replace("：", ":").replace("∥", "||").replace("‖", "||")

    entries = [e.strip() for e in re.split(r"\|\|", s) if e.strip()]
    result = []
    for entry in entries:
        if ":" in entry:
            brand, model = entry.split(":", 1)
            result.append((brand.strip(), model.strip()))
        elif "/" in entry and len(entry.split("/")) == 2:
            brand, model = entry.split("/", 1)
            result.append((brand.strip(), model.strip()))
        else:
            result.append(("", entry.strip()))
    return result


def detect_columns(ws, header_row):
    """
    扫描表格，自动识别各列用途
    返回 {col_index: {"letter": "A", "header": "...", "role": "name/qty/brand/other", "score": 0-100}}
    """
    max_col = ws.max_column
    data_rows = list(range(header_row + 1, min(header_row + 11, ws.max_row + 1)))
    result = {}

    for ci in range(1, max_col + 1):
        header_val = ws.cell(row=header_row, column=ci).value
        header_str = str(header_val).strip() if header_val else ""
        letter = get_column_letter(ci)

        sample_vals = [ws.cell(row=r, column=ci).value for r in data_rows]
        sample_strs = [str(v).strip() for v in sample_vals if v is not None]

        role = "other"
        score = 0

        # 检测品牌型号列：含 || 或 厂商:型号 模式
        brand_signals = sum(
            1 for v in sample_strs
            if "||" in v or (re.search(r"[A-Za-z0-9\u4e00-\u9fff]+:[A-Za-z0-9]", v))
        )
        if brand_signals >= 2 or "品牌" in header_str or "型号" in header_str:
            role = "brand"
            score = brand_signals * 20 + (30 if "品牌" in header_str or "型号" in header_str else 0)

        # 检测用量列：数值，表头含用量/数量/qty
        numeric_count = sum(1 for v in sample_vals if v is not None and str(v).replace(".", "").isdigit())
        qty_keywords = ["用量", "数量", "qty", "quantity", "用料", "需求量"]
        if any(k in header_str.lower() for k in qty_keywords):
            role = "qty"
            score = 80
        elif numeric_count >= len(data_rows) * 0.6 and role == "other":
            role = "qty"
            score = numeric_count * 10

        # 检测物料名称列：长文本，表头含名/描述
        avg_len = sum(len(v) for v in sample_strs) / max(len(sample_strs), 1)
        name_keywords = ["名称", "品名", "物料名", "描述", "名", "description", "name", "规格"]
        if any(k in header_str for k in name_keywords):
            if role == "other":
                role = "name"
                score = 70
        elif avg_len > 8 and role == "other":
            role = "name"
            score = int(avg_len * 2)

        result[ci] = {
            "letter": letter,
            "header": header_str,
            "role": role,
            "score": score,
            "sample": sample_strs[:3],
        }

    # 每种 role 只保留得分最高的
    best = {}
    for ci, info in result.items():
        r = info["role"]
        if r != "other":
            if r not in best or info["score"] > best[r]["score"]:
                best[r] = {"ci": ci, **info}

    return result, best


def write_review_bom(rows, output_file, project_name):
    wb = Workbook()
    ws = wb.active
    ws.title = "SW节点整机BOM配置"

    GREEN = "92D050"; YELLOW = "FFFF00"; ORANGE = "FFC000"
    thin = Side(style="thin")
    bdr = Border(left=thin, right=thin, top=thin, bottom=thin)

    def S(cell, bold=False, bg=None, color="000000", h="center", v="center", size=11, wrap=False):
        cell.font = Font(bold=bold, color=color, size=size)
        if bg: cell.fill = PatternFill("solid", start_color=bg)
        cell.alignment = Alignment(horizontal=h, vertical=v, wrap_text=wrap)

    ws.merge_cells("A1:A2"); ws["A1"] = "项目名称"; S(ws["A1"], bold=True, bg=GREEN, size=14)
    ws.merge_cells("B1:B2"); ws["B1"] = project_name; S(ws["B1"], bold=True, bg=GREEN, size=14)
    ws.merge_cells("E1:I2"); ws["E1"] = "整机BOM配置表"; S(ws["E1"], bold=True, bg=GREEN, size=16)
    ws["J1"] = "配置说明"; S(ws["J1"], bold=True, bg=GREEN)
    ws["K1"] = "TBD"; S(ws["K1"], bg="BDD7EE")
    ws.row_dimensions[1].height = 30

    ws.merge_cells("A3:I3"); ws["A3"] = "SW节点HQ SN"
    S(ws["A3"], bold=True, bg=YELLOW, color="FF0000", size=12)
    ws["K3"] = ""; S(ws["K3"], bg=ORANGE, color="FF0000")
    ws.row_dimensions[3].height = 20

    headers = ["序号", "组件子类", "虚拟层/物料", "物料类型", "HQ PN",
               "物料名称", "厂商型号", "厂商", "主二供", "", "用量"]
    for ci, h in enumerate(headers, 1):
        c = ws.cell(row=4, column=ci, value=h)
        S(c, bold=True, bg="D9D9D9"); c.border = bdr
    ws.row_dimensions[4].height = 22

    dr = 5
    for item in rows:
        for si, (brand, model, qty) in enumerate(item["suppliers"]):
            label = SUPPLIER_LABELS[si] if si < len(SUPPLIER_LABELS) else f"{si+1}供"
            for ci, val in enumerate([item["seq"], "", "", "", "", item["name"],
                                       model, brand, label, "", qty], 1):
                c = ws.cell(row=dr, column=ci, value=val)
                c.border = bdr
                c.alignment = Alignment(horizontal="center", vertical="center")
            dr += 1

    for i, w in enumerate([6, 10, 12, 10, 18, 35, 30, 20, 8, 6, 8], 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    wb.save(output_file)
    return dr - 5


# ───────────────────────────── GUI ─────────────────────────────

class BomApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("BOM 转换工具")
        self.geometry("780x680")
        self.resizable(True, True)
        self.configure(bg="#f5f5f5")

        # 状态变量
        self.wb = None
        self.ws = None
        self.col_name_var  = tk.StringVar()
        self.col_qty_var   = tk.StringVar()
        self.col_brand_var = tk.StringVar()
        self.header_row_var = tk.IntVar(value=1)
        self.project_var   = tk.StringVar()
        self.output_var    = tk.StringVar(value="内部评审BOM.xlsx")
        self.input_path    = tk.StringVar()
        self.sheet_var     = tk.StringVar()

        self._build_ui()

    # ── 构建界面 ──────────────────────────────────────────────

    def _build_ui(self):
        pad = {"padx": 12, "pady": 6}

        # 标题
        tk.Label(self, text="BOM 转换工具", font=("Arial", 16, "bold"),
                 bg="#2d6cdf", fg="white").pack(fill="x", ipady=10)

        nb = ttk.Notebook(self)
        nb.pack(fill="both", expand=True, padx=10, pady=10)

        self.tab1 = ttk.Frame(nb); nb.add(self.tab1, text="  第一步：选择文件  ")
        self.tab2 = ttk.Frame(nb); nb.add(self.tab2, text="  第二步：列映射  ")
        self.tab3 = ttk.Frame(nb); nb.add(self.tab3, text="  第三步：输出设置  ")
        self.tab4 = ttk.Frame(nb); nb.add(self.tab4, text="  日志  ")
        self.nb = nb

        self._build_tab1()
        self._build_tab2()
        self._build_tab3()
        self._build_tab4()

    def _section(self, parent, title):
        f = ttk.LabelFrame(parent, text=title, padding=10)
        f.pack(fill="x", padx=12, pady=6)
        return f

    # ── Tab1：选择文件 ────────────────────────────────────────

    def _build_tab1(self):
        p = self.tab1

        f1 = self._section(p, "客户 BOM 文件")
        tk.Label(f1, text="文件路径：").grid(row=0, column=0, sticky="w")
        e = ttk.Entry(f1, textvariable=self.input_path, width=50)
        e.grid(row=0, column=1, padx=6)
        ttk.Button(f1, text="浏览...", command=self._browse_file).grid(row=0, column=2)

        f2 = self._section(p, "Sheet 选择")
        tk.Label(f2, text="Sheet：").grid(row=0, column=0, sticky="w")
        self.sheet_cb = ttk.Combobox(f2, textvariable=self.sheet_var, width=30, state="readonly")
        self.sheet_cb.grid(row=0, column=1, padx=6, sticky="w")
        self.sheet_cb.bind("<<ComboboxSelected>>", lambda e: self._load_sheet())

        f3 = self._section(p, "表头行")
        tk.Label(f3, text="表头在第几行：").grid(row=0, column=0, sticky="w")
        ttk.Spinbox(f3, from_=1, to=10, textvariable=self.header_row_var, width=6).grid(row=0, column=1, sticky="w", padx=6)
        ttk.Button(f3, text="重新扫描", command=self._scan_columns).grid(row=0, column=2, padx=6)

        # 预览
        f4 = self._section(p, "文件预览（前5行）")
        self.preview_tree = ttk.Treeview(f4, height=5, show="headings")
        sb = ttk.Scrollbar(f4, orient="horizontal", command=self.preview_tree.xview)
        self.preview_tree.configure(xscrollcommand=sb.set)
        self.preview_tree.pack(fill="x")
        sb.pack(fill="x")

        ttk.Button(p, text="下一步：确认列映射 →", command=lambda: self.nb.select(1))\
            .pack(anchor="e", padx=12, pady=8)

    # ── Tab2：列映射 ─────────────────────────────────────────

    def _build_tab2(self):
        p = self.tab2
        tk.Label(p, text="脚本已自动识别以下列，请确认或手动修改（填列字母，如 A / B / G）",
                 wraplength=700, justify="left", fg="#555").pack(anchor="w", padx=12, pady=8)

        f = self._section(p, "列映射")
        rows = [
            ("物料名称列", self.col_name_var,  "物料的中文品名/描述所在列"),
            ("用量列",     self.col_qty_var,   "每个物料的用量/数量所在列"),
            ("品牌型号列", self.col_brand_var, "品牌:型号||替代料 所在列"),
        ]
        for i, (label, var, hint) in enumerate(rows):
            tk.Label(f, text=label + "：", anchor="w", width=12).grid(row=i, column=0, sticky="w", pady=4)
            ttk.Entry(f, textvariable=var, width=8).grid(row=i, column=1, padx=6)
            tk.Label(f, text=hint, fg="#888").grid(row=i, column=2, sticky="w", padx=6)

        # 自动检测结果展示
        f2 = self._section(p, "自动检测结果")
        self.detect_text = tk.Text(f2, height=10, font=("Consolas", 9), state="disabled",
                                   bg="#fafafa", relief="flat")
        self.detect_text.pack(fill="x")

        ttk.Button(p, text="下一步：设置输出 →", command=lambda: self.nb.select(2))\
            .pack(anchor="e", padx=12, pady=8)

    # ── Tab3：输出设置 ────────────────────────────────────────

    def _build_tab3(self):
        p = self.tab3

        f1 = self._section(p, "项目信息")
        tk.Label(f1, text="项目名称：").grid(row=0, column=0, sticky="w")
        ttk.Entry(f1, textvariable=self.project_var, width=40).grid(row=0, column=1, padx=6, sticky="w")

        f2 = self._section(p, "输出文件")
        tk.Label(f2, text="输出文件名：").grid(row=0, column=0, sticky="w")
        ttk.Entry(f2, textvariable=self.output_var, width=40).grid(row=0, column=1, padx=6, sticky="w")
        ttk.Button(f2, text="另存为...", command=self._browse_output).grid(row=0, column=2, padx=6)

        # 大按钮
        self.run_btn = tk.Button(p, text="开始转换", font=("Arial", 13, "bold"),
                                 bg="#2d6cdf", fg="white", relief="flat",
                                 padx=20, pady=10, command=self._run_convert)
        self.run_btn.pack(pady=20)

        self.status_label = tk.Label(p, text="", font=("Arial", 11), fg="#2d6cdf", bg="#f5f5f5")
        self.status_label.pack()

    # ── Tab4：日志 ────────────────────────────────────────────

    def _build_tab4(self):
        p = self.tab4
        self.log = scrolledtext.ScrolledText(p, font=("Consolas", 9), state="disabled",
                                             bg="#1e1e1e", fg="#d4d4d4", relief="flat")
        self.log.pack(fill="both", expand=True, padx=8, pady=8)
        ttk.Button(p, text="清空日志", command=self._clear_log).pack(anchor="e", padx=8, pady=4)

    # ── 事件处理 ──────────────────────────────────────────────

    def _browse_file(self):
        path = filedialog.askopenfilename(
            title="选择客户BOM文件",
            filetypes=[("Excel文件", "*.xlsx *.xlsm *.xls"), ("所有文件", "*.*")]
        )
        if not path:
            return
        self.input_path.set(path)
        self._log(f"已选择文件：{path}")
        try:
            self.wb = openpyxl.load_workbook(path, data_only=True)
            sheets = self.wb.sheetnames
            self.sheet_cb["values"] = sheets
            self.sheet_var.set(sheets[0])
            self._load_sheet()
        except Exception as e:
            messagebox.showerror("错误", f"无法打开文件：\n{e}")

    def _load_sheet(self):
        if not self.wb:
            return
        name = self.sheet_var.get()
        self.ws = self.wb[name]
        self._log(f"已加载 Sheet：{name}（共 {self.ws.max_row} 行，{self.ws.max_column} 列）")
        self._update_preview()
        self._scan_columns()

    def _update_preview(self):
        tree = self.preview_tree
        tree.delete(*tree.get_children())
        if not self.ws:
            return
        max_col = min(self.ws.max_column, 10)
        cols = [get_column_letter(i) for i in range(1, max_col + 1)]
        tree["columns"] = cols
        for c in cols:
            tree.heading(c, text=c)
            tree.column(c, width=100, anchor="w")
        for ri in range(1, min(6, self.ws.max_row + 1)):
            vals = [str(self.ws.cell(row=ri, column=ci).value or "")[:30]
                    for ci in range(1, max_col + 1)]
            tree.insert("", "end", values=vals)

    def _scan_columns(self):
        if not self.ws:
            messagebox.showwarning("提示", "请先选择文件和Sheet")
            return
        hr = self.header_row_var.get()
        _, best = detect_columns(self.ws, hr)

        self.detect_text.configure(state="normal")
        self.detect_text.delete("1.0", "end")

        lines = ["列字母  表头名称          用途识别    样本值\n" + "-" * 65]
        all_cols, _ = detect_columns(self.ws, hr)
        for ci in sorted(all_cols.keys()):
            info = all_cols[ci]
            role_zh = {"name": "✅ 物料名称", "qty": "✅ 用量", "brand": "✅ 品牌型号", "other": "  -"}.get(info["role"], "")
            sample = " | ".join(info["sample"][:2])[:40]
            lines.append(f"  {info['letter']:<6}  {info['header']:<18}  {role_zh:<14}  {sample}")

        self.detect_text.insert("end", "\n".join(lines))
        self.detect_text.configure(state="disabled")

        # 自动填入识别结果
        if "name" in best:
            self.col_name_var.set(best["name"]["letter"])
        if "qty" in best:
            self.col_qty_var.set(best["qty"]["letter"])
        if "brand" in best:
            self.col_brand_var.set(best["brand"]["letter"])

        self._log(f"列扫描完成（表头行={hr}）：名称={self.col_name_var.get()}，"
                  f"用量={self.col_qty_var.get()}，品牌型号={self.col_brand_var.get()}")

    def _browse_output(self):
        path = filedialog.asksaveasfilename(
            title="保存为",
            defaultextension=".xlsx",
            filetypes=[("Excel文件", "*.xlsx")]
        )
        if path:
            self.output_var.set(path)

    def _run_convert(self):
        # 校验
        if not self.ws:
            messagebox.showerror("错误", "请先选择输入文件"); return
        if not self.col_name_var.get() or not self.col_qty_var.get() or not self.col_brand_var.get():
            messagebox.showerror("错误", "请确认列映射（第二步）"); return
        if not self.project_var.get().strip():
            messagebox.showerror("错误", "请填写项目名称"); return
        if not self.output_var.get().strip():
            messagebox.showerror("错误", "请填写输出文件名"); return

        self.run_btn.configure(state="disabled")
        self.status_label.configure(text="转换中...")
        self.nb.select(3)  # 切到日志页

        threading.Thread(target=self._do_convert, daemon=True).start()

    def _do_convert(self):
        try:
            hr         = self.header_row_var.get()
            col_name   = column_index_from_string(self.col_name_var.get().upper())
            col_qty    = column_index_from_string(self.col_qty_var.get().upper())
            col_brand  = column_index_from_string(self.col_brand_var.get().upper())
            project    = self.project_var.get().strip()
            out_file   = self.output_var.get().strip()

            self._log(f"\n开始转换...")
            self._log(f"  列映射：名称={self.col_name_var.get()}, "
                      f"用量={self.col_qty_var.get()}, 品牌型号={self.col_brand_var.get()}")

            rows = []
            seq = 0
            skipped = 0

            for ri in range(hr + 1, self.ws.max_row + 1):
                name_val  = self.ws.cell(row=ri, column=col_name).value
                qty_val   = self.ws.cell(row=ri, column=col_qty).value
                brand_val = self.ws.cell(row=ri, column=col_brand).value

                if not name_val and not brand_val:
                    skipped += 1
                    continue

                suppliers_raw = parse_brand_model(brand_val) or [("", "")]
                try:
                    main_qty = float(qty_val) if qty_val not in (None, "") else 0
                    main_qty = int(main_qty) if main_qty == int(main_qty) else main_qty
                except:
                    main_qty = qty_val

                suppliers = [
                    (b, m, main_qty if i == 0 else 0)
                    for i, (b, m) in enumerate(suppliers_raw)
                ]

                seq += 1
                rows.append({"seq": seq, "name": str(name_val).strip(), "suppliers": suppliers})

            self._log(f"  解析完成：{len(rows)} 个物料，跳过空行 {skipped} 行")

            total_rows = write_review_bom(rows, out_file, project)
            abs_path = os.path.abspath(out_file)

            self._log(f"  输出完成：{abs_path}")
            self._log(f"  共写入 {total_rows} 行（含替代料展开）")
            self._log(f"\n✅ 转换成功！")

            self.after(0, lambda: self.status_label.configure(
                text=f"✅ 完成！共 {len(rows)} 个物料，{total_rows} 行", fg="#2a8a2a"))
            self.after(0, lambda: messagebox.showinfo(
                "完成", f"转换成功！\n输出文件：\n{abs_path}"))

        except Exception as e:
            self._log(f"\n❌ 错误：{e}")
            import traceback
            self._log(traceback.format_exc())
            self.after(0, lambda: self.status_label.configure(text="❌ 转换失败，请查看日志", fg="red"))
            self.after(0, lambda: messagebox.showerror("错误", str(e)))
        finally:
            self.after(0, lambda: self.run_btn.configure(state="normal"))

    def _log(self, msg):
        def _write():
            self.log.configure(state="normal")
            self.log.insert("end", msg + "\n")
            self.log.see("end")
            self.log.configure(state="disabled")
        self.after(0, _write)

    def _clear_log(self):
        self.log.configure(state="normal")
        self.log.delete("1.0", "end")
        self.log.configure(state="disabled")


if __name__ == "__main__":
    app = BomApp()
    app.mainloop()
