# -*- coding: utf-8 -*-
"""
BOM转换工具 - 图形界面版
支持两种客户格式：
  格式A：品牌型号合并列（MURATA:GRM188||SAMSUNG:CL10）
  格式B：厂家、型号分开列，分号分隔（YAGEO;KOA  /  RC0805JR;RK73Z2）
运行方式：python bom_gui.py
依赖：pip install openpyxl
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter, column_index_from_string
import os, re, threading

# ───────────────────────────── 解析逻辑 ─────────────────────────────

SUPPLIER_LABELS = ["主供", "二供", "三供", "四供", "五供",
                   "六供", "七供", "八供", "九供", "十供"]


def parse_combined(raw):
    """
    格式A：MURATA:GRM188||SAMSUNG:CL10
    支持全角冒号、∥等变体
    返回 [(厂商, 型号), ...]
    """
    if not raw or str(raw).strip() == "":
        return []
    s = str(raw).strip().replace("：", ":").replace("∥", "||").replace("‖", "||")
    result = []
    for entry in [e.strip() for e in re.split(r"\|\|", s) if e.strip()]:
        if ":" in entry:
            b, m = entry.split(":", 1)
            result.append((b.strip(), m.strip()))
        elif "/" in entry and len(entry.split("/")) == 2:
            b, m = entry.split("/", 1)
            result.append((b.strip(), m.strip()))
        else:
            result.append(("", entry.strip()))
    return result


def parse_split(brand_raw, model_raw):
    """
    格式B：厂家列 YAGEO;KOA;FENGHUA / 型号列 RC0805JR;RK73Z2;RS-05000
    返回 [(厂商, 型号), ...]
    """
    brands = [b.strip() for b in str(brand_raw or "").split(";") if b.strip()] if brand_raw else []
    models = [m.strip() for m in str(model_raw or "").split(";") if m.strip()] if model_raw else []

    max_len = max(len(brands), len(models), 1)
    result = []
    for i in range(max_len):
        b = brands[i] if i < len(brands) else ""
        m = models[i] if i < len(models) else ""
        if b or m:
            result.append((b, m))
    return result


def detect_columns(ws, header_row):
    """扫描表格，自动识别各列用途，返回 (all_cols, best)"""
    max_col = ws.max_column
    data_rows = list(range(header_row + 1, min(header_row + 11, ws.max_row + 1)))
    all_cols = {}

    for ci in range(1, max_col + 1):
        header_val = ws.cell(row=header_row, column=ci).value
        header_str = str(header_val).strip() if header_val else ""
        letter = get_column_letter(ci)
        sample_vals = [ws.cell(row=r, column=ci).value for r in data_rows]
        sample_strs = [str(v).strip() for v in sample_vals if v is not None]

        role = "other"
        score = 0

        # 品牌型号合并列（格式A）：含 || 或 厂商:型号
        brand_combined = sum(1 for v in sample_strs if "||" in v or re.search(r"[A-Za-z0-9]+:[A-Za-z0-9]", v))
        if brand_combined >= 2 or "品牌型号" in header_str:
            role = "brand_combined"
            score = brand_combined * 20 + (40 if "品牌型号" in header_str else 0)

        # 厂家列（格式B）：分号分隔、表头含"厂"
        brand_split = sum(1 for v in sample_strs if ";" in v and not re.search(r"[A-Za-z0-9]+:[A-Za-z0-9]", v))
        if any(k in header_str for k in ["厂家", "厂商", "制造商", "Manufacturer", "Brand"]):
            role = "brand_split"
            score = 80
        elif brand_split >= 3 and role == "other":
            role = "brand_split"
            score = brand_split * 15

        # 型号列（格式B）：分号分隔、表头含"型号"但不含"品牌"
        model_split = sum(1 for v in sample_strs if ";" in v)
        if "型号" in header_str and "品牌" not in header_str:
            role = "model_split"
            score = 80
        elif model_split >= 3 and role == "other":
            role = "model_split"
            score = model_split * 12

        # 用量列
        numeric = sum(1 for v in sample_vals if v is not None and str(v).replace(".", "").isdigit())
        if any(k in header_str for k in ["用量", "数量", "qty", "quantity", "Quantity"]):
            role = "qty"; score = 85
        elif numeric >= len(data_rows) * 0.6 and role == "other":
            role = "qty"; score = numeric * 10

        # 物料名称列
        avg_len = sum(len(v) for v in sample_strs) / max(len(sample_strs), 1)
        if any(k in header_str for k in ["名称", "品名", "物料名", "描述", "项目描述", "description", "Description"]):
            if role == "other":
                role = "name"; score = 75
        elif avg_len > 8 and role == "other":
            role = "name"; score = int(avg_len * 2)

        all_cols[ci] = {
            "letter": letter, "header": header_str, "role": role,
            "score": score, "sample": sample_strs[:3],
        }

    # 每种 role 保留得分最高
    best = {}
    for ci, info in all_cols.items():
        r = info["role"]
        if r != "other" and (r not in best or info["score"] > best[r]["score"]):
            best[r] = {"ci": ci, **info}

    return all_cols, best


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

ROLE_LABEL = {
    "brand_combined": "✅ 品牌型号（合并）",
    "brand_split":    "✅ 厂家（分开）",
    "model_split":    "✅ 型号（分开）",
    "qty":            "✅ 用量",
    "name":           "✅ 物料名称",
    "other":          "  -",
}

class BomApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("BOM 转换工具  v2")
        self.geometry("820x720")
        self.resizable(True, True)
        self.configure(bg="#f5f5f5")

        self.wb = None; self.ws = None
        self.col_name_var    = tk.StringVar()
        self.col_qty_var     = tk.StringVar()
        self.col_brand_var   = tk.StringVar()   # 格式A合并列 / 格式B厂家列
        self.col_model_var   = tk.StringVar()   # 格式B专用型号列（留空=格式A）
        self.header_row_var  = tk.IntVar(value=1)
        self.project_var     = tk.StringVar()
        self.output_var      = tk.StringVar(value="内部评审BOM.xlsx")
        self.input_path      = tk.StringVar()
        self.sheet_var       = tk.StringVar()
        self.fmt_var         = tk.StringVar(value="auto")  # auto / A / B

        self._build_ui()

    # ── 构建界面 ──────────────────────────────────────────────

    def _build_ui(self):
        tk.Label(self, text="BOM 转换工具", font=("Arial", 15, "bold"),
                 bg="#2d6cdf", fg="white").pack(fill="x", ipady=10)

        nb = ttk.Notebook(self)
        nb.pack(fill="both", expand=True, padx=10, pady=8)
        self.nb = nb

        self.tab1 = ttk.Frame(nb); nb.add(self.tab1, text="  第一步：选择文件  ")
        self.tab2 = ttk.Frame(nb); nb.add(self.tab2, text="  第二步：列映射  ")
        self.tab3 = ttk.Frame(nb); nb.add(self.tab3, text="  第三步：输出设置  ")
        self.tab4 = ttk.Frame(nb); nb.add(self.tab4, text="  日志  ")

        self._build_tab1()
        self._build_tab2()
        self._build_tab3()
        self._build_tab4()

    def _section(self, parent, title):
        f = ttk.LabelFrame(parent, text=title, padding=10)
        f.pack(fill="x", padx=12, pady=5)
        return f

    # ── Tab1：选择文件 ────────────────────────────────────────

    def _build_tab1(self):
        p = self.tab1

        f1 = self._section(p, "客户 BOM 文件")
        tk.Label(f1, text="文件路径：").grid(row=0, column=0, sticky="w")
        ttk.Entry(f1, textvariable=self.input_path, width=52).grid(row=0, column=1, padx=6)
        ttk.Button(f1, text="浏览...", command=self._browse_file).grid(row=0, column=2)

        f2 = self._section(p, "Sheet / 表头行")
        tk.Label(f2, text="Sheet：").grid(row=0, column=0, sticky="w")
        self.sheet_cb = ttk.Combobox(f2, textvariable=self.sheet_var, width=28, state="readonly")
        self.sheet_cb.grid(row=0, column=1, padx=6, sticky="w")
        self.sheet_cb.bind("<<ComboboxSelected>>", lambda e: self._load_sheet())
        tk.Label(f2, text="  表头行：").grid(row=0, column=2, sticky="w")
        ttk.Spinbox(f2, from_=1, to=10, textvariable=self.header_row_var, width=5).grid(row=0, column=3)
        ttk.Button(f2, text="重新扫描", command=self._scan_columns).grid(row=0, column=4, padx=8)

        f3 = self._section(p, "文件预览（前5行）")
        self.preview_tree = ttk.Treeview(f3, height=5, show="headings")
        sx = ttk.Scrollbar(f3, orient="horizontal", command=self.preview_tree.xview)
        self.preview_tree.configure(xscrollcommand=sx.set)
        self.preview_tree.pack(fill="x"); sx.pack(fill="x")

        ttk.Button(p, text="下一步：确认列映射 →",
                   command=lambda: self.nb.select(1)).pack(anchor="e", padx=12, pady=8)

    # ── Tab2：列映射 ─────────────────────────────────────────

    def _build_tab2(self):
        p = self.tab2

        # 格式选择
        ff = self._section(p, "品牌型号格式")
        tk.Label(ff, text="格式：").grid(row=0, column=0, sticky="w")
        for val, txt in [("auto", "自动识别"),
                          ("A", "格式A：合并列（MURATA:GRM188||SAMSUNG:CL10）"),
                          ("B", "格式B：分开列（厂家 / 型号 各一列，分号分隔）")]:
            rb = ttk.Radiobutton(ff, text=txt, variable=self.fmt_var, value=val,
                                 command=self._update_fmt_ui)
            rb.grid(row=0 if val == "auto" else (1 if val == "A" else 2),
                    column=1, sticky="w", pady=2)

        # 列映射输入
        fm = self._section(p, "列映射（填列字母，如 A / D / E / F）")
        fields = [
            ("物料名称列",      self.col_name_var,  "物料品名/描述所在列"),
            ("用量列",          self.col_qty_var,   "用量/数量所在列"),
            ("品牌型号列\n（格式A）", self.col_brand_var, "格式A：厂商:型号||... 合并列"),
            ("厂家列\n（格式B）",     self.col_brand_var, "格式B：YAGEO;KOA;FENGHUA"),
            ("型号列\n（格式B）",     self.col_model_var, "格式B：RC0805JR;RK73Z2;RS-05000（留空=格式A）"),
        ]
        # 实际只显示：名称、用量、品牌/厂家、型号（格式B）
        labels_vars = [
            ("物料名称列",  self.col_name_var,  "物料品名/描述所在列"),
            ("用量列",      self.col_qty_var,   "用量/数量所在列"),
            ("品牌/厂家列", self.col_brand_var, "格式A=品牌型号合并列；格式B=厂家列"),
            ("型号列",      self.col_model_var, "仅格式B需要填写，格式A留空"),
        ]
        self.model_row_widgets = []
        for i, (lbl, var, hint) in enumerate(labels_vars):
            tk.Label(fm, text=lbl + "：", anchor="w", width=14).grid(row=i, column=0, sticky="w", pady=3)
            e = ttk.Entry(fm, textvariable=var, width=8)
            e.grid(row=i, column=1, padx=6)
            lh = tk.Label(fm, text=hint, fg="#777")
            lh.grid(row=i, column=2, sticky="w", padx=6)
            if i == 3:
                self.model_row_widgets = [e, lh]

        # 自动检测结果
        f2 = self._section(p, "自动检测结果")
        self.detect_text = tk.Text(f2, height=9, font=("Consolas", 9),
                                   state="disabled", bg="#fafafa", relief="flat")
        self.detect_text.pack(fill="x")

        ttk.Button(p, text="下一步：设置输出 →",
                   command=lambda: self.nb.select(2)).pack(anchor="e", padx=12, pady=8)

    def _update_fmt_ui(self):
        fmt = self.fmt_var.get()
        # 格式B才需要型号列
        state = "normal" if fmt in ("auto", "B") else "disabled"
        for w in self.model_row_widgets:
            try: w.configure(state=state)
            except: pass

    # ── Tab3：输出设置 ────────────────────────────────────────

    def _build_tab3(self):
        p = self.tab3
        f1 = self._section(p, "项目信息")
        tk.Label(f1, text="项目名称：").grid(row=0, column=0, sticky="w")
        ttk.Entry(f1, textvariable=self.project_var, width=42).grid(row=0, column=1, padx=6, sticky="w")

        f2 = self._section(p, "输出文件")
        tk.Label(f2, text="输出文件名：").grid(row=0, column=0, sticky="w")
        ttk.Entry(f2, textvariable=self.output_var, width=42).grid(row=0, column=1, padx=6, sticky="w")
        ttk.Button(f2, text="另存为...", command=self._browse_output).grid(row=0, column=2, padx=6)

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
        if not path: return
        self.input_path.set(path)
        self._log(f"已选择文件：{path}")
        try:
            self.wb = openpyxl.load_workbook(path, data_only=True)
            self.sheet_cb["values"] = self.wb.sheetnames
            self.sheet_var.set(self.wb.sheetnames[0])
            self._load_sheet()
        except Exception as e:
            messagebox.showerror("错误", f"无法打开文件：\n{e}")

    def _load_sheet(self):
        if not self.wb: return
        self.ws = self.wb[self.sheet_var.get()]
        self._log(f"Sheet：{self.ws.title}（{self.ws.max_row}行 × {self.ws.max_column}列）")
        self._update_preview()
        self._scan_columns()

    def _update_preview(self):
        tree = self.preview_tree
        tree.delete(*tree.get_children())
        if not self.ws: return
        mc = min(self.ws.max_column, 10)
        cols = [get_column_letter(i) for i in range(1, mc + 1)]
        tree["columns"] = cols
        for c in cols:
            tree.heading(c, text=c); tree.column(c, width=100, anchor="w")
        for ri in range(1, min(6, self.ws.max_row + 1)):
            vals = [str(self.ws.cell(row=ri, column=ci).value or "")[:28]
                    for ci in range(1, mc + 1)]
            tree.insert("", "end", values=vals)

    def _scan_columns(self):
        if not self.ws:
            messagebox.showwarning("提示", "请先选择文件和Sheet"); return
        hr = self.header_row_var.get()
        all_cols, best = detect_columns(self.ws, hr)

        self.detect_text.configure(state="normal")
        self.detect_text.delete("1.0", "end")
        lines = ["列    表头                  识别用途              样本\n" + "-"*72]
        for ci in sorted(all_cols.keys()):
            info = all_cols[ci]
            role_zh = ROLE_LABEL.get(info["role"], "  -")
            sample  = " | ".join(info["sample"][:2])[:35]
            lines.append(f"  {info['letter']:<5} {info['header']:<22} {role_zh:<22} {sample}")
        self.detect_text.insert("end", "\n".join(lines))
        self.detect_text.configure(state="disabled")

        # 自动填入
        if "name" in best:
            self.col_name_var.set(best["name"]["letter"])
        if "qty" in best:
            self.col_qty_var.set(best["qty"]["letter"])

        # 品牌型号列自动判断格式
        if "brand_combined" in best:
            self.col_brand_var.set(best["brand_combined"]["letter"])
            self.col_model_var.set("")
            self.fmt_var.set("A")
        elif "brand_split" in best:
            self.col_brand_var.set(best["brand_split"]["letter"])
            if "model_split" in best:
                self.col_model_var.set(best["model_split"]["letter"])
            self.fmt_var.set("B")
        self._update_fmt_ui()

        self._log(f"扫描完成（表头行={hr}）：名称={self.col_name_var.get()}，"
                  f"用量={self.col_qty_var.get()}，品牌/厂家={self.col_brand_var.get()}，"
                  f"型号={self.col_model_var.get() or '(格式A不需要)'}，格式={self.fmt_var.get()}")

    def _browse_output(self):
        path = filedialog.asksaveasfilename(
            defaultextension=".xlsx", filetypes=[("Excel文件", "*.xlsx")])
        if path: self.output_var.set(path)

    def _run_convert(self):
        if not self.ws:
            messagebox.showerror("错误", "请先选择输入文件"); return
        if not self.col_name_var.get() or not self.col_qty_var.get() or not self.col_brand_var.get():
            messagebox.showerror("错误", "请确认列映射（第二步）"); return
        if not self.project_var.get().strip():
            messagebox.showerror("错误", "请填写项目名称（第三步）"); return
        self.run_btn.configure(state="disabled")
        self.status_label.configure(text="转换中...", fg="#2d6cdf")
        self.nb.select(3)
        threading.Thread(target=self._do_convert, daemon=True).start()

    def _do_convert(self):
        try:
            hr        = self.header_row_var.get()
            col_name  = column_index_from_string(self.col_name_var.get().upper())
            col_qty   = column_index_from_string(self.col_qty_var.get().upper())
            col_brand = column_index_from_string(self.col_brand_var.get().upper())
            col_model = column_index_from_string(self.col_model_var.get().upper()) \
                        if self.col_model_var.get().strip() else None
            fmt       = self.fmt_var.get()
            project   = self.project_var.get().strip()
            out_file  = self.output_var.get().strip()

            use_split = (fmt == "B") or (fmt == "auto" and col_model is not None)
            self._log(f"\n开始转换（{'格式B：厂家/型号分开' if use_split else '格式A：合并列'}）")

            rows = []; seq = 0; skipped = 0

            for ri in range(hr + 1, self.ws.max_row + 1):
                name_val  = self.ws.cell(row=ri, column=col_name).value
                qty_val   = self.ws.cell(row=ri, column=col_qty).value
                brand_val = self.ws.cell(row=ri, column=col_brand).value
                model_val = self.ws.cell(row=ri, column=col_model).value if col_model else None

                if not name_val and not brand_val:
                    skipped += 1; continue

                if use_split:
                    suppliers_raw = parse_split(brand_val, model_val)
                else:
                    suppliers_raw = parse_combined(brand_val)
                if not suppliers_raw:
                    suppliers_raw = [("", "")]

                try:
                    mq = float(qty_val) if qty_val not in (None, "") else 0
                    mq = int(mq) if mq == int(mq) else mq
                except:
                    mq = qty_val

                suppliers = [(b, m, mq if i == 0 else 0)
                             for i, (b, m) in enumerate(suppliers_raw)]
                seq += 1
                rows.append({"seq": seq, "name": str(name_val).strip(), "suppliers": suppliers})

            self._log(f"  解析：{len(rows)} 个物料，跳过空行 {skipped}")
            total = write_review_bom(rows, out_file, project)
            abs_path = os.path.abspath(out_file)
            self._log(f"  输出：{abs_path}")
            self._log(f"  共写入 {total} 行（含替代料展开）\n✅ 转换成功！")

            self.after(0, lambda: self.status_label.configure(
                text=f"✅ 完成！{len(rows)} 个物料，{total} 行", fg="#2a8a2a"))
            self.after(0, lambda: messagebox.showinfo("完成", f"转换成功！\n{abs_path}"))

        except Exception as e:
            import traceback
            self._log(f"\n❌ 错误：{e}\n{traceback.format_exc()}")
            self.after(0, lambda: self.status_label.configure(text="❌ 转换失败，请查看日志", fg="red"))
            self.after(0, lambda: messagebox.showerror("错误", str(e)))
        finally:
            self.after(0, lambda: self.run_btn.configure(state="normal"))

    def _log(self, msg):
        def _w():
            self.log.configure(state="normal")
            self.log.insert("end", msg + "\n")
            self.log.see("end")
            self.log.configure(state="disabled")
        self.after(0, _w)

    def _clear_log(self):
        self.log.configure(state="normal")
        self.log.delete("1.0", "end")
        self.log.configure(state="disabled")


if __name__ == "__main__":
    app = BomApp()
    app.mainloop()
