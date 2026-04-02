# -*- coding: utf-8 -*-
"""
BOM 转换工具 v3 - 科技风深色界面
依赖安装：pip install customtkinter openpyxl
运行方式：python bom_gui.py
"""

import customtkinter as ctk
from tkinter import filedialog, messagebox
import tkinter as tk
from tkinter import ttk
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter, column_index_from_string
import os, re, threading

# ── 主题配置 ──────────────────────────────────────────────────
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

C_BG      = "#0d1117"   # 主背景
C_SURFACE = "#161b22"   # 卡片背景
C_BORDER  = "#30363d"   # 边框
C_ACCENT  = "#1f6feb"   # 蓝色主色
C_ACCENT2 = "#58a6ff"   # 浅蓝高亮
C_GREEN   = "#3fb950"   # 成功绿
C_RED     = "#f85149"   # 错误红
C_YELLOW  = "#d29922"   # 警告黄
C_TEXT    = "#e6edf3"   # 主文字
C_MUTED   = "#8b949e"   # 次要文字

# ───────────────────── 解析 & 输出逻辑 ──────────────────────

SUPPLIER_LABELS = ["主供","二供","三供","四供","五供","六供","七供","八供","九供","十供"]

def parse_combined(raw):
    if not raw or str(raw).strip() == "": return []
    s = str(raw).strip().replace("：",":").replace("∥","||").replace("‖","||")
    result = []
    for entry in [e.strip() for e in re.split(r"\|\|", s) if e.strip()]:
        if ":" in entry:
            b,m = entry.split(":",1); result.append((b.strip(),m.strip()))
        elif "/" in entry and len(entry.split("/"))==2:
            b,m = entry.split("/",1); result.append((b.strip(),m.strip()))
        else:
            result.append(("",entry.strip()))
    return result

def parse_split(brand_raw, model_raw):
    brands = [b.strip() for b in str(brand_raw or "").split(";") if b.strip()] if brand_raw else []
    models = [m.strip() for m in str(model_raw or "").split(";") if m.strip()] if model_raw else []
    result = []
    for i in range(max(len(brands),len(models),1)):
        b = brands[i] if i < len(brands) else ""
        m = models[i] if i < len(models) else ""
        if b or m: result.append((b,m))
    return result

def detect_columns(ws, header_row):
    data_rows = list(range(header_row+1, min(header_row+11, ws.max_row+1)))
    all_cols = {}
    for ci in range(1, ws.max_column+1):
        hv = ws.cell(row=header_row, column=ci).value
        hs = str(hv).strip() if hv else ""
        letter = get_column_letter(ci)
        samples = [ws.cell(row=r, column=ci).value for r in data_rows]
        strs = [str(v).strip() for v in samples if v is not None]
        role="other"; score=0
        b_comb = sum(1 for v in strs if "||" in v or re.search(r"[A-Za-z0-9]+:[A-Za-z0-9]",v))
        if b_comb>=2 or "品牌型号" in hs: role="brand_combined"; score=b_comb*20+(40 if "品牌型号" in hs else 0)
        b_split = sum(1 for v in strs if ";" in v and not re.search(r"[A-Za-z0-9]+:[A-Za-z0-9]",v))
        if any(k in hs for k in ["厂家","厂商","制造商","Manufacturer","Brand"]): role="brand_split"; score=80
        elif b_split>=3 and role=="other": role="brand_split"; score=b_split*15
        m_split = sum(1 for v in strs if ";" in v)
        if "型号" in hs and "品牌" not in hs: role="model_split"; score=80
        elif m_split>=3 and role=="other": role="model_split"; score=m_split*12
        numeric = sum(1 for v in samples if v is not None and str(v).replace(".","").isdigit())
        if any(k in hs for k in ["用量","数量","qty","quantity","Quantity"]): role="qty"; score=85
        elif numeric>=len(data_rows)*0.6 and role=="other": role="qty"; score=numeric*10
        avg_len = sum(len(v) for v in strs)/max(len(strs),1)
        if any(k in hs for k in ["名称","品名","物料名","描述","项目描述","description","Description"]):
            if role=="other": role="name"; score=75
        elif avg_len>8 and role=="other": role="name"; score=int(avg_len*2)
        all_cols[ci]={"letter":letter,"header":hs,"role":role,"score":score,"sample":strs[:3]}
    best={}
    for ci,info in all_cols.items():
        r=info["role"]
        if r!="other" and (r not in best or info["score"]>best[r]["score"]): best[r]={"ci":ci,**info}
    return all_cols, best

def write_review_bom(rows, output_file, project_name):
    wb=Workbook(); ws=wb.active; ws.title="SW节点整机BOM配置"
    GREEN="92D050"; YELLOW="FFFF00"; ORANGE="FFC000"
    thin=Side(style="thin"); bdr=Border(left=thin,right=thin,top=thin,bottom=thin)
    def S(cell,bold=False,bg=None,color="000000",h="center",v="center",size=11,wrap=False):
        cell.font=Font(bold=bold,color=color,size=size)
        if bg: cell.fill=PatternFill("solid",start_color=bg)
        cell.alignment=Alignment(horizontal=h,vertical=v,wrap_text=wrap)
    ws.merge_cells("A1:A2"); ws["A1"]="项目名称"; S(ws["A1"],bold=True,bg=GREEN,size=14)
    ws.merge_cells("B1:B2"); ws["B1"]=project_name; S(ws["B1"],bold=True,bg=GREEN,size=14)
    ws.merge_cells("E1:I2"); ws["E1"]="整机BOM配置表"; S(ws["E1"],bold=True,bg=GREEN,size=16)
    ws["J1"]="配置说明"; S(ws["J1"],bold=True,bg=GREEN)
    ws["K1"]="TBD"; S(ws["K1"],bg="BDD7EE")
    ws.row_dimensions[1].height=30
    ws.merge_cells("A3:I3"); ws["A3"]="SW节点HQ SN"
    S(ws["A3"],bold=True,bg=YELLOW,color="FF0000",size=12)
    ws["K3"]=""; S(ws["K3"],bg=ORANGE,color="FF0000"); ws.row_dimensions[3].height=20
    headers=["序号","组件子类","虚拟层/物料","物料类型","HQ PN","物料名称","厂商型号","厂商","主二供","","用量"]
    for ci,h in enumerate(headers,1):
        c=ws.cell(row=4,column=ci,value=h); S(c,bold=True,bg="D9D9D9"); c.border=bdr
    ws.row_dimensions[4].height=22
    dr=5
    for item in rows:
        for si,(brand,model,qty) in enumerate(item["suppliers"]):
            label=SUPPLIER_LABELS[si] if si<len(SUPPLIER_LABELS) else f"{si+1}供"
            for ci,val in enumerate([item["seq"],"","","","",item["name"],model,brand,label,"",qty],1):
                c=ws.cell(row=dr,column=ci,value=val); c.border=bdr
                c.alignment=Alignment(horizontal="center",vertical="center")
            dr+=1
    for i,w in enumerate([6,10,12,10,18,35,30,20,8,6,8],1):
        ws.column_dimensions[get_column_letter(i)].width=w
    wb.save(output_file); return dr-5

# ───────────────────── GUI ───────────────────────────────────

class BomApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("BOM 转换工具  v3.0")
        self.geometry("960x680")
        self.configure(fg_color=C_BG)
        self.resizable(True, True)

        self.wb=None; self.ws=None
        self.col_name_var  = ctk.StringVar()
        self.col_qty_var   = ctk.StringVar()
        self.col_brand_var = ctk.StringVar()
        self.col_model_var = ctk.StringVar()
        self.header_row_var= ctk.IntVar(value=1)
        self.project_var   = ctk.StringVar()
        self.output_var    = ctk.StringVar(value="内部评审BOM.xlsx")
        self.input_path    = ctk.StringVar()
        self.sheet_var     = ctk.StringVar()
        self.fmt_var       = ctk.StringVar(value="auto")
        self._active_step  = 0

        self._build()

    # ── 顶栏 ──────────────────────────────────────────────────

    def _build(self):
        # 顶部标题栏
        hdr = ctk.CTkFrame(self, fg_color=C_SURFACE, height=56, corner_radius=0)
        hdr.pack(fill="x")
        hdr.pack_propagate(False)
        ctk.CTkLabel(hdr, text="◈  BOM 转换工具", font=ctk.CTkFont("Arial",18,"bold"),
                     text_color=C_ACCENT2).pack(side="left", padx=20)
        ctk.CTkLabel(hdr, text="v3.0  |  HQ内部评审格式生成器",
                     font=ctk.CTkFont("Arial",11), text_color=C_MUTED).pack(side="left")
        self.status_chip = ctk.CTkLabel(hdr, text="● 就绪", font=ctk.CTkFont("Arial",11),
                                         text_color=C_GREEN)
        self.status_chip.pack(side="right", padx=20)

        # 主体：左侧导航 + 右侧内容
        body = ctk.CTkFrame(self, fg_color=C_BG)
        body.pack(fill="both", expand=True)

        self._build_sidebar(body)

        self.content = ctk.CTkFrame(body, fg_color=C_BG)
        self.content.pack(side="left", fill="both", expand=True, padx=(0,16), pady=12)

        # 页面
        self.pages = {}
        for i, fn in enumerate([self._page_file, self._page_cols,
                                 self._page_output, self._page_log]):
            f = ctk.CTkFrame(self.content, fg_color=C_BG)
            fn(f)
            self.pages[i] = f

        self._show_step(0)

        # 状态栏
        bar = ctk.CTkFrame(self, fg_color=C_SURFACE, height=28, corner_radius=0)
        bar.pack(fill="x", side="bottom")
        bar.pack_propagate(False)
        self.bar_label = ctk.CTkLabel(bar, text="请先选择客户BOM文件",
                                       font=ctk.CTkFont("Arial",10), text_color=C_MUTED)
        self.bar_label.pack(side="left", padx=16)

    # ── 左侧导航 ─────────────────────────────────────────────

    def _build_sidebar(self, parent):
        nav = ctk.CTkFrame(parent, fg_color=C_SURFACE, width=180, corner_radius=0)
        nav.pack(side="left", fill="y", pady=0)
        nav.pack_propagate(False)

        ctk.CTkLabel(nav, text="STEPS", font=ctk.CTkFont("Arial",10,"bold"),
                     text_color=C_MUTED).pack(anchor="w", padx=20, pady=(20,8))

        steps = ["01  选择文件", "02  列映射", "03  输出设置", "04  日志"]
        self.nav_btns = []
        for i, label in enumerate(steps):
            btn = ctk.CTkButton(
                nav, text=label, anchor="w",
                font=ctk.CTkFont("Arial", 13),
                fg_color="transparent", hover_color="#21262d",
                text_color=C_MUTED, corner_radius=6, height=40,
                command=lambda x=i: self._show_step(x)
            )
            btn.pack(fill="x", padx=10, pady=2)
            self.nav_btns.append(btn)

        # 底部转换按钮
        ctk.CTkFrame(nav, fg_color=C_BORDER, height=1).pack(fill="x", padx=10, pady=12)
        self.run_btn = ctk.CTkButton(
            nav, text="▶  开始转换",
            font=ctk.CTkFont("Arial",13,"bold"),
            fg_color=C_ACCENT, hover_color="#388bfd",
            height=42, corner_radius=8,
            command=self._run_convert
        )
        self.run_btn.pack(fill="x", padx=10, pady=4)

    def _show_step(self, idx):
        for p in self.pages.values(): p.pack_forget()
        self.pages[idx].pack(fill="both", expand=True)
        self._active_step = idx
        for i, btn in enumerate(self.nav_btns):
            if i == idx:
                btn.configure(text_color=C_ACCENT2, fg_color="#1c2d3e")
            else:
                btn.configure(text_color=C_MUTED, fg_color="transparent")

    # ── 通用组件 ─────────────────────────────────────────────

    def _card(self, parent, title, **kwargs):
        outer = ctk.CTkFrame(parent, fg_color=C_SURFACE, corner_radius=10, **kwargs)
        outer.pack(fill="x", pady=6)
        if title:
            ctk.CTkLabel(outer, text=title, font=ctk.CTkFont("Arial",11,"bold"),
                         text_color=C_ACCENT2).pack(anchor="w", padx=16, pady=(12,4))
            ctk.CTkFrame(outer, fg_color=C_BORDER, height=1).pack(fill="x", padx=16, pady=(0,8))
        return outer

    def _field_row(self, parent, label, var, hint="", width=180, placeholder=""):
        row = ctk.CTkFrame(parent, fg_color="transparent")
        row.pack(fill="x", padx=16, pady=4)
        ctk.CTkLabel(row, text=label, font=ctk.CTkFont("Arial",11),
                     text_color=C_TEXT, width=120, anchor="w").pack(side="left")
        e = ctk.CTkEntry(row, textvariable=var, width=width,
                         placeholder_text=placeholder,
                         fg_color="#0d1117", border_color=C_BORDER,
                         text_color=C_TEXT)
        e.pack(side="left", padx=8)
        if hint:
            ctk.CTkLabel(row, text=hint, font=ctk.CTkFont("Arial",10),
                         text_color=C_MUTED).pack(side="left", padx=4)
        return e

    # ── 页面1：选择文件 ───────────────────────────────────────

    def _page_file(self, p):
        ctk.CTkLabel(p, text="选择客户 BOM 文件",
                     font=ctk.CTkFont("Arial",16,"bold"),
                     text_color=C_TEXT).pack(anchor="w", pady=(0,12))

        c1 = self._card(p, "文件路径")
        row = ctk.CTkFrame(c1, fg_color="transparent")
        row.pack(fill="x", padx=16, pady=(0,12))
        ctk.CTkEntry(row, textvariable=self.input_path, width=480,
                     fg_color="#0d1117", border_color=C_BORDER,
                     text_color=C_TEXT, placeholder_text="点击「浏览」选择 .xlsx 文件")\
            .pack(side="left", padx=(0,8))
        ctk.CTkButton(row, text="浏览", width=80, fg_color=C_ACCENT,
                      hover_color="#388bfd", command=self._browse_file).pack(side="left")

        c2 = self._card(p, "Sheet 与表头行")
        row2 = ctk.CTkFrame(c2, fg_color="transparent")
        row2.pack(fill="x", padx=16, pady=(0,8))
        ctk.CTkLabel(row2, text="Sheet：", text_color=C_TEXT,
                     font=ctk.CTkFont("Arial",11)).pack(side="left")
        self.sheet_cb = ctk.CTkComboBox(row2, variable=self.sheet_var, width=220,
                                         values=[], state="readonly",
                                         fg_color="#0d1117", border_color=C_BORDER,
                                         text_color=C_TEXT, button_color=C_ACCENT,
                                         command=lambda _: self._load_sheet())
        self.sheet_cb.pack(side="left", padx=8)
        ctk.CTkLabel(row2, text="表头行：", text_color=C_TEXT,
                     font=ctk.CTkFont("Arial",11)).pack(side="left", padx=(16,0))
        ctk.CTkEntry(row2, textvariable=self.header_row_var, width=50,
                     fg_color="#0d1117", border_color=C_BORDER,
                     text_color=C_TEXT).pack(side="left", padx=8)
        ctk.CTkButton(row2, text="重新扫描", width=90, fg_color="#21262d",
                      hover_color="#30363d", border_width=1, border_color=C_BORDER,
                      command=self._scan_columns).pack(side="left", padx=8)

        c3 = self._card(p, "文件预览（前5行）")
        # 用原生 ttk.Treeview（customtkinter 暂无表格组件）
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Tech.Treeview", background="#0d1117", foreground=C_TEXT,
                        fieldbackground="#0d1117", borderwidth=0, rowheight=22,
                        font=("Consolas",9))
        style.configure("Tech.Treeview.Heading", background=C_SURFACE,
                        foreground=C_ACCENT2, font=("Arial",9,"bold"))
        style.map("Tech.Treeview", background=[("selected","#1c2d3e")])

        frame_tv = ctk.CTkFrame(c3, fg_color="#0d1117", corner_radius=6)
        frame_tv.pack(fill="x", padx=16, pady=(0,12))
        self.preview_tree = ttk.Treeview(frame_tv, style="Tech.Treeview",
                                          height=5, show="headings")
        sx = ttk.Scrollbar(frame_tv, orient="horizontal",
                           command=self.preview_tree.xview)
        self.preview_tree.configure(xscrollcommand=sx.set)
        self.preview_tree.pack(fill="x"); sx.pack(fill="x")

        ctk.CTkButton(p, text="下一步 →", width=120, fg_color=C_ACCENT,
                      hover_color="#388bfd",
                      command=lambda: self._show_step(1)).pack(anchor="e", pady=8)

    # ── 页面2：列映射 ─────────────────────────────────────────

    def _page_cols(self, p):
        ctk.CTkLabel(p, text="列映射配置",
                     font=ctk.CTkFont("Arial",16,"bold"),
                     text_color=C_TEXT).pack(anchor="w", pady=(0,12))

        c1 = self._card(p, "品牌型号格式")
        row = ctk.CTkFrame(c1, fg_color="transparent")
        row.pack(fill="x", padx=16, pady=(0,12))
        for val, txt in [("auto","自动识别"),
                          ("A","格式A：合并列（MURATA:GRM188||SAMSUNG:CL10）"),
                          ("B","格式B：分开列（厂家 / 型号 各一列，分号分隔）")]:
            ctk.CTkRadioButton(row, text=txt, variable=self.fmt_var, value=val,
                               text_color=C_TEXT, fg_color=C_ACCENT,
                               hover_color="#388bfd",
                               command=self._update_fmt_ui).pack(side="left", padx=12)

        c2 = self._card(p, "列位置（填列字母，如 A / D / G）")
        self._field_row(c2, "物料名称列", self.col_name_var, "物料品名/描述", 80, "如 D")
        self._field_row(c2, "用量列",     self.col_qty_var,  "数量/用量",      80, "如 E")
        self._field_row(c2, "品牌/厂家列",self.col_brand_var,"格式A合并 / 格式B厂家", 80, "如 G")
        self.model_row = self._field_row(c2, "型号列（格式B）", self.col_model_var,
                                          "格式B专用，格式A留空", 80, "如 H")
        ctk.CTkFrame(c2, fg_color="transparent", height=8).pack()

        c3 = self._card(p, "自动扫描结果")
        self.detect_text = ctk.CTkTextbox(c3, height=180, font=ctk.CTkFont("Consolas",10),
                                           fg_color="#0d1117", text_color=C_TEXT,
                                           border_color=C_BORDER, border_width=1,
                                           state="disabled")
        self.detect_text.pack(fill="x", padx=16, pady=(0,12))

        ctk.CTkButton(p, text="下一步 →", width=120, fg_color=C_ACCENT,
                      hover_color="#388bfd",
                      command=lambda: self._show_step(2)).pack(anchor="e", pady=8)

    def _update_fmt_ui(self):
        pass  # 格式提示用，不禁用

    # ── 页面3：输出设置 ───────────────────────────────────────

    def _page_output(self, p):
        ctk.CTkLabel(p, text="输出设置",
                     font=ctk.CTkFont("Arial",16,"bold"),
                     text_color=C_TEXT).pack(anchor="w", pady=(0,12))

        c1 = self._card(p, "项目信息")
        self._field_row(c1, "项目名称", self.project_var,
                        "", 320, "如 TPL / HQ-2024-001")
        ctk.CTkFrame(c1, fg_color="transparent", height=8).pack()

        c2 = self._card(p, "输出文件")
        row = ctk.CTkFrame(c2, fg_color="transparent")
        row.pack(fill="x", padx=16, pady=(0,12))
        ctk.CTkLabel(row, text="保存路径：", text_color=C_TEXT,
                     font=ctk.CTkFont("Arial",11), width=120).pack(side="left")
        ctk.CTkEntry(row, textvariable=self.output_var, width=360,
                     fg_color="#0d1117", border_color=C_BORDER,
                     text_color=C_TEXT).pack(side="left", padx=8)
        ctk.CTkButton(row, text="另存为", width=80, fg_color="#21262d",
                      hover_color="#30363d", border_width=1, border_color=C_BORDER,
                      command=self._browse_output).pack(side="left")

        # 格式说明卡片
        c3 = self._card(p, "输出格式说明")
        lines = [
            "→  序号相同的行 = 同一物料的主供 / 替代料",
            "→  第一行：主供（用量=实际数值）",
            "→  其余行：二供/三供...（用量=0）",
            "→  HQ PN 列留空，评审时由工程师填写",
        ]
        for line in lines:
            ctk.CTkLabel(c3, text=line, font=ctk.CTkFont("Consolas",10),
                         text_color=C_MUTED, anchor="w").pack(anchor="w", padx=16, pady=1)
        ctk.CTkFrame(c3, fg_color="transparent", height=8).pack()

    # ── 页面4：日志 ───────────────────────────────────────────

    def _page_log(self, p):
        ctk.CTkLabel(p, text="转换日志",
                     font=ctk.CTkFont("Arial",16,"bold"),
                     text_color=C_TEXT).pack(anchor="w", pady=(0,8))

        self.log_box = ctk.CTkTextbox(p, font=ctk.CTkFont("Consolas",10),
                                       fg_color="#0d1117", text_color="#79c0ff",
                                       border_color=C_BORDER, border_width=1,
                                       state="disabled")
        self.log_box.pack(fill="both", expand=True)

        ctk.CTkButton(p, text="清空", width=80, fg_color="#21262d",
                      hover_color="#30363d", border_width=1, border_color=C_BORDER,
                      command=self._clear_log).pack(anchor="e", pady=6)

    # ── 事件 ─────────────────────────────────────────────────

    def _browse_file(self):
        path = filedialog.askopenfilename(
            title="选择客户BOM",
            filetypes=[("Excel文件","*.xlsx *.xlsm *.xls"),("所有文件","*.*")])
        if not path: return
        self.input_path.set(path)
        self._log(f"文件：{path}")
        try:
            self.wb = openpyxl.load_workbook(path, data_only=True)
            sheets = self.wb.sheetnames
            self.sheet_cb.configure(values=sheets)
            self.sheet_var.set(sheets[0])
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
        mc = min(self.ws.max_column, 9)
        cols = [get_column_letter(i) for i in range(1, mc+1)]
        tree["columns"] = cols
        for c in cols:
            tree.heading(c, text=c); tree.column(c, width=110, anchor="w")
        for ri in range(1, min(6, self.ws.max_row+1)):
            vals = [str(self.ws.cell(row=ri,column=ci).value or "")[:26]
                    for ci in range(1, mc+1)]
            tree.insert("","end",values=vals)

    def _scan_columns(self):
        if not self.ws:
            messagebox.showwarning("提示","请先选择文件"); return
        hr = self.header_row_var.get()
        all_cols, best = detect_columns(self.ws, hr)

        role_label = {
            "brand_combined": "✅ 品牌型号(合并)",
            "brand_split":    "✅ 厂家(分开)",
            "model_split":    "✅ 型号(分开)",
            "qty":            "✅ 用量",
            "name":           "✅ 物料名称",
            "other":          "   -",
        }
        self.detect_text.configure(state="normal")
        self.detect_text.delete("0.0","end")
        lines = [f"{'列':<4} {'表头':<22} {'识别用途':<24} 样本\n" + "─"*70]
        for ci in sorted(all_cols.keys()):
            info = all_cols[ci]
            rz = role_label.get(info["role"],"   -")
            sample = " | ".join(info["sample"][:2])[:32]
            lines.append(f" {info['letter']:<4} {info['header']:<22} {rz:<24} {sample}")
        self.detect_text.insert("0.0", "\n".join(lines))
        self.detect_text.configure(state="disabled")

        if "name"  in best: self.col_name_var.set(best["name"]["letter"])
        if "qty"   in best: self.col_qty_var.set(best["qty"]["letter"])
        if "brand_combined" in best:
            self.col_brand_var.set(best["brand_combined"]["letter"])
            self.col_model_var.set(""); self.fmt_var.set("A")
        elif "brand_split" in best:
            self.col_brand_var.set(best["brand_split"]["letter"])
            if "model_split" in best: self.col_model_var.set(best["model_split"]["letter"])
            self.fmt_var.set("B")

        self._log(f"扫描完成 → 名称={self.col_name_var.get()} 用量={self.col_qty_var.get()} "
                  f"品牌/厂家={self.col_brand_var.get()} 型号={self.col_model_var.get() or '-'} "
                  f"格式={self.fmt_var.get()}")
        self._bar(f"扫描完成，请到「列映射」步骤确认列位置")

    def _browse_output(self):
        path = filedialog.asksaveasfilename(
            defaultextension=".xlsx", filetypes=[("Excel文件","*.xlsx")])
        if path: self.output_var.set(path)

    def _run_convert(self):
        if not self.ws:
            messagebox.showerror("错误","请先选择输入文件"); return
        if not self.col_name_var.get() or not self.col_qty_var.get() or not self.col_brand_var.get():
            messagebox.showerror("错误","请确认列映射（步骤02）"); return
        if not self.project_var.get().strip():
            messagebox.showerror("错误","请填写项目名称（步骤03）"); return
        self.run_btn.configure(state="disabled", text="转换中...")
        self._chip("● 转换中...", C_YELLOW)
        self._show_step(3)
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
            use_split = (fmt=="B") or (fmt=="auto" and col_model is not None)

            self._log(f"\n{'='*50}")
            self._log(f"开始转换  {'[格式B：厂家/型号分开列]' if use_split else '[格式A：合并列]'}")
            self._log(f"列映射：名称={self.col_name_var.get()} 用量={self.col_qty_var.get()} "
                      f"品牌/厂家={self.col_brand_var.get()}" +
                      (f" 型号={self.col_model_var.get()}" if use_split else ""))

            rows=[]; seq=0; skipped=0
            for ri in range(hr+1, self.ws.max_row+1):
                nv  = self.ws.cell(row=ri,column=col_name).value
                qv  = self.ws.cell(row=ri,column=col_qty).value
                bv  = self.ws.cell(row=ri,column=col_brand).value
                mv  = self.ws.cell(row=ri,column=col_model).value if col_model else None
                if not nv and not bv: skipped+=1; continue
                sr = parse_split(bv,mv) if use_split else parse_combined(bv)
                if not sr: sr=[("","")]
                try:
                    mq=float(qv) if qv not in (None,"") else 0
                    mq=int(mq) if mq==int(mq) else mq
                except: mq=qv
                suppliers=[(b,m,mq if i==0 else 0) for i,(b,m) in enumerate(sr)]
                seq+=1; rows.append({"seq":seq,"name":str(nv).strip(),"suppliers":suppliers})

            self._log(f"解析完成：{len(rows)} 个物料（跳过空行 {skipped}）")
            total = write_review_bom(rows, out_file, project)
            abs_path = os.path.abspath(out_file)
            self._log(f"输出文件：{abs_path}")
            self._log(f"写入 {total} 行（含替代料展开）")
            self._log(f"{'='*50}\n✅  转换成功！")

            self.after(0, lambda: self._chip("● 转换成功", C_GREEN))
            self.after(0, lambda: self._bar(f"✅ 完成  {len(rows)} 个物料  {total} 行  →  {out_file}"))
            self.after(0, lambda: messagebox.showinfo("完成",f"转换成功！\n输出文件：\n{abs_path}"))
        except Exception as e:
            import traceback
            self._log(f"\n❌ 错误：{e}\n{traceback.format_exc()}")
            self.after(0, lambda: self._chip("● 转换失败", C_RED))
            self.after(0, lambda: messagebox.showerror("错误", str(e)))
        finally:
            self.after(0, lambda: self.run_btn.configure(state="normal", text="▶  开始转换"))

    def _log(self, msg):
        def _w():
            self.log_box.configure(state="normal")
            self.log_box.insert("end", msg+"\n")
            self.log_box.see("end")
            self.log_box.configure(state="disabled")
        self.after(0, _w)

    def _clear_log(self):
        self.log_box.configure(state="normal")
        self.log_box.delete("0.0","end")
        self.log_box.configure(state="disabled")

    def _chip(self, text, color):
        self.status_chip.configure(text=text, text_color=color)

    def _bar(self, text):
        self.bar_label.configure(text=text)


if __name__ == "__main__":
    app = BomApp()
    app.mainloop()
