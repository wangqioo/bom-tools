# -*- coding: utf-8 -*-
"""
PSTX 原理图分析工具 v1.0
支持 Cadence Packager-XL 导出的 pstxprt.dat / pstxnet.dat

功能：BOM 管理 / 网络拓扑 / DRC / 电容降额 / 元件查询 / Excel 导出

依赖：pip install openpyxl
运行：python pstx_analyzer.py
"""

import sys
import subprocess

# 自动安装依赖
try:
    import openpyxl
except ImportError:
    print("未检测到 openpyxl，正在自动安装...")
    subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'openpyxl'])
    import openpyxl

import os
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext

# 同目录下的模块
_HERE = os.path.dirname(os.path.abspath(__file__))
if _HERE not in sys.path:
    sys.path.insert(0, _HERE)

import parser as _parser
import analysis as _analysis
import exporter as _exporter


# ═══════════════════════════════════════════════════════════
# 辅助：带滚动条的 Treeview
# ═══════════════════════════════════════════════════════════

def make_tree(parent, columns, height=12):
    """创建带双向滚动条的 Treeview，返回 (frame, tree)"""
    frame = tk.Frame(parent)
    tree = ttk.Treeview(frame, columns=columns, show='headings', height=height)
    vsb = ttk.Scrollbar(frame, orient='vertical',   command=tree.yview)
    hsb = ttk.Scrollbar(frame, orient='horizontal',  command=tree.xview)
    tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

    tree.grid(row=0, column=0, sticky='nsew')
    vsb.grid(row=0, column=1, sticky='ns')
    hsb.grid(row=1, column=0, sticky='ew')
    frame.grid_rowconfigure(0, weight=1)
    frame.grid_columnconfigure(0, weight=1)
    return frame, tree


def fill_tree(tree, rows: list, columns: list = None):
    """清空并填充 Treeview"""
    tree.delete(*tree.get_children())
    if not rows:
        return
    cols = columns or list(rows[0].keys())
    tree['columns'] = cols
    for c in cols:
        tree.heading(c, text=c, anchor='w')
        w = max(len(c) * 9, 80)
        tree.column(c, width=min(w, 220), anchor='w', stretch=True)
    for row in rows:
        tree.insert('', 'end', values=[str(row.get(c, '')) for c in cols])


# ═══════════════════════════════════════════════════════════
# 主应用
# ═══════════════════════════════════════════════════════════

class PstxApp(tk.Tk):

    def __init__(self):
        super().__init__()
        self.title('PSTX 原理图分析工具 v1.0')
        self.geometry('1040x720')
        self.minsize(900, 600)
        self.resizable(True, True)

        # 数据
        self._components = {}
        self._nets       = {}
        self._bom_detail_normal  = []
        self._bom_detail_depop   = []
        self._bom_merged_normal  = []
        self._bom_merged_depop   = []
        self._net_analysis       = {}
        self._drc                = {}
        self._derating           = []

        # StringVar
        self.prt_path   = tk.StringVar()
        self.net_path   = tk.StringVar()
        self.out_path   = tk.StringVar(value='pstx_分析报告.xlsx')
        self.project_var = tk.StringVar()
        self.bom_filter  = tk.StringVar(value='贴装')
        self.bom_search  = tk.StringVar()
        self.query_mode  = tk.StringVar(value='位号')
        self.query_text  = tk.StringVar()
        self.derating_ratio = tk.DoubleVar(value=2.0)
        self.custom_volt_text = tk.StringVar()

        self.bom_filter.trace_add('write', lambda *_: self._refresh_bom())
        self.bom_search.trace_add('write', lambda *_: self._refresh_bom())

        self._build_ui()

    # ─────────────────────────────────────────────────────
    # UI 构建
    # ─────────────────────────────────────────────────────

    def _section(self, parent, title):
        f = ttk.LabelFrame(parent, text=title, padding=8)
        f.pack(fill='x', padx=10, pady=4)
        return f

    def _build_ui(self):
        nb = ttk.Notebook(self)
        nb.pack(fill='both', expand=True, padx=8, pady=6)
        self.nb = nb

        tabs = [
            ('  文件加载  ', self._build_tab_load),
            ('  BOM 管理  ', self._build_tab_bom),
            ('  网络分析  ', self._build_tab_net),
            ('  设计检查  ', self._build_tab_drc),
            ('  电容降额  ', self._build_tab_derating),
            ('  元件查询  ', self._build_tab_query),
            ('  日志      ', self._build_tab_log),
        ]
        for text, builder in tabs:
            frame = ttk.Frame(nb)
            nb.add(frame, text=text)
            builder(frame)

    # ── Tab 1：文件加载 ───────────────────────────────────

    def _build_tab_load(self, p):
        f1 = self._section(p, 'pstxprt.dat（元件属性）')
        tk.Label(f1, text='文件路径：').grid(row=0, column=0, sticky='w')
        ttk.Entry(f1, textvariable=self.prt_path, width=60).grid(row=0, column=1, padx=6)
        ttk.Button(f1, text='浏览…',
                   command=lambda: self._browse_dat(self.prt_path)).grid(row=0, column=2)

        f2 = self._section(p, 'pstxnet.dat（网络连接）')
        tk.Label(f2, text='文件路径：').grid(row=0, column=0, sticky='w')
        ttk.Entry(f2, textvariable=self.net_path, width=60).grid(row=0, column=1, padx=6)
        ttk.Button(f2, text='浏览…',
                   command=lambda: self._browse_dat(self.net_path)).grid(row=0, column=2)

        f3 = self._section(p, '项目名称（导出用）')
        tk.Label(f3, text='项目名称：').grid(row=0, column=0, sticky='w')
        ttk.Entry(f3, textvariable=self.project_var, width=40).grid(row=0, column=1, padx=6, sticky='w')

        # 解析按钮 + 状态
        btn_row = tk.Frame(p)
        btn_row.pack(pady=14)
        self.parse_btn = tk.Button(
            btn_row, text='开始解析', font=('Arial', 13, 'bold'),
            bg='#2d6cdf', fg='white', relief='flat',
            padx=24, pady=10, command=self._run_parse)
        self.parse_btn.pack(side='left', padx=8)
        self.load_status = tk.Label(btn_row, text='', font=('Arial', 11))
        self.load_status.pack(side='left', padx=8)

        # 概览信息框
        f4 = self._section(p, '解析概览')
        self.overview_text = tk.Text(f4, height=8, font=('Consolas', 9),
                                      state='disabled', bg='#f5f5f5', relief='flat')
        self.overview_text.pack(fill='x')

    # ── Tab 2：BOM 管理 ───────────────────────────────────

    def _build_tab_bom(self, p):
        ctrl = self._section(p, '筛选 / 搜索')
        for val, txt in [('贴装', '贴装元件'), ('DEPOP', 'DEPOP'), ('全部', '全部')]:
            ttk.Radiobutton(ctrl, text=txt, variable=self.bom_filter,
                            value=val).pack(side='left', padx=10)
        ttk.Separator(ctrl, orient='vertical').pack(side='left', fill='y', padx=6)
        tk.Label(ctrl, text='搜索：').pack(side='left')
        ttk.Entry(ctrl, textvariable=self.bom_search, width=26).pack(side='left', padx=4)

        # 表格
        tree_frame = tk.Frame(p)
        tree_frame.pack(fill='both', expand=True, padx=10, pady=4)
        cols = ['位号', '料号', '值', '封装', '类型', '页面']
        _, self.bom_tree = make_tree(tree_frame, cols, height=16)
        self.bom_tree.pack(fill='both', expand=True)

        # 底部工具栏
        bot = tk.Frame(p)
        bot.pack(fill='x', padx=10, pady=4)
        self.bom_count_lbl = tk.Label(bot, text='', fg='#444')
        self.bom_count_lbl.pack(side='left')
        ttk.Button(bot, text='导出 Excel', command=self._export_excel).pack(side='right')

    # ── Tab 3：网络分析 ───────────────────────────────────

    def _build_tab_net(self, p):
        stat_frame = self._section(p, '汇总')
        self.net_stat_text = tk.Text(stat_frame, height=4, font=('Consolas', 9),
                                      state='disabled', bg='#f5f5f5', relief='flat')
        self.net_stat_text.pack(fill='x')

        sub = ttk.Notebook(p)
        sub.pack(fill='both', expand=True, padx=10, pady=4)
        self.net_sub = sub

        net_tabs = [
            ('电源网络',       ['网络名', '节点数'],           '_tree_power'),
            ('GND 网络',       ['网络名', '节点数'],           '_tree_gnd'),
            ('差分对',         ['基础名', 'P端网络', 'N端网络'], '_tree_diff'),
            ('单端网络',       ['网络名', '连接元件', '引脚'],  '_tree_single'),
            ('各页面元件数',   ['页面', '元件数'],             '_tree_pages'),
        ]
        for title, cols, attr in net_tabs:
            f = ttk.Frame(sub)
            sub.add(f, text=f'  {title}  ')
            _, tree = make_tree(f, cols, height=14)
            tree.pack(fill='both', expand=True)
            setattr(self, attr, tree)

    # ── Tab 4：设计检查 ───────────────────────────────────

    def _build_tab_drc(self, p):
        sub = ttk.Notebook(p)
        sub.pack(fill='both', expand=True, padx=10, pady=8)

        drc_tabs = [
            ('缺料号',       ['位号', '类型', '页面'],                                         '_tree_drc_hq'),
            ('缺 VALUE',     ['位号', '类型', '页面'],                                         '_tree_drc_val'),
            ('缺封装',       ['位号', '类型', '页面'],                                         '_tree_drc_pkg'),
            ('TBD 属性',     ['位号', '属性', '当前值', '类型', '页面'],                        '_tree_drc_tbd'),
            ('单端网络',     ['网络名', '连接元件', '引脚', '页面'],                            '_tree_drc_single'),
            ('BOM_OPTION',   ['实际填写值', '疑似应为', '编辑距离', '使用该值的位号', '风险'], '_tree_drc_opt'),
        ]
        for title, cols, attr in drc_tabs:
            f = ttk.Frame(sub)
            sub.add(f, text=f'  {title}  ')
            _, tree = make_tree(f, cols, height=15)
            tree.pack(fill='both', expand=True)
            setattr(self, attr, tree)

    # ── Tab 5：电容降额 ───────────────────────────────────

    def _build_tab_derating(self, p):
        ctrl = self._section(p, '参数设置')
        tk.Label(ctrl, text=f'降额系数（额定/工作 ≥ X）：').grid(row=0, column=0, sticky='w')
        self.ratio_scale = ttk.Scale(ctrl, from_=1.0, to=5.0, orient='horizontal',
                                     variable=self.derating_ratio, length=200)
        self.ratio_scale.grid(row=0, column=1, padx=8)
        self.ratio_lbl = tk.Label(ctrl, text='2.0', width=5, font=('Arial', 11, 'bold'))
        self.ratio_lbl.grid(row=0, column=2)
        self.derating_ratio.trace_add('write', self._on_ratio_change)

        tk.Label(ctrl, text='自定义电压映射\n（每行 NET前缀=电压V）：',
                 justify='left').grid(row=1, column=0, sticky='nw', pady=6)
        self.volt_map_entry = tk.Text(ctrl, height=3, width=40, font=('Consolas', 9))
        self.volt_map_entry.grid(row=1, column=1, columnspan=2, padx=8, sticky='w')
        self.volt_map_entry.insert('1.0', '# 示例：VBUS=5.0\n# P12V_AUX=12.0')

        ttk.Button(ctrl, text='重新计算',
                   command=self._recalc_derating).grid(row=2, column=1, sticky='w', pady=4)

        cols = ['位号', '值', '封装', '类型', '额定电压', '推断工作电压(V)',
                '推断来源网络', '降额比', '状态', '页面', 'DEPOP']
        tree_f = tk.Frame(p)
        tree_f.pack(fill='both', expand=True, padx=10, pady=4)
        _, self.derating_tree = make_tree(tree_f, cols, height=14)
        self.derating_tree.pack(fill='both', expand=True)

        self.derating_stat_lbl = tk.Label(p, text='', fg='#555')
        self.derating_stat_lbl.pack(anchor='w', padx=10, pady=2)

    # ── Tab 6：元件查询 ───────────────────────────────────

    def _build_tab_query(self, p):
        ctrl = self._section(p, '查询')
        for val, txt in [('位号', '按位号 (refdes)'), ('网络名', '按网络名')]:
            ttk.Radiobutton(ctrl, text=txt, variable=self.query_mode,
                            value=val).pack(side='left', padx=10)
        ttk.Separator(ctrl, orient='vertical').pack(side='left', fill='y', padx=6)
        ttk.Entry(ctrl, textvariable=self.query_text, width=30).pack(side='left', padx=4)
        ttk.Button(ctrl, text='查询', command=self._do_query).pack(side='left', padx=4)

        f2 = self._section(p, '查询结果')
        self.query_result = scrolledtext.ScrolledText(
            f2, font=('Consolas', 10), state='disabled',
            bg='#1e1e1e', fg='#d4d4d4', relief='flat', height=22)
        self.query_result.pack(fill='both', expand=True)

    # ── Tab 7：日志 ───────────────────────────────────────

    def _build_tab_log(self, p):
        self.log = scrolledtext.ScrolledText(
            p, font=('Consolas', 9), state='disabled',
            bg='#1e1e1e', fg='#d4d4d4', relief='flat')
        self.log.pack(fill='both', expand=True, padx=8, pady=8)
        ttk.Button(p, text='清空日志',
                   command=self._clear_log).pack(anchor='e', padx=8, pady=4)

    # ─────────────────────────────────────────────────────
    # 事件
    # ─────────────────────────────────────────────────────

    def _browse_dat(self, var: tk.StringVar):
        path = filedialog.askopenfilename(
            title='选择 .dat 文件',
            filetypes=[('DAT 文件', '*.dat'), ('所有文件', '*.*')])
        if path:
            var.set(path)
            self._log(f'选择文件：{path}')

    def _on_ratio_change(self, *_):
        v = self.derating_ratio.get()
        self.ratio_lbl.configure(text=f'{v:.1f}')

    # ─────────────────────────────────────────────────────
    # 解析流程
    # ─────────────────────────────────────────────────────

    def _run_parse(self):
        prt = self.prt_path.get().strip()
        net = self.net_path.get().strip()
        if not prt or not net:
            messagebox.showerror('错误', '请先选择 pstxprt.dat 和 pstxnet.dat')
            return
        self.parse_btn.configure(state='disabled')
        self._start_spinner('解析中')
        threading.Thread(target=self._do_parse, args=(prt, net), daemon=True).start()

    def _do_parse(self, prt_path: str, net_path: str):
        try:
            self._log('\n── 开始解析 ──────────────────')

            with open(prt_path, encoding='utf-8', errors='replace') as f:
                prt_content = f.read()
            with open(net_path, encoding='utf-8', errors='replace') as f:
                net_content = f.read()

            self._log(f'  pstxprt.dat：{len(prt_content):,} 字节')
            self._log(f'  pstxnet.dat：{len(net_content):,} 字节')

            components, nets, _ = _parser.parse_all(prt_content, net_content)
            self._log(f'  解析元件：{len(components)} 个')
            self._log(f'  解析网络：{len(nets)} 个')

            # 运行所有分析
            dn, dd, mn, md = _analysis.build_bom(components)
            na   = _analysis.analyze_networks(nets, components)
            drc  = _analysis.check_drc(components, nets)
            drt  = _analysis.analyze_derating(components, nets, self.derating_ratio.get(),
                                               self._parse_volt_map())

            # 存储结果
            self._components        = components
            self._nets              = nets
            self._bom_detail_normal = dn
            self._bom_detail_depop  = dd
            self._bom_merged_normal = mn
            self._bom_merged_depop  = md
            self._net_analysis      = na
            self._drc               = drc
            self._derating          = drt

            self._log('  分析完成')
            self._log(f'  贴装 {len(mn)} 种 / {sum(r.get("数量",0) for r in mn)} 个')
            self._log(f'  DEPOP {len(md)} 种')
            self._log(f'  DRC 问题：{sum(len(v) for v in drc.values() if isinstance(v,list))} 项')

            self.after(0, self._on_parse_done)

        except Exception as e:
            import traceback
            msg = f'解析失败：{e}\n{traceback.format_exc()}'
            self._log(msg)
            self.after(0, lambda: self._stop_spinner('❌ 解析失败'))
            self.after(0, lambda: messagebox.showerror('错误', str(e)))
        finally:
            self.after(0, lambda: self.parse_btn.configure(state='normal'))

    def _on_parse_done(self):
        self._stop_spinner('✅ 解析完成')
        self._refresh_all()
        self.nb.select(1)   # 跳转到 BOM 页

    def _refresh_all(self):
        self._update_overview()
        self._refresh_bom()
        self._refresh_net()
        self._refresh_drc()
        self._refresh_derating()

    # ─────────────────────────────────────────────────────
    # 概览文本
    # ─────────────────────────────────────────────────────

    def _update_overview(self):
        na  = self._net_analysis
        drc = self._drc
        mn  = self._bom_merged_normal
        md  = self._bom_merged_depop
        drt = self._derating

        drc_total = sum(len(v) for v in drc.values() if isinstance(v, list))
        fail = sum(1 for r in drt if r.get('状态', '').startswith('❌'))

        lines = [
            f'贴装元件：{len(mn)} 种 / {sum(r.get("数量",0) for r in mn)} 个',
            f'DEPOP：   {len(md)} 种 / {sum(r.get("数量",0) for r in md)} 个',
            f'网络总数：{na.get("total", 0)}   电源网络：{len(na.get("power_nets",{}))}   GND：{len(na.get("gnd_nets",{}))}   差分对：{len(na.get("diff_pairs",{}))}',
            f'单端网络（疑似漏连）：{len(na.get("single_node", {}))}',
            f'DRC 问题总数：{drc_total}',
            f'电容降额不合格：{fail}',
        ]
        self.overview_text.configure(state='normal')
        self.overview_text.delete('1.0', 'end')
        self.overview_text.insert('end', '\n'.join(lines))
        self.overview_text.configure(state='disabled')

    # ─────────────────────────────────────────────────────
    # BOM 刷新
    # ─────────────────────────────────────────────────────

    def _refresh_bom(self):
        mode   = self.bom_filter.get()
        kw     = self.bom_search.get().strip().lower()

        if mode == '贴装':
            source = self._bom_detail_normal
        elif mode == 'DEPOP':
            source = self._bom_detail_depop
        else:
            source = self._bom_detail_normal + self._bom_detail_depop

        if kw:
            source = [r for r in source
                      if any(kw in str(v).lower() for v in r.values())]

        cols = ['位号', '料号', '值', '封装', '类型', '页面']
        fill_tree(self.bom_tree, source, cols)
        self.bom_count_lbl.configure(text=f'共 {len(source)} 行')

    # ─────────────────────────────────────────────────────
    # 网络分析刷新
    # ─────────────────────────────────────────────────────

    def _refresh_net(self):
        na = self._net_analysis

        # 文本统计
        lines = [
            f'网络总数：{na.get("total", 0)}',
            f'电源网络：{len(na.get("power_nets", {}))}   GND：{len(na.get("gnd_nets", {}))}',
            f'差分对：{len(na.get("diff_pairs", {}))}   单端网络：{len(na.get("single_node", {}))}',
        ]
        self.net_stat_text.configure(state='normal')
        self.net_stat_text.delete('1.0', 'end')
        self.net_stat_text.insert('end', '    '.join(lines))
        self.net_stat_text.configure(state='disabled')

        # 电源网络
        power = [{'网络名': k, '节点数': len(v)}
                 for k, v in sorted(na.get('power_nets', {}).items(), key=lambda x: -len(x[1]))]
        fill_tree(self._tree_power, power, ['网络名', '节点数'])

        # GND
        gnd = [{'网络名': k, '节点数': len(v)}
               for k, v in sorted(na.get('gnd_nets', {}).items(), key=lambda x: -len(x[1]))]
        fill_tree(self._tree_gnd, gnd, ['网络名', '节点数'])

        # 差分对
        diff = [{'基础名': b, 'P端网络': pr['P'], 'N端网络': pr['N']}
                for b, pr in sorted(na.get('diff_pairs', {}).items())]
        fill_tree(self._tree_diff, diff, ['基础名', 'P端网络', 'N端网络'])

        # 单端网络
        single = [{'网络名': k, '连接元件': v[0]['refdes'], '引脚': v[0]['pin_name']}
                  for k, v in sorted(na.get('single_node', {}).items())]
        fill_tree(self._tree_single, single, ['网络名', '连接元件', '引脚'])

        # 各页面
        pages = [{'页面': pg, '元件数': cnt}
                 for pg, cnt in sorted(na.get('page_counter', {}).items())]
        fill_tree(self._tree_pages, pages, ['页面', '元件数'])

    # ─────────────────────────────────────────────────────
    # DRC 刷新
    # ─────────────────────────────────────────────────────

    def _refresh_drc(self):
        drc = self._drc
        fill_tree(self._tree_drc_hq,     drc.get('missing_hq_code', []),  ['位号', '类型', '页面'])
        fill_tree(self._tree_drc_val,    drc.get('missing_value', []),    ['位号', '类型', '页面'])
        fill_tree(self._tree_drc_pkg,    drc.get('missing_package', []),  ['位号', '类型', '页面'])
        fill_tree(self._tree_drc_tbd,    drc.get('tbd_attrs', []),        ['位号', '属性', '当前值', '类型', '页面'])
        fill_tree(self._tree_drc_single, drc.get('single_pin_nets', []),  ['网络名', '连接元件', '引脚', '页面'])
        fill_tree(self._tree_drc_opt,    drc.get('bom_option_typos', []),
                  ['实际填写值', '疑似应为', '编辑距离', '使用该值的位号', '风险'])

    # ─────────────────────────────────────────────────────
    # 降额刷新
    # ─────────────────────────────────────────────────────

    def _refresh_derating(self):
        cols = ['位号', '值', '封装', '类型', '额定电压', '推断工作电压(V)',
                '推断来源网络', '降额比', '状态', '页面', 'DEPOP']
        fill_tree(self.derating_tree, self._derating, cols)

        total  = len(self._derating)
        fail   = sum(1 for r in self._derating if r.get('状态', '').startswith('❌'))
        ok     = sum(1 for r in self._derating if r.get('状态', '').startswith('✅'))
        unk    = total - fail - ok
        self.derating_stat_lbl.configure(
            text=f'共 {total} 个电容  |  ✅ 合格 {ok}  |  ❌ 不合格 {fail}  |  ⚪ 无法判断 {unk}')

    def _recalc_derating(self):
        if not self._components:
            messagebox.showwarning('提示', '请先解析文件')
            return
        self._derating = _analysis.analyze_derating(
            self._components, self._nets,
            self.derating_ratio.get(),
            self._parse_volt_map())
        self._refresh_derating()
        self._log(f'降额重新计算完成（系数={self.derating_ratio.get():.1f}）')

    def _parse_volt_map(self):
        """从文本框解析自定义电压映射"""
        result = {}
        text = self.volt_map_entry.get('1.0', 'end')
        for line in text.splitlines():
            line = line.strip()
            if not line or line.startswith('#'):
                continue
            if '=' in line:
                k, _, v = line.partition('=')
                try:
                    result[k.strip()] = float(v.strip())
                except ValueError:
                    pass
        return result or None

    # ─────────────────────────────────────────────────────
    # 元件查询
    # ─────────────────────────────────────────────────────

    def _do_query(self):
        kw   = self.query_text.get().strip()
        mode = self.query_mode.get()
        if not kw:
            return
        if not self._components:
            messagebox.showwarning('提示', '请先解析文件')
            return

        lines = []
        if mode == '位号':
            comp = self._components.get(kw)
            if comp:
                lines.append(f'═══ 元件：{kw} ═══')
                for k, v in comp.items():
                    if k == 'nets':
                        continue
                    lines.append(f'  {k:<16} {v}')
                lines.append('')
                lines.append('  引脚 → 网络：')
                for pin, net in sorted(comp.get('nets', {}).items()):
                    lines.append(f'    pin {pin:<6} → {net}')
            else:
                # 模糊匹配
                matched = [r for r in self._components if kw.upper() in r.upper()]
                if matched:
                    lines.append(f'未找到精确匹配，模糊结果：')
                    lines.extend(f'  {r}' for r in sorted(matched)[:50])
                else:
                    lines.append(f'未找到位号：{kw}')

        else:  # 按网络名
            # 精确匹配
            nodes = self._nets.get(kw)
            if nodes:
                lines.append(f'═══ 网络：{kw}（{len(nodes)} 个节点）═══')
                for n in nodes:
                    comp = self._components.get(n['refdes'], {})
                    desc = comp.get('value', '') or comp.get('part_name', '')
                    lines.append(f'  {n["refdes"]:<10} pin {n["pin"]:<6} ({n["pin_name"]:<12}) {desc}')
            else:
                # 模糊匹配
                matched = [k for k in self._nets if kw.upper() in k.upper()]
                if matched:
                    lines.append(f'未找到精确匹配，模糊结果（前50）：')
                    for nm in sorted(matched)[:50]:
                        lines.append(f'  {nm}  ({len(self._nets[nm])} nodes)')
                else:
                    lines.append(f'未找到网络：{kw}')

        self.query_result.configure(state='normal')
        self.query_result.delete('1.0', 'end')
        self.query_result.insert('end', '\n'.join(lines))
        self.query_result.configure(state='disabled')

    # ─────────────────────────────────────────────────────
    # Excel 导出
    # ─────────────────────────────────────────────────────

    def _export_excel(self):
        if not self._components:
            messagebox.showwarning('提示', '请先解析文件')
            return
        out = filedialog.asksaveasfilename(
            title='保存分析报告',
            initialfile=self.out_path.get(),
            defaultextension='.xlsx',
            filetypes=[('Excel 文件', '*.xlsx')])
        if not out:
            return
        self.out_path.set(out)

        self._log('\n导出 Excel 中…')
        threading.Thread(target=self._do_export, args=(out,), daemon=True).start()

    def _do_export(self, path: str):
        try:
            data = {
                'project_name':       self.project_var.get().strip() or '未命名项目',
                'bom_normal_detail':  self._bom_detail_normal,
                'bom_depop_detail':   self._bom_detail_depop,
                'bom_normal_merged':  self._bom_merged_normal,
                'bom_depop_merged':   self._bom_merged_depop,
                'net_analysis':       self._net_analysis,
                'drc':                self._drc,
                'derating':           self._derating,
                'components':         self._components,
            }
            actual = _exporter.export_to_excel(data, path)
            self._log(f'Excel 已保存：{actual}')
            self.after(0, lambda: messagebox.showinfo('完成', f'导出成功！\n{actual}'))
            # 打开所在文件夹
            folder = os.path.dirname(os.path.abspath(actual))
            try:
                if sys.platform == 'win32':
                    os.startfile(folder)
                elif sys.platform == 'darwin':
                    subprocess.Popen(['open', folder])
                else:
                    subprocess.Popen(['xdg-open', folder])
            except Exception:
                pass
        except Exception as e:
            import traceback
            self._log(f'导出失败：{e}\n{traceback.format_exc()}')
            self.after(0, lambda: messagebox.showerror('错误', str(e)))

    # ─────────────────────────────────────────────────────
    # Spinner + 日志
    # ─────────────────────────────────────────────────────

    def _start_spinner(self, label='处理中'):
        self._spinning   = True
        self._spin_step  = 0
        self._spin_label = label
        self._spin()

    def _spin(self):
        if not self._spinning:
            return
        frames = ['◐', '◓', '◑', '◒']
        f = frames[self._spin_step % 4]
        self.load_status.configure(
            text=f'{f} {self._spin_label}，请稍候…', fg='#2d6cdf')
        self._spin_step += 1
        self._spin_job = self.after(200, self._spin)

    def _stop_spinner(self, msg=''):
        self._spinning = False
        if hasattr(self, '_spin_job'):
            self.after_cancel(self._spin_job)
        color = '#2a8a2a' if msg.startswith('✅') else ('red' if '❌' in msg else '#333')
        self.load_status.configure(text=msg, fg=color)

    def _log(self, msg: str):
        def _w():
            self.log.configure(state='normal')
            self.log.insert('end', msg + '\n')
            self.log.see('end')
            self.log.configure(state='disabled')
        self.after(0, _w)

    def _clear_log(self):
        self.log.configure(state='normal')
        self.log.delete('1.0', 'end')
        self.log.configure(state='disabled')


# ═══════════════════════════════════════════════════════════
if __name__ == '__main__':
    app = PstxApp()
    app.mainloop()
