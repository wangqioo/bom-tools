# CHANGELOG

所有对本仓库 `.py` 文件的修改均记录在此文件中。
格式遵循 [Keep a Changelog](https://keepachangelog.com/zh-CN/1.0.0/)，版本号遵循语义化版本。

---

## [Unreleased]

---

## [1.4.0] - 2026-04-23

### pstx_analyzer.py

**Added**
- 电容降额新增「板级最高电压快速 pass」逻辑：
  - 解析时先扫全板所有网络名，求出板级最高工作电压
  - 若 `额定电压 × 降额比 ≥ 板级最高电压`，该电容直接标记为 `✅ 板级直通`，无需逐网检查
  - 示例：板级最高 5V，额定 50V 电容 → 50×70%=35V ≥ 5V → 直通
  - 仅额定电压较低（降额后低于板级最高电压）的电容才做精确逐网推断
- 新增辅助函数 `_calc_board_max_voltage(nets, custom_volt_map) -> float`

---

## [1.3.1] - 2026-04-22

### pstx_analyzer.py

**Fixed**
- PHYS_PAGE 解析：改为仅对直接放置在顶层（路径中 SCH_1 深度 ≤ 1）的元件使用 PHYS_PAGE
  - 层次化设计中子模块内元件的 PHYS_PAGE 是子模块内页码，不是主图页码，不再使用
  - 顶层元件优先用 PHYS_PAGE（工程师印刷图上实际看到的页码），子层级回退到逻辑页号

---

## [1.3.0] - 2026-04-22

### pstx_analyzer.py

**Fixed**
- PHYS_PAGE 作为主图页码来源：从 `PHYS_PAGE` 属性提取物理页码（工程师实际使用的页码），替代 Cadence 内部逻辑页号（如 PAGE23）
- `resolve_component_pages()`：若元件已有 `page_real`（来自 PHYS_PAGE），跳过覆盖

---

## [1.2.2] - 2026-04-21

### pstx_analyzer.py

**Fixed**
- BOM_OPTION 两个子 Tab（拼写检查 + 元件清单）合并为单一 Tab，新增 `拼写风险` 列避免信息重复

---

## [1.2.1] - 2026-04-21

### pstx_analyzer.py

**Fixed**
- DRC 面板缺少「未命名网络」子 Tab：`check_drc()` 已计算、Excel 已导出，但 GUI Notebook 漏建该子 Tab

---

## [1.2.0] - 2026-04-20

### pstx_analyzer.py

**Added**（来自 dehdl_review 分支功能整合）
- AC 耦合电容识别：差分对分析，AC 耦合电容不参与降额判断
- PG/OD 信号网络识别：跳过 OD 输出网络的电压推断，避免误报
- 电阻检查 Tab：上拉 / 下拉 / 串阻检测 + 分压风险分级
- 自然排序：所有 Treeview 列排序改为自然序（C1 < C2 < C10）
- 降额比改为百分比显示，默认 70%，可在界面调整

---

## [1.1.0] - 2026-04-19

### pstx_analyzer.py

**Added**
- 文件夹自动检测：选择 worklib 工程目录后自动填入 pstxprt / pstxnet / pstxref 路径
- pstxref.dat 可选输入，提供元件描述补充
- 降额规则可折叠展示面板
- 所有 Treeview 表头支持点击排序

**Fixed**
- 合并为单文件（原 4 文件拆分导致 ModuleNotFoundError）
- `_make_tree()` 返回值修复（frame + tree 正确 pack，滚动条正常显示）

---

## [1.0.0] - 2026-04-18

### pstx_analyzer.py

**Added**（初始版本，基于 bom-tools 风格重写 pstx-schematic-analyzer）
- 单文件 tkinter 桌面应用
- 8 个功能 Tab：BOM 管理、DRC 设计检查、元件查询、网络查询、网络拓扑分析、电容降额分析、电阻检查、Excel 导出
- 解析 pstxprt.dat / pstxnet.dat / pstxref.dat（可选）
- openpyxl 多 Sheet 彩色 Excel 导出
- 多编码自动回退（utf-8-sig → utf-16 → gb18030）
- Levenshtein 编辑距离用于 BOM_OPTION 拼写风险检测
