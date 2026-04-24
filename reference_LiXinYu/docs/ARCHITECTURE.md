# 架构说明

## 模块边界

- `pstx_analyzer.py`
  - PSTX 文本解析
  - BOM / 网络 / DRC / 降额 / 电阻规则分析
  - Excel 导出
  - 只保留少量页码兼容包装，具体页码决策不再散落在这个文件里
- `pstx_page_logic.py`
  - `SECTION_NUMBER` / `C_PATH` / `P_PATH` 路径解析
  - `page.map` 与 `page*.csv` 的逻辑页 / 真实页索引
  - `module_order(.dat)` 的子模块映射主模块真实页计算
- `pstx_web.py`
  - Flask Web UI
  - 项目根路径读取
  - 报告、查询、导出接口
- `pstx_local_ui.py`
  - 本地桌面壳
  - 复用 `pstx_web.create_app()`，不再维护第二套业务逻辑

## 页码模型

组件默认显示页优先级：

1. `P_PATH` 顶层 `SCH_1` 页
2. `page.map` 映射出的真实页
3. `page*.csv` 映射出的真实页

逻辑页来源优先级：

1. `SECTION_NUMBER 1` 路径
2. `C_PATH`
3. `DRAWING`

`module_order(.dat)` 规则：

- key 优先按 `SECTION_NUMBER / C_PATH` 的逻辑路径构造
- `P_PATH` 只作为保守回退
- 子模块偏移页优先使用 `P_PATH` 中子模块本地真实页
- 映射公式为：

```text
子模块映射主模块真实页 = start_real_page + 子模块本地真实页 - 1
```

## 降额补丁规则

- `analyze_derating()` 先尝试扫描整板可识别的最大正电压
- 当最大已识别正电压 `<=12V` 且电容额定耐压 `>=50V` 时，直接放行为合格
- 若无法识别整板最大正电压，仍回到单颗电容原有推断逻辑

## 测试入口

- `tests/test_pstx_analyzer.py`
  - 主分析链、页码模型、电阻与降额规则
- `tests/test_pstx_web.py`
  - Web UI 后端流程
- `tests/test_pstx_local_ui.py`
  - 本地桌面壳入口
