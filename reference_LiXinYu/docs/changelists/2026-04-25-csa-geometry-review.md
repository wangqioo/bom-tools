# CSA 几何规范检查接入

## 背景

参考包 `dehdl_csa_review_branch_package.zip` 提供了面向 Cadence DE HDL `page*.csa` 的几何规范检查能力。本轮将其中可验证的核心规则提炼到当前 PSTX 工具中，并保持几何检查与电气规则边界分离。

## 变更

- 新增 `pstx_csa_geometry.py`，扫描 `<project_root>/sch_1/page*.csa`。
- 支持检测带 `DOT` 的四向十字交叉，T 型连接和无 DOT 十字不报。
- 支持输出 `CIRCLE` 画圈对象，并默认把 `ARC` 三点拟合结果作为需人工确认的画圈候选。
- `analyze_project_contents()` 新增 `csa_geometry` 结果。
- Web 报告新增“规范检查”分区，包含 CSA 页级汇总、DOT 四向十字交叉、画圈对象三张表。
- Excel 导出新增“规范检查”工作表。

## 规则边界

- CSA 检查只基于几何对象和原始行号，不推断真实网络短接。
- 画圈对象只作为 review 线索，不自动判断被圈内容是否违规。
- 缺少 `sch_1/page*.csa` 时不阻断主分析。

## 验证

- `python3 -m unittest tests.test_pstx_analyzer.CsaGeometryTests -v`
- `python3 -m unittest tests.test_pstx_web.WebUiTests.test_web_report_includes_csa_geometry_review_section -v`
- `python3 -m unittest discover -s tests -v`
