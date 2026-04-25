# 2026-04-25 CSA 几何规范检查接入 Issues

## Issues

### Issue 1

- ID: csa-geometry-1
- 标题: 提炼 CSA 几何检查模块
- 范围: 新增 `pstx_csa_geometry.py`，覆盖 page*.csa 收集、WIRE/DOT/CIRCLE/ARC 解析、DOT 四向十字交叉检测、页级与明细结果输出
- 依赖: 无
- 验收标准: demo CSA 样例能得到 cross=2、circle=3；T 型和无 DOT 十字不误报
- 状态: completed
- 验证方式: `python3 -m unittest tests.test_pstx_analyzer.CsaGeometryTests -v`
- commit: pending

### Issue 2

- ID: csa-geometry-2
- 标题: 接入分析结果、Web 报告和 Excel 导出
- 范围: `pstx_analyzer.py`、`pstx_web.py`、测试、文档
- 依赖: csa-geometry-1
- 验收标准: Web 报告出现“规范检查”分区；Excel 含规范检查工作表；缺少 CSA 文件不影响主分析
- 状态: completed
- 验证方式: `python3 -m unittest tests.test_pstx_web.WebUiTests.test_web_report_includes_csa_geometry_review_section -v`；`python3 -m unittest tests.test_pstx_analyzer.ExportTests.test_export_to_excel_includes_resistor_analysis_sheet -v`
- commit: pending

## Self Review

- [x] CSA 几何检查和电气规则边界清晰
- [x] 缺失 CSA 文件不阻断现有项目分析
- [x] Web / Excel 输出字段可直接给工程 review 使用
- [x] 新增测试覆盖参考代码中的关键正反样例
