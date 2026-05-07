# 2026-04-25 CSA 几何规范检查接入设计

## 背景

用户提供 `dehdl_csa_review_branch_package.zip`，其中参考代码面向 Cadence DE HDL `page*.csa` 文本页文件，检查两类原理图规范性问题：

- 带 `DOT` 的四向十字交叉；
- `CIRCLE` / 可选 `ARC` 画圈标注对象。

当前 PSTX 工具主链路已经围绕 `pstxprt.dat` / `pstxnet.dat` / `sch_1/page*.csv` / `module_order` 完成元件、网络、页码、电阻、电容降额和 Web 报告展示。CSA 检查属于几何画布层能力，不应混入电气连通性规则。

## 目标

- 新增独立 CSA 几何规范检查模块，扫描 `<project_root>/sch_1/page*.csa`。
- 在分析结果中输出页级汇总、DOT 四向十字交叉明细、画圈对象明细。
- Web 报告新增“规范检查”分区，Excel 导出新增对应工作表。
- 保持检查结论边界：几何规范候选，不推断网络短接或真实电气错误。

## 非目标

- 不解析 `.csb` 二进制文件。
- 不把画圈对象附近内容和网络/元件做空间关联。
- 不把 dotless cross 或 T 型连接上升为问题。
- 不修改既有 PSTX 解析、页码映射、电阻/降额算法。

## 数据模型

`csa_geometry` 返回结构：

```text
{
  enabled: bool,
  root: str,
  page_count: int,
  cross_count: int,
  circle_count: int,
  error_count: int,
  summary_rows: list[dict],
  dot_cross_rows: list[dict],
  circle_rows: list[dict],
  warnings: list[str],
}
```

## 接入点

- `pstx_csa_geometry.py`
  - 独立解析与检查逻辑。
- `pstx_analyzer.analyze_project_contents()`
  - 在已有项目根路径可用时调用 CSA 检查。
- `pstx_web._build_report_payload()`
  - 新增指标和 `csa` 分区。
- `pstx_analyzer.export_to_excel()`
  - 新增 `规范检查` 工作表。

## 风险控制

- CSA 缺失时静默返回空结果，不阻断主流程。
- CSA 单页解析异常进入 `warnings` 和页级 `错误` 字段，不阻断其它页面。
- `ARC` 默认作为候选解析，说明需人工确认。
