# Aster 浮窗审查增强

- Date: 2026-04-26
- Complexity: L1
- Status: final

## Background

当前 Aster 入口嵌在报告概览区，功能定位更像“摘要按钮”。用户希望把 AI 能力做成更强的审查助手：第一步转为浮窗，同时让更多已有审查要素进入 AI 检查上下文。

## Goal

- 报告页新增可展开/收起的 Aster 浮窗助手。
- Aster payload 不只包含原始分区表，还要显式提供审查覆盖范围、关键发现和人工复核边界。
- Aster 输出结构增加 checklist 和人工复核项，Web 能直接展示。
- 保持 secret/token/apiKey 不进前端回显、不进日志明文。

## Non-goals

- 不新增真实电路规则算法。
- 不让前端直接调用 Aster。
- 不上传原始 PSTX 文件。

## Solution

- `pstx_aster_client.py`
  - `build_report_brief()` 增加 `review_scope`、`key_findings`、`manual_review_boundaries`。
  - `build_aster_prompt()` 扩展输出 schema：`review_checklist` 和 `manual_review`。
  - `normalize_aster_answer()` 标准化新增字段。
- `pstx_aster_mock.py`
  - mock 摘要也返回 checklist / manual_review，便于离线 UI 验证。
- `web/templates/report.html`、`web/static/app.js`、`web/static/app.css`
  - Aster 区块改为浮窗面板和右下角 launcher。
  - 结果展示 checklist、优先级、分区焦点和人工复核边界。

## Verification Plan

- 单元测试确认 Aster brief 包含新增审查上下文，prompt 包含新 schema。
- Web 静态测试确认浮窗 DOM / JS / CSS 存在。
- 跑 Aster client、mock、Web targeted 测试和全量回归。
