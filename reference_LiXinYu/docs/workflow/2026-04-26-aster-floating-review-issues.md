# Aster 浮窗审查增强 Issues

- Date: 2026-04-26
- Complexity: L1
- Related design: `docs/workflow/2026-04-26-aster-floating-review-design.md`

## Task Overview

- Goal: 将 Aster 从内嵌摘要升级为浮窗审查助手，并扩展 AI 审查上下文。
- Ordering rule: Complete issues in sequence.
- Current status: complete

## Issue List

- [x] issue-1 增强 Aster 审查上下文与输出 schema
- [x] issue-2 将 Aster 报告页入口改为浮窗助手

## issue-1

- ID: issue-1
- 标题: 增强 Aster 审查上下文与输出 schema
- 范围: `pstx_aster_client.py`、`pstx_aster_mock.py`、Aster tests
- 依赖: none
- 验收标准: Aster brief 显式包含 review_scope/key_findings/manual_review_boundaries；返回 payload 支持 review_checklist/manual_review
- 状态: done
- 验证方式: `python3 -m unittest tests.test_pstx_aster_client tests.test_pstx_aster_mock -v`
- commit: `feat(issue-1): add aster floating review assistant`

## issue-2

- ID: issue-2
- 标题: 将 Aster 报告页入口改为浮窗助手
- 范围: `web/templates/report.html`、`web/static/app.js`、`web/static/app.css`、Web tests、docs/changelist/AGENTS
- 依赖: issue-1
- 验收标准: 报告页有可展开/收起浮窗；结果展示 checklist 和人工复核项；移动端可用
- 状态: done
- 验证方式: targeted Web tests 与全量测试
- commit: `feat(issue-1): add aster floating review assistant`

## Self Review

- [x] Requirement alignment: AI 入口已浮窗化，更多审查要素进入 AI 上下文
- [x] Regression risk: 现有 Aster mock/live/off 调用保持兼容
- [x] Test coverage: brief/schema/UI 静态和全量测试均覆盖
- [x] Dirty changes: unrelated modifications were avoided
- [x] Docs sync: workflow、changelist、AGENTS 已同步
