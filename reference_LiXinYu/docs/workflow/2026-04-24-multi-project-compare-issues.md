# 多项目管理与两两对比 Issues

- Date: 2026-04-24
- Complexity: L1
- Related design: `docs/workflow/2026-04-24-multi-project-compare-design.md`

## Task Overview

- Goal: 在 Web UI 中支持会话内多个项目管理，并提供两两差异对比。
- Ordering rule: Complete issues in sequence.
- Current status: completed

## Issue List

- [x] issue-1 后端项目库与对比 API
- [x] issue-2 前端项目管理与对比展示
- [x] issue-3 测试、文档和变更记录

## issue-1

- ID: issue-1
- 标题: 后端项目库与对比 API
- 范围: `pstx_web.py`
- 依赖: none
- 验收标准: `/api/projects` 返回当前会话项目摘要；`/api/compare` 可对两个 `run_id` 返回指标、元件、网络和结果表差异。
- 状态: done
- 验证方式: Web 单测覆盖接口结构和典型新增/变化结果。
- commit: pending

## issue-2

- ID: issue-2
- 标题: 前端项目管理与对比展示
- 范围: `web/templates/index.html`, `web/templates/report.html`, `web/static/app.js`, `web/static/app.css`
- 依赖: issue-1
- 验收标准: 首页和报告页可看到项目管理入口；选择两个项目后可查看差异概览和明细表。
- 状态: done
- 验证方式: 静态资源测试检查入口、JS 函数和 CSS class，接口单测检查页面包含挂载点。
- commit: pending

## issue-3

- ID: issue-3
- 标题: 测试、文档和变更记录
- 范围: `tests/test_pstx_web.py`, `AGENTS.md`, `docs/changelists/`
- 依赖: issue-2
- 验收标准: 新增 changelist 并更新 AGENTS；相关测试通过。
- 状态: done
- 验证方式: `python -m unittest discover -s tests -p test_pstx_web.py -v`、全量测试、`python -m compileall -q .`、`node --check web/static/app.js`。
- commit: pending
