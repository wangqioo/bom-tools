# Aster 前端临时凭据覆盖 Issues

- Date: 2026-04-26
- Complexity: L2
- Related design: `docs/workflow/2026-04-26-aster-runtime-credentials-design.md`

## Task Overview

- Goal: 允许前端临时设置 Aster 凭据，同时避免明文回显、磁盘保存和 localStorage 保存。
- Ordering rule: Complete issues in sequence.
- Current status: completed

## Issue List

- [x] issue-1 增加后端 runtime credential override
- [x] issue-2 接入前端表单、文档和验证

## issue-1

- ID: issue-1
- 标题: 增加后端 runtime credential override
- 范围: `pstx_aster_service.py`、`pstx_web.py`、service/Web 单元测试
- 依赖: none
- 验收标准: POST 可设置当前进程覆盖项；DELETE 可清除；status 不回显 secret
- 状态: done
- 验证方式: `python3 -m unittest tests.test_pstx_aster_service tests.test_pstx_web.WebUiTests.test_aster_runtime_config_can_set_and_clear_without_echoing_secret -v`
- commit: c4851b4

## issue-2

- ID: issue-2
- 标题: 接入前端表单、文档和验证
- 范围: `web/templates/report.html`、`web/static/app.js`、`web/static/app.css`、README/ARCHITECTURE/AGENTS/changelist
- 依赖: issue-1
- 验收标准: 报告页可输入临时凭据；提交后清空密码框并刷新状态；文档明确风险和作用域
- 状态: done
- 验证方式: targeted Web 测试与全量测试
- commit: feat(issue-2): add aster runtime credential form

## Self Review

- [x] Requirement alignment: 前端可以临时接触并设置 Aster 凭据
- [x] Regression risk: 环境变量配置和 mock/live 摘要调用保持兼容
- [x] Test coverage: 设置、清除、脱敏、前端资源均有覆盖
- [x] Dirty changes: unrelated modifications were avoided
- [x] Docs sync: workflow、changelist、AGENTS、README、ARCHITECTURE 已同步
