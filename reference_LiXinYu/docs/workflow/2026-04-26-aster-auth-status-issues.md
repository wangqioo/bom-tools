# Aster 认证状态展示 Issues

- Date: 2026-04-26
- Complexity: L1
- Related design: `docs/workflow/2026-04-26-aster-auth-status-design.md`

## Task Overview

- Goal: 在前端展示 Aster 认证配置状态，同时避免泄露任何 secret/token/apiKey。
- Ordering rule: Complete issues in sequence.
- Current status: completed

## Issue List

- [x] issue-1 增加 Aster 安全状态接口
- [x] issue-2 接入报告页 UI、文档和验证

## issue-1

- ID: issue-1
- 标题: 增加 Aster 安全状态接口
- 范围: `pstx_aster_service.py`、`pstx_web.py`、service/Web 单元测试
- 依赖: none
- 验收标准: `/api/aster/status` 返回 mode/backend/变量配置状态；不返回 secret 原文或片段
- 状态: done
- 验证方式: `python3 -m unittest tests.test_pstx_aster_service tests.test_pstx_web.WebUiTests.test_aster_status_endpoint_redacts_credentials -v`
- commit: dc81a2a

## issue-2

- ID: issue-2
- 标题: 接入报告页 UI、文档和验证
- 范围: `web/templates/report.html`、`web/static/app.js`、`web/static/app.css`、README/ARCHITECTURE/AGENTS/changelist
- 依赖: issue-1
- 验收标准: 报告页展示 Aster 认证状态；文档明确只连接 Aster，不直连 Dify
- 状态: done
- 验证方式: targeted Web 测试与全量测试
- commit: feat(issue-2): add aster auth status ui

## Self Review

- [x] Requirement alignment: 只连接 Aster，不新增原生 Dify
- [x] Regression risk: mock/live 摘要调用保持兼容
- [x] Test coverage: 状态接口、脱敏、前端资源均有覆盖
- [x] Dirty changes: unrelated modifications were avoided
- [x] Docs sync: workflow、changelist、AGENTS、README、ARCHITECTURE 已同步
