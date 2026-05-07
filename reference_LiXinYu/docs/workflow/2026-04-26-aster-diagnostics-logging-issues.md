# Aster 诊断日志增强 Issues

- Date: 2026-04-26
- Complexity: L1
- Related design: `docs/workflow/2026-04-26-aster-diagnostics-logging-design.md`

## Task Overview

- Goal: 增强 Aster live 调用错误排查能力，同时避免日志泄露敏感凭据。
- Ordering rule: Complete issues in sequence.
- Current status: complete

## Issue List

- [x] issue-1 增加 Aster 安全诊断日志
- [x] issue-2 接入 Web 错误诊断展示和文档

## issue-1

- ID: issue-1
- 标题: 增加 Aster 安全诊断日志
- 范围: `pstx_aster_client.py`、`tests/test_pstx_aster_client.py`
- 依赖: none
- 验收标准: 401 时写入 JSONL 日志；异常 diagnostics 包含 request_id/status/log_file；日志不包含 API Key 明文
- 状态: done
- 验证方式: `python3 -m unittest tests.test_pstx_aster_client -v`
- commit: `feat(issue-1): add aster diagnostics logging`

## issue-2

- ID: issue-2
- 标题: 接入 Web 错误诊断展示和文档
- 范围: `pstx_aster_service.py`、`web/static/app.js`、README/ARCHITECTURE/AGENTS/changelist、相关测试
- 依赖: issue-1
- 验收标准: Aster 错误 payload 和前端展示包含安全 diagnostics/hints/log_file；文档说明日志路径和开关
- 状态: done
- 验证方式: targeted Web/service 测试与全量测试
- commit: `feat(issue-1): add aster diagnostics logging`

## Self Review

- [x] Requirement alignment: 增加足够日志用于定位 401
- [x] Regression risk: mock/live 摘要调用保持兼容
- [x] Test coverage: 401 日志、脱敏、错误 payload、前端资源均有覆盖
- [x] Dirty changes: unrelated modifications were avoided
- [x] Docs sync: workflow、changelist、AGENTS、README、ARCHITECTURE 已同步
