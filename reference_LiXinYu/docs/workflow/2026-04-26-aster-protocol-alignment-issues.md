# Aster 官方协议对齐 Issues

- Date: 2026-04-26
- Complexity: L1
- Related design: `docs/workflow/2026-04-26-aster-protocol-alignment-design.md`

## Task Overview

- Goal: 将 Room/Auth 链路补齐到官方文档要求的 token validate 流程，同时保持 ChatFlow 不走 accessToken。
- Ordering rule: Complete issues in sequence.
- Current status: complete

## Issue List

- [x] issue-1 对齐 Room/Auth token validate 协议
- [x] issue-2 接入配置、UI 和文档

## issue-1

- ID: issue-1
- 标题: 对齐 Room/Auth token validate 协议
- 范围: `pstx_aster_client.py`、`tests/test_pstx_aster_client.py`
- 依赖: none
- 验收标准: Room 后端在 create/chat 前完成 token 获取和 validate；validate 失败会 force renew；日志保持脱敏
- 状态: done
- 验证方式: `python3 -m unittest tests.test_pstx_aster_client -v`
- commit: `fix(issue-1): align aster room auth protocol`

## issue-2

- ID: issue-2
- 标题: 接入配置、UI 和文档
- 范围: `pstx_aster_service.py`、`web/templates/report.html`、README/ARCHITECTURE/AGENTS/changelist、相关测试
- 依赖: issue-1
- 验收标准: `ASTER_ORIGIN` 可通过环境变量/临时表单设置；状态接口脱敏展示；文档明确 ChatFlow 与 Room/Auth 边界
- 状态: done
- 验证方式: targeted service/web tests 与全量测试
- commit: `fix(issue-1): align aster room auth protocol`

## Self Review

- [x] Requirement alignment: 官方文档中的 token validate 要求已落实
- [x] Regression risk: ChatFlow 路径仍使用 apiKey + empNo，不引入 accessToken
- [x] Test coverage: room/auth 成功、validate 失败重取 token、状态配置均有覆盖
- [x] Dirty changes: unrelated modifications were avoided
- [x] Docs sync: workflow、changelist、AGENTS、README、ARCHITECTURE 已同步
