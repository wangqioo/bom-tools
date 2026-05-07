# Aster Live 生产接入 Issues

- Date: 2026-04-25
- Complexity: L2
- Related design: `docs/workflow/2026-04-25-aster-live-production-design.md`

## Task Overview

- Goal: 将 Aster mock 摘要升级为可通过环境变量切换到真实生产 ChatFlow 的后端能力。
- Ordering rule: Complete issues in sequence.
- Current status: completed

## Issue List

- [x] issue-1 实现真实 Aster 客户端与摘要标准化
- [x] issue-2 接入 Web API、文档和生产部署说明

## issue-1

- ID: issue-1
- 标题: 实现真实 Aster 客户端与摘要标准化
- 范围: `pstx_aster_client.py`、客户端单元测试
- 依赖: none
- 验收标准: 能对模拟 ChatFlow 发起真实 HTTP 请求；能解析 JSON 或文本回答；配置缺失有明确错误
- 状态: done
- 验证方式: `python3 -m unittest tests.test_pstx_aster_client -v`
- commit: d4b342b

## issue-2

- ID: issue-2
- 标题: 接入 Web API、文档和生产部署说明
- 范围: `pstx_aster_service.py`、`pstx_web.py`、Web 测试、README/ARCHITECTURE/AGENTS/changelist
- 依赖: issue-1
- 验收标准: 默认 mock 不变；`PSTX_ASTER_MODE=live` 可调用 live service；错误状态可被前端显示；生产部署说明完整
- 状态: done
- 验证方式: Web 单元测试与全量测试
- commit: feat(issue-2): integrate aster live web mode

## Self Review

- [x] Requirement alignment: implementation supports real production live mode without frontend secrets
- [x] Regression risk: default mock behavior and existing Web flow remain compatible
- [x] Test coverage: live ChatFlow, missing config, mock compatibility and Web route are covered
- [x] Dirty changes: unrelated modifications were avoided
- [x] Docs sync: workflow docs, changelist, AGENTS, README and architecture docs are updated
