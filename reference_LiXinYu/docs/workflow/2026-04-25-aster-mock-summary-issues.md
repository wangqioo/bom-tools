# 2026-04-25 Aster Mock 辅助审查摘要 Issues

## Issues

### Issue 1

- ID: aster-mock-1
- 标题: 梳理 Aster harness 实验包
- 范围: `docs/analysis/aster_harness_review.md`
- 依赖: 无
- 验收标准: 文档覆盖模块职责、安全边界、mock-only 接入策略和后续真实接入建议
- 状态: completed
- 验证方式: 人工 review 文档内容
- commit: pending

### Issue 2

- ID: aster-mock-2
- 标题: 增加 mock-only 辅助审查摘要 API
- 范围: 新增后端模块、`pstx_web.py` API 路由、Web 报告页入口、测试
- 依赖: aster-mock-1
- 验收标准: 报告页可请求 mock 摘要；API 返回 `mode=mock`；不依赖真实 Aster 环境变量
- 状态: completed
- 验证方式: Web 单元测试与全量测试
- commit: pending

## Self Review

- [x] 不访问真实 Aster 内网地址
- [x] 前端不接触 secret/token/apiKey
- [x] mock 输出可解释且能指向已有报告分区
- [x] 后续 live 接入边界记录清楚
