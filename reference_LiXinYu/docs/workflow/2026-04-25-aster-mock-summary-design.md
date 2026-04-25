# 2026-04-25 Aster Mock 辅助审查摘要设计

## 背景

用户提供 Aster Codex Harness 实验包，希望把 Aster 相关能力作为实验方向接入当前 PSTX 工具。该 harness 明确要求外部 Codex 环境不要访问真实 Aster 内网端点，因此本轮只做 mock-only 集成。

## 目标

- 落地实验包解析文档：`docs/analysis/aster_harness_review.md`。
- 新增 mock-only 的 Aster 辅助审查摘要后端能力。
- Web 报告页增加“AI 辅助审查摘要（实验）”入口。
- 新增 API 测试，验证不依赖真实 Aster。

## 非目标

- 不调用真实 Aster 内网服务。
- 不实现 Aster token 加密、登录、ChatFlow 网络调用。
- 不上传本地文件或完整报告到外部服务。
- 不实现开放式本地 tool calling。

## 接口草图

```text
GET /api/report/<run_id>/aster-summary
```

返回：

```text
{
  ok: true,
  mode: "mock",
  provider: "local-aster-mock",
  summary: "...",
  priorities: [...],
  section_focus: [...],
  safeguards: [...]
}
```

## 边界

- 输入来源为已生成的 `report` payload 和 `bundle` 内的聚合计数。
- 输出是本地规则生成的 mock 摘要，用来验证产品体验。
- 后续真实接入必须走后端配置、显式开关和脱敏预览。
