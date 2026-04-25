# Aster 官方协议对齐

- Date: 2026-04-26
- Complexity: L1
- Status: final

## Background

新的 Aster 接口说明明确区分了两类调用：

- ChatFlow / AgentFlow：通过 `apiKey + empNo` 查询参数调用 `/aster/flow-api/run/...`，不走 accessToken。
- 普通智能体 Room：先通过 `/auth/api/v1/generateOrProlongToken` 获取或续期 accessToken，再用 `Authorization: Bearer <token>` 创建会话和问答，并建议在问答前验证 token 状态。

当前代码已有 ChatFlow 和 Room/Auth 两条链路，但 Room/Auth 只做生成/续期，没有接入 `/auth/js-sdk/validateAccessToken`。这会导致 token 失效、设备不匹配、API 未开启或权限变化时，错误延后到 room create / chat send，表现成不够明确的 401。

## Goal

- 保持 ChatFlow 为 `apiKey + empNo`，不引入 accessToken。
- 为 Room/Auth 增加 token validate 步骤，符合“问答前确保登录态正常”的文档要求。
- 增加 `ASTER_ORIGIN` 配置，用作 validate 接口的 `aigc-origin` 和加解密 key。
- 继续保证日志、状态接口和前端不回显 secret/token/API Key 明文。

## Non-goals

- 不接入 AgentFlow/Workflow 输出解析。
- 不上传文件。
- 不在当前开发机调用真实 Aster 内网。

## Solution

- `pstx_aster_client.py`
  - 增加 raw encrypted request helper。
  - 增加 validate token 加密请求、解密响应和失败 diagnostics。
  - Room 保护接口调用前执行 `ensure_valid_access_token()`。
  - 增加 `ASTER_ORIGIN` 和 `PSTX_ASTER_VALIDATE_TOKEN`。
- `pstx_aster_service.py`
  - Runtime config / status 接入 `ASTER_ORIGIN`。
- `web/templates/report.html`
  - 临时凭据表单允许填写 Room Validate Origin。
- 文档
  - README 和架构文档明确 ChatFlow 与 Room/Auth 的协议边界。

## Verification Plan

- 新增 room/auth mock，覆盖 token 生成、validate、room create、chat send。
- 覆盖 validate 失败时 force renew 后再 validate 的路径。
- 跑 Aster client/service/web 相关测试和全量回归。
