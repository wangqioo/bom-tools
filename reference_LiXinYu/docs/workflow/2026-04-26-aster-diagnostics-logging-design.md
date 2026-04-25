# Aster 诊断日志增强

- Date: 2026-04-26
- Complexity: L1
- Status: final

## Background

生产调用 Aster 时出现 `Aster HTTP 401`，返回体包含 `Access token is invalid`。当前 UI 只展示错误字符串，缺少请求链路、后端模式、配置来源、请求 ID 和可定位的本地日志。

## Goal

- 为 Aster live 调用写入持久化 JSONL 诊断日志。
- 错误响应返回安全 diagnostics，方便 UI 展示 request_id 和日志位置。
- 不记录 `apiKey`、`appSecret`、`accessToken`、`ciphertext` 明文。
- 对 401 / access token 类错误给出排查提示。

## Non-goals

- 不在当前开发机调用真实 Aster。
- 不把完整 PSTX 报告 prompt 默认写入日志。
- 不改变 Aster wrapper 的真实调用协议。

## Solution

- `pstx_aster_client.py`
  - 增加安全脱敏、请求 ID、JSONL 日志写入。
  - HTTP 请求开始、HTTP 错误、JSON 解析错误、Room/Auth token 获取关键节点写日志。
  - 异常携带 sanitized diagnostics。
- `pstx_aster_service.py`
  - 错误 payload 带 diagnostics、diagnostic_hints 和 log_file。
- `web/static/app.js`
  - 摘要失败时渲染诊断信息，而不是只显示一行错误。

## Impact

- 默认日志路径：`logs/aster_debug.log`。
- 可通过 `PSTX_ASTER_LOG_FILE` 修改日志路径。
- 可通过 `PSTX_ASTER_LOG_PAYLOAD=1` 额外记录裁剪后的请求 payload；默认只记录长度和 hash。

## Risks

- 如果开启 payload 日志，可能记录工程摘要信息，生产需谨慎。
- 日志不包含密钥明文，因此仍需要结合配置状态面板确认密钥是否填对。

## Verification Plan

- 单元测试模拟 401，确认日志写入且不包含 API Key 明文。
- 单元测试确认错误 payload 带 diagnostics 和 401 排查提示。
- Web 资源测试确认错误诊断渲染逻辑存在。
- 跑全量 unittest、Python 编译、JS 语法检查和 diff 检查。
