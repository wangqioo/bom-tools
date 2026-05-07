# Aster 实验包解析与接入基线

## 来源

- 实验包：`/Users/rounder/Downloads/aster-codex-harness.zip`
- 解包阅读路径：`/tmp/aster-codex-harness-ref/aster-codex-harness`
- 验证方式：`node --test tests/*.test.js`
- 验证结果：4 个测试全部通过

## 实验包定位

该包是 Aster 内网模型 API 的本地集成 harness，不是当前 PSTX 工具的直接业务模块。它提供了 Node.js 侧的客户端、mock server 和本地 skill 编排样例，用于在无法访问内网 Aster 的外部开发环境中验证接口形状。

## 核心模块

- `src/crypto.js`
  - 实现 Aster 文档要求的 AES-256-ECB + PKCS7 padding + 双 Base64 加解密。
  - `genCipherKey()` 使用 `*` 左填充并截取右侧 32 字符。
- `src/auth-client.js`
  - 封装 `/auth/api/v1/generateOrProlongToken`。
  - 负责 `accessToken` 缓存、续期和 token 有效期判断。
  - `appSecret` 必须只存在后端。
- `src/room-client.js`
  - 封装普通智能体 room 创建和 room 流式问答。
  - 使用 `accessToken` 和 `/aster/room/chat/stream/send`。
- `src/flow-client.js`
  - 封装 AgentFlow、ChatFlow、ChatFlow SSE 和文件上传。
  - 使用 `apiKey`、`empNo`、`conversationId`。
- `src/sse-parser.js`
  - 区分普通 JSON line stream 和标准 SSE stream。
- `src/skill-harness.js`
  - 让模型输出 JSON 形式的本地工具请求，再由应用代码执行白名单函数。
  - 这不是 Aster 平台原生 tool calling，只能作为应用层实验模式。
- `src/mock-server.js`
  - 提供本地 mock Aster 服务，覆盖 token、room、chatFlow、SSE、upload-file 等接口。

## 安全边界

- 当前外部 Codex / macOS 开发环境不访问真实 Aster 内网地址。
- 不把 `ASTER_APP_SECRET`、`ASTER_API_KEY` 或 token 下发到前端。
- Web UI 只能请求本机后端，由后端决定是否调用 mock 或未来的真实内网服务。
- 本地 skill harness 必须保持白名单、参数校验和只读优先。
- 涉及上传文件、发送报告内容到真实 Aster、保存凭据等行为，后续必须单独增加确认和配置开关。

## 当前仓库接入策略

当前仓库已经保留 mock，并新增 live 生产调用能力：

- `PSTX_ASTER_MODE=mock`：调用 `pstx_aster_mock.py`，只基于本地报告指标生成摘要。
- `PSTX_ASTER_MODE=live`：调用 `pstx_aster_client.py`，由后端请求真实 Aster。
- 默认 live 后端为 ChatFlow：`PSTX_ASTER_BACKEND=chat-flow`。
- ChatFlow 协议使用 `apiKey + empNo`，不走 accessToken。
- Room/Auth 协议也已保留，但生产环境需要安装 AES 依赖 `pycryptodome`、`pycryptodomex` 或 `cryptography`。
- Room/Auth 调用顺序为生成/续期 token -> validate token -> create room -> send question。
- Web 前端仍只请求本机后端接口，不接触任何 Aster secret/token。
- 发送给真实 Aster 的内容是裁剪后的报告摘要，不上传完整 PSTX 原始文件。
- Web 通过 `/api/aster/status` 展示 Aster 认证状态，只显示配置项是否存在，不显示 secret 原文或掩码片段。
- Web 也支持 `POST /api/aster/runtime-config` 临时提交 Aster 凭据到后端内存，适合本地/受控内网临时调试。
- 当前工具只连接 Aster wrapper；若底层由 Dify 承载，Dify 应用发布和 API Key 生成在 Aster/Dify 平台侧完成。
- live 调用会写入脱敏 JSONL 诊断日志，默认 `logs/aster_debug.log`，用于定位 401/网络错误/返回格式错误。
- 诊断日志默认只记录 payload 长度和 hash；`PSTX_ASTER_LOG_PAYLOAD=1` 才会额外记录裁剪后的 payload，生产排障后应关闭。

## 生产环境变量

- `PSTX_ASTER_MODE=mock|live|off`
- `PSTX_ASTER_BACKEND=chat-flow|room`
- `ASTER_BASE_URL`
- `ASTER_EMP_NO`
- `ASTER_API_KEY`，ChatFlow 必需
- `ASTER_APP_ID` / `ASTER_APP_SECRET`，Room/Auth 必需
- `ASTER_ORIGIN`，Room/Auth validate 可选，不填时从 `ASTER_BASE_URL` 推导
- `PSTX_ASTER_VALIDATE_TOKEN`
- `PSTX_ASTER_VALIDATE_AUTH_HEADER`
- `PSTX_ASTER_TIMEOUT_SECONDS`
- `PSTX_ASTER_MAX_ROWS_PER_TABLE`
- `PSTX_ASTER_MAX_PAYLOAD_CHARS`
- `PSTX_ASTER_REDACT_PATHS`
- `PSTX_ASTER_LOG_FILE`
- `PSTX_ASTER_LOG_PAYLOAD`
- `PSTX_ASTER_LOG_MAX_BYTES`

## 后续建议

1. 在真实内网环境用 Aster 提供的测试 flow 做一次端到端验收。
2. 根据内部安全要求决定是否开启路径脱敏、最大样例行数和最大 payload 字符数。
3. 如果需要上传文件型 flow，再单独增加文件上传开关和上传前预览。
