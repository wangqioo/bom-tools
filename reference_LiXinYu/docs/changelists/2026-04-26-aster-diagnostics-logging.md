# Aster 诊断日志增强

## 背景

真实 Aster 调用出现 `Aster HTTP 401`，错误体包含 `Access token is invalid`。原先前端只展示错误字符串，无法快速确认请求属于 ChatFlow 还是 Room、是哪一次请求、日志在哪里，也不方便判断是否把 accessToken/API Key 类型混用了。

## 变更

- `pstx_aster_client.py`
  - 为 Aster live 请求增加 `request_id` 和脱敏 JSONL 日志。
  - 默认日志路径为 `logs/aster_debug.log`，可通过 `PSTX_ASTER_LOG_FILE` 覆盖。
  - 记录 `request.start`、`request.success`、`request.http_error`、`request.url_error`、`request.json_error`、`request.response_error` 等事件。
  - 默认只记录请求 body 的键、长度和 hash；`PSTX_ASTER_LOG_PAYLOAD=1` 时才记录裁剪后的 payload。
  - `apiKey`、`appSecret`、`accessToken`、`Authorization`、`ciphertext` 在日志和 diagnostics 中脱敏。
- `pstx_aster_service.py`
  - Aster 上游错误 payload 增加 `diagnostics`、`diagnostic_hints` 和 `log_file`。
  - 401 / unauthorized / access token invalid 时提示检查 backend、ChatFlow API Key、Room 凭据和 `ASTER_BASE_URL`。
- `web/static/app.js`
  - 报告页 Aster 摘要失败时显示请求 ID、操作、后端、HTTP 状态、日志文件和排查建议。
- 文档
  - README、架构文档和 Aster harness 解析文档补充诊断日志路径、开关和 401 排查方式。
- `.gitignore`
  - 忽略默认 `logs/` 目录，避免诊断日志误提交。

## 验证

- `python3 -m unittest tests.test_pstx_aster_client -v`
- `python3 -m unittest tests.test_pstx_aster_service -v`
- `python3 -m unittest tests.test_pstx_web.WebUiTests.test_localhost_web_flow_uses_real_page_for_page_summary_and_keeps_query_pages_split tests.test_pstx_web.WebUiTests.test_aster_live_mode_missing_config_reports_displayable_error -v`
- `node --check web/static/app.js`
- `python3 -m unittest discover -s tests -v`
- `python3 -m compileall -q .`
- `git diff --check`
