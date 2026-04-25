# Aster 官方协议对齐

## 背景

新的 Aster 接口说明明确了 ChatFlow 与普通智能体 Room 的认证差异。ChatFlow 通过 `apiKey + empNo` 调用，不使用 accessToken；Room/Auth 需要先生成或续期 accessToken，并在问答前校验 token。

## 变更

- `pstx_aster_client.py`
  - 新增 `/auth/js-sdk/validateAccessToken` 支持，按官方示例使用 AES-256-ECB + PKCS7 + 双 Base64 加密请求体和解密响应。
  - Room/Auth 在创建 room 和发送问答前默认执行 token validate。
  - validate 返回无效时会强制重新鉴权并再次 validate，仍失败才抛出明确错误。
  - 增加 `ASTER_ORIGIN`，作为 validate 接口的 `aigc-origin` 和加解密 key；不填时从 `ASTER_BASE_URL` 推导。
  - 增加 `PSTX_ASTER_VALIDATE_TOKEN` 和 `PSTX_ASTER_VALIDATE_AUTH_HEADER`，用于受控环境调试。
  - 兼容 `cryptography`、`pycryptodome` 和官方示例中的 `pycryptodomex` 导入路径。
  - 修复 Room 流式响应中 `data.content` 结构无法抽取 answer 的问题。
- `pstx_aster_service.py`
  - Runtime config 和状态接口接入 `ASTER_ORIGIN`，不作为 secret 回显。
- `web/templates/report.html`
  - 临时 Aster 凭据表单新增 Room Validate Origin。
- `requirements.txt`
  - 增加生产依赖入口，包含 Flask、openpyxl 和 pycryptodome。

## 验证

- `python3 -m unittest tests.test_pstx_aster_client -v`
- `python3 -m unittest tests.test_pstx_aster_service tests.test_pstx_web.WebUiTests.test_localhost_web_flow_uses_real_page_for_page_summary_and_keeps_query_pages_split tests.test_pstx_web.WebUiTests.test_aster_runtime_config_can_set_and_clear_without_echoing_secret -v`
- `node --check web/static/app.js`
- `python3 -m unittest discover -s tests -v`
- `python3 -m compileall -q .`
- `git diff --check`
