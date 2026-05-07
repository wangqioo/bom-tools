# Aster 认证状态展示

## 背景

用户希望前端能看到 Aster 相关认证信息，同时明确当前工具只连接 Aster，不直接连接原生 Dify。

## 变更

- 新增 `build_aster_status()`，汇总 `PSTX_ASTER_MODE`、`PSTX_ASTER_BACKEND`、`ASTER_BASE_URL`、`ASTER_EMP_NO`、`ASTER_API_KEY`、`ASTER_APP_ID`、`ASTER_APP_SECRET` 的配置状态。
- 新增 `GET /api/aster/status`。
- 报告页 Aster 面板新增认证状态区域，自动展示当前模式、后端、ready/missing 状态和每个配置项是否已配置。
- 前端对 secret 类字段只显示“已配置（隐藏）”或“未配置”，不显示任何密钥原文或片段。
- 文档补充：若 Aster 底层由 Dify 承载，Dify 应用发布、API Key、模型供应商和应用变量在 Aster/Dify 平台侧配置，本工具仍只连接 Aster wrapper。

## 安全边界

- 不新增前端 secret 输入框。
- 不把 `ASTER_API_KEY`、`ASTER_APP_SECRET`、`accessToken` 返回给浏览器。
- 不新增原生 Dify 直连后端。

## 验证

- `python3 -m unittest tests.test_pstx_aster_service tests.test_pstx_web.WebUiTests.test_aster_status_endpoint_redacts_credentials -v`
- `python3 -m unittest tests.test_pstx_web.WebUiTests.test_localhost_web_flow_uses_real_page_for_page_summary_and_keeps_query_pages_split tests.test_pstx_web.WebUiTests.test_aster_status_endpoint_redacts_credentials -v`
- `python3 -m unittest discover -s tests -v`
- `python3 -m compileall -q .`
- `node --check web/static/app.js`
- `git diff --check`
