# Aster Mock 辅助审查摘要

## 背景

基于 `aster-codex-harness.zip` 实验包，本轮先做 mock-only 接入，验证 Aster 辅助审查摘要的产品形态和后端 API 形状，不访问真实内网 Aster。

## 变更

- 新增 `docs/analysis/aster_harness_review.md`，整理实验包模块、接口边界、安全风险和后续接入建议。
- 新增 `pstx_aster_mock.py`，根据当前报告聚合指标生成本地 mock 辅助摘要。
- 新增 `GET /api/report/<run_id>/aster-summary`。
- 报告页概览区新增“AI 辅助审查摘要”实验入口。
- 前端新增 mock 摘要、优先级建议和分区关注度渲染逻辑与样式。
- 新增单元测试覆盖 mock 摘要和 Web API。

## 安全边界

- 不访问真实 Aster 内网地址。
- 不读取或下发 `ASTER_APP_SECRET`、`ASTER_API_KEY`、`accessToken`。
- mock 摘要只使用当前报告聚合指标，不上传文件或完整项目内容。
- 真实 Aster 接入需另行增加显式开关、服务端凭据配置和脱敏预览。

## 验证

- `python3 -m unittest tests.test_pstx_aster_mock -v`
- `python3 -m unittest tests.test_pstx_web.WebUiTests.test_localhost_web_flow_uses_real_page_for_page_summary_and_keeps_query_pages_split -v`
- `python3 -m unittest discover -s tests -v`
- `python3 -m compileall -q .`
- `node --check web/static/app.js`
- `git diff --check`
