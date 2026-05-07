# Aster 浮窗审查助手

## 背景

Aster 已能调用远程服务后，原来的入口仍是概览区内嵌摘要按钮，承载不了更强的 AI 审查工作流。需要先把它升级为可持续驻留的浮窗，并让更多已有规则结果进入 AI 审查上下文。

## 变更

- `pstx_aster_client.py`
  - Aster brief 增加 `review_scope`，显式覆盖 BOM/DEPOP、网络/页码映射、DRC、芯片 Pin/电阻、降额和 CSA 规范。
  - Aster brief 增加 `key_findings`，按风险提示聚合非空结果表和代表性样例行。
  - Aster brief 增加 `manual_review_boundaries`，提醒 AI 不得把电压 token、OD/OC、AC 耦合、DEPOP 排除和 CSA 几何候选误写成确定结论。
  - Prompt schema 增加 `review_checklist` 和 `manual_review`。
  - 返回 payload 标准化 `review_checklist` 和 `manual_review`。
- `pstx_aster_mock.py`
  - 本地 mock 同步返回审查清单和人工复核项，方便离线验证 UI。
- `web/templates/report.html`
  - Aster 区块改为右下角可展开/收起浮窗助手。
- `web/static/app.js` / `web/static/app.css`
  - 新增浮窗 launcher、收起状态、审查清单展示和人工复核展示。

## 验证

- `python3 -m unittest tests.test_pstx_aster_client tests.test_pstx_aster_mock -v`
- `python3 -m unittest tests.test_pstx_web.WebUiTests.test_localhost_web_flow_uses_real_page_for_page_summary_and_keeps_query_pages_split -v`
- `node --check web/static/app.js`
- `python3 -m unittest discover -s tests -v`
- `python3 -m compileall -q .`
- `git diff --check`
