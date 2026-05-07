# Web 单项目报告打开动效页

- Date: 2026-04-24
- Scope: Web UI debug page

## 变更摘要

- 新增 `/debug/report-open` 路由，用于单独调试“打开某个项目分析报告”的动效。
- 新增 `web/templates/debug_report_open.html`，模拟项目卡片、加载态、报告工作台、KPI 和首个表格。
- 新增 `web/static/debug_report_open.js`，用 `data-phase` 驱动选择项目、加载报告、工作台就绪、数据入场四个阶段。
- 扩展 `web/static/app.css`，增加 `.report-open-*` 样式和 `report-open-rise` / `table-row-rise` 动画。
- 首页调试入口拆成“综合动效”和“单项目打开”；综合动效页也增加跳转入口。
- 补充 Web 测试，覆盖新路由、入口、专用脚本和样式锚点。

## 边界

- 不修改真实 `/report/<run_id>` 报告页逻辑。
- 不调用 `/api/analyze` 或 `/api/report`。
- 不读取项目目录或 PSTX 文件。

## 验证

- `node --check web\static\app.js`
- `node --check web\static\debug_effects.js`
- `node --check web\static\debug_report_open.js`
- `python -m unittest discover -s tests -p test_pstx_web.py -v`
- `python -m unittest discover -s tests -v`
