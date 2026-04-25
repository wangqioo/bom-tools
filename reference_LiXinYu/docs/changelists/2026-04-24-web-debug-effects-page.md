# Web Debug 动效模拟页

- Date: 2026-04-24
- Scope: Web UI debug page

## 变更摘要

- 新增 `/debug/effects` 路由，渲染本地 Debug 动效模拟页。
- 新增 `web/templates/debug_effects.html`，集中展示概览入场、表格展开、筛选弹层和项目对比结果动画。
- 新增 `web/static/debug_effects.js`，使用模拟 DOM 回放动效，不触发真实 PSTX 解析或报告缓存写入。
- 在首页增加“动效模拟页”入口，便于 UI 调试时快速进入。
- 增加 `.debug-*` 样式，保持与现有 Web 报告视觉体系一致，并兼容减少动态效果设置。
- 补充 Web 测试，覆盖 debug 路由、静态脚本和样式锚点。

## 边界

- 不修改分析 API。
- 不修改报告数据结构。
- 不读取项目目录或 PSTX 文件。

## 验证

- `node --check web\static\app.js`
- `node --check web\static\debug_effects.js`
- `python -m unittest discover -s tests -p test_pstx_web.py -v`
- `python -m unittest discover -s tests -v`
