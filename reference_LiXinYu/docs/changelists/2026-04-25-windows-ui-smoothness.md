# Windows UI 显示与流畅性优化

- Date: 2026-04-25
- Scope: Web UI display and performance

## 变更摘要

- 新增 `bootRuntimeHints()`，浏览器端识别 Windows 并在 `<html>` 上添加 `is-windows` 样式标记。
- 将 UI 字体栈调整为 Windows 中文优先：`Microsoft YaHei UI` / `Segoe UI`，数据和代码继续保留 Cascadia / Consolas。
- 降低大表首次渲染数量：`TABLE_INITIAL_RENDER_LIMIT` 和 `TABLE_RENDER_STEP` 从 320 调整为 220。
- 新增 `runWhenBrowserIsIdle()`，近视口懒加载表格优先在浏览器空闲期挂载，降低页面打开时的主线程压力。
- 限制全局 stagger 赋值数量，避免大报告中对过多节点做无意义样式写入。
- 优化表格滚动容器：稳定滚动条占位、限制 overscroll、增加 Edge/Chrome 友好的滚动条样式。
- Windows 报告页禁用持续背景漂移和毛玻璃效果，减少 GPU 合成压力。
- 报告页数据就绪动画只强调导航、检查器和首个结果区，避免大量表格同时动画。
- 强化 `prefers-reduced-motion`，确保背景漂移等连续动画真正关闭。

## 边界

- 不修改 PSTX 解析逻辑。
- 不修改 Flask API。
- 不改变 Web 页面信息架构。

## 验证

- `node --check web\static\app.js`
- `node --check web\static\debug_effects.js`
- `node --check web\static\debug_report_open.js`
- `python -m unittest discover -s tests -p test_pstx_web.py -v`
- `python -m unittest discover -s tests -v`
