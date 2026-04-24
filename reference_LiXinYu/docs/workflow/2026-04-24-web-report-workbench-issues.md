# 2026-04-24 Web 报告工作台 UI 优化 Issues

## Issue 1: 报告页重构为三栏工作台

- 状态：已完成
- 范围：`web/templates/report.html`、`web/static/app.css`
- 内容：用 sticky 顶部状态栏替换大 hero，保留项目身份、生成时间、本地运行提示、新建任务和导出入口。
- 内容：报告布局改为左侧导航、中间主工作区、右侧检查器三栏结构。
- 验证：报告模板包含 `report-topbar`、`report-inspector`，Chrome headless 能渲染并截图。

## Issue 2: 表格交互和长文本显示优化

- 状态：已完成
- 范围：`web/static/app.js`、`web/static/app.css`
- 内容：每个分区首个有数据的表格自动展开，避免用户进入分区后先看到一组关闭面板。
- 内容：表格新增紧凑/舒展行距切换，默认单行省略长文本并把完整文本放入单元格 `title`。
- 内容：横向滚动区域增加边缘阴影提示，列隐藏、排序、筛选和列宽能力继续保留。
- 验证：静态资源测试检查 `toolbar-density`，Chrome headless DOM 检查 `report-data-table` 和 `toolbar-density`。

## Issue 3: Reveal 动画可读性兜底

- 状态：已完成
- 范围：`web/static/app.css`、`web/static/app.js`
- 内容：默认 `[data-reveal]` 不再隐藏内容，只有 JS 初始化 `reveal-enabled` 后才应用入场透明度。
- 内容：首屏节点通过 `requestAnimationFrame` 做立即可见兜底，避免 headless 或弱时序下整页发白。
- 内容：初始展开表格改为同步挂载，避免出现 `is-open` 但工具栏和数据表尚未渲染的脱节状态。
- 验证：Chrome headless 截图确认首屏对比度正常，DOM 检查 `is-visible`、`report-data-table`、`toolbar-density` 均存在。
