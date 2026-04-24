# 2026-04-24 Web 报告工作台 UI 优化

## 变更摘要

- 报告页改为三栏工作台：顶部状态栏、左侧可收起导航、中间主审查区、右侧上下文检查器。
- 表格默认打开每个分区首个有效结果，减少“全是折叠面板”的凌乱感。
- 表格新增紧凑/舒展行距切换；长文本默认单行省略，完整内容通过 tooltip 保留。
- 横向表格增加滚动阴影，继续保留列隐藏、排序、筛选和列宽调整能力。
- reveal 动画改为渐进增强，默认内容可见，避免 headless 或弱时序下整页发白。
- 初始展开表格改为同步挂载，避免展开外壳存在但工具栏和数据表尚未出现。

## Bug 与修复手段

- Bug：报告首屏原结构偏像静态说明页，提交项目路径后的分析结果入口分散，用户需要下滑和展开多个区域才能开始审查。
- 修复：用 `report-topbar`、`report-layout`、`report-inspector` 重组页面，形成明确的“状态 + 导航 + 工作区 + 上下文”信息架构。
- Bug：长字段会把表格撑得难以阅读，虽然已有列宽能力，但默认状态仍容易显得拥挤。
- 修复：默认 `td` 单行省略并设置 `title`，新增舒展行距模式用于查看完整内容。
- Bug：首个表格自动展开原本依赖异步帧挂载，在 headless 渲染中可能出现已展开但表格 body 为空。
- 修复：初始展开时同步调用 `mountTable()`，点击展开仍按需挂载。
- Bug：`[data-reveal]` 默认 opacity 为 0，若 IntersectionObserver 或截图时序不触发，会导致页面整体发白。
- 修复：只有 `reveal-enabled` 激活后才应用隐藏态，并对首屏节点添加立即可见兜底。

## 验证

- `python -m unittest discover -s tests -p test_pstx_web.py -v`
- `python -m unittest discover -s tests -p test_pstx_analyzer.py -v`
- `python -m unittest discover -s tests -v`
- `node --check web\static\app.js`
- `python -m compileall -q .`
- Chrome headless 临时项目烟测：检查 `report-topbar`、`report-inspector`、`table-block is-open`、`toolbar-density`、`report-data-table`，并生成截图。
