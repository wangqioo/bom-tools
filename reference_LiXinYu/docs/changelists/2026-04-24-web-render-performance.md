# 2026-04-24 Web 渲染性能优化

## 变更摘要

- 首个展开表格改为近视口懒加载，不再在报告初始化时同步挂载所有分区的表格 DOM。
- 大表改为渐进渲染，默认先渲染前 320 行，保留筛选结果总量，并提供“继续渲染更多”按钮追加后续行。
- 表格筛选输入增加轻量 debounce，避免每个键入事件立即重建表格。
- 列宽拖拽改为 `requestAnimationFrame` 节流，避免 pointermove 高频触发布局重算。
- 横向滚动阴影改为帧节流更新，降低滚动过程中 class 切换频率。
- 默认原始顺序且未选择排序字段时跳过 `sort()`，减少大表初次渲染的计算量。
- 报告分区增加 `content-visibility: auto` 和 `contain-intrinsic-size`，让浏览器跳过离屏分区绘制。
- 报告页取消高频区域的 `backdrop-filter`，改用更轻的平面层级，降低合成与模糊成本。

## Bug 与修复手段

- Bug：所有分区首个有效表格会在页面加载时同步执行 `mountTable()`，大项目下会一次性创建大量 DOM 节点。
- 修复：初始展开只保留占位提示，通过 `IntersectionObserver` 在表格进入视口附近时再挂载。
- Bug：表格筛选、排序、列隐藏会完整渲染全部匹配行，数千行时首屏和交互明显变慢。
- 修复：`applyTableState()` 只渲染当前批次行，并通过 footer 显示渲染进度和继续追加入口。
- Bug：列宽拖拽每个 pointermove 都执行列宽同步，容易造成连续 reflow。
- 修复：拖拽期间只在动画帧内同步一次列宽，结束时再落盘持久化。

## 验证

- `python -m unittest discover -s tests -p test_pstx_web.py -v`
- `node --check web\static\app.js`
- Chrome headless 临时项目烟测：检查懒加载表格最终挂载、分批渲染 footer、截图输出。
