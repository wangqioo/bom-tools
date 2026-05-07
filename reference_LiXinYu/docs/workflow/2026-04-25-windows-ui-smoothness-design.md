# Windows UI 显示与流畅性优化

- Date: 2026-04-25
- Complexity: L1
- Status: final

## Background

项目主要在 Windows 环境部署和使用，但当前开发发生在 macOS。前一轮已为 Web UI 增加全站动效；本轮需要确保这些动效在 Windows 浏览器、中文字体、大表格和较弱 GPU 环境下仍然清晰、稳定、流畅。

## Goal

提升 Windows 下 Web UI 的显示质量和交互流畅性，重点覆盖字体、滚动、表格渲染、动效负载和报告页高密度数据阅读体验。

## Non-goals

- 不修改 PSTX 解析逻辑。
- 不修改分析 API 或报告数据结构。
- 不引入新的前端框架或构建工具。
- 不改变现有页面信息架构。

## Solution

- 增加 Windows runtime hint，在浏览器端识别 Windows 并启用 Windows 优先字体与低负载报告页样式。
- 将中文 UI 字体优先切到 `Microsoft YaHei UI` / `Segoe UI`，代码和长数据仍保留 Cascadia / Consolas 体系。
- 降低大表首次渲染行数，并用 `requestIdleCallback` 在浏览器空闲期挂载懒加载表格。
- 优化滚动容器：稳定滚动条占位、收敛 overscroll、增加 Windows/Edge 友好的滚动条样式。
- 报告页禁用持续背景漂移动画和毛玻璃，减少 GPU/合成压力。
- 将报告页表格入场动画限制到首个结果区，避免大量表格同时动画。
- 保留 `prefers-reduced-motion`，并确保连续背景动画在减少动态效果时真正关闭。

## Impact

- Windows 下中文显示更清晰，表格和表单字体回退更稳定。
- 大报告首次打开更轻，近视口表格挂载更平滑。
- 报告页静态阅读更稳，Debug 页仍保留足够展示感。

## Risks

- 首次渲染行数降低后，用户看到“继续渲染更多”的频率可能略高，但换取报告初始流畅度。
- Windows runtime hint 依赖浏览器 UA，仅用于优化样式，不影响功能判断。

## Verification Plan

- `node --check web\static\app.js`
- `node --check web\static\debug_effects.js`
- `node --check web\static\debug_report_open.js`
- `python -m unittest discover -s tests -p test_pstx_web.py -v`
- `python -m unittest discover -s tests -v`
- `python -m compileall -q .`
- `git diff --check`
