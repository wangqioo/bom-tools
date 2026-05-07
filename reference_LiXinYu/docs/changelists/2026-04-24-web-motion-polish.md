# Web UI 动效优化

- Date: 2026-04-24
- Scope: Web UI animation layer

## 变更摘要

- 新增通用前端动效 helper：`prefersReducedMotion()`、`restartMotion()`、`staggerChildren()` 和 `animateTableOpen()`。
- 优化报告概览的 KPI、重点提示、分区卡片入场节奏，避免所有信息同时跳出。
- 优化表格展开反馈：展开状态增加边框/阴影变化、箭头旋转和表体进入动画。
- 优化列显示和多列筛选面板：打开时增加轻量弹出动画。
- 优化项目管理和两两对比：项目列表刷新、对比结果出现、差异统计卡片和差异块使用分层 stagger 动画。
- 保留 `prefers-reduced-motion` 支持，减少动效偏好下禁用新增动画。

## 边界

- 不修改报告数据结构。
- 不修改 `/api/projects`、`/api/compare` 或解析分析逻辑。
- 不引入新的前端依赖。

## 验证

- `node --check web\static\app.js`
- `python -m unittest discover -s tests -p test_pstx_web.py -v`
- `python -m unittest discover -s tests -v`
