# Web 全站动效优化

- Date: 2026-04-24
- Scope: Web UI motion system

## 变更摘要

- 新增全站 motion 初始化：`bootPageMotion()`、`applyGlobalStaggers()`、`showLoadingMask()`、`hideLoadingMask()`。
- 首页增加页面级入场、表单字段 stagger、按钮/输入焦点反馈和加载遮罩收束动画。
- 报告页增加报告数据就绪后的侧栏、顶栏、检查器、导航和表格入场节奏。
- 查询结果增加加载呼吸态和结果卡片入场动画。
- 项目管理、项目对比、综合 Debug 页、单项目报告打开 Debug 页统一到相同动效节奏。
- 新增 ambient 背景漂移、页面面板入场和加载面板入场 keyframes。
- 保留 `prefers-reduced-motion` 兼容，减少动态效果时禁用新增动画。

## 边界

- 不修改 PSTX 解析逻辑。
- 不修改分析 API 或报告数据结构。
- 不引入新的前端依赖。

## 验证

- `node --check web\static\app.js`
- `node --check web\static\debug_effects.js`
- `node --check web\static\debug_report_open.js`
- `python -m unittest discover -s tests -p test_pstx_web.py -v`
- `python -m unittest discover -s tests -v`
