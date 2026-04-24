# AGENTS

## 运行入口

- 本地 UI：`python pstx_local_ui.py`
- Web UI：`python pstx_web.py`
- 兼容入口：`python pstx_analyzer.py`

## 代码边界

- 页码解析只放在 `pstx_page_logic.py`
- 规则分析和 Excel 导出放在 `pstx_analyzer.py`
- Web 展示放在 `pstx_web.py` 和 `web/`
- 本地桌面入口只做 Web UI 壳，不复制业务逻辑

## 文档路径

- `docs/ARCHITECTURE.md`
  - 当前模块边界和页码模型
- `docs/REVIEW.md`
  - 当前版本 review 摘要
- `docs/reviews/`
  - 规则逻辑 review 与人工判断边界记录
- `docs/workflow/`
  - 每轮设计和 issue 拆分记录
- `docs/changelists/`
  - 每轮变更摘要

## 最新变更

- `docs/changelists/2026-04-24-web-multi-column-filtering.md`
  - Web 表格多列组合筛选：支持多字段、多条件 AND 筛选并保留原有关键字和排序能力
- `docs/changelists/2026-04-24-web-render-performance.md`
  - Web 报告页渲染性能优化：近视口懒加载、大表渐进渲染、拖拽/滚动节流和离屏绘制隔离
- `docs/changelists/2026-04-24-web-engineering-ui-pass.md`
  - 对齐生成概念图后的 Web 工程审美优化：固定侧栏、项目状态条、KPI dock、表格优先
- `docs/changelists/2026-04-24-web-report-workbench.md`
  - Web 报告页三栏工作台、表格密度切换、长文本显示和 reveal 可读性修复
- `docs/reviews/2026-04-24-logic-rule-review.md`
  - 规则逻辑 review 记录：已修复问题、保守保持项和后续样本需求
- `docs/changelists/2026-04-24-logic-review-hardening.md`
  - 规则逻辑 review 后的电阻值解析与多路径串阻搜索修复
- `docs/changelists/2026-04-24-code-review-hardening.md`
  - 全代码 review 后的 module_order 去重与 Web 输入解码硬化修复
- `docs/changelists/2026-04-24-submodule-real-page-mapping.md`
  - `module_order(.dat)` 子模块映射真实页修复
- `docs/changelists/2026-04-24-derating-50v-override.md`
  - 电容降额的 12V / 50V 低压直通规则

## 测试

```bash
python -m unittest discover -s tests -q
```

新增页码或 UI 行为时，优先跑：

- `tests/test_pstx_analyzer.py`
- `tests/test_pstx_web.py`
- `tests/test_pstx_local_ui.py`
