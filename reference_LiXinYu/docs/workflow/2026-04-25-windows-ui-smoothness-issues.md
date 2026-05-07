# Windows UI 显示与流畅性优化 Issues

- Date: 2026-04-25
- Complexity: L1
- Related design: `docs/workflow/2026-04-25-windows-ui-smoothness-design.md`

## Task Overview

- Goal: 面向 Windows 部署环境优化 Web UI 显示质量和流畅性。
- Ordering rule: Complete issues in sequence.
- Current status: done

## Issue List

- [x] issue-1 优化 Windows runtime hint 与大表挂载节奏
- [x] issue-2 优化 Windows 字体、滚动、报告页动效负载
- [x] issue-3 补充测试和文档索引

## issue-1

- ID: issue-1
- 标题: 优化 Windows runtime hint 与大表挂载节奏
- 范围: `web/static/app.js`
- 依赖: none
- 验收标准: 浏览器端可标记 Windows 环境；大表懒加载优先在浏览器空闲期挂载；首次渲染压力降低
- 状态: done
- 验证方式: `node --check web\static\app.js`
- commit: pending

## issue-2

- ID: issue-2
- 标题: 优化 Windows 字体、滚动、报告页动效负载
- 范围: `web/static/app.css`
- 依赖: issue-1
- 验收标准: Windows 字体栈、滚动条稳定性、表格滚动容器、报告页低负载样式和减少动态效果兼容均落地
- 状态: done
- 验证方式: `python -m unittest discover -s tests -p test_pstx_web.py -v`
- commit: pending

## issue-3

- ID: issue-3
- 标题: 补充测试和文档索引
- 范围: `tests/test_pstx_web.py`、`docs/changelists/`、`AGENTS.md`
- 依赖: issue-2
- 验收标准: 测试覆盖新增 Windows / smoothness 关键锚点，changelist 和 AGENTS 记录本轮优化
- 状态: done
- 验证方式: `python -m unittest discover -s tests -v`
- commit: pending

## Self Review

- [x] Requirement alignment: 明确面向 Windows 部署使用场景优化 UI 显示和流畅性
- [x] Regression risk: 不改解析/API，前端改动以样式和渲染节奏为主
- [x] Test coverage: 补充静态资源锚点并计划运行 Web 专项与全量测试
- [x] Dirty changes: 当前仓库已有前序未提交改动，本轮只追加 Windows UI smoothness 相关变更
- [x] Docs sync: 已新增设计、issue 和 changelist
