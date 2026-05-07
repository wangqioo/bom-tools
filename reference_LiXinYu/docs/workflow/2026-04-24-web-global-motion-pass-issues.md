# Web 全站动效优化 Issues

- Date: 2026-04-24
- Complexity: L1
- Related design: none

## Task Overview

- Goal: 给所有 Web 相关页面补齐统一、克制、工程感更强的动效，包括首页、报告页、查询、项目管理、对比和 Debug 页面。
- Ordering rule: Complete issues in sequence.
- Current status: done

## Issue List

- [x] issue-1 扩展全站 motion helper
- [x] issue-2 补齐页面级 CSS 动效
- [x] issue-3 补充测试和文档索引

## issue-1

- ID: issue-1
- 标题: 扩展全站 motion helper
- 范围: `web/static/app.js`
- 依赖: none
- 验收标准: 页面启动、加载遮罩、查询结果和报告数据就绪均有统一动效状态入口
- 状态: done
- 验证方式: `node --check web\static\app.js`
- commit: pending

## issue-2

- ID: issue-2
- 标题: 补齐页面级 CSS 动效
- 范围: `web/static/app.css`
- 依赖: issue-1
- 验收标准: 首页、报告页、Debug 页、表单、按钮、查询结果和加载态均具备一致的入场/反馈动画，并兼容减少动态效果
- 状态: done
- 验证方式: `python -m unittest discover -s tests -p test_pstx_web.py -v`
- commit: pending

## issue-3

- ID: issue-3
- 标题: 补充测试和文档索引
- 范围: `tests/test_pstx_web.py`、`docs/changelists/`、`AGENTS.md`
- 依赖: issue-2
- 验收标准: Web 静态资源测试覆盖新增 motion helper 和 CSS 动画锚点，changelist 与 AGENTS 可追踪本轮变更
- 状态: done
- 验证方式: `python -m unittest discover -s tests -v`
- commit: pending

## Self Review

- [x] Requirement alignment: 覆盖所有 Web 相关页面，不修改真实解析和分析 API
- [x] Regression risk: 动效以 class / CSS animation 叠加为主，原 DOM 与数据结构保持不变
- [x] Test coverage: 增加静态资源锚点，并计划运行 Web 专项与全量测试
- [x] Dirty changes: 当前仓库已有前序未提交改动，本轮只追加全站动效相关变更
- [x] Docs sync: 已新增 workflow 记录并准备同步 changelist / AGENTS
