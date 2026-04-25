# Web UI 动效优化 Issues

- Date: 2026-04-24
- Complexity: L0
- Related design: none

## Task Overview

- Goal: 在不改变报告数据结构和页面主布局的前提下，为 Web UI 增加更清晰、更克制的状态动效。
- Ordering rule: Complete issues in sequence.
- Current status: done

## Design Note

本轮只做轻量 motion layer：动效用于页面入场、表格展开、筛选面板打开、项目列表刷新和项目对比结果出现，不引入额外前端依赖，不改变后端 API。

## Issue List

- [x] issue-1 增加 Web UI motion layer
- [x] issue-2 补充测试锚点和文档索引

## issue-1

- ID: issue-1
- 标题: 增加 Web UI motion layer
- 范围: `web/static/app.js` 与 `web/static/app.css`
- 依赖: none
- 验收标准: 页面 reveal、表格展开、列面板、项目列表和对比结果均具备轻量动效，并尊重 `prefers-reduced-motion`
- 状态: done
- 验证方式: `node --check web\static\app.js`；`python -m unittest discover -s tests -p test_pstx_web.py -v`
- commit: pending

## issue-2

- ID: issue-2
- 标题: 补充测试锚点和文档索引
- 范围: `tests/test_pstx_web.py`、`docs/changelists/`、`AGENTS.md`
- 依赖: issue-1
- 验收标准: 静态资源测试能覆盖新增动效入口，changelist 与 AGENTS 索引指向本轮变更
- 状态: done
- 验证方式: `python -m unittest discover -s tests -p test_pstx_web.py -v`
- commit: pending

## Self Review

- [x] Requirement alignment: 动效只覆盖用户要求的 Web UI 动画优化，不改变业务解析逻辑
- [x] Regression risk: 保留原有 DOM 结构和 API，仅新增 class/helper 与 CSS animation
- [x] Test coverage: 补充了静态资源断言并计划运行 Web 测试与全量测试
- [x] Dirty changes: 当前仓库已有前序未提交改动，本轮只追加 UI motion 相关变更
- [x] Docs sync: 已新增 workflow 记录并准备同步 changelist / AGENTS
