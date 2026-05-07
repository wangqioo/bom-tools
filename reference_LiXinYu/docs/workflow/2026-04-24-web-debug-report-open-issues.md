# Web 单项目报告打开动效页 Issues

- Date: 2026-04-24
- Complexity: L0
- Related design: none

## Task Overview

- Goal: 增加一个独立 Debug 页面，用模拟数据展示“单个项目打开分析报告”的首屏过渡动效。
- Ordering rule: Complete issues in sequence.
- Current status: done

## Design Note

本轮新增 `/debug/report-open`。页面模拟从项目卡片进入报告工作台的状态流：选择项目、加载报告、工作台就绪、数据入场。它不读取真实项目、不调用分析 API，只用于调试打开报告时的动效节奏。

## Issue List

- [x] issue-1 新增单项目报告打开动效页
- [x] issue-2 补充入口、测试和文档索引

## issue-1

- ID: issue-1
- 标题: 新增单项目报告打开动效页
- 范围: `pstx_web.py`、`web/templates/debug_report_open.html`、`web/static/debug_report_open.js`、`web/static/app.css`
- 依赖: none
- 验收标准: `/debug/report-open` 可打开，并能播放项目卡片到报告工作台首屏的模拟动效
- 状态: done
- 验证方式: `node --check web\static\debug_report_open.js`；`python -m unittest discover -s tests -p test_pstx_web.py -v`
- commit: pending

## issue-2

- ID: issue-2
- 标题: 补充入口、测试和文档索引
- 范围: `web/templates/index.html`、`web/templates/debug_effects.html`、`tests/test_pstx_web.py`、`docs/changelists/`、`AGENTS.md`
- 依赖: issue-1
- 验收标准: 首页与综合动效页可进入单项目打开页，测试覆盖路由、静态脚本和样式锚点
- 状态: done
- 验证方式: `python -m unittest discover -s tests -p test_pstx_web.py -v`
- commit: pending

## Self Review

- [x] Requirement alignment: 页面专注单项目打开报告动效，不混入多项目对比或真实解析
- [x] Regression risk: 新路由和新脚本独立，现有报告页行为不变
- [x] Test coverage: 增加路由、入口、JS 和 CSS 锚点测试
- [x] Dirty changes: 当前仓库已有前序未提交改动，本轮只追加 report-open debug 相关变更
- [x] Docs sync: 已新增 workflow 记录并准备同步 changelist / AGENTS
