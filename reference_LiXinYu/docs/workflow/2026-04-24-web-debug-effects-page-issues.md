# Web Debug 动效模拟页 Issues

- Date: 2026-04-24
- Complexity: L0
- Related design: none

## Task Overview

- Goal: 增加一个本地 Debug 动效模拟页，用固定模拟 DOM 快速观察 Web UI 动画效果。
- Ordering rule: Complete issues in sequence.
- Current status: done

## Design Note

本轮新增 `/debug/effects`，定位为 motion lab。页面只展示模拟数据，不调用真实 PSTX 解析接口，不写入分析缓存；用于快速回放概览入场、表格展开、筛选弹层和项目对比结果动画。

## Issue List

- [x] issue-1 新增 Debug 动效模拟页
- [x] issue-2 补充入口、测试和文档索引

## issue-1

- ID: issue-1
- 标题: 新增 Debug 动效模拟页
- 范围: `pstx_web.py`、`web/templates/debug_effects.html`、`web/static/debug_effects.js`、`web/static/app.css`
- 依赖: none
- 验收标准: `/debug/effects` 可打开，页面能回放概览、表格、筛选弹层和对比结果模拟动画
- 状态: done
- 验证方式: `node --check web\static\debug_effects.js`；`python -m unittest discover -s tests -p test_pstx_web.py -v`
- commit: pending

## issue-2

- ID: issue-2
- 标题: 补充入口、测试和文档索引
- 范围: `web/templates/index.html`、`tests/test_pstx_web.py`、`docs/changelists/`、`AGENTS.md`
- 依赖: issue-1
- 验收标准: 首页可进入动效模拟页，测试覆盖路由、静态脚本和 debug 样式锚点
- 状态: done
- 验证方式: `python -m unittest discover -s tests -p test_pstx_web.py -v`
- commit: pending

## Self Review

- [x] Requirement alignment: 新增页面用于 Debug 模拟动效，不改变真实报告逻辑
- [x] Regression risk: 新路由与新静态脚本独立，首页只新增入口链接
- [x] Test coverage: 增加路由、脚本和样式锚点测试
- [x] Dirty changes: 当前仓库已有前序未提交改动，本轮只追加 debug motion 页面相关变更
- [x] Docs sync: 已新增 workflow 记录并准备同步 changelist / AGENTS
