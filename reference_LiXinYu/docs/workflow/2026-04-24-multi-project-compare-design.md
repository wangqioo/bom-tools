# 多项目管理与两两对比

- Date: 2026-04-24
- Complexity: L1
- Status: final

## Background

当前 Web UI 以单次分析报告为中心，后端通过 `RUN_CACHE` 临时保存最近分析结果。用户希望在一次 localhost 会话中管理多个项目，并能选择任意两个已分析项目做差异对比。

## Goal

提供会话内项目列表、最近项目选择、两两对比 API 和 Web 展示。对比结果应覆盖指标、元件/BOM、网络连接、主要检查结果表，并清晰标注新增、删除、变化。

## Non-goals

本轮不引入数据库、不把项目历史持久化到磁盘、不做跨进程恢复，也不改变底层 PSTX 解析和规则分析逻辑。

## Solution

复用现有 `RUN_CACHE` 作为会话级项目库，新增项目摘要构建函数和 `/api/projects` 列表接口。新增 `/api/compare` 接口，接收两个 `run_id`，返回指标变化、元件变化、网络变化和结果表变化。前端在首页和报告页渲染“项目管理 / 对比”面板，用户可选择两个项目并查看差异摘要和明细表。

## Impact

后端只新增 Web 层数据整理，不影响 `analyze_project_contents()` 的输出。前端新增可复用项目管理面板，继续保持 localhost-only 运行模式。

## Risks

对比结果基于当前会话缓存，服务重启后历史项目消失。大项目之间逐项比较可能返回较多差异，因此接口会对明细行做上限截断并保留总数。

## Verification Plan

补充 Web 单测覆盖项目列表接口、对比接口和页面静态入口。运行 `python -m unittest discover -s tests -p test_pstx_web.py -v`、全量单测、`python -m compileall -q .` 和 `node --check web/static/app.js`。
