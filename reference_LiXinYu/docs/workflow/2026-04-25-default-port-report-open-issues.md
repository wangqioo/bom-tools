# 2026-04-25 默认端口与单项目打开修复

## Design Note

本轮为 L0 修复：将 localhost Web / 本地 UI 默认端口从 `8765` 切换到 `44441`，规避当前开发机上 `8765` 被其他服务占用导致 `/debug/report-open` 打到错误服务的问题；同时为单项目打开动效页补上服务端初始状态，避免 JS 加载前处于未定义阶段。

## Issues

### Issue 1

- ID: default-port-report-open-1
- 标题: 切换默认端口并修复单项目打开初始状态
- 范围: Web 入口、本地 UI 入口、单项目打开 debug 模板、README、测试
- 依赖: 无
- 验收标准: 默认端口为 `44441`；`/debug/report-open` 可通过新端口访问；页面初始 `data-phase` 明确为 `pick`
- 状态: completed
- 验证方式: `python3 -m unittest discover -s tests -p test_pstx_web.py -v`；`python3 -m unittest discover -s tests -p test_pstx_local_ui.py -v`；实际启动到 `127.0.0.1:44441`
- commit: pending

## Self Review

- [x] 未修改解析/规则分析业务逻辑
- [x] 默认端口在 Web、本地 UI、README、测试中保持一致
- [x] 单项目打开页保留原有动效，只增加初始状态兜底
