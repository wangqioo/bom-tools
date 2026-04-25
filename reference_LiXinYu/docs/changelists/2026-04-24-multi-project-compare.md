# Web 多项目管理与两两对比

## 变更摘要

- 新增会话内项目库接口 `/api/projects`，列出当前 localhost 服务中最近分析过的项目。
- 新增两两对比接口 `/api/compare`，支持通过两个 `run_id` 对比项目差异。
- 对比范围覆盖指标变化、元件变化、网络节点变化，以及 BOM、网络、DRC、电阻、电容降额等报告结果表差异。
- 首页和报告页新增“项目管理 / 对比”面板，可直接选择两个项目查看差异。
- 当前实现复用内存中的 `RUN_CACHE`，不引入数据库，也不跨服务重启持久化。

## 差异类型

- `新增`：右侧项目存在、左侧项目不存在。
- `删除`：左侧项目存在、右侧项目不存在。
- `变化`：两侧对象同名但关键字段或行内容不同。

## 验证

- `python -m unittest discover -s tests -p test_pstx_web.py -v`
- `python -m unittest discover -s tests -v`
- `python -m compileall -q .`
- `node --check web/static/app.js`
