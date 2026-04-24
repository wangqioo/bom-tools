# 子模块映射真实页修复

- Date: 2026-04-24
- Complexity: L1
- Status: final

## Background

当前 `module_order` 解析和匹配逻辑存在两处偏差：

- 文件发现只覆盖 `module_order`，未覆盖用户当前工程里的 `module_order.dat`
- lookup key 现状优先基于 `P_PATH` 构造，但用户给出的最新样例说明 `module_order` 的路径 key 应与 `C_PATH / SECTION_NUMBER` 的逻辑路径对齐

这会导致子模块映射主模块真实页在部分工程中完全失配，表现为 `page_submodule_mapped` 为空或匹配到错误条目。

## Goal

- 支持读取 `module_order.dat`
- 将 `module_order` key 的主匹配逻辑改为优先使用逻辑路径
- 保留子模块本地真实页参与 `start_real_page + local_page - 1` 偏移计算
- 用用户提供的 `PEX90144_CBB_V1` 样例固化测试

## Non-goals

- 不重做顶层真实页 `P_PATH / page.map / page*.csv` 的决策优先级
- 不在本轮扩展到新的页码数据源
- 不重构 Web UI 展示结构

## Solution

- 在 `pstx_page_logic.py` 中扩展 `build_module_order_index()`，同时扫描 `module_order.dat` 与 `module_order`
- 调整 `build_module_order_lookup_candidates()`：
  - 先基于逻辑路径构造 key
  - 再以真实路径作为保守回退
- 保持 `page_submodule_real` 仍然来自 `P_PATH` 的子模块本地真实页
- 修正测试样例里旧的 `module_order` 假设，使其符合“逻辑路径命中、真实页偏移”的现状

## Impact

- `page_submodule_mapped`
- `module_order_key`
- `module_order_state`
- 依赖这些字段的 DRC / 查询 / Web 报表

## Risks

- 逻辑路径优先后，若存在同一实例同时生成逻辑 key 与真实 key，需要确认优先级是否会遮蔽旧工程的特殊格式
- 新增 `.dat` 扫描后，若两个文件同时存在且内容不一致，可能引入冲突条目

## Verification Plan

- 运行 `tests/test_pstx_analyzer.py` 中页码模型相关用例
- 增加 `module_order.dat` 文件发现测试
- 增加用户给出的 `PEX90144_CBB_V1` 样例测试
- 跑 `python -m unittest discover -s tests -v`
