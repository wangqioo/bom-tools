# 规则逻辑 Review 硬化修复

- Date: 2026-04-24
- Complexity: L1
- Status: final

## Background

上一轮 review 主要处理代码实现稳定性。本轮转向业务规则可信度，重点检查“能正常运行但会把规则判断漏掉或降级为无法判断”的逻辑点。

## Goal

修复电阻相关规则中的两个确定性逻辑漏洞：电阻值单位解析漏掉常见 `OHM/OHMS/KOHM` 写法，以及多级串阻搜索在多路径场景下只保留第一条到达路径。

## Non-goals

不改变电压网络保守识别原则，不把 OD/OC、AC 耦合、电容降额等候选规则升级为确定结论，不扩大 Web UI 展示结构。

## Solution

增强 `_parse_ohms()` 的单位归一化，先处理 `OHM/OHMS/Ω/欧` 后缀，再把 `KOHM/MOHM/GOHM` 等折算成原有 K/M/G 单位。调整 `_walk_series_paths()`，把全局 visited 改为路径本地防环，并增加最大跳数和结果数保护，允许发现同一远端上下拉的多条独立串阻路径。

## Impact

串阻分压风险和芯片 pin 隔串阻上下拉状态会覆盖更多真实拓扑；常见带 `OHM` 文本单位的电阻不再被当成阻值缺失。

## Risks

路径本地搜索比全局 visited 覆盖更全，在复杂电阻网中结果数量可能增加。已增加跳数和结果数上限，避免异常网表导致搜索爆炸。

## Verification Plan

新增电阻值单位解析测试和菱形多路径串阻测试，随后运行完整 unittest 和 compileall。
