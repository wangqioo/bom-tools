# 2026-04-24 规则逻辑 Review 硬化修复

## 变更摘要

- 修复电阻值解析对 `OHM/OHMS/KOHM` 等文本单位支持不完整的问题，避免串阻/上下拉阻值被误判为缺失。
- 修复多级串阻搜索使用全局 visited 导致的多路径漏判问题。现在同一信号通过多条独立串阻链路到达同一远端上下拉时，会保留每条独立路径。
- 新增菱形多路径串阻拓扑测试，覆盖隔串阻上拉和串阻分压风险结果。

## Bug 与修复手段

- Bug: `_parse_ohms()` 先替换 `OHM` 再替换 `OHMS`，`10OHMS` 会变成无法解析的 `10RS`；`10KOHM` 也无法正确折算。
- 修复: 统一处理 `OHMS?/Ω/欧` 后缀，并把 `K/M/G + R` 归一为 K/M/G 单位。
- Bug: `_walk_series_paths()` 使用全局 visited，菱形拓扑中第二条到同一远端节点的有效串阻路径会被剪掉。
- 修复: 改为路径本地防环，按路径签名去重，并增加最大跳数和最大结果数保护。

## 验证

- `python -m unittest discover -s tests -p test_pstx_analyzer.py -v`
- `python -m unittest discover -s tests -v`
- `python -m compileall -q .`
