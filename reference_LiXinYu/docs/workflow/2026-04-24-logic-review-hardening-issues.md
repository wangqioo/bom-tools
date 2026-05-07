# 规则逻辑 Review 硬化修复 Issues

- Date: 2026-04-24
- Complexity: L1
- Related design: `docs/workflow/2026-04-24-logic-review-hardening-design.md`

## Task Overview

- Goal: 针对规则逻辑 review 中发现的确定性漏判修复并补测试。
- Ordering rule: Complete issues in sequence.
- Current status: done

## Issue List

- [x] issue-1 修复电阻值 OHM/OHMS 单位解析
- [x] issue-2 修复多级串阻多路径漏判

## issue-1

- ID: issue-1
- 标题: 修复电阻值 OHM/OHMS 单位解析
- 范围: `_parse_ohms()` 和电阻值解析回归测试
- 依赖: none
- 验收标准: `10OHM`、`10OHMS`、`10KOHM`、`4.7KΩ` 均可解析为正确欧姆值。
- 状态: done
- 验证方式: `python -m unittest discover -s tests -p test_pstx_analyzer.py -v`
- commit: fix(issue-1,issue-2): harden resistor logic review

## issue-2

- ID: issue-2
- 标题: 修复多级串阻多路径漏判
- 范围: `_walk_series_paths()`、隔串阻上下拉和分压风险回归测试
- 依赖: issue-1
- 验收标准: 同一芯片 pin 通过两条独立多级串阻路径连接到同一个远端上拉时，两条路径都出现在隔串阻上拉和分压风险中。
- 状态: done
- 验证方式: `python -m unittest discover -s tests -p test_pstx_analyzer.py -v`
- commit: fix(issue-1,issue-2): harden resistor logic review

## Self Review Checklist

- [x] Requirement alignment: 修复的是规则层漏判，不是代码风格。
- [x] Regression risk: 串阻搜索保留路径本地防环，并设置跳数/结果上限。
- [x] Test coverage: 新增单位解析和多路径拓扑回归测试。
- [x] Dirty changes: 未改动 Web UI 或页码模型。
- [x] Docs sync: 已新增 workflow 与 changelist。
