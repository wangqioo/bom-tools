# 电容降额 50V 低压直通 Issues

- Date: 2026-04-24
- Complexity: L0
- Related design: none

## Task Overview

- Goal: 在降额检查中新增“全局最大已识别电压不超过 12V 时，额定耐压至少 50V 的电容直接通过”补丁逻辑。
- Ordering rule: Complete issues in sequence.
- Current status: done

## Design Note

> 这是一个局部规则补丁。实现上先全局扫描可识别的最大正电压，再在单颗电容判定前做快速放行，不改动现有 `P_PATH / page.map / page*.csv` 页码逻辑，也不改动原有单电容电压推断主流程。

## Issue List

- [x] issue-1 新增全局最大电压扫描与 50V 低压直通规则
- [x] issue-2 补测试并验证
- [x] issue-3 同步变更记录

## issue-1

- ID: issue-1
- 标题: 新增全局最大电压扫描与 50V 低压直通规则
- 范围: `pstx_analyzer.py`
- 依赖: none
- 验收标准:
  - 先扫描整板可识别的最大正电压
  - 当最大电压 `<=12V` 且电容额定耐压 `>=50V` 时直接判通过
  - 超过 `12V` 时不触发该规则
- 状态: done
- 验证方式: 语法检查 + 单元测试
- commit: not possible - not a git repository

## issue-2

- ID: issue-2
- 标题: 补测试并验证
- 范围: `tests/test_pstx_analyzer.py`
- 依赖: issue-1
- 验收标准:
  - 覆盖 `<=12V` 直通场景
  - 覆盖 `>12V` 不直通场景
- 状态: done
- 验证方式: `python -m unittest discover -s tests -p test_pstx_analyzer.py -v`
- commit: not possible - not a git repository

## issue-3

- ID: issue-3
- 标题: 同步变更记录
- 范围: `docs/changelists/`、`docs/ARCHITECTURE.md`、`AGENTS.md`
- 依赖: issue-2
- 验收标准:
  - changelist 已记录
  - 架构说明补充该规则
  - AGENTS 索引更新
- 状态: done
- 验证方式: 人工复核
- commit: not possible - not a git repository
