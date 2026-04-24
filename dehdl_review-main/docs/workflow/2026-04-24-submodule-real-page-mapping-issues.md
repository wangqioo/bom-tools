# 子模块映射真实页修复 Issues

- Date: 2026-04-24
- Complexity: L1
- Related design: `docs/workflow/2026-04-24-submodule-real-page-mapping-design.md`

## Task Overview

- Goal: 修复 `module_order(.dat)` 驱动的子模块映射主模块真实页逻辑，使其与 `C_PATH / SECTION_NUMBER` 的逻辑路径对齐，并保留子模块本地真实页偏移。
- Ordering rule: Complete issues in sequence.
- Current status: done

## Issue List

- [x] issue-1 修正 `module_order(.dat)` 文件发现与 lookup key 逻辑
- [x] issue-2 补齐样例测试并回归验证
- [x] issue-3 同步文档与变更记录

## issue-1

- ID: issue-1
- 标题: 修正 `module_order(.dat)` 文件发现与 lookup key 逻辑
- 范围: `pstx_page_logic.py`
- 依赖: none
- 验收标准:
  - 能读取 `module_order.dat`
  - `module_order` key 优先按逻辑路径匹配
  - 子模块本地真实页仍参与偏移计算
- 状态: done
- 验证方式: 新增/更新单元测试并人工对照用户样例
- commit: not possible - not a git repository

## issue-2

- ID: issue-2
- 标题: 补齐样例测试并回归验证
- 范围: `tests/test_pstx_analyzer.py`
- 依赖: issue-1
- 验收标准:
  - 覆盖 `module_order.dat` 发现
  - 覆盖 `PEX90144_CBB_V1` 样例
  - 更新旧测试中对 `P_PATH` key 的错误假设
- 状态: done
- 验证方式: `python -m unittest discover -s tests -v`
- commit: not possible - not a git repository

## issue-3

- ID: issue-3
- 标题: 同步文档与变更记录
- 范围: `docs/ARCHITECTURE.md`、`docs/README.md`、`AGENTS.md`、`docs/changelists/`
- 依赖: issue-2
- 验收标准:
  - 架构文档更新为“逻辑路径优先命中 module_order”
  - changelist 记录本轮修复内容
  - 入口索引不再与实际目录冲突
- 状态: done
- 验证方式: 人工复核文档路径与内容
- commit: not possible - not a git repository
