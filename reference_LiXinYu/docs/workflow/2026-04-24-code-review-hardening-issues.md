# 全代码 Review 硬化修复 Issues

- Date: 2026-04-24
- Complexity: L1
- Related design: `docs/workflow/2026-04-24-code-review-hardening-design.md`

## Task Overview

- Goal: 修复 review 中发现的静默误判风险，并补齐回归测试。
- Ordering rule: Complete issues in sequence.
- Current status: done

## Issue List

- [x] issue-1 去重 module_order 双文件重复映射
- [x] issue-2 强化 Web 文本输入解码

## issue-1

- ID: issue-1
- 标题: 去重 module_order 双文件重复映射
- 范围: `pstx_page_logic.build_module_order_index()` 和页码回归测试
- 依赖: none
- 验收标准: `module_order.dat` 与 `module_order` 内容相同时，子模块映射仍为 `unique`，不会被误判为 ambiguous。
- 状态: done
- 验证方式: `python -m unittest discover -s tests -p test_pstx_analyzer.py -v`
- commit: 392665d

## issue-2

- ID: issue-2
- 标题: 强化 Web 文本输入解码
- 范围: `pstx_web.py` 本地文件读取、上传文件读取和 Web 回归测试
- 依赖: issue-1
- 验收标准: GB18030 编码的 PSTX 文本可以被正确读取，元数据记录采用编码，不再依赖 UTF-8 replacement。
- 状态: done
- 验证方式: `python -m unittest discover -s tests -p test_pstx_web.py -v`
- commit: fix(issue-2): harden web input decoding

## Self Review Checklist

- [x] Requirement alignment: 修复范围对应 review 发现的真实风险点。
- [x] Regression risk: 双文件重复只去重等价映射，不吞掉真实冲突。
- [x] Test coverage: 新增页码和 Web 输入回归测试。
- [x] Dirty changes: 未触碰无关业务规则。
- [x] Docs sync: 已新增 workflow 与 changelist。
