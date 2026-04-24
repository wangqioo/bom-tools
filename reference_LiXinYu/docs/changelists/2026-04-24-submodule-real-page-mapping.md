# 2026-04-24 子模块映射真实页修复

## 变更摘要

- `pstx_page_logic.py`
  - 新增 `module_order.dat` 文件发现
  - `module_order` lookup key 改为优先按逻辑路径构造
  - 子模块页偏移改为优先使用 `P_PATH` 里的子模块本地真实页
- `tests/test_pstx_analyzer.py`
  - 修正旧测试里“`module_order` key 走 `P_PATH`”的错误假设
  - 新增 `PEX90144_CBB_V1` 样例
  - 新增 `module_order.dat` 发现与映射验证

## 修复点

1. 修复 `module_order.dat` 未被读取的问题
2. 修复 `module_order` key 错误优先使用 `P_PATH` 的问题
3. 修复子模块偏移页使用候选 key 页号而非子模块本地真实页的问题

## 验证

- `python -m py_compile E:\codex_use\dehdl_review\pstx_page_logic.py E:\codex_use\dehdl_review\pstx_analyzer.py`
- `python -m unittest discover -s tests -p test_pstx_analyzer.py -v`
- `python -m unittest discover -s tests -v`

## 备注

- 当前目录不是 git 仓库，无法在本地执行 `pull`、`commit` 或 `push`
- 远端同步需要在恢复 `.git` 元数据后再执行
