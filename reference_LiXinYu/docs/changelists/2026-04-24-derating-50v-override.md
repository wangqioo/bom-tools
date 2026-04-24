# 2026-04-24 电容降额 50V 低压直通

## 变更摘要

- `pstx_analyzer.py`
  - 新增全局最大已识别电压扫描
  - 当整板最大已识别正电压 `<=12V` 且电容额定耐压 `>=50V` 时，直接判定为通过
- `tests/test_pstx_analyzer.py`
  - 新增低压直通场景测试
  - 新增超过 12V 不触发直通场景测试

## 规则说明

- 这是一个补丁性质的快速放行规则
- 只在“全局最大已识别正电压明确且不超过 12V”时触发
- 没有识别出全局最大电压，或全局最大电压大于 12V 时，仍走原有单电容降额判断

## 验证

- `python -m py_compile E:\codex_use\dehdl_review\pstx_analyzer.py E:\codex_use\dehdl_review\tests\test_pstx_analyzer.py`
- `python -m unittest discover -s tests -p test_pstx_analyzer.py -v`
