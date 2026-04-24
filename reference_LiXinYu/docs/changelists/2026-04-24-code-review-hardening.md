# 2026-04-24 全代码 Review 硬化修复

## 变更摘要

- 修复 `module_order.dat` 与 `module_order` 同时存在且内容相同时，等价映射被重复计入并导致子模块页码误判为 `ambiguous` 的问题。
- Web 输入读取新增统一字节解码逻辑，支持 UTF-8/UTF-16/GB18030/CP936 等常见编码，并用 PSTX 关键字进行轻量评分，降低中文属性和层级路径被 replacement 静默破坏的风险。
- 本地文件和上传文件的输入元数据新增 `encoding` 字段，便于后续排查输入文件质量。

## Bug 与修复手段

- Bug: 双入口 `module_order.dat` / `module_order` 可能保存同一份 module_order 内容，旧逻辑把同 key 的完全相同记录视为多条候选，最终 `_resolve_module_order_entry()` 返回 `ambiguous`。
- 修复: 在 `build_module_order_index()` 中按 `path_key + start_real_page + page_count + flag` 去重，只保留真正不同的同 key 映射作为歧义。
- Bug: Web 读取 PSTX 文本只使用 `utf-8(errors='replace')`，GBK/GB18030 文件不会报错但会破坏中文、路径或属性值。
- 修复: 新增 `_decode_text_bytes()`，多编码尝试并按 PSTX 关键字和控制字符评分选择解码结果。

## 验证

- `python -m unittest discover -s tests -p test_pstx_analyzer.py -v`
- `python -m unittest discover -s tests -p test_pstx_web.py -v`
- `python -m unittest discover -s tests -v`
