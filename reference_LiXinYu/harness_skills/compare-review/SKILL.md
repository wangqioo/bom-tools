---
name: compare-review
title: 项目对比深度取证
description: 用于 A/B 项目差异、关键器件、Pin/Net、BOM/飞书、Cadence 页级语义对比。
triggers: [对比, 差异, A/B, 新增, 删除, pin, net, Cadence, sch_1, page]
capability_profiles: [compare_quick_scan, compare_key_devices, compare_pin_net, compare_bom_feishu, compare_cadence_pages, compare_full_review]
playbooks: [compare_diff_batch_lookup, cadence_page_semantic_compare, compare_bom_feishu_material]
allowed_tools: [batch_query_compare_diff, batch_get_compare_rows, compare_cadence_page_semantics, get_compare_row, get_cadence_page_object, summarize_compare_risks]
output_rules: [不要只读首屏 preview 断言没有差异, 页范围必须按用户看到的页码理解]
---

## Instructions

复合对比问题优先批量查 diff，再针对高风险差异读取 row/object detail。页级问题先 resolve page range，再做 Cadence 语义比对。

最终结论必须引用 compare/cadence evidence；如果只是历史经验或摘要推断，标记为待复核。
