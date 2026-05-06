---
name: feishu-material
title: 飞书物料缓存问答
description: 用于 HQ 料号、规格型号、PI、选型顺序、飞书缓存来源库和 Sheet 取证。
triggers: [飞书, HQ料号, HQ编码, 物料编码, 规格型号, PI, 选型顺序]
capability_profiles: [feishu_bom_qa, dfmea_prep]
playbooks: [feishu_material_qa, compare_bom_feishu_material]
allowed_tools: [list_feishu_cache_libraries, search_feishu_cache_rows, batch_search_feishu_cache_rows, get_feishu_cache_row]
output_rules: [没有缓存命中时说明无命中, 不凭经验补全物料字段]
---

## Instructions

回答物料问题必须引用本地飞书缓存 evidence。多个 HQ/型号/PI 关键词优先用批量检索。若缓存为空或无命中，给出建议关键词和需要用户补充的信息。
