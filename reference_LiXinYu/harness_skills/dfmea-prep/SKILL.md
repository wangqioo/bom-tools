---
name: dfmea-prep
title: DFMEA 准备取证
description: 用于 DFMEA 准备度、元件身份、飞书物料、datasheet 证据缺口和用户补充问题。
triggers: [dfmea, 失效, 失效模式, 规格书, 芯片类别, 测试方案]
capability_profiles: [dfmea_prep, datasheet_qa]
playbooks: [dfmea_preparation, datasheet_pdf_qa, feishu_material_qa]
allowed_tools: [summarize_dfmea_readiness, batch_get_component_identity_cards, batch_match_component_datasheets, batch_search_datasheet_chunks, batch_search_feishu_cache_rows, get_datasheet_chunk]
output_rules: [只输出准备度和证据缺口，不生成正式 DFMEA 表, 缺规格或关键网络时优先结构化追问]
---

## Instructions

先用身份卡、飞书缓存和 datasheet chunk 取证。若证据不足，不要硬猜失效模式；返回 `needs_user_input`，问题要具体到位号、字段和缺失原因。

定量参数、极限值、推荐工作条件必须读取 detail chunk/page 后再回答。
