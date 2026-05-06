---
name: schematic-datasheet-connection-review
title: 原理图连接 × Datasheet 反查
description: 用于把用户问题拆成原理图/网表 evidence、元件身份、MinerU-backed datasheet detail，再反查电源、接口、复位、时钟和 strap 连接风险。
triggers:
  - datasheet 连接
  - 规格书连接
  - MinerU
  - 反查连接
  - 连接是否有问题
  - 网表证据
  - 接口电平
  - 电源域
  - level shifter
  - reset timing
  - power sequence
capability_profiles: [connection_datasheet_review, chip_topology, datasheet_qa, dfmea_prep]
playbooks: [schematic_datasheet_connection_review, chip_level_topology, datasheet_pdf_qa, dfmea_preparation]
allowed_tools:
  - list_datasheet_sources
  - list_datasheet_review_templates
  - get_datasheet_review_template
  - summarize_llm_topology_netlist
  - summarize_topology_review_tasks
  - batch_query_llm_topology_netlist
  - get_llm_topology_edge
  - get_llm_topology_node
  - batch_get_component_identity_cards
  - get_component_identity_card
  - batch_match_component_datasheets
  - match_component_datasheets
  - search_datasheet_parameters
  - get_datasheet_parameter
  - batch_search_datasheet_chunks
  - get_datasheet_chunk
  - trace_project_source
output_rules:
  - 先定位用户问题中的位号、网络、接口、电源 rail 或页码目标
  - 先取原理图/网表/拓扑 evidence，再取 datasheet detail
  - MinerU search/snippet 只能作为 locator，定量结论必须 detail
  - 用 datasheet fact 回查原理图连接，只输出 evidence-backed risk/gap
---

## Purpose

这张 skill 用于“datasheet 反查原理图连接”的复合审查。Agent 不应把它拆成单纯 PDF 摘要，也不应只读拓扑边后就下结论；正确动作是先拿连接 evidence，再拿元件 datasheet evidence，最后把两者逐项映射。

## Evidence Chain

1. `解读用户问题`
   - 提取 refdes、网络名、rail、接口组、页码、风险关键词。
   - 如果用户只说“这个连接是否有问题”，先用 topology/report entity 工具定位候选对象。

2. `原理图/网表 evidence`
   - 优先 `batch_query_llm_topology_netlist`，必要时 `summarize_llm_topology_netlist` / `summarize_topology_review_tasks`。
   - 对关键边使用 `get_llm_topology_edge`，对关键芯片使用 `get_llm_topology_node`。
   - 需要底层文件时使用 `trace_project_source`，引用 line-number excerpt。

3. `元件身份 evidence`
   - 用 `batch_get_component_identity_cards` 确认 HQ、型号、封装、pin-net、power nets、interface nets。
   - 身份不确定时先列 gap，不要继续把 datasheet 事实硬套到该 refdes。

4. `MinerU-backed datasheet evidence`
   - 先 `list_datasheet_sources` 检查本地 MinerU 索引是否可用。
   - 用 `batch_match_component_datasheets` / `match_component_datasheets` 找到对应 PDF。
   - 定量参数优先 `search_datasheet_parameters` + `get_datasheet_parameter`。
   - 章节事实用 `batch_search_datasheet_chunks` + `get_datasheet_chunk`。

5. `反查连接`
   - 电源 rail：recommended operating、absolute maximum、上电/掉电顺序、EN/PGOOD/RESET 关系。
   - 接口：两端电源域、IO threshold、OD/OC/open-drain、level shifter 或兼容阈值。
   - 时钟/复位/strap：时序、默认电平、采样窗口、上下拉要求。

## Output Shape

回答固定拆为：

1. `用户问题解读`：目标 refdes/net/interface/rail。
2. `原理图连接 evidence`：拓扑边、pin-net、source trace。
3. `Datasheet evidence`：doc/page/chunk/parameter，标注 MinerU-backed。
4. `反查结论`：evidence-backed risk、pass-like observation 或 needs_manual_review。
5. `缺口`：缺 PDF、缺 MinerU 索引、缺 detail、身份不确定或拓扑不够完整。

## Hard Rules

- 不能把 search snippet 当最终 datasheet 事实。
- 不能把 absolute maximum 当 recommended operating。
- 不能把 LLM topology 当完整电气签核网表。
- 不能在 datasheet 未命中时编造型号、电压或接口阈值。
- 不能输出正式 DFMEA 表；这里只做连接反查和复核建议。
