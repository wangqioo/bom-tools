---
name: datasheet-key-info
title: Datasheet 关键信息取证
description: 用于基于 MinerU 索引的 datasheet 关键信息抽取、参数核验、64144 类复杂芯片阅读和原理图 evidence 映射。
triggers:
  - datasheet
  - MinerU
  - 规格书
  - 手册
  - 64144
  - 关键参数
  - 参数卡
  - recommended operating
  - absolute maximum
  - electrical characteristics
  - power sequence
  - thermal
capability_profiles: [datasheet_qa, dfmea_prep, compare_datasheet_qa]
playbooks: [datasheet_pdf_qa, compare_datasheet_pdf_qa, dfmea_preparation]
allowed_tools:
  - list_datasheet_sources
  - list_datasheet_review_templates
  - get_datasheet_review_template
  - list_datasheet_documents
  - search_datasheet_chunks
  - batch_search_datasheet_chunks
  - search_datasheet_parameters
  - get_datasheet_parameter
  - get_datasheet_chunk
  - get_datasheet_page_excerpt
  - search_datasheets
  - get_datasheet_excerpt
  - get_evidence_pack
  - batch_query_report_entities
  - summarize_llm_topology_netlist
  - query_llm_topology_netlist
  - batch_query_llm_topology_netlist
  - list_compare_sections
  - query_compare_diff
  - batch_query_compare_diff
  - get_compare_row
output_rules:
  - 默认承认 MinerU 是主抽取路径
  - search/snippet 只能当 locator
  - 定量结论必须读取 parameter 或 chunk detail
  - 每个 datasheet 事实都要映射到原理图或标记缺口
  - 不把未知项折叠成 pass/fail
---

## Purpose

用这张 skill 处理 datasheet 关键信息问题，尤其是 64144 这类大芯片/复杂 SoC/多电源域器件。目标不是“读完整 PDF 后写摘要”，而是快速定位对原理图 review 真正有用的证据：电源、时序、接口电平、复位/strap、热和绝对极限，并把这些事实映射回原理图。

## Output Shape

回答时固定拆成三块：

1. `已确认的 datasheet 事实`：列参数名、值/范围、条件、doc/page/chunk/parameter locator。
2. `已映射的原理图 evidence`：列 refdes、pin、net、rail、页码、拓扑边或 compare diff。
3. `缺失或人工复核项`：列缺失字段、需要继续调用的工具、为什么不能下结论。

## Hard Rules

- 不从 `datasheet-search` snippet 直接回答电压/电流/温度/时序结论。
- 不混淆 recommended operating 与 absolute maximum。
- 不把型号相似、HQ 相似、封装相似的多个 PDF 合并；身份未确认时保持分开。
- 不把 `needs_manual_review` 的 PDF 当作 MinerU 已可靠解析。
- 不生成正式 DFMEA 表；DFMEA 只输出准备度、缺口和需要补充的证据。

## Evidence Order

1. 先确认索引状态。默认认为 PDF 通过 MinerU 抽取；如果状态显示未配置、MinerU 失败、`needs_manual_review` 或 chunk/parameter 为空，先报告证据不可用，不要把 fallback/snippet 当可靠结论。
2. 先读模板。复杂芯片优先 `complex_chip`；电源芯片、level shifter 等按器件类别选模板。模板决定回答结构和必须检查的原理图 evidence。
3. 先拿参数卡。电压、电流、温度、时序、接口阈值、绝对最大值、推荐工作条件，优先用 `search_datasheet_parameters` / `datasheet-parameters` 定位。
4. 再读 detail。任何会影响设计判断的数值结论，都必须继续读取 `get_datasheet_parameter`、`get_datasheet_chunk` 或 `get_datasheet_page_excerpt`。搜索结果和 chunk snippet 只是 locator。
5. 最后映射原理图。用报告 evidence、拓扑、页级 Cadence 语义或 Compare 差异，把 datasheet 事实回连到具体 refdes、pin、net、rail、page、接口边或差异项。

## 64144-Style Complex Chip Checklist

优先检索这些 evidence group，再组织结论：

- `recommended operating conditions`：推荐工作电压、电源域范围、允许工作温度。
- `absolute maximum ratings`：绝对极限，明确只能作为风险边界，不能当推荐值。
- `power rail voltage`：各 rail 名称、电压范围、容差、模拟/数字/PLL/IO 域。
- `power consumption current`：典型/最大电流、模式条件、是否按 rail 分列。
- `power up sequence` / `power down sequence`：上电/掉电顺序、延时、PGOOD/EN/RESET 条件。
- `reset timing`：RESET 输入/输出、POR、最小脉宽、释放时序。
- `pin description voltage domain`：pin 所属电源域、方向、复用功能、默认状态。
- `IO threshold`：输入高低阈值、输出电平、OD/OC/open-drain 如有必须引用原文。
- `clock requirements`：时钟频率、精度、抖动、启动条件。
- `strap boot mode`：strap 默认电平、采样时刻、上/下拉建议。
- `thermal characteristics` / `junction temperature`：热阻、Tj/Tcase、功耗条件和散热边界。

## Mapping Back To Schematic

- 电源要求必须映射到 rail net、PMIC/VR 输出、滤波/电容、使能和 PGOOD/RESET 关系。
- 时序要求必须映射到 EN/PGOOD/RESET/CLK pin-net evidence；如果原理图缺少对应证据，明确列为 missing。
- 接口电平必须映射接口两端器件电源域；跨电平必须检查 level shifter 或兼容阈值证据。
- 绝对最大值只用于“不应超过”的边界判断，不替代 recommended operating。
- Compare 场景要先定位 A/B 差异中的 HQ、型号、pin/net 或页码，再读 datasheet evidence；不要只凭一个项目的 PDF 推断另一个项目。
