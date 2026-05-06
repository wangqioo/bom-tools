---
name: topology-review
title: 芯片级拓扑 Review
description: 用于大芯片、连接器、电平转换、时钟/reset/I2C/SPI/PCIe 等芯片级连接关系和 review 任务。
triggers: [拓扑, 网表, 芯片级, 连接关系, 电平转换, level shifter, PCIE, PCE, P5E, clock, reset]
capability_profiles: [chip_topology, full_review]
playbooks: [chip_level_topology]
allowed_tools: [summarize_llm_topology_netlist, query_llm_topology_netlist, batch_query_llm_topology_netlist, summarize_topology_review_tasks, get_llm_topology_node, get_llm_topology_edge]
output_rules: [拓扑边默认无方向, 不把 RCL 当主节点, 高风险连接要回拉 edge/node detail]
---

## Instructions

把 LLM 拓扑网表当作“人工 review 定位索引”，不是电气仿真网表。回答时优先讲芯片/连接器之间的关系、接口类型、页码和 review hints。

若用户问某颗芯片关联对象，优先 `query_llm_topology_netlist` 或批量查询；若要下结论，继续读取 node/edge detail。
