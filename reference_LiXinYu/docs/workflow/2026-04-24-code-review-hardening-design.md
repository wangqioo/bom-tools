# 全代码 Review 硬化修复

- Date: 2026-04-24
- Complexity: L1
- Status: final

## Background

本轮按现有实现做全代码 review，重点不是新增业务能力，而是找出会造成静默误判或难以定位问题的实现缺口。当前最需要优先处理的是页码映射输入源重复和 Web 本地文件读取编码过窄。

## Goal

修复 `module_order.dat` 与 `module_order` 同时存在时的重复映射误判，并让 Web 输入层按多编码安全读取 PSTX 文本文件，避免 GBK/GB18030 数据被 UTF-8 replacement 静默破坏。

## Non-goals

不重写页码模型，不调整现有 DRC/电阻/电容规则，不改变 Web UI 信息架构。

## Solution

`pstx_page_logic.build_module_order_index()` 对等价的 module_order 行做去重，只保留真正不同的同 key 映射作为歧义。`pstx_web` 增加共享字节解码入口，对本地文件和上传文件统一尝试 UTF/GB 系列编码，并用 PSTX 关键字和控制字符做轻量评分选择最可信解码。

## Impact

页码解析在双文件并存场景下会保持唯一映射，不再因为相同内容重复出现而失败。Web 分析入口能更稳地读取中文属性和路径字段，输入文件元数据也会记录实际采用的编码。

## Risks

编码评分是启发式，极少数无 PSTX 关键字的短文本可能依赖编码优先级。`module_order` 只去重完全等价的 key/start/count/flag，不会吞掉同 key 但指向不同页段的真实冲突。

## Verification Plan

运行针对页码和 Web 输入的新单测，再运行完整 unittest 套件，确认现有 89 个以上测试和新增测试全部通过。
