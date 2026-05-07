# Aster Live 生产接入

- Date: 2026-04-25
- Complexity: L2
- Status: final

## Background

当前 Aster 功能只提供 `pstx_aster_mock.py` 本地 mock 摘要。用户要求实现可直接用于真实生产环境的版本，但当前 macOS 开发机不部署生产服务，也不应访问真实内网 Aster。

## Goal

- 实现真实 Aster 后端客户端，支持 ChatFlow 生产调用。
- 保留 Room/Auth 协议实现边界，供需要 accessToken 的生产形态使用。
- 通过环境变量控制 `mock` / `live` 模式，部署环境无需改代码。
- Web API 返回统一摘要结构，前端继续使用现有实验入口。
- 所有 secret/token 只留在后端，不进入前端和日志。

## Non-goals

- 不在当前机器调用真实 Aster 内网地址。
- 不把完整 PSTX 原始文件上传到 Aster。
- 不实现前端直连 Aster。
- 不保存 accessToken 到磁盘。

## Solution

- 新增 `pstx_aster_client.py`：
  - 读取并校验 Aster 环境配置。
  - 实现 Aster ChatFlow HTTP 调用。
  - 实现 Auth/Room 协议能力，Auth 加密依赖可用 AES 后端。
  - 构造脱敏后的报告摘要 payload 和 JSON-only prompt。
  - 将模型 JSON 或纯文本回答标准化为现有摘要结构。
- 新增 `pstx_aster_service.py`：
  - `PSTX_ASTER_MODE=mock` 时调用本地 mock。
  - `PSTX_ASTER_MODE=live` 时调用真实 Aster。
  - 对配置错误、HTTP 错误和模型返回错误给出可显示的 JSON 错误。
- 修改 `pstx_web.py`：
  - `/api/report/<run_id>/aster-summary` 统一走 service。
  - live 失败时返回非 2xx 状态和明确错误字段。

## Impact

- 默认行为仍是 mock，不影响当前本地使用和测试。
- 生产环境通过环境变量开启 live。
- 报告页可以直接请求真实 Aster 生成辅助审查摘要。

## Risks

- 真实 Aster 返回格式可能不稳定，因此需要模型 JSON 解析失败的文本兜底。
- Room/Auth AES 依赖在生产环境必须安装 `cryptography` 或 `pycryptodome`。
- 发送给 Aster 的报告摘要仍可能包含位号、网络名等项目信息，生产部署需按内部安全要求确认。

## Verification Plan

- 本地单元测试用临时 HTTP server 模拟 ChatFlow。
- 测试 live 配置缺失时返回明确错误。
- 测试默认 mock 仍保持兼容。
- 跑 Web 相关测试、全量 unittest、Python 编译和 JS 语法检查。
