# Aster 认证状态展示

- Date: 2026-04-26
- Complexity: L1
- Status: final

## Background

用户希望在前端添加 Aster 相关认证信息。由于 Aster 凭据包含 `ASTER_API_KEY`、`ASTER_APP_SECRET` 和 token 类敏感信息，前端不能展示或保存明文认证信息，只能展示安全的配置状态。

## Goal

- 报告页展示 Aster 当前模式、后端、必需环境变量是否已配置。
- 后端提供安全状态接口，不返回 secret/token/apiKey 原文或片段。
- 继续只连接 Aster wrapper，不新增原生 Dify 直连。
- 文档说明 Dify 只需在 Aster 服务侧配置；本工具只配置 Aster API。

## Non-goals

- 不在前端输入或保存 Aster secret。
- 不接原生 Dify API。
- 不在当前开发机调用真实 Aster。

## Solution

- `pstx_aster_service.py` 新增 `build_aster_status()`，返回脱敏状态。
- `pstx_web.py` 新增 `GET /api/aster/status`。
- 报告页 Aster 面板新增状态容器，前端加载并渲染配置项。
- 文档与 changelist 同步安全边界。

## Impact

- 默认 mock 使用不变。
- 生产用户能在页面看到 live 是否缺少关键环境变量。
- 不改变 `/api/report/<run_id>/aster-summary` 的调用方式。

## Risks

- 状态接口必须避免输出 secret 值，因此只返回 configured/missing 布尔状态和变量名。
- 页面错误提示不能诱导用户把 secret 填入浏览器。

## Verification Plan

- 单元测试覆盖 mock/live 状态、缺失变量和不泄露 secret。
- Web 测试覆盖 `/api/aster/status` 和前端资源。
- 跑全量 unittest、Python 编译、JS 语法检查和 diff 检查。
