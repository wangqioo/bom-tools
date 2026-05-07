# Aster 前端临时凭据覆盖

- Date: 2026-04-26
- Complexity: L2
- Status: final

## Background

用户希望前端可以接触并设置 Aster 的 `secret/token/apiKey`。长期把 secret 写入前端 JS/HTML/localStorage 风险很高，因此采用“前端临时输入、后端内存覆盖”的折中方案。

## Goal

- 报告页提供 Aster 临时凭据设置表单。
- 支持设置当前 Python 进程内的 `PSTX_ASTER_MODE`、`PSTX_ASTER_BACKEND`、`ASTER_BASE_URL`、`ASTER_EMP_NO`、`ASTER_API_KEY`、`ASTER_APP_ID`、`ASTER_APP_SECRET`。
- 提交后状态接口立即反映覆盖项。
- 不把 secret 明文返回给前端，不写入磁盘，不写入 localStorage。
- 支持一键清除当前进程覆盖项。

## Non-goals

- 不支持浏览器长期保存 secret。
- 不支持前端手动注入 accessToken；Room/Auth 仍由后端用 App Secret 换取 token。
- 不新增原生 Dify 直连。

## Solution

- `pstx_aster_service.py` 增加线程安全的 runtime override 存储。
- `pstx_web.py` 增加：
  - `POST /api/aster/runtime-config`
  - `DELETE /api/aster/runtime-config`
- 报告页 Aster 面板增加临时凭据表单。
- 前端提交后清空密码字段并刷新认证状态。

## Impact

- 环境变量仍是生产推荐配置。
- 前端临时设置仅用于本地/当前进程调试或临时生产会话。
- 重启服务后覆盖项消失。

## Risks

- 浏览器输入 secret 仍存在肩窥、浏览器插件、调试工具等风险。
- 如果部署不是 localhost，应优先禁用该表单或只在受控内网使用。

## Verification Plan

- 测试 runtime override 能覆盖环境变量并参与状态判断。
- 测试 secret 不在 status 响应中回显。
- 测试 Web POST/DELETE 接口。
- 测试前端资源包含提交/清除逻辑。
- 跑全量 unittest、Python 编译、JS 语法检查和 diff 检查。
