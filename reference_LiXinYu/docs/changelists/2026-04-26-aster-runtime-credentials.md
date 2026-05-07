# Aster 前端临时凭据覆盖

## 背景

用户希望前端可以接触并设置 Aster 的 API Key / Secret。为降低泄露风险，本轮不把 secret 写入 JS、HTML、localStorage 或磁盘，而是增加“当前进程内存覆盖”能力。

## 变更

- 新增后端 runtime override：
  - `POST /api/aster/runtime-config`
  - `DELETE /api/aster/runtime-config`
- 可临时设置：
  - `PSTX_ASTER_MODE`
  - `PSTX_ASTER_BACKEND`
  - `ASTER_BASE_URL`
  - `ASTER_EMP_NO`
  - `ASTER_API_KEY`
  - `ASTER_APP_ID`
  - `ASTER_APP_SECRET`
- 报告页 Aster 面板新增“临时 Aster 凭据”表单。
- 提交后清空 `api_key` / `app_secret` 密码框。
- `/api/aster/status` 显示 runtime 覆盖是否启用和覆盖项名称，但不回显 secret。

## 安全边界

- 不支持浏览器长期保存 secret。
- 不支持前端手动传入 accessToken；Room/Auth 仍由后端用 App Secret 换取 token。
- Runtime 覆盖项只保存在当前 Python 进程内存，重启服务后消失。
- 长期生产部署仍推荐使用环境变量。

## 验证

- `python3 -m unittest tests.test_pstx_aster_service tests.test_pstx_web.WebUiTests.test_aster_runtime_config_can_set_and_clear_without_echoing_secret -v`
- `python3 -m unittest discover -s tests -v`
- `python3 -m compileall -q .`
- `node --check web/static/app.js`
- `git diff --check`
