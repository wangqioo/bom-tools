# Aster Live 生产接入

## 背景

在 mock-only 摘要基础上，新增可直接用于真实生产环境的 Aster 后端调用能力。当前开发机不访问真实内网 Aster，生产部署通过环境变量启用 live 模式。

## 变更

- 新增 `pstx_aster_client.py`：
  - 支持真实 Aster ChatFlow HTTP 调用。
  - 保留 Room/Auth 客户端和 AES-256-ECB + PKCS7 + 双 Base64 协议实现。
  - 构造裁剪后的报告摘要 prompt，并把模型 JSON 或纯文本回答标准化为前端摘要结构。
- 新增 `pstx_aster_service.py`：
  - 支持 `PSTX_ASTER_MODE=mock|live|off`。
  - live 失败时返回 Web 可显示错误，不泄露 secret/token。
- Web 报告页文案从 mock-only 改为 Aster 摘要入口，按钮根据后端配置生成 mock 或 live 摘要。
- README、ARCHITECTURE、AGENTS 和 Aster 解析文档补充生产部署说明。
- 新增客户端和 service 单元测试。

## 生产配置

ChatFlow 生产模式：

```bash
export PSTX_ASTER_MODE=live
export PSTX_ASTER_BACKEND=chat-flow
export ASTER_BASE_URL="https://aigc.huaqin.com"
export ASTER_EMP_NO="100019100"
export ASTER_API_KEY="flow_api_key_xxx"
python pstx_web.py
```

Room/Auth 模式额外需要：

```bash
export PSTX_ASTER_BACKEND=room
export ASTER_APP_ID="ag_xxx"
export ASTER_APP_SECRET="***"
python -m pip install cryptography
```

## 安全边界

- 前端不读取、不渲染、不保存 `ASTER_APP_SECRET`、`ASTER_API_KEY`、`accessToken`。
- 默认只发送报告摘要和少量表格样例行，不上传原始 PSTX 文件。
- `PSTX_ASTER_MODE=mock` 可作为生产回退。
- 当前开发机未调用真实 Aster 内网地址。

## 验证

- `python3 -m unittest tests.test_pstx_aster_client -v`
- `python3 -m unittest tests.test_pstx_aster_service -v`
- `python3 -m unittest tests.test_pstx_web.WebUiTests.test_aster_live_mode_missing_config_reports_displayable_error -v`
- `python3 -m unittest discover -s tests -v`
- `python3 -m compileall -q .`
- `node --check web/static/app.js`
- `git diff --check`
