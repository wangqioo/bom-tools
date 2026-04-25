# PSTX 原理图审查工具

解析 Cadence Packager-XL 导出的 `pstxprt.dat` / `pstxnet.dat`，结合 `sch_1` 页码与 CSA 几何对象，生成 BOM、网络、DRC、规范检查、电容降额、电阻规则、芯片 Pin 状态、页码映射、Aster mock 辅助摘要与 Excel 报告。

## 运行入口

首次部署建议安装依赖：

```bash
python -m pip install -r requirements.txt
```

### 本地桌面 UI（保留）

```bash
python pstx_local_ui.py
```

- 默认启动 localhost Web 服务，并尝试用 `pywebview` 套壳显示。
- 如果当前环境没有 `pywebview`，会自动退回系统浏览器，不影响功能。
- 想强制浏览器模式：

```bash
python pstx_local_ui.py --browser
```

### Web UI

```bash
python pstx_web.py
```

- 只监听 `127.0.0.1`。
- 默认端口 `44441`，被占用时自动顺延。
- 支持输入项目根路径、上传 PSTX 文件、分区浏览、查询、导出 Excel。

### 兼容入口

```bash
python pstx_analyzer.py
```

会转到本地桌面 UI 入口。

## 当前核心逻辑

页码解析已统一到 `pstx_page_logic.py`：

- `C_PATH` / `SECTION_NUMBER`：作为逻辑页来源。
- `P_PATH`：优先作为顶层真实页来源。
- `sch_1/page.map`：用于逻辑页和真实页交叉验证。
- `sch_1/page*.csv`：作为另一条真实页映射校验来源。
- `module_order`：用于计算子模块本地页映射到主模块真实页。

对子模块页：

```text
子模块映射主模块真实页 = module_order.start_real_page + 子模块本地真实页 - 1
```

嵌套复用场景会优先匹配最深层可命中的 `module_order` key；如果最深层缺失，再回退到外层 key，避免完全丢失映射。

CSA 几何规范检查集中在 `pstx_csa_geometry.py`：

- 扫描 `<project_root>/sch_1/page*.csa`。
- 检测带 `DOT` 的四向十字交叉；T 型连接和无 DOT 十字不报。
- 检测 `CIRCLE` 画圈对象，并把 `ARC` 拟合圆作为需人工确认的画圈候选。
- 该检查只输出几何候选，不推断网络短接或电气错误。

Aster 摘要链路：

- 默认使用 `pstx_aster_mock.py` 生成本地 mock 摘要。
- 生产模式使用 `pstx_aster_service.py` 切换到 `pstx_aster_client.py` 的真实 Aster 调用。
- 报告页以右下角浮窗形式提供 Aster 审查助手，输出摘要、优先级、审查清单、分区焦点和人工复核边界。
- 前端默认只请求本地 `/api/report/<run_id>/aster-summary` 和 `/api/aster/status`；状态接口不回显 secret/token/apiKey。
- 发送给 Aster 的内容是后端裁剪后的报告摘要和审查上下文，不上传原始 PSTX 文件。
- `/api/aster/status` 只显示环境变量是否已配置，不显示密钥原文或片段。
- 报告页也支持临时输入 Aster API Key / App Secret；这些值只写入当前 Python 进程内存，提交后密码框会清空，重启服务后失效。

### Aster 生产模式

默认是本地 mock：

```bash
python pstx_web.py
```

需要真实 Aster 时，在部署机器配置环境变量：

```bash
export PSTX_ASTER_MODE=live
export PSTX_ASTER_BACKEND=chat-flow
export ASTER_BASE_URL="https://aigc.huaqin.com"
export ASTER_EMP_NO="100019100"
export ASTER_API_KEY="flow_api_key_xxx"
python pstx_web.py
```

可选生产参数：

```bash
export PSTX_ASTER_TIMEOUT_SECONDS=45
export PSTX_ASTER_MAX_ROWS_PER_TABLE=16
export PSTX_ASTER_MAX_PAYLOAD_CHARS=60000
export PSTX_ASTER_REDACT_PATHS=1
export PSTX_ASTER_LOG_FILE="./logs/aster_debug.log"
export PSTX_ASTER_LOG_PAYLOAD=0
```

Aster live 调用会写入 JSONL 诊断日志，默认路径是 `logs/aster_debug.log`。日志会记录 `request_id`、调用后端、URL 脱敏版本、HTTP 状态、响应摘要和错误体片段；`apiKey`、`appSecret`、`accessToken`、`Authorization`、`ciphertext` 会被脱敏，不写明文。只有排查 prompt 内容时才建议临时设置 `PSTX_ASTER_LOG_PAYLOAD=1`，因为它会额外记录裁剪后的请求 payload。

如果生产环境需要 Room/Auth 后端：

```bash
export PSTX_ASTER_BACKEND=room
export ASTER_APP_ID="ag_xxx"
export ASTER_APP_SECRET="***"
export ASTER_ORIGIN="test-aigc-api.huaqin.com"
export PSTX_ASTER_VALIDATE_TOKEN=1
python -m pip install pycryptodome
```

Room/Auth 使用 AES-256-ECB + PKCS7 + 双 Base64，与 Aster 官方示例保持一致。该链路会先调用 `/auth/api/v1/generateOrProlongToken` 获取或续期 token，再调用 `/auth/js-sdk/validateAccessToken` 校验 token，最后创建 room 并发送问答。`ASTER_ORIGIN` 用作 validate 接口的 `aigc-origin` 和加解密 key；不填时默认从 `ASTER_BASE_URL` 推导域名。生产故障时可把 `PSTX_ASTER_MODE` 改回 `mock`，不影响本地报告生成。

如果不方便改环境变量，也可以在报告页的“临时 Aster 凭据”表单里设置：

- `mode=live`
- `backend=chat-flow`
- `base_url`
- `emp_no`
- `api_key`
- `origin`，仅 Room/Auth token validate 需要，默认可留空

这只适合本地或受控内网临时使用。长期生产部署仍推荐使用环境变量，避免浏览器插件、调试工具或屏幕录制接触 secret。

如果 Aster 背后实际由 Dify 承载，本工具仍然只连接 Aster wrapper，不直连原生 Dify。需要在 Aster/Dify 平台侧完成：

- 发布对应的 ChatFlow / Agent 应用。
- 生成给 Aster wrapper 使用的 API Key。
- 配置模型供应商、知识库、系统提示词和应用变量。
- 确认 Aster 暴露的 `ASTER_BASE_URL`、`ASTER_API_KEY`、`ASTER_EMP_NO` 与该应用匹配。

如果看到 `Aster HTTP 401`、`unauthorized` 或 `Access token is invalid`，优先检查：

- `PSTX_ASTER_BACKEND` 是否选对，ChatFlow 使用 `chat-flow`，普通智能体/Room 使用 `room`。
- `backend=chat-flow` 时，`ASTER_API_KEY` 应该是对应 ChatFlow 的 API Key；ChatFlow 不使用 accessToken。
- `backend=room` 时，确认 `ASTER_APP_ID`、`ASTER_APP_SECRET`、`ASTER_EMP_NO` 能正常换取 accessToken，并确认 validate 返回 `isValid=true/statusCode=1`。
- `ASTER_BASE_URL` 是否是 Aster wrapper 根地址，而不是原生 Dify 地址或带错路径的地址。
- `ASTER_ORIGIN` 是否与 Aster 环境要求一致；如果 validate 失败，可尝试按接口文档中的虚拟域名填写。
- 报告页失败卡片里的 `请求 ID` 和 `日志文件`，可以和 `logs/aster_debug.log` 中同一个 `request_id` 对上。

## 精简后的结构

```text
pstx_analyzer.py       核心解析、规则分析、Excel 导出
pstx_page_logic.py     页码解析、page.map/page*.csv/module_order 映射
pstx_csa_geometry.py   sch_1/page*.csa 几何规范检查
pstx_aster_mock.py     Aster mock-only 辅助审查摘要
pstx_aster_client.py   Aster 真实生产客户端
pstx_aster_service.py  Aster mock/live/off 模式切换
pstx_web.py            localhost Web UI
pstx_local_ui.py       本地桌面套壳入口
web/                   HTML/CSS/JS 前端资源
tests/                 回归测试
docs/                  精简后的设计与 review 说明
```

## 测试

```bash
PYTHONPATH=/opt/pyvenv/lib/python3.13/site-packages:. python -S -m unittest discover -s tests -q
```

在普通本地环境里也可以直接运行：

```bash
python -m unittest discover -s tests -q
```
