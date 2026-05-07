# 架构说明

## 模块边界

- `pstx_analyzer.py`
  - PSTX 文本解析
  - BOM / 网络 / DRC / 降额 / 电阻规则分析
  - Excel 导出
  - 只保留少量页码兼容包装，具体页码决策不再散落在这个文件里
- `pstx_page_logic.py`
  - `SECTION_NUMBER` / `C_PATH` / `P_PATH` 路径解析
  - `page.map` 与 `page*.csv` 的逻辑页 / 真实页索引
  - `module_order(.dat)` 的子模块映射主模块真实页计算
- `pstx_csa_geometry.py`
  - `sch_1/page*.csa` 几何对象解析
  - 带 `DOT` 的四向十字交叉检测
  - `CIRCLE` / `ARC` 画圈对象候选输出
  - 只提供几何规范候选，不推断电气连通性结论
- `pstx_aster_mock.py`
  - Aster 实验链路的 mock-only 辅助审查摘要
  - 只读取当前报告聚合指标和分区行数
  - 不访问真实 Aster 内网地址，不读取或下发 secret/token/apiKey
- `pstx_aster_client.py`
  - 真实 Aster ChatFlow / Room 后端客户端
  - 构造裁剪后的报告摘要 prompt，标准化模型 JSON 或纯文本回答
  - 支持 `ASTER_BASE_URL`、`ASTER_API_KEY`、`ASTER_EMP_NO` 等生产环境变量
  - live 请求写入脱敏 JSONL 诊断日志，异常携带 request_id/status/log_file
  - Room/Auth 后端按官方流程生成/续期 token，并在 room create/chat 前调用 `validateAccessToken`
- `pstx_aster_service.py`
  - 读取 `PSTX_ASTER_MODE` 并在 mock / live / off 间切换
  - 将配置错误、上游错误转换为 Web 可显示的 JSON 错误
  - 对外提供脱敏认证状态，前端只看到 configured/missing，不看到 secret/token/apiKey
  - 支持当前 Python 进程内的 runtime credential override，供前端临时输入 Aster 凭据
- `pstx_web.py`
  - Flask Web UI
  - 项目根路径读取
  - 报告、查询、导出接口
- `pstx_local_ui.py`
  - 本地桌面壳
  - 复用 `pstx_web.create_app()`，不再维护第二套业务逻辑

## 页码模型

组件默认显示页优先级：

1. `P_PATH` 顶层 `SCH_1` 页
2. `page.map` 映射出的真实页
3. `page*.csv` 映射出的真实页

逻辑页来源优先级：

1. `SECTION_NUMBER 1` 路径
2. `C_PATH`
3. `DRAWING`

`module_order(.dat)` 规则：

- key 优先按 `SECTION_NUMBER / C_PATH` 的逻辑路径构造
- `P_PATH` 只作为保守回退
- 子模块偏移页优先使用 `P_PATH` 中子模块本地真实页
- 映射公式为：

```text
子模块映射主模块真实页 = start_real_page + 子模块本地真实页 - 1
```

## 降额补丁规则

- `analyze_derating()` 先尝试扫描整板可识别的最大正电压
- 当最大已识别正电压 `<=12V` 且电容额定耐压 `>=50V` 时，直接放行为合格
- 若无法识别整板最大正电压，仍回到单颗电容原有推断逻辑

## CSA 几何规范检查

- 默认扫描项目根目录下的 `sch_1/page*.csa`
- 文件名中的 `pageX.csa` 作为真实页显示为 `PAGEX`
- `SET PAGE_NUMBER` 保留为 `CSA页名`，用于辅助对照原始页文件
- DOT 四向十字交叉的命中条件是同一 DOT 坐标同时存在 left/right/up/down 四个方向的正交 WIRE
- T 型连接、无 DOT 的视觉十字、斜线不报
- `ARC` 拟合圆只作为画圈候选，报告中保留解析说明并要求人工确认

## Aster 摘要链路

- 默认提供 mock-only 辅助审查摘要，用于本地离线使用
- 生产环境可通过 `PSTX_ASTER_MODE=live` 切到真实 Aster
- Web 前端调用本地 `GET /api/report/<run_id>/aster-summary`
- Web 前端通过 `GET /api/aster/status` 展示 Aster 认证配置状态
- Web 前端可通过 `POST /api/aster/runtime-config` 设置当前进程临时凭据，并通过 `DELETE /api/aster/runtime-config` 清除
- 报告页 Aster 入口是浮窗助手，展示摘要、优先级、审查清单、分区焦点和人工复核边界
- 后端摘要输入限制为当前报告 payload 裁剪摘要，不上传原始 PSTX 文件
- 后端会额外整理 `review_scope`、`key_findings` 和 `manual_review_boundaries`，显式覆盖 BOM/DEPOP、网络/页码映射、DRC、芯片 Pin/电阻、降额和 CSA 规范
- 默认 live 后端为 ChatFlow：`PSTX_ASTER_BACKEND=chat-flow`
- ChatFlow 路径只使用 `apiKey + empNo`，不使用 accessToken
- Room/Auth 后端需要部署环境安装 AES 依赖 `pycryptodome`、`pycryptodomex` 或 `cryptography`
- Room/Auth 路径使用 `ASTER_APP_ID` / `ASTER_APP_SECRET` 生成或续期 token，使用 `ASTER_ORIGIN` 作为 validate 接口 `aigc-origin` 和加解密 key，不填时从 `ASTER_BASE_URL` 推导
- Room/Auth 默认开启 `PSTX_ASTER_VALIDATE_TOKEN=1`，在创建 room 和发送问答前校验 token；可临时设为 `0` 跳过校验
- live 请求默认诊断日志为 `logs/aster_debug.log`，可通过 `PSTX_ASTER_LOG_FILE` 修改；默认只记录 payload 长度/hash，`PSTX_ASTER_LOG_PAYLOAD=1` 才记录裁剪后的 payload
- 诊断日志和 Web 错误 payload 会脱敏 `apiKey`、`appSecret`、`accessToken`、`Authorization`、`ciphertext`
- 401 / unauthorized / access token invalid 会在 Web 错误 payload 中返回排查提示，重点区分 ChatFlow API Key 和 Room accessToken
- mock 可作为生产回退：`PSTX_ASTER_MODE=mock`
- 若 Aster 底层由 Dify 承载，Dify 的应用发布、API Key 和模型供应商配置在 Aster/Dify 平台侧完成；本工具只连接 Aster wrapper
- 前端临时输入的 API Key / App Secret 只保存在后端内存，状态接口不回显；生产长期配置仍推荐使用环境变量

## 测试入口

- `tests/test_pstx_analyzer.py`
  - 主分析链、页码模型、CSA 几何检查、电阻与降额规则
- `tests/test_pstx_web.py`
  - Web UI 后端流程
- `tests/test_pstx_local_ui.py`
  - 本地桌面壳入口
