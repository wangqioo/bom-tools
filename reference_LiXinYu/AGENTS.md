# AGENTS

## 运行入口

- 本地 UI：`python pstx_local_ui.py`
- Web UI：`python pstx_web.py`
- 兼容入口：`python pstx_analyzer.py`

## 代码边界

- 页码解析只放在 `pstx_page_logic.py`
- CSA 几何规范检查只放在 `pstx_csa_geometry.py`
- Aster 摘要模式切换放在 `pstx_aster_service.py`，真实 Aster 协议放在 `pstx_aster_client.py`，本地 mock 保留在 `pstx_aster_mock.py`
- Aster 长期凭据优先用后端环境变量；前端只允许通过报告页临时表单提交到当前进程内存，不得回显或持久化 secret/token/apiKey
- 规则分析和 Excel 导出放在 `pstx_analyzer.py`
- Web 展示放在 `pstx_web.py` 和 `web/`
- 本地桌面入口只做 Web UI 壳，不复制业务逻辑

## 文档路径

- `docs/ARCHITECTURE.md`
  - 当前模块边界和页码模型
- `docs/REVIEW.md`
  - 当前版本 review 摘要
- `docs/analysis/aster_harness_review.md`
  - Aster 实验包解析、安全边界和 mock-only 接入基线
- `docs/reviews/`
  - 规则逻辑 review 与人工判断边界记录
- `docs/workflow/`
  - 每轮设计和 issue 拆分记录
- `docs/changelists/`
  - 每轮变更摘要

## 最新变更

- `docs/changelists/2026-04-26-aster-floating-review-assistant.md`
  - Aster 升级为报告页浮窗审查助手，并将 BOM/DEPOP、网络/页码映射、DRC、电阻、降额和 CSA 等审查域加入 AI 上下文
- `docs/changelists/2026-04-26-aster-protocol-alignment.md`
  - 按 Aster 官方文档对齐 ChatFlow 与 Room/Auth 协议边界，Room 默认生成/续期并 validate token 后再问答
- `docs/changelists/2026-04-26-aster-diagnostics-logging.md`
  - Aster live 调用新增脱敏 JSONL 诊断日志、request_id、401 排查提示和报告页错误详情展示
- `docs/changelists/2026-04-26-aster-runtime-credentials.md`
  - 报告页支持前端临时输入 Aster 凭据并写入后端当前进程内存，提交后清空密码框，不写磁盘、不回显 secret
- `docs/changelists/2026-04-26-aster-auth-status-ui.md`
  - 报告页新增 Aster 认证状态展示和 `/api/aster/status`，只显示配置状态，不暴露 secret/token/apiKey；保持只连接 Aster wrapper，不直连原生 Dify
- `docs/changelists/2026-04-25-aster-live-production.md`
  - Aster 摘要支持 `PSTX_ASTER_MODE=live` 真实生产 ChatFlow 调用，保留 mock 回退和服务端凭据边界
- `docs/changelists/2026-04-25-aster-mock-summary.md`
  - 基于 Aster harness 增加 mock-only 辅助审查摘要，报告页新增实验入口，后端不访问真实 Aster 且不读取凭据
- `docs/changelists/2026-04-25-csa-geometry-review.md`
  - 接入参考包中的 CSA 几何规范检查能力，新增 DOT 四向十字交叉和 CIRCLE/ARC 画圈对象检查，并在 Web/Excel 报告中新增“规范检查”分区
- `docs/changelists/2026-04-25-default-port-44441-report-open-fix.md`
  - 默认启动端口切换为 `44441`，并为 `/debug/report-open` 单项目打开页补充初始状态兜底，规避 `8765` 端口冲突导致访问到错误服务
- `docs/changelists/2026-04-25-windows-ui-smoothness.md`
  - 面向 Windows 部署优化 Web UI 显示和流畅性：字体、滚动条、大表懒加载、报告页低负载动效和减少动态效果兼容
- `docs/changelists/2026-04-24-web-global-motion-pass.md`
  - Web 全站动效优化：覆盖首页、报告页、查询、项目管理、项目对比和 Debug 页面，统一页面级入场与状态反馈
- `docs/changelists/2026-04-24-web-debug-report-open.md`
  - Web 新增 `/debug/report-open` 单项目报告打开动效页，用模拟数据展示项目卡片到报告工作台首屏的过渡
- `docs/changelists/2026-04-24-web-debug-effects-page.md`
  - Web 新增 `/debug/effects` 动效模拟页，用固定模拟数据回放概览、表格、筛选弹层和项目对比动画
- `docs/changelists/2026-04-24-web-motion-polish.md`
  - Web UI 动效优化：概览卡片、表格展开、列面板、项目列表和项目对比结果增加轻量分层动画
- `docs/changelists/2026-04-24-multi-project-compare.md`
  - Web 新增会话内多项目管理和两两对比，可对比指标、元件、网络与主要结果表差异
- `docs/changelists/2026-04-24-ground-net-series-resistor-fix.md`
  - AGND/GNDA/VSS/0V 类地网在串阻路径搜索中统一按 GND 终止点处理，避免误判为普通信号
- `docs/changelists/2026-04-24-web-multi-column-filtering.md`
  - Web 表格多列组合筛选：支持多字段、多条件 AND 筛选并保留原有关键字和排序能力
- `docs/changelists/2026-04-24-web-render-performance.md`
  - Web 报告页渲染性能优化：近视口懒加载、大表渐进渲染、拖拽/滚动节流和离屏绘制隔离
- `docs/changelists/2026-04-24-web-engineering-ui-pass.md`
  - 对齐生成概念图后的 Web 工程审美优化：固定侧栏、项目状态条、KPI dock、表格优先
- `docs/changelists/2026-04-24-web-report-workbench.md`
  - Web 报告页三栏工作台、表格密度切换、长文本显示和 reveal 可读性修复
- `docs/reviews/2026-04-24-logic-rule-review.md`
  - 规则逻辑 review 记录：已修复问题、保守保持项和后续样本需求
- `docs/changelists/2026-04-24-logic-review-hardening.md`
  - 规则逻辑 review 后的电阻值解析与多路径串阻搜索修复
- `docs/changelists/2026-04-24-code-review-hardening.md`
  - 全代码 review 后的 module_order 去重与 Web 输入解码硬化修复
- `docs/changelists/2026-04-24-submodule-real-page-mapping.md`
  - `module_order(.dat)` 子模块映射真实页修复
- `docs/changelists/2026-04-24-derating-50v-override.md`
  - 电容降额的 12V / 50V 低压直通规则

## 测试

```bash
python -m unittest discover -s tests -q
```

新增页码或 UI 行为时，优先跑：

- `tests/test_pstx_analyzer.py`
- `tests/test_pstx_web.py`
- `tests/test_pstx_local_ui.py`
