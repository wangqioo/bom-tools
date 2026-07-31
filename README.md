# BOM Tools

BOM Tools 是面向硬件研发、BOM 审核、物料选型和 PLM 上传流程的内部辅助平台。项目以 Flask Web 应用为主入口，整合 Excel 处理、飞书表格匹配、BOM 对比、PLM 格式转换、缺陷反馈和需求收集等工具。

## 功能概览

- BOM 格式转换：识别客户 BOM 的品牌、型号、制造商等字段，并展开为标准多行格式。
- 飞书优选库匹配：按本地 BOM 字段和飞书表格字段做多键 AND 匹配，支持缓存在线表格数据。
- BOM 优选率查询：基于已缓存的优选库统计 HQ 料号优选率并导出标色 Excel。
- PLM 上传格式转换：把整机 BOM 配置表转换为 PLM 导入格式。
- PLM 网页自动化：基于 Playwright 辅助查询、上传和导出 PLM 数据。
- BOM 对比工具集：支持通用 BOM、客户 BOM 与 HQ BOM、HQ BOM 版本、整机 HQ BOM 版本、Cadence BOM 与 HQ BOM 对比。
- 厂商命名映射：维护客户厂商别名到 HQ 标准厂商名的映射。
- Bug 与需求工单：在 Web 页面提交问题、附件、需求和状态流转信息。
- 小工具合集：当前包含本地文件 MD5 计算。

更细的业务规则见 [功能说明.md](功能说明.md) 和 [doc/](doc/)。

## 快速启动

### 开发环境

```powershell
cd C:\Users\100448405\bom-tools
python -m venv venv
.\venv\Scripts\activate
python -m pip install -r web_app2\requirements.txt
python web_app2\app.py
```

访问：

```text
http://localhost:5000
```

默认启用登录。测试环境会自动关闭登录校验；本地临时调试也可以设置：

```powershell
$env:BOM_TOOLS_AUTH_REQUIRED = "0"
python web_app2\app.py
```

### 一键部署和启动

在项目根目录运行：

```cmd
deploy_one_click.bat
```

该脚本会创建虚拟环境、安装依赖，并优先使用 `deploy_bundle/wheels/` 中的离线 wheel 包。若发现 `bom-tools_offline_*.zip`，会询问是否先部署再启动。

## 测试与检查

当前测试使用标准库 `unittest`，不强依赖 pytest：

```powershell
python -m unittest discover -s tests
```

发布或交付前建议跑完整预检：

```powershell
python scripts\preflight_check.py
```

预检会依次执行：

- UTF-8 源文件检查
- `web_app2` Python 编译检查
- 平台/工具版本号变更检查
- 全量 `unittest`

如需单独检查版本号：

```powershell
python scripts\check_version_bumps.py --root .
```

版本号定义在 [web_app2/shared.py](web_app2/shared.py)：

- `PLATFORM_VERSION`：Web 平台壳版本。
- `TOOL_VERSIONS`：各功能工具版本。

当修改平台入口、前端壳或具体工具代码时，应同步提升对应版本号，避免用户浏览器缓存导致页面和后端能力不一致。

## 离线发布包

生成离线发布包：

```powershell
powershell -ExecutionPolicy Bypass -File scripts\export_offline_release.ps1
```

安装离线发布包：

```powershell
powershell -ExecutionPolicy Bypass -File scripts\install_offline_release.ps1 -PackagePath .\deploy_bundle\bom-tools_offline_YYYYMMDD_HHMMSS.zip -InstallDir C:\path\to\bom-tools
```

发布脚本会排除运行时数据，并在安装时备份、恢复这些目录：

- `web_app2/auth_data`
- `web_app2/cache`
- `web_app2/uploads`
- `web_app2/outputs`
- `web_app2/logs`
- `web_app2/bug_reports`
- `web_app2/feature_requests`
- `web_app2/manufacturer_aliases`

## 目录结构

```text
bom-tools/
  web_app2/                     Flask Web 应用
    app.py                      Web 入口与蓝图注册
    shared.py                   公共工具、路径、版本号
    auth.py                     登录、用户和权限
    bom/                        BOM 格式转换
    feishu/                     飞书匹配与优选率
    plm/                        PLM 转换与自动化
    bom_compare/                BOM 对比工具集
    manufacturer_alias/         厂商别名映射
    bug_report/                 Bug 工单
    feature_request/            需求工单
    templates/                  页面模板
    static/                     前端 CSS/JS
  scripts/                      开发、检查、发布和独立工具脚本
  tests/                        unittest 测试
  doc/                          业务规则说明
  deploy_bundle/                离线部署资源和 wheel 包
  manufacturer_mapping_extracts/ 厂商映射历史资料
```

## 依赖说明

核心依赖：

- Python 3.10+
- Flask
- openpyxl
- xlrd（读取并转换标准 `.xls` 旧格式文件）
- requests
- waitress
- Playwright

PLM 自动化依赖 Chromium 运行时。离线环境应随包提供 `ms-playwright/`，或设置 `PLAYWRIGHT_BROWSERS_PATH` 指向已有浏览器目录。

## Excel 导入开发约定

Web 新功能的 Excel 上传必须复用 [`web_app2/shared.py`](web_app2/shared.py) 的统一上传入口，不能直接保存后用 `openpyxl` 打开。独立桌面脚本必须复用 [`scripts/excel_compat.py`](scripts/excel_compat.py) 的 `open_workbook_compat`。两类入口都会将标准 `.xls` 自动转换为 `.xlsx`，保证各功能的格式兼容行为一致。

实现方式、前端限制、测试要求和加密文件边界见 [doc/10_Excel导入开发规范.md](doc/10_Excel导入开发规范.md)。

## 运维注意事项

- Web 上传和导出文件会写入 `web_app2/uploads`、`web_app2/outputs`，后台清理任务会定期删除旧文件。
- 飞书缓存写入 `web_app2/cache`，由用户手动刷新，不随普通文件清理删除。
- 登录用户数据保存在 `web_app2/auth_data`。
- 工单和厂商映射数据保存在对应运行时目录，发布包默认不覆盖。
- `deploy_bundle/` 内可能包含一份部署用代码副本，日常开发优先修改根目录 `web_app2/`，再通过发布脚本生成部署包。
