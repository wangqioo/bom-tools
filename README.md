# BOM Tools

企业内部硬件设计辅助工具集，提供 Web 统一界面，支持多人在线使用。

---

## 快速启动

### 有网环境

```bash
cd web_app2
pip install -r requirements.txt
python app.py        # 访问 http://localhost:5000
```

### 无网 / Windows 离线部署

将 `deploy_bundle/` 文件夹整体复制到目标机器，双击运行：

```
deploy_bundle/install_and_run.bat
```

脚本将自动创建虚拟环境、离线安装依赖并启动服务器。

> **前置要求：** Python 3.10+（从 [python.org](https://www.python.org/downloads/) 安装，勾选 "Add Python to PATH"）

---

## 工具功能

### 1. BOM 格式转换

将客户提供的多种格式 BOM 展开为标准多行格式，支持：

- **格式 A** — 品牌型号合并在一列（`||` 或空格分隔）
- **格式 B** — 品牌列与型号列分开，分号分隔多供应商
- **格式 C** — 制造商含内部编号，冒号分隔

输出模式：原格式展开 / 转为整机 BOM 配置表。

---

### 2. 飞书优选库 + 关系库匹配

连接飞书内部 API 网关，从在线库中批量查找物料信息。

**核心特性：**
- 预置 15 个库（14 优选库 + 1 对应关系库），分组显示
- 本地键与飞书键多对多 AND 匹配（可动态增删键数量）
- 提取列映射：标准输出列名（HQ料号 / HQ规格型号 / HQ制造商 / 优选等级 / HQ描述）与各 sheet 实际列名一一对应
- 服务端数据缓存，支持一键批量缓存所有启用 sheet
- 配置导出 / 恢复默认，所有设置自动持久化

**匹配逻辑：**
- 本地键为空的行自动跳过，不参与匹配，原行保留
- 未匹配行原样保留，已匹配行附加提取列数据

---

### 3. 查询 BOM 优选率

上传含 HQ料号 的 BOM，在所有已缓存优选库中查询优选等级。

**优选判定：**

| 优选等级值 | 是否优选 |
|---|---|
| 文字含「优选」 | ✅ |
| 数字 7 / 8 / 9 | ✅ |
| 数字 1–6 或其他 | ❌ |

**优选率公式：优选料数 ÷ 已匹配料数**（未匹配行不参与计算）

输出 Excel 色彩标注：绿色 = 优选 / 黄色 = 匹配到但非优选 / 灰色 = 未匹配。

---

### 4. 转换为上传 PLM 系统格式

包含两个子功能：

#### 4a. 整机 BOM 配置表转换

将整机 BOM 配置表转换为 PLM 系统标准导入格式（25 列）：序号、料号、单耗、替代关系、主辅 BOM 标记等。

#### 4b. 规格型号提取

上传 BOM → 选择任意一列 → 提取全部值（自动去除空格）→ 输出单列 Excel（列名「规格型号」）。

---

## 目录结构

```
bom-tools/
├── web_app2/               主 Web 应用（Flask）
│   ├── app.py              入口
│   ├── shared.py           公共工具
│   ├── bom/                BOM 格式转换 Blueprint
│   ├── feishu/             飞书匹配 + 优选率 Blueprint
│   ├── plm/                PLM 转换 Blueprint
│   ├── templates/
│   │   └── index.html      单页前端（所有 JS 内联）
│   └── default_config.json 飞书库默认配置（含完整 sheet 映射）
├── deploy_bundle/          Windows 离线部署包
│   ├── install_and_run.bat 一键安装 + 启动
│   ├── requirements.txt    依赖声明
│   ├── wheels/             离线 .whl 文件（Python 3.10–3.14）
│   └── web_app2/           同上，随包自带
├── scripts/                CLI 脚本（单机运行，独立使用）
│   ├── bom_gui.py
│   ├── feishu_multi_matcher.py
│   ├── plm_upload.py
│   ├── pstx_analyzer.py
│   └── csa_checker.py
└── CHANGELOG.md
```

---

## 依赖

```
Flask >= 3.0
openpyxl >= 3.1
requests >= 2.28
```

---

## 配置持久化

Web 版所有飞书相关配置（Token、匹配键、提取列映射、缓存记录）均保存在**浏览器 localStorage** 中，刷新页面或重启服务器不会丢失。

服务器内置默认配置（`default_config.json`），新用户首次打开后可点击「↩ 恢复默认配置」一键加载。
