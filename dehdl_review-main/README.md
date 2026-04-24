# PSTX 原理图审查工具

解析 Cadence Packager-XL 导出的 `pstxprt.dat` / `pstxnet.dat`，生成 BOM、网络、DRC、电容降额、电阻规则、芯片 Pin 状态、页码映射与 Excel 报告。

## 运行入口

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
- 默认端口 `8765`，被占用时自动顺延。
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

## 精简后的结构

```text
pstx_analyzer.py       核心解析、规则分析、Excel 导出
pstx_page_logic.py     页码解析、page.map/page*.csv/module_order 映射
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
