# 默认端口 44441 与单项目打开修复

## 背景

开发机上 `8765` 已被其他 Python 服务占用，导致直接访问原默认端口时可能命中错误服务，表现为单项目打开动效页 `/debug/report-open` 不存在或返回 404。

## 变更

- Web UI 默认端口从 `8765` 改为 `44441`。
- 本地 UI 套壳默认端口同步改为 `44441`。
- 单项目打开动效页服务端模板增加 `data-phase="pick"` 初始状态，避免 JS 加载前页面状态未定义。
- README 和测试断言同步更新默认端口。

## 验证

- `python3 -m unittest discover -s tests -p test_pstx_web.py -v`
- `python3 -m unittest discover -s tests -p test_pstx_local_ui.py -v`
- 通过 `http://127.0.0.1:44441/debug/report-open` 验证页面可访问
