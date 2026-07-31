# Excel 导入开发规范

## 适用范围

本规范适用于所有需要读取本地 Excel 文件的新接口、页面预览、批量处理、导入功能和独立桌面脚本。目标是让标准 `.xls` 与 `.xlsx` 在所有功能中走统一的兼容流程。

`.xlsm` 如作为输入，需在功能设计和测试中明确是否保留 VBA；现有公共读取流程面向表格数据，不承诺保留宏。

## 必须使用的公共入口

在 Blueprint 中从 `shared` 导入以下函数：

```python
import uuid

from shared import (
    _open_workbook,
    _save_uploaded_excel,
    _save_or_reuse_uploaded_excel,
)
```

首次上传并立即处理时：

```python
uid = str(uuid.uuid4())[:8]
input_path = _save_uploaded_excel(request.files.get("file"), "feature_in", uid)
workbook = _open_workbook(input_path, data_only=True)
```

预览后复用同一次上传的文件时：

```python
uid, input_path = _save_or_reuse_uploaded_excel(
    request.files.get("file"),
    "feature_preview",
    request.form.get("uid", ""),
)
workbook = _open_workbook(input_path, data_only=True)
```

`prefix` 必须是当前功能专用、稳定的 ASCII 标识，例如 `feature_in`、`feature_preview_left`。不要让不同文件输入共用同一个 `prefix`，否则并发请求可能覆盖上传文件。

## 禁止的实现方式

- 不要对 `FileStorage` 直接 `save()` 后立刻调用 `openpyxl.load_workbook()`。
- 不要在新功能中复制 PowerShell / Excel COM 转换代码。
- 不要直接调用 `_convert_xls`、`_convert_xls_with_xlrd` 或 `_convert_xls_with_excel`；这些是公共入口内部的实现细节。
- 不要仅在前端的 `accept` 属性增加 `.xls` 就认为已经支持旧格式。后端必须通过上述入口处理。
- 不要把 `.xls` 当作 `.xlsx` 直接传给 `openpyxl`。

## 公共入口行为

`_save_uploaded_excel` 会按文件类型处理：

| 输入文件 | 处理结果 |
|---|---|
| `.xlsx` | 保存后返回该 `.xlsx` 路径。 |
| 标准 `.xls` | 先使用 `xlrd` 读取并转换为临时 `.xlsx`；失败时在 Windows 上尝试已安装 Excel 的 COM 转换；返回转换后的 `.xlsx` 路径。 |
| 损坏、密码保护或企业加密 `.xls` | 返回统一的可读错误信息，要求在具有授权的 Excel 中打开并另存为 `.xlsx`。 |

因此，调用方得到的 `input_path` 可以始终交给 `_open_workbook` 或 `openpyxl` 处理，无需针对 `.xls` 分支。

旧式 `.xls` 的转换以读取业务数据为目标，不保证保留公式计算结果以外的样式、宏、嵌入对象、批注或其他 Excel 专有内容。需要保真处理这些内容的功能，应明确只接受 `.xlsx`，并在界面和接口中说明原因。

## 前端与接口约定

- 可读取 Excel 的文件选择器应包含 `accept=".xlsx,.xlsm,.xls"`；后端仍是格式兼容的唯一保障。
- 接口错误必须将 `_save_uploaded_excel` 和 `_open_workbook` 抛出的 `ValueError` 原样返回为用户可见错误，不能改写为“服务器未安装 Excel”。
- 输出文件继续使用 `.xlsx`，不要重新生成 `.xls`。

## 测试与发布检查

每个新增 Excel 导入功能至少覆盖以下用例：

1. `.xlsx` 上传、预览或处理成功。
2. `.xls` 上传会经过公共入口，且后续读取成功。
3. `.xls` 转换失败时，接口返回可操作的错误信息，不泄露服务器路径或堆栈。
4. 若支持预览复用，重新提交 `uid` 能正确找到转换后的文件。

提交前运行：

```powershell
python scripts\preflight_check.py
```

修改 `web_app2/shared.py` 或依赖清单时，还必须同步更新 `deploy_bundle/` 对应文件，并确认离线 wheel 包包含 `xlrd`。

## 独立脚本

Web 功能必须使用 `web_app2/shared.py` 的上传入口。独立桌面或命令行脚本必须使用 `scripts/excel_compat.py` 的 `open_workbook_compat`，不能因为文件选择框允许 `.xls` 就直接调用 `openpyxl.load_workbook`。该兼容层优先使用 `xlrd`，必要时在 Windows 上使用授权 Excel 转换；转换失败时必须给出可操作的错误信息。
