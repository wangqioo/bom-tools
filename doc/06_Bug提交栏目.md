# Bug提交栏目实现逻辑

## 工具定位

Bug 提交栏目用于收集团队成员使用工具时遇到的问题，支持附件、公开列表、详情查看、状态修改和示例案例。

前端入口：`data-tool="bug-report"`。
后端模块：`web_app2/bug_report/__init__.py`。

## 数据存储

使用 SQLite：`web_app2/bug_reports/reports.sqlite3`。

附件目录：`web_app2/bug_reports/attachments/`。

表字段包括：id、提交时间、提交人、工号、模块、严重程度、状态、标题、描述、复现步骤、期望结果、附件 JSON。

SQLite 开启 WAL 和 busy timeout，适合局域网多人轻量使用。

## 内置示例

`_seed_reports()` 会用 `INSERT OR IGNORE` 自动补齐 3 条固定 ID 示例，不重复插入，不覆盖真实记录。

示例包括 BOM 表头问题、飞书缓存超时、PLM ZIP 文件清单需求。

## 主要接口

### `GET /api/bug_reports`

返回所有问题记录，按提交时间倒序。

### `POST /api/bug_reports`

提交新问题。

必填：姓名、工号、问题标题、问题描述。

可选：影响模块、严重程度、复现步骤、期望结果、附件。

默认状态：`待处理`。

### `POST /api/bug_reports/<report_id>/status`

修改处理状态。

允许状态：待处理、处理中、已修复、已关闭、暂缓、无法复现。非法状态会被拒绝。

### `GET /bug_attachments/<filename>`

打开或下载附件。使用 `safe_join()` 防止路径穿越。

## 附件处理

支持图片、Excel、CSV、TXT、LOG、ZIP、RAR、7Z。

保存文件名格式：`<report_id>_<随机8位>_<安全文件名>`。

## 前端处理逻辑

1. `initBugReport()` 初始化页面。
2. `bugLoadReports()` 拉取记录并渲染卡片。
3. 卡片点击后 `bugOpenReport()` 打开详情弹窗。
4. 附件链接阻止冒泡，点击附件不会误打开详情。
5. 弹窗内可选择处理状态。
6. `bugSaveStatus()` 调用后端接口保存状态，成功后刷新列表。

## 状态颜色

`bugStatusClass()` 将状态映射为不同颜色：

- 待处理：橙色。
- 处理中：蓝色。
- 已修复：绿色。
- 已关闭：灰色。
- 暂缓：紫色。
- 无法复现：红色。

列表卡片和详情弹窗都会显示颜色。
