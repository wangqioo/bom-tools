# 查询BOM优选率工具实现逻辑

## 工具定位

该工具用于按本地 BOM 的 HQ 料号查询飞书优选库缓存，判断物料是否优选并计算优选率。

前端入口：`data-tool="pref-rate"`。
后端接口：`web_app2/feishu/__init__.py` 中 `/api/feishu/pref_rate`。

## 前置条件

需要先在“飞书优选库+关系库匹配”中缓存优选库 Sheet。优选率工具不主动拉飞书，而是读取 `web_app2/cache/` 下的缓存文件。

## 接口输入

`/api/feishu/pref_rate` 接收：

- `file`：本地 BOM。
- `config`：JSON，包含 Sheet 名、表头行、本地 HQ 料号列、飞书表格缓存配置。

## 数据处理流程

1. 保存本地 BOM。
2. 打开 Excel，选择 Sheet。
3. 校验表头行和本地 HQ 料号列。
4. 遍历配置中的飞书缓存。
5. 从缓存中读取飞书表头和数据。
6. 通过字段别名找到飞书 `HQ料号` 和 `优选等级` 列。
7. 构建综合查询表：`HQ料号 -> {pref, source}`。
8. 遍历本地 BOM 有效行，按 HQ 料号查询。
9. 输出原始 BOM 行，并追加 `优选等级`、`来源` 两列。
10. 返回统计结果和下载链接。

## 优选判断规则

函数：`_is_preferred_level(value)`。

规则：

- 空值不是优选。
- 包含 `非优选/不优选/not preferred` 等负向词时不是优选。
- 包含 `优选/preferred` 时是优选。
- 数字值大于等于 7 时是优选。
- 其他情况不是优选。

## 输出颜色

- 绿色：匹配到且为优选。
- 黄色：匹配到但非优选。
- 灰色：未匹配。

## 统计口径

返回字段：

- `total`：本地 BOM 有效行数。
- `matched`：在缓存中匹配到的行数。
- `unmatched`：未匹配行数。
- `preferred`：匹配到且为优选的行数。
- `non_preferred`：匹配到但非优选的行数。
- `rate`：优选率。

优选率公式：`preferred / matched * 100%`。未匹配项不参与分母。

## 输出文件

输出文件：`web_app2/outputs/pref_rate_<uid>.xlsx`。
Sheet 名：`优选率查询`。
