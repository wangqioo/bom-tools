# BOM比对工具合集实现逻辑

## 工具定位

BOM 比对工具合集包含三个子功能：客户 BOM 对比 HQ BOM、同项目 HQ BOM 版本对比、Cadence BOM 对比 HQ BOM。

前端入口：`data-tool="bom-compare"`。
后端模块：`web_app2/bom_compare/__init__.py`。

## 通用 BOM 比对

适用于客户 BOM 对比 HQ BOM 和 Cadence BOM 对比 HQ BOM。

接口：

- `/api/bom_compare/generic_sheets`：读取左右文件 Sheet 和表头，自动推荐匹配键。
- `/api/bom_compare/generic`：执行比对并导出报告。

处理流程：

1. 上传左右两份 BOM。
2. 分别选择 Sheet 和表头行。
3. 后端读取两边表头。
4. `_detect_common_key()` 根据常见字段推荐匹配键，例如位号、REFDES、客户料号、HQ 料号、PN、型号。
5. 前端提供同名字段和自定义字段映射。
6. 后端按匹配键加载左右数据，空键跳过，重复键记录。
7. 逐个匹配键比较字段值。
8. 统计仅左侧存在、仅右侧存在、字段变更、一致、重复键和空键数量。
9. 输出差异报告。

## Cadence BOM 对比 HQ BOM

该功能复用通用比对接口，`compare_type` 为 `cadence_hq`。默认匹配键检测优先考虑 `REFDES/reference` 对 HQ BOM 的 `位号`。

## 同项目 HQ BOM 版本对比

接口：

- `/api/bom_compare/local_sheets`：读取标准 HQ BOM 的 Sheet 和表头。
- `/api/bom_compare/hq_version`：比对基准版本和对比版本。

标准 HQ BOM 校验：

- 第 1-2 行为项目信息。
- 第 3 行为表头。
- 必须包含：序号、料号、型号、物料描述、单耗、替代关系、位号、生产厂家。

比对流程：

1. 上传基准版本和对比版本 HQ BOM。
2. 校验两份文件是否为标准 HQ BOM。
3. 读取元信息，例如版本、BOM 名称、项目配置名。
4. 按用户选择的匹配键读取数据。
5. 对比键集合，判断新增、删除。
6. 对共同键比较用户选择的字段，判断变更或未变更。
7. 记录重复键。
8. 输出版本差异报告。

## 差异类型

通用比对：仅左侧存在、仅右侧存在、字段变更、一致。

HQ 版本比对：新增、删除、变更、未变更。

## 输出报告

通用比对输出：差异总览、差异明细、仅左侧存在、仅右侧存在、字段变更、重复键。

HQ 版本比对输出：差异总览、差异明细、新增物料、删除物料、变更物料、重复料号。

## 输出文件名

- `客户BOM对比HQ_BOM_<uid>.xlsx`
- `Cadence_BOM对比HQ_BOM_<uid>.xlsx`
- `HQ_BOM版本差异_<uid>.xlsx`

## 异常处理

缺少文件、表头行非法、未选择匹配键、未选择比对字段、HQ BOM 格式不标准都会返回明确错误。
