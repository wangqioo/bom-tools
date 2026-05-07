# 本轮 review 摘要

## 已修复

1. 恢复本地 UI
   - 新增 `pstx_local_ui.py`。
   - 本地 UI 是 Web UI 的桌面套壳，不再维护独立 Tk 逻辑。
   - `python pstx_analyzer.py` 兼容转入本地 UI。

2. 消除页码逻辑重复
   - `pstx_analyzer.py` 中旧的 page 解析函数已裁剪。
   - 页码解析统一委托到 `pstx_page_logic.py`。

3. 修复 import 阶段副作用
   - `pstx_web.py` 不再 import 时立即检查 / 安装 Flask。
   - Flask 只在创建 Web app 时加载。

4. 强化 `module_order` 映射
   - 两层复用继续支持。
   - 多层复用优先匹配最深层 `module_order` key。
   - `P_PATH` 存在时不再回退到 `C_PATH` 匹配。
   - 子模块本地页超过 `page_count` 时保持空映射并标记状态。

5. 恢复真实页反向冲突检查
   - 当多个逻辑页指向同一个真实页时，映射检查会标记为 `真实页对应多个逻辑页`。

6. 裁剪文档
   - 移除了历史 analysis / changelist / workflow 堆叠文档。
   - 保留 README、架构说明和本轮 review 摘要。

## 验证

- 语法检查通过。
- 回归测试通过：`Ran 86 tests ... OK`。
