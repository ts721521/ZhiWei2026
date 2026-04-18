# office_converter.py 拆分计划（2026-02-24）

## 背景

`office_converter.py` 当前约 7k+ 行，职责集中在单文件中，已出现以下维护成本：

- 变更影响面大，回归风险高。
- 新功能接入点不清晰，排查耗时。
- 单元测试很难做到按职责隔离。

目标是在不破坏现有 CLI/GUI 行为的前提下，完成“低风险、可回滚”的分层拆分。

## 目标与非目标

### 目标

- 降低单文件复杂度，明确模块边界。
- 保持 `office_converter.py` 对外入口兼容。
- 每个阶段都可独立回归并可随时停止。
- 让测试可以按模块增量扩展。

### 非目标

- 本计划不改动现有功能语义。
- 不在同一阶段引入新业务能力。
- 不一次性重写全部逻辑。

## 目标结构（建议）

建议新增 `converter/` 包，逐步承接逻辑：

- `converter/config.py`
  - 配置加载、默认值、运行参数规范化。
- `converter/models.py`
  - 运行态数据结构（计划、结果、统计、错误信息）。
- `converter/errors.py`
  - 错误分类与可重试策略。
- `converter/scan.py`
  - 源目录扫描、过滤、去重、增量判定入口。
- `converter/convert.py`
  - Office 转换执行（含并发/断点恢复协调）。
- `converter/merge.py`
  - PDF 合并、索引页、映射文件输出。
- `converter/delivery.py`
  - `_LLM_UPLOAD`、manifest、去重与产物汇总。
- `converter/mshelp.py`
  - CAB/MSHelp 模式流程。
- `converter/runtime.py`
  - `OfficeConverter` 编排层（调用上述模块）。

`office_converter.py` 初期只做兼容入口和桥接，最终变为薄 facade。

## 分阶段实施

## Phase 0：基线冻结与保护

- 固化回归命令：`python -m unittest discover -s tests -p "test_*.py" -v`。
- 增加最小烟囱回归：
  - convert_only
  - merge_only
  - convert_then_merge
  - task runtime config 路径
- 记录当前关键产物路径契约（`_MERGED`、`_AI`、`_FAILED_FILES`）。

完成标准：

- 基线测试稳定通过。
- 输出目录结构与文件命名行为有文档化契约。

## Phase 1：抽离“纯函数/无状态”模块

- 先拆 `errors.py`（错误分类）、`models.py`（结构体）、`config.py`（默认配置与规范化函数）。
- 仅移动代码，不改语义；旧调用点改为新模块导入。

完成标准：

- 无功能变更。
- 全量测试通过。
- 新模块具备最小单测覆盖。

## Phase 2：抽离扫描与增量判定

- 将扫描、过滤、增量差异计算迁入 `scan.py`。
- 统一输出扫描结果对象，减少运行中共享可变状态。

完成标准：

- 增量相关现有测试全绿（含 skip/retry 路径）。
- 扫描结果在 CLI/GUI 路径一致。

## Phase 3：抽离转换执行层

- 将串行/并发转换与断点恢复迁入 `convert.py`。
- `runtime.py` 只负责编排和生命周期。

完成标准：

- 并发与断点恢复行为与历史一致。
- 失败文件记录与报告行为不回退。

## Phase 4：抽离合并与交付层

- `merge.py` 承接 PDF 合并、映射、索引输出。
- `delivery.py` 承接 `_LLM_UPLOAD`、manifest、去重与汇总。

完成标准：

- 合并产物命名、索引、映射字段保持兼容。
- upload 清单字段和路径契约不变。

## Phase 5：抽离 MSHelp 与最终瘦身

- 将 MSHelp 逻辑迁入 `mshelp.py`。
- `office_converter.py` 收敛为：
  - 参数入口
  - `OfficeConverter` facade
  - 对外兼容导出

完成标准：

- 业务功能与 CLI 参数对外兼容。
- `office_converter.py` 控制在“入口+兼容层”规模。

## 测试策略

- 每阶段必须执行全量 `unittest`。
- 每抽离一个模块，补该模块最小单测，不等最后一起补。
- 关键路径优先新增回归：
  - 任务模式运行态合成 -> converter 执行
  - 合并映射定位链路
  - 增量更新包与 manifest

## 风险与回滚

- 风险：拆分中引入隐式状态差异，导致 GUI 与 CLI 行为不一致。
- 控制：
  - 每阶段小步提交，禁止跨阶段混改。
  - 通过 facade 保持旧 API，不一次性切断旧路径。
  - 出现回归时只回滚当前阶段，不回滚已验证阶段。

## 交付节奏（建议）

- 第 1 周：Phase 0-1
- 第 2 周：Phase 2
- 第 3 周：Phase 3
- 第 4 周：Phase 4-5 + 文档收尾

## 验收标准

- 全量自动化测试持续通过。
- 关键手工冒烟路径可复现通过。
- 主文件职责明显收敛，模块边界清晰。
- README/使用说明/测试报告同步更新。

