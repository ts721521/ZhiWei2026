# V2 / V3 开发规划（待 V1 稳定后启动）

最后更新：2026-02-11  
状态：Draft（评审版）

## 1. 背景与前提
- 当前已完成 V1：性能可观测（阶段耗时统计）+ Office 进程复用/周期重启开关。
- 本文仅规划 V2、V3，不触发立即开发；是否启动由 V1 运行稳定性决定。
- 近期基线样本（单次日志）约为：89 文件 / 736.54s（约 8.28s/文件）。

## 2. 启动门槛（V1 稳定性验收）
满足以下条件后，才建议启动 V2：
1. 连续 5 次真实批量任务运行（每次 >= 50 文件）无崩溃。
2. 超时率 <= 2%，失败率 <= 1%（剔除损坏源文件）。
3. 开启 `office_reuse_app=true` 后，结果文件与 V1 之前版本无功能回归。
4. 日志中的性能统计完整输出（扫描/转换/后处理/总耗时）。

## 3. V2 规划（低风险提速，默认主线）

### 3.1 目标
- 在保持 Office 转换串行的前提下，通过“后处理并发”提升总体吞吐。
- 默认保持稳定优先，提供一键关闭开关，确保可快速回退。

### 3.2 范围
- 包含：
  - Markdown 导出并发化。
  - Excel JSON / Records JSON / 质量报告等后处理任务异步化。
  - 队列与线程池调度（仅后处理，不并发 Office COM 转换）。
  - 线程安全改造（共享列表、统计计数、日志上下文）。
- 不包含：
  - Office COM 多线程/多进程并发转换（属于 V3）。

### 3.3 架构方案
- 主线程（转换线程）继续执行 `run_batch` 串行 Office 转换。
- 新增后处理任务队列（Producer/Consumer）：
  - Producer：每个成功 PDF 产生后处理任务。
  - Consumer：`ThreadPoolExecutor` 处理 Markdown/结构化导出。
- 增加线程安全措施：
  - 对 `generated_*`、`markdown_quality_records` 等共享集合加锁。
  - 对 `perf_metrics` 的并发累计统一封装（保留现有 `_add_perf_seconds`）。
- 关键配置（建议新增）：
  - `enable_async_postprocess`（默认 `true`）。
  - `postprocess_workers`（默认 `2`，范围建议 `1~4`）。
  - `postprocess_queue_size`（默认 `64`）。

### 3.4 里程碑
1. M1 设计与骨架：
   - 新增后处理执行器与任务模型，默认 `workers=1` 行为等价。
2. M2 功能接入：
   - 接入 Markdown/JSON/质量报告路径，补齐线程安全。
3. M3 观测增强：
   - 新增队列积压、后处理耗时分项日志。
4. M4 验证与发布：
   - 压测、回归、故障回退演练。

### 3.5 验收标准
1. 功能一致性：输出文件数量、命名、索引字段与串行模式一致。
2. 稳定性：连续 3 次 100+ 文件任务无死锁、无数据丢失。
3. 性能收益：PDF 占比较高场景下总耗时改善目标 20%~50%（与硬件相关）。
4. 回退有效：`enable_async_postprocess=false` 后恢复串行后处理。

### 3.6 风险与缓解
- 风险：共享状态竞争导致统计不准或记录错序。  
  缓解：统一锁与不可变任务快照。
- 风险：后处理过快放大 IO 压力。  
  缓解：限制 `workers<=4` + 有界队列。
- 风险：日志顺序可读性下降。  
  缓解：任务 ID + 文件名双标识。

## 4. V3 规划（实验性高收益，高风险）

### 4.1 目标
- 探索 Office 转换并行能力，提升 Office 文件占比高场景的吞吐上限。
- 明确定位为“实验模式”，默认关闭。

### 4.2 范围
- 包含：
  - 多进程转换调度（非多线程）。
  - 每进程独立 COM 初始化、独立临时目录、独立 Office 生命周期。
  - 结果聚合、失败重试与自动降级策略。
- 不包含：
  - 默认主线启用；V3 必须通过显式开关开启。

### 4.3 架构方案
- Orchestrator 进程：
  - 扫描与任务切分。
  - 分发到 `ProcessPoolExecutor`（建议上限 `office_workers=2`）。
- Worker 进程：
  - 独立加载配置快照。
  - 独立 `pythoncom.CoInitialize()` / `CoUninitialize()`。
  - 独立 sandbox 与失败目录隔离。
- 关键限制：
  - 并发模式下禁用全局 `taskkill` 策略（避免进程互杀）。
  - `office_workers` 超过阈值自动拒绝（例如 >2）。
  - 异常率超阈值自动降级回串行。

### 4.4 关键配置（建议新增）
- `enable_experimental_office_parallel`（默认 `false`）。
- `office_workers`（默认 `1`，实验最大 `2`）。
- `office_parallel_auto_fallback`（默认 `true`）。
- `office_parallel_fail_threshold`（默认 `3`，连续失败触发降级）。

### 4.5 里程碑
1. M1 PoC：
   - 仅支持 Word 或单一类型并发验证，建立稳定性边界。
2. M2 类型扩展：
   - 增加 Excel/PPT 支持，补齐结果聚合。
3. M3 保护机制：
   - 自动降级、熔断、重试与详细遥测。
4. M4 灰度验证：
   - 小范围用户可选试用，收集崩溃率/收益数据。

### 4.6 验收标准
1. 稳定性：实验模式下连续 3 次 100+ 文件任务可完成且无崩溃。
2. 回退：任意时刻可自动或手动退回串行，不影响结果完整性。
3. 性能：Office 文件占比高场景目标提升 30%~100%（机器依赖强）。
4. 兼容性：WPS/MS 两种引擎至少覆盖一条稳定路径（先 WPS 或先 MS 二选一）。

### 4.7 风险与缓解
- 风险：COM 与 Office 自动化并发不稳定（卡死/句柄泄漏）。  
  缓解：多进程隔离、最大并发 2、自动熔断降级。
- 风险：全局清理策略误杀其他 worker。  
  缓解：并发模式改为“进程内清理”，禁用全局杀进程。
- 风险：调试复杂度高。  
  缓解：统一任务 ID、分进程日志、最小可复现实验集。

## 5. 建议排期（仅参考）
- V2：5~8 个工作日（含回归与压测）。
- V3：10~15 个工作日（含 PoC、灰度、熔断/回退机制）。

## 6. 决策建议
1. 先观察 V1 一周运行数据（按第 2 章门槛评估）。
2. 若 V1 稳定，优先开发 V2（低风险、可快速见效）。
3. V2 上线稳定后，再决定是否投入 V3（实验模式）。


---

## 7. V1.1 Pre-Phase (Must Complete Before V2/V3)
To align with latest user priorities, V2/V3 should be gated by a V1.1 usability milestone:

1. LLM upload output must be centralized in a single folder (`_LLM_UPLOAD`).
2. Sandbox free-space protection must be available before large runs.
3. Both capabilities must pass real-world large-batch verification before V2 parallel optimization starts.

Recommendation:
- Treat this as `V1.1` and deliver before enabling any new V2/V3 performance branch.
