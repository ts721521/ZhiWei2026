# ZhiWei 2026 代码审查与优化建议

基于对 2026 仓库的代码审查，从功能、Bug、性能、代码质量与可维护性等多方面给出优化建议。

## 一、Bug 与健壮性

### 1.1 异常处理过宽（except Exception 吞错）

- **现状**：全仓库大量使用 `except Exception:`，部分位置未记录日志或未区分可恢复/不可恢复错误，易掩盖真实故障。
- **重点文件**（建议优先收紧）：
  - [office_converter.py](2026/office_converter.py)：约 14 处（含导入 HAS_* 与运行时逻辑）。
  - [office_gui.py](2026/office_gui.py)：约 12 处。
  - [gui/mixins/gui_task_workflow_mixin.py](2026/gui/mixins/gui_task_workflow_mixin.py)：约 27 处。
  - [gui/mixins/gui_run_mode_state_mixin.py](2026/gui/mixins/gui_run_mode_state_mixin.py)：约 30 处。
  - [converter/run_workflow.py](2026/converter/run_workflow.py)：342、348 行等。
- **建议**：
  - 在关键路径（转换执行、任务启停、配置保存）改为捕获具体异常类型（如 `OSError`、`json.JSONDecodeError`），并至少 `logging.exception` 或 `logging.error` 记录。
  - 导入可选依赖（如 chromadb、bs4、docx）处的 `except Exception` 可保留，但建议统一为 `except ImportError` 或明确注释“仅用于可选依赖探测”。

### 1.2 路径与编码一致性

- **现状**：大量使用 `os.path.join`，少数模块使用 `pathlib`；[converter/cab_extract.py](2026/converter/cab_extract.py) 中混用 `encoding="gbk"` 与 `encoding="utf-8"`，在非中文 Windows 或跨平台时可能出问题。
- **建议**：
  - 新代码与重构时统一采用 `pathlib.Path` 做路径拼接与 `open()`，便于跨平台与 Unicode 路径。
  - 对 cab_extract 等明确依赖系统编码的场景，在文档或代码注释中说明编码策略，并考虑用 `locale.getpreferredencoding()` 或配置项替代硬编码 gbk。

### 1.3 测试历史不稳定项

- **现状**：[docs/test-reports/TEST_REPORT_SUMMARY.md](2026/docs/test-reports/TEST_REPORT_SUMMARY.md) 记录过 `test_gui_task_mode.TestGuiTaskTabVisibility.test_classic_mode_hides_task_tab` 断言失败（经典模式任务 Tab 可见性），以及 4 个 pypdf 相关测试曾因未安装 pypdf 报错。
- **建议**：
  - 在 CI/本地文档中明确要求安装 [requirements.txt](2026/requirements.txt) 及“按需”依赖（pypdf、openpyxl 等）后再跑全量测试；或在测试中对可选依赖做 skip 条件，避免误报。
  - 对 `test_classic_mode_hides_task_tab` 做一次稳定复现与修复（UI 状态或测试时机），并记录在 TEST_REPORT_SUMMARY 中。

---

## 二、功能与完整性

### 2.1 配置校验与 Schema

- **现状**：[converter/config_defaults.py](2026/converter/config_defaults.py) 仅做 `setdefault` 规范化，无严格类型或取值范围校验；[tests/test_default_config_schema.py](2026/tests/test_default_config_schema.py) 只验证默认 config 的 key 存在性。
- **建议**：
  - 在加载用户配置后（如 [converter/config_load.py](2026/converter/config_load.py) 或 GUI 应用配置前）增加一层校验：数值范围（如 `parallel_workers`、`timeout_seconds`）、枚举值（如 `run_mode`、`content_strategy`）、路径存在性（可选）。可复用或引用现有 `test_default_config_schema` 中的期望结构，抽成共享 schema 或校验函数。
  - 校验失败时返回明确错误信息并记录日志，避免无效配置进入运行流程。

### 2.2 文档同步校验脚本

- **现状**：[AGENTS.md](2026/AGENTS.md) 要求变更关键路径时同步更新文档，并建议使用 `scripts/check_doc_sync.py` 做提交前校验；该脚本已存在。
- **建议**：
  - 在 README 或 AGENTS.md 中明确写出“提交前执行”示例命令（如 `python scripts/check_doc_sync.py --staged`），并在贡献流程中推荐接入 pre-commit 或 CI，确保文档同步规则可执行。

### 2.3 错误反馈与可观测性

- **现状**：转换失败、重试、跳过等依赖日志与少量 UI 提示；非致命错误有 [test_nonfatal_ui_error_reporting](2026/tests/test_nonfatal_ui_error_reporting.py) 覆盖。
- **建议**：
  - 对“失败报告”“重试列表”等关键产出，在 UI 或日志中提供明确入口（如“打开失败报告”“导出重试列表”），便于用户自助排查。
  - 考虑在关键阶段（扫描结束、批处理开始/结束、Chroma 导出）打点结构化日志或简单指标（如耗时、文件数），便于后续做性能与稳定性分析。

---

## 三、性能

### 3.1 大列表与内存

- **现状**（来自探索结论）：
  - [converter/scan_convert_candidates.py](2026/converter/scan_convert_candidates.py)：`os.walk` 结果一次性 append 成列表返回，源目录极大时占用高。
  - [converter/collect_index.py](2026/converter/collect_index.py)：多源根全量收集后再去重、写 Excel，`unique_records` 整份在内存。
  - [converter/chromadb_docs.py](2026/converter/chromadb_docs.py)：对每个 md 路径读全文件再分块，全部 append 到 `docs` 列表。
- **建议**：
  - `scan_convert_candidates`：改为生成器接口（yield 每个匹配路径），或在内部按目录/批次 yield，由调用方（如 [converter/run_workflow.py](2026/converter/run_workflow.py)）按批消费；若调用方强依赖“总数量”，可先做一次轻量计数再流式处理。
  - `collect_index`：在保持当前“全量去重”语义的前提下，可考虑按源根或按批写入 Excel（追加写入或分批写 sheet），降低单次内存峰值。
  - `chromadb_docs`：按文件或按批 yield 文档块，与 [converter/chromadb_export.py](2026/converter/chromadb_export.py) 已有的 batch_size=200 对接，实现流式或分批构建，避免一次性构建完整 `docs` 列表。

### 3.2 并行任务提交方式

- **现状**：[converter/batch_parallel.py](2026/converter/batch_parallel.py) 对 `file_list` 一次性 `executor.submit` 全部任务，若文件数极大会产生大量 pending futures 与回调。
- **建议**：改为按批提交（例如每 N 个文件一批 submit，结合 `as_completed` 按批消费），或限制最大 pending 数量，以控制内存与调度开销；需保持与现有 checkpoint 间隔、线程 CoInitialize 等行为兼容。

### 3.3 并发模型说明

- **现状**：全部使用多线程（ThreadPoolExecutor + 单文件 Thread），无 multiprocessing；与 COM/Office 单线程单元兼容，但受 GIL 限制。
- **建议**：在文档或注释中明确“当前为线程模型，因 Office COM 限制；若未来引入多进程需考虑 COM 与序列化”，避免后续误用多进程导致难以排查的问题。

---

## 四、代码质量与可维护性

### 4.1 office_converter 与 GUI 耦合

- **现状**：GUI 大量从 [office_converter.py](2026/office_converter.py) 导入常量和函数，office_converter 作为“大接口”，变更影响面大；[docs/plans/2026-02-24-office-converter-split-handover.md](2026/docs/plans/2026-02-24-office-converter-split-handover.md) 已记录 office_converter 行数控制在 2000 以内。
- **建议**：
  - 将“模式/策略/常量”等 GUI 与 converter 共用部分抽到独立模块（如 `converter/constants.py` 已存在，可扩展），GUI 仅从该模块与少量 office_converter 入口导入，减少对单文件大类的依赖。
  - 继续按交接文档对 office_converter 做“薄委托”拆分，优先处理 `merge_markdowns`、`_run_merge_mode_pipeline`、`_convert_on_mac` 等 50 行以上函数。

### 4.2 路径处理统一

- **建议**：新代码与重构时统一使用 `pathlib.Path`，逐步替换热点路径上的 `os.path.join`，减少拼接错误与编码问题；可先在 `converter/run_workflow.py`、`converter/scan_convert_candidates.py`、`converter/merge_pdfs.py` 等新改动较多的模块试点。

### 4.3 类型注解与静态检查

- **建议**：对 `task_manager`、`converter/checkpoint_utils`、`converter/config_defaults` 等无 GUI 依赖的模块，逐步增加类型注解并在 CI 中启用 pyright 或 mypy（仅对已标注文件或目录），降低接口误用和回归风险。

### 4.4 日志级别与一致性

- **现状**：各模块使用 `logging.debug/info/warning/error` 不统一，部分关键分支无日志。
- **建议**：约定“用户可见进度/结果用 info，可恢复异常用 warning 并带上下文，不可恢复用 error/exception”；对转换入口、批处理起止、失败分支补充必要日志，便于线上或用户环境排查。

---

## 五、优先级与实施顺序建议

| 优先级 | 方向 | 建议实施内容 |
|--------|------|--------------|
| P0 | Bug/健壮性 | 关键路径异常收紧（office_converter/run_workflow/execution mixin）+ 测试稳定性（pypdf skip/classic_mode 断言） |
| P1 | 功能 | 配置加载后 schema/范围校验 + 文档同步脚本接入说明与 CI |
| P2 | 性能 | scan_convert_candidates 生成器化 + batch_parallel 分批提交 |
| P3 | 质量 | 路径 pathlib 试点 + office_converter 继续薄委托 + 日志约定 |

---

## 六、无需修改或低优先级项

- **文件资源**：未发现“open 不用 with 也不 close”的用法；可保持现状。
- **可变默认参数**：未发现 `def f(x=[])` 等反模式；无需改动。
- **循环依赖**：当前为单向链 `office_gui → office_converter → converter`，无循环；仅需在拆分时注意不要引入反向依赖。

以上建议可根据排期与人力按优先级分阶段落地，每轮改动后按 AGENTS.md 执行全量测试并更新 TEST_REPORT_SUMMARY 与相关交接文档。
