# office_converter 拆分交接文档（2026-02-24）

## 1. 停止标准与当前状态

建议“双阈值”作为一期完成标准：

- 硬阈值：`office_converter.py <= 2000` 行。
- 质量阈值：除 `__init__` 外，无 `>120` 行函数。

当前状态（2026-02-24 本轮结束）：

- `office_converter.py`：`1990` 行。
- 结论：已达到硬阈值；质量阈值也满足。

## 2. 本轮新增拆分（本次续拆）

- `converter/error_recording.py`
  - `record_detailed_error`
  - `record_scan_access_skip`
- `converter/index_runtime.py`
  - `write_merge_map`
  - `append_conversion_index_record`
  - `write_conversion_index_workbook`
- `converter/markdown_render.py`
  - 新增 `table_to_markdown_lines`

`office_converter.py` 本轮改为委托：

- `record_detailed_error`
- `_record_scan_access_skip`
- `_write_merge_map`
- `_append_conversion_index_record`
- `_write_conversion_index_workbook`
- `_table_to_markdown_lines`

## 3. 本轮新增测试

- `tests/test_converter_error_recording_module.py`
- `tests/test_converter_index_runtime_module.py`
- `tests/test_converter_markdown_render_module.py`（补充表格渲染与委托测试）

## 4. 全量回归结果（必须记录）

执行命令：

```bash
python -m unittest discover -s tests -p "test_*.py" -v
```

结果（2026-02-24）：

- `Ran 229 tests in 6.692s`
- `OK`

## 5. 当前大函数概览（便于下一轮）

按体量前几项：

1. `__init__`：101 行
2. `merge_markdowns`：55 行
3. `_run_merge_mode_pipeline`：54 行
4. `_convert_on_mac`：53 行
5. `confirm_config_in_terminal`：44 行

说明：当前体量已达一期目标；后续如继续拆分，可优先处理上面 2~5 项（`__init__` 可最后处理）。

## 6. 续拆固定流程（保持不变）

1. 新建模块函数（依赖注入）。
2. `office_converter.py` 保留薄委托。
3. 新增测试：模块核心行为 + 委托测试。
4. 先定向，再全量 `unittest discover`。
5. 同步更新本交接文档与 `docs/test-reports/TEST_REPORT_SUMMARY.md`。

---

## 7. 2026-02-24 本轮续拆（merge_markdowns / merge_mode_pipeline / trace_map 聚合）

### 7.1 本轮新增拆分

- `converter/merge_markdown.py`
  - `merge_markdowns`
- `converter/merge_mode_pipeline.py`
  - `run_merge_mode_pipeline`
- `converter/traceability.py`
  - 新增 `write_trace_map_for_converter`

`office_converter.py` 本轮改为薄委托：

- `merge_markdowns` -> `merge_markdowns_impl`
- `_run_merge_mode_pipeline` -> `run_merge_mode_pipeline_impl`
- `_write_trace_map` -> `write_trace_map_for_converter_impl`

### 7.2 本轮新增/更新测试

- 新增 `tests/test_converter_merge_markdown_module.py`
- 新增 `tests/test_converter_merge_mode_pipeline_module.py`
- 更新 `tests/test_converter_traceability_module.py`

### 7.3 回归结果

执行命令：

```bash
python -m unittest discover -s tests -p "test_*.py" -v
```

结果（2026-02-24）：

- `Ran 245 tests in 7.709s`
- `OK`

### 7.4 当前状态

- `office_converter.py`：`1986` 行（重新回到 `<= 2000` 阈值内）。
- 除 `__init__` 外，当前最大函数仍显著低于早期高风险阶段；后续可继续优先拆分：
  - `_convert_on_mac`
  - `confirm_config_in_terminal`
  - `_get_local_app`

---

## 8. 2026-02-24 本轮续拆（_convert_on_mac / confirm_config_in_terminal / _get_local_app）

### 8.1 本轮新增拆分

- `converter/mac_convert.py`
  - `convert_on_mac`
- `converter/config_terminal.py`
  - `confirm_config_in_terminal`
- `converter/local_office_app.py`
  - `get_local_app`

`office_converter.py` 本轮改为薄委托：

- `_convert_on_mac` -> `convert_on_mac_impl`
- `confirm_config_in_terminal` -> `confirm_config_in_terminal_impl`
- `_get_local_app` -> `get_local_app_impl`

### 8.2 本轮新增测试

- `tests/test_converter_mac_convert_module.py`
- `tests/test_converter_config_terminal_module.py`
- `tests/test_converter_local_office_app_module.py`

### 8.3 回归结果

执行命令：

```bash
python -m unittest discover -s tests -p "test_*.py" -v
```

结果（2026-02-24）：

- `Ran 251 tests in 7.778s`
- `OK`

### 8.4 当前状态（更新）

- `office_converter.py`：`1890` 行。
- 体量前列函数：
  1. `__init__`：107 行
  2. `_extract_sheet_charts`：36 行
  3. `_write_excel_structured_json_exports`：34 行
  4. `_setup_excel_pages`：33 行
  5. `convert_logic_in_thread`：32 行

---

## 9. 2026-02-24 本轮续拆（_extract_sheet_charts / _setup_excel_pages / _write_excel_structured_json_exports）

### 9.1 本轮新增拆分

- `converter/excel_chart_extract.py`
  - `extract_sheet_charts`
- `converter/excel_page_setup.py`
  - `setup_excel_pages`
- `converter/excel_json_batch.py`
  - `write_excel_structured_json_exports`

`office_converter.py` 本轮改为薄委托：

- `_extract_sheet_charts` -> `extract_sheet_charts_impl`
- `_setup_excel_pages` -> `setup_excel_pages_impl`
- `_write_excel_structured_json_exports` -> `write_excel_structured_json_exports_impl`

### 9.2 本轮新增测试

- `tests/test_converter_excel_chart_extract_module.py`
- `tests/test_converter_excel_page_setup_module.py`
- `tests/test_converter_excel_json_batch_module.py`

### 9.3 回归结果

执行命令：

```bash
python -m unittest discover -s tests -p "test_*.py" -v
```

结果（2026-02-24）：

- `Ran 257 tests in 11.077s`
- `OK`

### 9.4 当前状态（更新）

- `office_converter.py`：`1818` 行。
- 体量前列函数：
  1. `__init__`：107 行
  2. `convert_logic_in_thread`：32 行
  3. `_export_pdf_markdown`：30 行
  4. `_build_ai_output_path_from_source`：30 行
  5. `_write_chromadb_export`：29 行

---

## 10. 2026-02-24 本轮续拆（_export_pdf_markdown / _write_chromadb_export）

### 10.1 本轮新增拆分

- `converter/pdf_markdown_runtime.py`
  - `export_pdf_markdown_for_converter`
- `converter/chromadb_runtime.py`
  - `write_chromadb_export_for_converter`

`office_converter.py` 本轮改为薄委托：

- `_export_pdf_markdown` -> `export_pdf_markdown_for_converter_impl`
- `_write_chromadb_export` -> `write_chromadb_export_for_converter_impl`

### 10.2 本轮新增测试

- `tests/test_converter_pdf_markdown_runtime_module.py`
- `tests/test_converter_chromadb_runtime_module.py`

### 10.3 回归结果

执行命令：

```bash
python -m unittest discover -s tests -p "test_*.py" -v
```

结果（2026-02-24）：

- `Ran 261 tests in 7.324s`
- `OK`

### 10.4 当前状态（更新）

- `office_converter.py`：`1790` 行。
- 体量前列函数：
  1. `__init__`：107 行
  2. `convert_logic_in_thread`：32 行
  3. `_build_ai_output_path_from_source`：30 行
  4. `scan_excel_content_in_thread`：28 行
  5. `load_config`：28 行

---

## 11. 2026-02-24 本轮续拆（scan_excel_content_in_thread / convert_logic_in_thread / _build_ai_output_path_from_source / load_config）

### 11.1 本轮新增拆分

- `converter/excel_content_scan.py`
  - `scan_excel_content_in_thread`
- `converter/convert_thread_runtime.py`
  - `convert_logic_in_thread_for_converter`
- `converter/ai_paths_runtime.py`
  - `build_ai_output_path_from_source_for_converter`
- `converter/config_load.py`
  - `load_config`

`office_converter.py` 本轮改为薄委托：

- `scan_excel_content_in_thread` -> `scan_excel_content_in_thread_impl`
- `convert_logic_in_thread` -> `convert_logic_in_thread_for_converter_impl`
- `_build_ai_output_path_from_source` -> `build_ai_output_path_from_source_for_converter_impl`
- `load_config` -> `load_config_impl`

### 11.2 本轮新增测试

- `tests/test_converter_excel_content_scan_module.py`
- `tests/test_converter_convert_thread_runtime_module.py`
- `tests/test_converter_ai_paths_runtime_module.py`
- `tests/test_converter_config_load_module.py`

### 11.3 回归结果

执行命令：

```bash
python -m unittest discover -s tests -p "test_*.py" -v
```

结果（2026-02-24）：

- `Ran 269 tests in 7.794s`
- `OK`

### 11.4 当前状态（更新）

- `office_converter.py`：`1759` 行。
- 体量前列函数：
  1. `__init__`：107 行
  2. `_build_ai_output_path_from_source`：30 行
  3. `select_engine_mode`：26 行
  4. `get_target_path`：26 行
  5. `run`：25 行

---

## 12. 2026-02-24 本轮补齐（V6.0 风险项 8.1：markitdown 打包 MVP）

### 12.1 本轮新增内容

- 新增 `converter/markitdown_pack_probe.py`
  - `run_markitdown_probe`
- 新增 `scripts/markitdown_smoke_mvp.py`
  - 最小可运行 smoke 脚本（输入 Office 文件，输出 Markdown，返回码 0/1）
- 新增文档 `docs/notes/markitdown_pack_mvp.md`
  - PyInstaller 打包命令
  - 干净机验证清单
  - 失败兜底策略（docling/本地能力说明）

### 12.2 本轮新增测试

- `tests/test_converter_markitdown_pack_probe_module.py`

### 12.3 回归结果

执行命令：

```bash
python -m unittest discover -s tests -p "test_*.py" -v
```

结果（2026-02-24）：

- `Ran 271 tests in 7.447s`
- `OK`

### 12.4 V6.0 计划完成度核对

- 模块一（极速 MD）：已完成（直出 MD、_Knowledge_Bundle、流式拼接、仅 MD 模式）。
- 模块二（溯源锚点）：已完成（ZW-短 ID、MD frontmatter、trace_map.xlsx 增量更新、locate_source 兼容前缀）。
- 模块三（Prompt 包装）：已完成（模板类型 + Prompt_Ready.txt + 语料注入，流式拼接）。
- 模块四（GUI 同步）：已完成（3 个开关 + 模板下拉 + 前缀输入 + 配置持久化 + 运行态约束）。
- 风险项 8.1：代码与文档资产已补齐（MVP 脚本 + 打包流程），但“干净机实际运行验证”属于环境执行项，需在线下目标机完成一次验收。

---

## 13. 2026-02-24 本轮整理（GUI mixin 文件归档到 gui/mixins）

### 13.1 本轮调整

- 新建包结构：
  - `gui/__init__.py`
  - `gui/mixins/__init__.py`
- 将以下 GUI mixin 实现文件集中到 `gui/mixins/`：
  - `gui_config_compose_mixin.py`
  - `gui_config_dirty_mixin.py`
  - `gui_config_io_mixin.py`
  - `gui_config_logic_mixin.py`
  - `gui_config_save_mixin.py`
  - `gui_config_tab_mixin.py`
  - `gui_execution_mixin.py`
  - `gui_gdrive_mixin.py`
  - `gui_locator_mixin.py`
  - `gui_misc_ui_mixin.py`
  - `gui_profile_mixin.py`
  - `gui_runtime_status_mixin.py`
  - `gui_run_mode_state_mixin.py`
  - `gui_run_tab_mixin.py`
  - `gui_source_folder_mixin.py`
  - `gui_task_workflow_mixin.py`
  - `gui_tooltip_mixin.py`
  - `gui_tooltip_settings_mixin.py`
  - `gui_ui_shell_mixin.py`
- 在原根目录文件保留兼容转发（`from gui.mixins.xxx import *`），避免外部旧导入中断。
- `office_gui.py` 统一使用 `from gui.mixins...` 新路径导入。
- 更新路径敏感测试（读取源码文件路径）以优先读取新路径，并保留旧路径回退：
  - `tests/test_gui_run_mode_state_behavior.py`
  - `tests/test_task_selection_tab_preserve.py`
  - `tests/test_gui_task_mode.py`

### 13.2 回归结果

执行命令：

```bash
python -m unittest discover -s tests -p "test_*.py" -v
```

结果（2026-02-24）：

- `Ran 271 tests in 7.932s`
- `OK`

---

## 14. 2026-02-24 本轮整理（GUI 导入统一 + 根目录壳文件收口）

### 14.1 本轮调整

- 全仓统一旧导入：
  - 将测试中 `from gui_xxx_mixin import ...` 改为 `from gui.mixins.gui_xxx_mixin import ...`。
- 清理路径回退逻辑：
  - 去除测试中对根目录 `gui_xxx_mixin.py` 的 fallback，仅保留 `gui/mixins/*.py` 真实路径。
- 根目录壳文件收口：
  - 由于环境策略禁止直接删除，已将根目录 19 个 `gui_*_mixin.py` 兼容壳文件迁移至 `gui/legacy_shims/`，根目录不再散落同名文件。
- 测试稳定性修正：
  - `tests/test_task_detail_render.py`：修复字符串语法错误并改为稳定 ASCII 断言。
  - `tests/test_task_binding_summary.py`：修正历史乱码断言文本为实际返回文案。

### 14.2 回归结果

执行命令：

```bash
python -m unittest discover -s tests -p "test_*.py" -v
```

结果（2026-02-24）：

- `Ran 271 tests in 6.717s`
- `OK`

---

## 15. 2026-02-24 本轮整理（legacy_shims 归档出 gui 运行目录）

### 15.1 本轮调整

- 验证完成：仓内无代码再引用 `gui/legacy_shims`。
- 受环境策略限制（禁止删除文件），将 `gui/legacy_shims/` 整体归档移动至：
  - `docs/archive/gui_legacy_shims_2026-02-24/`
- 结果：`gui/` 目录仅保留
  - `gui/mixins/`
  - `gui/__init__.py`

### 15.2 回归结果

执行命令：

```bash
python -m unittest discover -s tests -p "test_*.py" -v
```

结果（2026-02-24）：

- `Ran 271 tests in 6.688s`
- `OK`

---

## 16. 2026-02-24 本轮整理（mixin 包入口统一导出）

### 16.1 本轮调整

- `gui/mixins/__init__.py`
  - 新增统一导出：集中导入 19 个 mixin class，并声明 `__all__`。
- `office_gui.py`
  - 从逐模块导入改为从 `gui.mixins` 统一导入，降低导入分散度，后续迁移只需维护包入口。

### 16.2 回归结果

执行命令：

```bash
python -m unittest discover -s tests -p "test_*.py" -v
```

结果（2026-02-24）：

- `Ran 271 tests in 6.964s`
- `OK`

---

## 17. 2026-02-24 本轮治理（文档同步规则落地 + 自动校验）

### 17.1 本轮调整

- 重写 `AGENTS.md`，将“文档与代码一致性”设为强制流程：
  - 定义代码变更触发条件。
  - 强制同步更新 `docs/plans/2026-02-24-office-converter-split-handover.md` 与 `docs/test-reports/TEST_REPORT_SUMMARY.md`。
  - 架构迁移类变更必须同步更新 `AGENTS.md`。
  - 固化回归命令与记录项（命令、Ran N、OK/FAILED、变更摘要）。
- 新增自动校验脚本：`scripts/check_doc_sync.py`
  - 支持 `--changed ...` 显式文件列表校验。
  - 支持 `--staged`（有 git 时）校验暂存区改动。
  - 校验失败会明确输出缺失文档项。
- 新增单测：`tests/test_doc_sync_checker.py`
  - 覆盖通过/缺文档/缺 AGENTS 三类场景。

### 17.2 回归结果

执行命令：

```bash
python -m unittest discover -s tests -p "test_*.py" -v
```

结果（2026-02-24）：

- `Ran 274 tests in 6.978s`
- `OK`

---

## 18. 2026-02-24 本轮治理（文档同步门禁接入 pre-commit + CI）

### 18.1 本轮调整

- 新增本地 pre-commit 模板：
  - `.githooks/pre-commit`
  - 提交前执行：`python scripts/check_doc_sync.py --staged`
- 新增钩子安装脚本：
  - `scripts/install_git_hook.py`
  - 支持向上查找 `.git` 目录并安装到对应 `hooks/pre-commit`。
- 新增 CI 工作流：
  - `.github/workflows/quality-gate.yml`
  - 流程包含：
    1. 计算变更文件
    2. 执行 `check_doc_sync.py` 文档同步校验
    3. 执行全量单测 `python -m unittest discover -s tests -p "test_*.py" -v`
- 新增执行说明文档：
  - `docs/notes/doc_sync_enforcement.md`
- 更新规范文档：
  - `AGENTS.md` 增补 pre-commit 与 CI 启用说明。
- 新增测试：
  - `tests/test_install_git_hook_script.py`（覆盖 `.git` 向上查找逻辑）

### 18.2 本轮校验与回归

执行：

```bash
python scripts/check_doc_sync.py --changed office_gui.py docs/plans/2026-02-24-office-converter-split-handover.md docs/test-reports/TEST_REPORT_SUMMARY.md AGENTS.md
```

结果：

- `OK: documentation sync check passed`

执行：

```bash
python -m unittest discover -s tests -p "test_*.py" -v
```

结果（2026-02-24）：

- `Ran 276 tests in 11.165s`
- `OK`

---

## 19. 2026-02-24 审查建议落地（性能与可维护性）

### 19.1 审查建议评估与采纳结论

- 已采纳（高收益、低风险、可立即验证）：
  1. `batch_parallel` 限制 pending 任务数量，避免一次性 `submit` 全量文件导致高内存占用。
  2. `scan_convert_candidates` 新增生成器接口，支持后续流式消费。
- 暂缓（需更大范围联动或策略确认）：
  1. 全仓 `except Exception` 收敛到细粒度异常（范围较大，需分批推进并补齐行为测试）。
  2. 全量 `pathlib.Path` 迁移（涉及广泛路径行为，建议按模块滚动迁移）。
  3. 配置 schema 强校验（需先确定运行时兼容策略与错误提示 UX）。

### 19.2 本轮代码改动

- `converter/batch_parallel.py`
  - 新增 `parallel_max_pending` 配置（默认 `parallel_workers * 2`，且不低于 `parallel_workers`）。
  - 提交策略改为“限流提交 + 完成即消费”，避免全量 pending futures。
  - 修正 elapsed 计时：从任务提交时开始计时，避免原先统计偏差。
- `converter/scan_convert_candidates.py`
  - 新增 `iter_convert_candidates(...)` 生成器接口。
  - 现有 `scan_convert_candidates(...)` 保持兼容，内部改为 `list(iter_convert_candidates(...))`。

### 19.3 本轮新增/更新测试

- 更新 `tests/test_converter_batch_parallel_module.py`
  - 新增 `test_batch_parallel_respects_small_pending_window`
- 更新 `tests/test_converter_scan_convert_candidates_module.py`
  - 覆盖 `iter_convert_candidates(...)` 与列表接口一致性

### 19.4 回归结果

执行命令：

```bash
python -m unittest discover -s tests -p "test_*.py" -v
```

结果（2026-02-24）：

- `Ran 277 tests in 7.654s`
- `OK`

---

## 20. 2026-02-24 审查建议落地（配置 Schema 校验）

### 20.1 本轮改动

- 新增 `converter/config_validation.py`
  - `validate_runtime_config_or_raise(cfg)`
  - 覆盖关键字段校验：
    - 枚举：`run_mode` / `collect_mode` / `content_strategy` / `default_engine` / `kill_process_mode` / `merge_mode` / `merge_convert_submode`
    - 数值范围：`timeout_seconds` / `pdf_wait_seconds` / `ppt_timeout_seconds` / `ppt_pdf_wait_seconds` / `parallel_workers` / `parallel_checkpoint_interval` / `parallel_max_pending`
    - 布尔类型：并行、合并、输出开关、fast_md 开关等
    - 结构类型：`source_folders` / `excluded_folders` / `price_keywords` / `allowed_extensions`
- 更新 `converter/config_load.py`
  - 在 `_apply_config_defaults()` 后执行 schema 校验。
  - 由于 `collect_mode/content_strategy` 可能仅在 runtime 值中体现，校验前构造 `effective_cfg`（融合 `converter` 运行态值）以避免误报。
  - 校验失败时输出明确错误并 `exit(1)`。

### 20.2 本轮新增/更新测试

- 新增 `tests/test_converter_config_validation_module.py`
  - 有效配置通过
  - 非法值组合拒绝并返回明确错误
- 更新 `tests/test_converter_config_load_module.py`
  - `test_config_load_exits_on_invalid_schema`
  - 调整原有 `Dummy._apply_config_defaults` stub，使其产出最小有效 schema

### 20.3 回归结果

执行命令：

```bash
python -m unittest discover -s tests -p "test_*.py" -v
```

结果（2026-02-24）：

- `Ran 280 tests in 8.234s`
- `OK`

---

## 21. 2026-02-24 审查建议续改（异常收敛与配置保存可观测性）

### 21.1 本轮改动

- `converter/config_load.py`
  - 收敛异常捕获：
    - 文件读取阶段：`OSError` / `UnicodeDecodeError`
    - JSON 解析阶段：`ValueError`
  - 区分错误信息：
    - 读取失败 -> `Failed to load config file`
    - JSON 语法失败 -> `Invalid JSON in config file`
  - 保持 schema 校验，并在校验前构造 `effective_cfg`（融合 runtime 字段）避免误报。
- `office_converter.py`
  - `save_config` 不再静默吞错：
    - 捕获 `OSError` / `TypeError` / `ValueError`
    - 记录 `logging.error("failed to save config: ...")`

### 21.2 本轮新增/更新测试

- 更新 `tests/test_converter_config_load_module.py`
  - 新增 `test_config_load_exits_on_invalid_json`
- 新增 `tests/test_converter_save_config_behavior.py`
  - 覆盖 `save_config` 在不可序列化配置下会记录错误日志

### 21.3 回归结果

执行命令：

```bash
python -m unittest discover -s tests -p "test_*.py" -v
```

结果（2026-02-24）：

- `Ran 282 tests in 7.325s`
- `OK`

---

## 22. 2026-02-24 审查建议续改（run_workflow / scan_convert_candidates 异常收敛）

### 22.1 本轮改动

- `converter/scan_convert_candidates.py`
  - 日期过滤分支异常收敛：
    - `except Exception` -> `except (OSError, OverflowError, ValueError)`
- `converter/run_workflow.py`
  - 将多处“可识别场景”从通用异常收敛为具体异常集合：
    - 沙箱空间预检、后处理导出、索引写入、更新包写入、失败报告导出等分支
    - `except Exception` -> `except (OSError, RuntimeError, ValueError)`
  - 打开输出目录分支：
    - `except Exception` -> `except (KeyError, OSError, AttributeError)`
  - 临时目录清理：
    - `except Exception` -> `except OSError`

### 22.2 回归结果

执行命令：

```bash
python -m unittest discover -s tests -p "test_*.py" -v
```

结果（2026-02-24）：

- `Ran 282 tests in 7.730s`
- `OK`

---

## 23. 2026-02-25 审查建议续改（TaskWorkflowMixin 首批异常收敛）

### 23.1 本轮改动

- `gui/mixins/gui_task_workflow_mixin.py`
  - 对低风险 UI 辅助分支进行首批异常收敛（减少 `except Exception`）：
    - `_set_text_widget_content`
    - `_report_nonfatal_ui_error`
    - `_task_list_filter_text`
    - `_task_list_status_filter`
    - `_task_list_sort_by`
    - `_task_scope_current_config_only`
    - `_refresh_task_status_filter_values`
    - `_normalize_config_for_compare`
    - `_safe_abs_path`
    - `_profile_record_abs_path`
  - 保持原行为：异常仍然按原逻辑回退（返回默认值/跳过 UI 更新/不中断主流程）。

### 23.2 本轮新增测试

- 新增 `tests/test_task_workflow_exception_narrowing.py`
  - 覆盖任务筛选变量异常回退
  - 覆盖状态筛选下拉框 `configure -> __setitem__` 回退路径
  - 覆盖 nonfatal 队列写入异常吞吐
  - 覆盖配置比较序列化失败回退
  - 覆盖 `_safe_abs_path` 在 `abspath` 抛 `ValueError` 时回退原值

### 23.3 回归结果

执行命令：

```bash
python -m unittest discover -s tests -p "test_*.py" -v
```

结果（2026-02-25）：

- `Ran 287 tests in 7.575s`
- `OK`

---

## 24. 2026-02-25 审查建议续改（TaskWorkflowMixin 第二批异常收敛，清零 bare `Exception`）

### 24.1 本轮改动

- `gui/mixins/gui_task_workflow_mixin.py`
  - 完成第二批异常收敛，覆盖任务列表刷新、配置绑定查询、向导保存、编辑/加载任务等路径。
  - 文件内 `except Exception` 已清零（仅保留具体异常类型）。
  - 主要替换为按场景分类的异常集：
    - UI 控件交互：`tk.TclError` / `RuntimeError` / `AttributeError`
    - 配置与值处理：`TypeError` / `ValueError`
    - 文件与存储：`OSError`
  - 保持原行为：仍执行“非致命回退 + 日志/弹窗反馈”，不改变业务分支。

### 24.2 本轮测试更新

- 更新 `tests/test_task_workflow_exception_narrowing.py`
  - 新增静态约束：`test_task_workflow_mixin_has_no_bare_except_exception`
  - 防止后续回归重新引入 `except Exception`。

### 24.3 回归结果

执行命令：

```bash
python -m unittest discover -s tests -p "test_*.py" -v
```

结果（2026-02-25）：

- `Ran 288 tests in 6.917s`
- `OK`

---

## 25. 2026-02-25 测试体验修复（测试期间禁止自动打开目录）

### 25.1 本轮改动

- `converter/run_workflow.py`
  - 自动打开输出目录前新增测试上下文判定：
    - `PYTEST_CURRENT_TEST` 环境变量存在
    - `ZW_TEST_MODE` 环境变量存在
    - `sys.argv` 包含 `unittest`
  - 命中测试上下文时跳过 `os.startfile(...)`，避免单测跑完弹出资源管理器窗口。
  - 保留显式配置项：`config.auto_open_output_dir`（默认 `True`）。
- `gui/mixins/gui_misc_ui_mixin.py`
  - `_open_path` 新增同样的测试上下文判定，防止 GUI 相关测试触发系统打开目录。

### 25.2 本轮测试更新

- 更新 `tests/test_converter_run_workflow_module.py`
  - 新增 `test_run_workflow_does_not_open_folder_in_unittest_context`
- 更新 `tests/test_gui_misc_ui_mixin.py`
  - 新增 `test_open_path_skips_open_in_unittest_context`

### 25.3 回归结果

执行命令：

```bash
python -m unittest discover -s tests -p "test_*.py" -v
```

结果（2026-02-25）：

- `Ran 290 tests in 6.623s`
- `OK`

---

## 26. 2026-02-25 未完成项续改（MiscUIMixin 异常收敛）

### 26.1 本轮改动

- `gui/mixins/gui_misc_ui_mixin.py`
  - 继续收敛剩余通用异常捕获，替换为场景化异常类型：
    - `_toggle_tooltip_advanced`: `AttributeError` / `RuntimeError` / `tk.TclError`
    - `_open_path` (Linux 分支): `OSError` / `subprocess.SubprocessError` / `ValueError`
    - `_poll_log_queue`: `winfo_exists` 与 `after` 回调相关异常收敛
  - 清理了 mixin 文件底部无实际运行价值的 `__main__` 调试入口块，避免该模块继续引入裸异常与无关依赖。
  - 保留测试期“禁自动打开目录”行为。

### 26.2 本轮测试更新

- 更新 `tests/test_gui_misc_ui_mixin.py`
  - 新增 `test_misc_ui_mixin_has_no_bare_except_exception`，约束该模块不回退到 `except Exception`。

### 26.3 回归结果

执行命令：

```bash
python -m unittest discover -s tests -p "test_*.py" -v
```

结果（2026-02-25）：

- `Ran 291 tests in 7.886s`
- `OK`

---

## 27. 2026-02-25 未完成项续改（artifact_meta 异常收敛）

### 27.1 本轮改动

- `converter/artifact_meta.py`
  - 收敛元数据处理中的通用异常捕获：
    - `os.path.relpath(...)`：`except Exception` -> `except (TypeError, ValueError, OSError)`
    - `compute_md5(...)`：`except Exception` -> `except (TypeError, ValueError, OSError)`
    - `compute_file_hash(...)`：`except Exception` -> `except (TypeError, ValueError, OSError)`
  - 保持原行为（失败时回退为空摘要或绝对路径），仅提高可维护性与异常语义清晰度。

### 27.2 回归结果

执行命令：

```bash
python -m unittest discover -s tests -p "test_*.py" -v
```

结果（2026-02-25）：

- `Ran 291 tests in 6.219s`
- `OK`

---

## 28. 2026-02-25 未完成项续改（ai_paths 异常收敛）

### 28.1 本轮改动

- `converter/ai_paths.py`
  - 收敛 AI 输出路径辅助函数中的通用异常捕获：
    - `os.path.relpath(..., target_root)`：`except Exception` -> `except (TypeError, ValueError, OSError)`
    - `source_root_resolver + relpath` 分支：`except Exception` -> `except (TypeError, ValueError, OSError)`
  - 保持原有回退语义：异常时退回 `basename` 逻辑，不影响产物目录结构。

### 28.2 本轮测试更新

- 更新 `tests/test_converter_ai_paths_module.py`
  - 新增 `test_build_ai_output_path_falls_back_when_relpath_raises`
  - 新增 `test_build_ai_output_path_from_source_ignores_resolver_error`
  - 新增 `test_ai_paths_module_has_no_bare_except_exception`

### 28.3 回归结果

执行命令：

```bash
python -m unittest discover -s tests -p "test_*.py" -v
```

结果（2026-02-25）：

- `Ran 294 tests in 7.750s`
- `OK`

---

## 29. 2026-02-25 未完成项续改（checkpoint_utils 异常收敛）

### 29.1 本轮改动

- `converter/checkpoint_utils.py`
  - 收敛 checkpoint 读写辅助中的通用异常捕获：
    - `save_checkpoint`: `except Exception` -> `except (OSError, TypeError, ValueError)`
    - `clear_checkpoint_file`: `except Exception` -> `except OSError`
  - 保留原有失败降级语义（warning 日志，不中断主流程）。

### 29.2 本轮测试更新

- 更新 `tests/test_converter_checkpoint_utils_module.py`
  - 新增 `test_checkpoint_utils_module_has_no_bare_except_exception` 静态约束。

### 29.3 回归结果

执行命令：

```bash
python -m unittest discover -s tests -p "test_*.py" -v
```

结果（2026-02-25）：

- `Ran 295 tests in 6.126s`
- `OK`

---

## 30. 2026-02-25 未完成项续改（config_load 最后裸异常收敛）

### 30.1 本轮改动

- `converter/config_load.py`
  - 将 schema 校验分支的最后一个通用捕获收敛：
    - `except Exception` -> `except (TypeError, ValueError, AttributeError)`
  - 行为保持不变：仍输出 `Invalid config schema` 并退出。

### 30.2 本轮测试更新

- 更新 `tests/test_converter_config_load_module.py`
  - 新增 `test_config_load_module_has_no_bare_except_exception` 静态约束。

### 30.3 回归结果

执行命令：

```bash
python -m unittest discover -s tests -p "test_*.py" -v
```

结果（2026-02-25）：

- `Ran 296 tests in 6.100s`
- `OK`

---

## 31. 2026-02-25 未完成项续改（cab_convert 异常收敛）

### 31.1 本轮改动

- `converter/cab_convert.py`
  - 清理 `finally` 清理分支中的通用异常捕获：
    - `except Exception` -> `except (OSError, TypeError, ValueError)`
  - 继续保持临时目录清理为“失败可忽略”语义，不影响主流程返回。

### 31.2 本轮测试更新

- 更新 `tests/test_converter_cab_convert_module.py`
  - 新增 `test_cab_convert_module_has_no_bare_except_exception` 静态约束。

### 31.3 回归结果

执行命令：

```bash
python -m unittest discover -s tests -p "test_*.py" -v
```

结果（2026-02-25）：

- `Ran 297 tests in 7.740s`
- `OK`

---

## 32. 2026-02-25 未完成项续改（default_config 异常收敛）

### 32.1 本轮改动

- `converter/default_config.py`
  - 将默认配置写盘入口的通用异常捕获收敛为具体类型：
    - `except Exception` -> `except (OSError, TypeError, ValueError)`
  - 保持原行为：失败时打印错误并返回 `False`。

### 32.2 本轮测试更新

- 更新 `tests/test_converter_default_config_module.py`
  - 新增 `test_default_config_module_has_no_bare_except_exception` 静态约束。

### 32.3 回归结果

执行命令：

```bash
python -m unittest discover -s tests -p "test_*.py" -v
```

结果（2026-02-25）：

- `Ran 298 tests in 5.903s`
- `OK`

---

## 33. 2026-02-25 未完成项续改（cab_extract 异常收敛）

### 33.1 本轮改动

- `converter/cab_extract.py`
  - Windows `expand` 分支异常收敛：
    - `except Exception` -> `except (OSError, RuntimeError, TypeError, ValueError)`
  - 行为保持：`expand` 失败时回落到 7z 流程，不改变原始容错策略。

### 33.2 本轮测试更新

- 更新 `tests/test_converter_cab_extract_module.py`
  - 新增 `test_cab_extract_module_has_no_bare_except_exception` 静态约束。

### 33.3 回归结果

执行命令：

```bash
python -m unittest discover -s tests -p "test_*.py" -v
```

结果（2026-02-25）：

- `Ran 299 tests in 6.056s`
- `OK`

---

## 34. 2026-02-25 未完成项续改（failure_stage 异常收敛）

### 34.1 本轮改动

- `converter/failure_stage.py`
  - 收敛失败阶段推断辅助中的通用异常捕获：
    - `get_failure_output_expectation`: `except Exception` -> `except (TypeError, ValueError, AttributeError, RuntimeError)`
    - `infer_failure_stage` 中 `expected_outputs_getter` 分支：`except Exception` -> `except (TypeError, ValueError, AttributeError, RuntimeError)`
  - 保持原有容错语义：获取输出期望失败时按空期望回退，不中断失败阶段判定。

### 34.2 本轮测试更新

- 更新 `tests/test_converter_failure_stage_module.py`
  - 新增 `test_failure_stage_module_has_no_bare_except_exception` 静态约束。

### 34.3 回归结果

执行命令：

```bash
python -m unittest discover -s tests -p "test_*.py" -v
```

结果（2026-02-25）：

- `Ran 300 tests in 6.299s`
- `OK`

---

## 35. 2026-02-25 未完成项续改（index_runtime 异常收敛）

### 35.1 本轮改动

- `converter/index_runtime.py`
  - 收敛转换索引记录追加逻辑中的通用异常捕获：
    - `compute_md5_fn(...)` 两处：`except Exception` -> `except (OSError, TypeError, ValueError, RuntimeError)`
    - `relpath_fn(...)` 两处：`except Exception` -> `except (OSError, TypeError, ValueError, RuntimeError)`
  - 保持原有容错语义：异常时对应 md5 置空、相对路径回退到绝对路径，不影响主流程。

### 35.2 本轮测试更新

- 更新 `tests/test_converter_index_runtime_module.py`
  - 新增 `test_index_runtime_module_has_no_bare_except_exception` 静态约束。

### 35.3 回归结果

执行命令：

```bash
python -m unittest discover -s tests -p "test_*.py" -v
```

结果（2026-02-25）：

- `Ran 301 tests in 5.951s`
- `OK`

---

## 36. 2026-02-25 未完成项续改（platform_utils 异常收敛）

### 36.1 本轮改动

- `converter/platform_utils.py`
  - 收敛 `clear_console` 中的通用异常捕获：
    - `except Exception` -> `except (AttributeError, OSError, TypeError, ValueError)`
  - 保持原有语义：控制台清理失败时静默忽略，不影响主流程。

### 36.2 本轮测试更新

- 更新 `tests/test_converter_platform_utils_module.py`
  - 新增 `test_platform_utils_module_has_no_bare_except_exception` 静态约束。

### 36.3 回归结果

执行命令：

```bash
python -m unittest discover -s tests -p "test_*.py" -v
```

结果（2026-02-25）：

- `Ran 302 tests in 6.140s`
- `OK`

---

## 37. 2026-02-25 未完成项续改（local_office_app 异常收敛）

### 37.1 本轮改动

- `converter/local_office_app.py`
  - 收敛 `_get_local_app` 对应模块实现中的通用异常捕获：
    - `Dispatch -> DispatchEx` 回退：`except Exception` -> `except (AttributeError, OSError, RuntimeError, TypeError, ValueError)`
    - COM 属性设置（`Visible/DisplayAlerts`）：`except Exception` -> `except (AttributeError, OSError, RuntimeError, TypeError, ValueError)`
    - Excel `AskToUpdateLinks` 设置：`except Exception` -> `except (AttributeError, OSError, RuntimeError, TypeError, ValueError)`
  - 语义保持：主流程仍优先容错，不影响 COM 启动回退与属性设置失败降级。

### 37.2 本轮测试更新

- 更新 `tests/test_converter_local_office_app_module.py`
  - 新增 `test_local_office_app_module_has_no_bare_except_exception` 静态约束。

### 37.3 回归结果

执行命令：

```bash
python -m unittest discover -s tests -p "test_*.py" -v
```

结果（2026-02-25）：

- `Ran 303 tests in 6.162s`
- `OK`

---

## 38. 2026-02-25 未完成项续改（callback_utils 异常收敛）

### 38.1 本轮改动

- `converter/callback_utils.py`
  - 收敛回调发射辅助中的通用异常捕获：
    - `emit_file_plan`: `except Exception` -> `except (TypeError, ValueError, AttributeError, RuntimeError)`
    - `emit_file_done`: `except Exception` -> `except (TypeError, ValueError, AttributeError, RuntimeError)`
  - 语义保持：回调失败仍仅记录 warning，不中断主流程。

### 38.2 本轮测试更新

- 更新 `tests/test_converter_callback_utils_module.py`
  - 新增 `test_callback_utils_module_has_no_bare_except_exception` 静态约束。

### 38.3 回归结果

执行命令：

```bash
python -m unittest discover -s tests -p "test_*.py" -v
```

结果（2026-02-25）：

- `Ran 304 tests in 6.126s`
- `OK`

---

## 39. 2026-02-25 未完成项续改（sandbox_guard 异常收敛）

### 39.1 本轮改动

- `converter/sandbox_guard.py`
  - 收敛沙箱空间检查中的通用异常捕获：
    - `sandbox_min_free_gb` 转换：`except Exception` -> `except (TypeError, ValueError)`
    - `disk_usage_fn(...)` 调用：`except Exception as e` -> `except (OSError, RuntimeError, TypeError, ValueError) as e`
  - 保持原语义：配置值异常时阈值降级为 `0`；磁盘空间探测失败时记录 warning 并返回，不中断主流程。

### 39.2 本轮测试更新

- 更新 `tests/test_converter_sandbox_guard_module.py`
  - 新增 `test_sandbox_guard_module_has_no_bare_except_exception` 静态约束。

### 39.3 回归结果

执行命令：

```bash
python -m unittest discover -s tests -p "test_*.py" -v
```

结果（2026-02-25）：

- `Ran 305 tests in 6.432s`
- `OK`

---

## 40. 2026-02-25 未完成项续改（interactive_prompts 异常收敛）

### 40.1 本轮改动

- `converter/interactive_prompts.py`
  - 收敛交互确认逻辑中的通用异常捕获：
    - `confirm_continue_missing_md_merge`：`except Exception` -> `except (EOFError, KeyboardInterrupt, OSError, RuntimeError, TypeError, ValueError)`
  - 保持原语义：交互输入异常时返回 `False`，不影响非交互默认继续策略。

### 40.2 本轮测试更新

- 更新 `tests/test_converter_interactive_prompts_module.py`
  - 新增 `test_interactive_prompts_module_has_no_bare_except_exception` 静态约束。

### 40.3 回归结果

执行命令：

```bash
python -m unittest discover -s tests -p "test_*.py" -v
```

结果（2026-02-25）：

- `Ran 306 tests in 6.414s`
- `OK`

---

## 41. 2026-02-25 未完成项续改（file_registry 异常收敛）

### 41.1 本轮改动

- `converter/file_registry.py`
  - 收敛 `load()` 中的通用异常捕获：
    - `except Exception` -> `except (OSError, json.JSONDecodeError, TypeError, ValueError, AttributeError)`
  - 保持原语义：加载失败时回退为空 `entries` 且 `version=1`，不影响主流程。

### 41.2 本轮测试更新

- 更新 `tests/test_converter_file_registry_module.py`
  - 新增 `test_file_registry_module_has_no_bare_except_exception` 静态约束。

### 41.3 回归结果

执行命令：

```bash
python -m unittest discover -s tests -p "test_*.py" -v
```

结果（2026-02-25）：

- `Ran 307 tests in 6.493s`
- `OK`

---

## 42. 2026-02-25 未完成项续改（office_cycle 异常收敛）

### 42.1 本轮改动

- `converter/office_cycle.py`
  - 收敛 `get_office_restart_every` 中的通用异常捕获：
    - `except Exception` -> `except (TypeError, ValueError, AttributeError)`
  - 保持原语义：解析失败回退默认值 `25`，并继续应用正数校验逻辑。

### 42.2 本轮测试更新

- 更新 `tests/test_converter_office_cycle_module.py`
  - 新增 `test_office_cycle_module_has_no_bare_except_exception` 静态约束。

### 42.3 回归结果

执行命令：

```bash
python -m unittest discover -s tests -p "test_*.py" -v
```

结果（2026-02-25）：

- `Ran 308 tests in 7.795s`
- `OK`

---

## 43. 2026-02-25 未完成项续改（source_roots 异常收敛）

### 43.1 本轮改动

- `converter/source_roots.py`
  - 收敛源目录探测中的通用异常捕获：
    - `probe_source_root_access`：`except Exception as e` -> `except (OSError, RuntimeError, TypeError, ValueError) as e`
  - 保持原语义：访问异常仍写入 `record_skip_fn` 并返回 `False`。

### 43.2 本轮测试更新

- 更新 `tests/test_converter_source_roots_module.py`
  - 新增 `test_source_roots_module_has_no_bare_except_exception` 静态约束。

### 43.3 回归结果

执行命令：

```bash
python -m unittest discover -s tests -p "test_*.py" -v
```

结果（2026-02-25）：

- `Ran 309 tests in 7.979s`
- `OK`

---

## 44. 2026-02-25 未完成项续改（traceability 异常收敛）

### 44.1 本轮改动

- `converter/traceability.py`
  - 收敛旧 `trace_map.xlsx` 读取分支的通用异常捕获：
    - `except Exception` -> `except (OSError, RuntimeError, TypeError, ValueError, KeyError, IndexError, AttributeError)`
  - 保持原语义：读取旧文件异常时回退为空 `existing_by_sid` 并继续用当前批次写出。

### 44.2 本轮测试更新

- 更新 `tests/test_converter_traceability_module.py`
  - 新增 `test_traceability_module_has_no_bare_except_exception` 静态约束。

### 44.3 回归结果

执行命令：

```bash
python -m unittest discover -s tests -p "test_*.py" -v
```

结果（2026-02-25）：

- `Ran 310 tests in 6.162s`
- `OK`

---

## 45. 2026-02-25 未完成项续改（批量异常收敛：file_ops / excel_*）

### 45.1 本轮改动

- `converter/file_ops.py`
  - 收敛 5 处通用异常捕获：
    - `unblock_file`（内外两层）
    - `copy_pdf_direct`
    - `quarantine_failed_file`
    - `handle_file_conflict` 覆盖分支
  - 保持语义：解锁失败静默、失败文件隔离失败返回 `None`、冲突覆盖失败返回 `overwrite_failed`。

- `converter/excel_sheet_utils.py`
  - `auto_fit_sheet` 单元格读取分支：
    - `except Exception` -> `except (TypeError, ValueError, AttributeError)`

- `converter/excel_chart_utils.py`
  - 收敛 4 处通用异常捕获（标题富文本解析、锚点解析、字符串兜底）。

- `converter/excel_defined_names.py`
  - 收敛 2 处通用异常捕获（defined names 容器遍历、destinations 读取）。

### 45.2 本轮测试更新

- 更新 `tests/test_converter_file_ops_module.py`
  - 新增 `test_file_ops_module_has_no_bare_except_exception`
- 更新 `tests/test_converter_excel_sheet_utils_module.py`
  - 新增 `test_excel_sheet_utils_module_has_no_bare_except_exception`
- 更新 `tests/test_converter_excel_chart_utils_module.py`
  - 新增 `test_excel_chart_utils_module_has_no_bare_except_exception`
- 更新 `tests/test_converter_excel_defined_names_module.py`
  - 新增 `test_excel_defined_names_module_has_no_bare_except_exception`

### 45.3 回归结果

执行命令：

```bash
python -m unittest discover -s tests -p "test_*.py" -v
```

结果（2026-02-25）：

- `Ran 314 tests in 6.554s`
- `OK`

---

## 46. 2026-02-25 未完成项续改（批量异常收敛：failure/mshelp/incremental/update）

### 46.1 本轮改动

- `converter/failure_report.py`
  - 收敛报告写盘分支中的通用异常捕获（JSON/TXT 各 1 处）。

- `converter/mshelp_records.py`
  - 收敛 `build_mshelp_record` 中路径推断与相对路径计算的通用异常捕获（2 处）。

- `converter/markitdown_pack_probe.py`
  - 收敛 markitdown 导入失败与探测执行失败分支（2 处）。

- `converter/incremental_registry_ops.py`
  - 收敛源哈希计算与输出 MD5 写回分支（2 处）。

- `converter/update_package_export.py`
  - 收敛 `_safe_compute_md5` 与打包复制失败分支（2 处）。

- `converter/excel_json_utils.py`
  - 收敛 `json_safe_value` 的 `isoformat` 兜底分支（1 处）。

### 46.2 本轮测试更新

- 更新 `tests/test_converter_failure_report_module.py`
  - 新增 `test_failure_report_module_has_no_bare_except_exception`
- 更新 `tests/test_converter_mshelp_records_module.py`
  - 新增 `test_mshelp_records_module_has_no_bare_except_exception`
- 更新 `tests/test_converter_markitdown_pack_probe_module.py`
  - 新增 `test_markitdown_pack_probe_module_has_no_bare_except_exception`
- 更新 `tests/test_converter_incremental_registry_ops_module.py`
  - 新增 `test_incremental_registry_ops_module_has_no_bare_except_exception`
- 更新 `tests/test_converter_update_package_export_module.py`
  - 新增 `test_update_package_export_module_has_no_bare_except_exception`
- 更新 `tests/test_converter_excel_json_utils_module.py`
  - 新增 `test_excel_json_utils_module_has_no_bare_except_exception`

### 46.3 回归结果

执行命令：

```bash
python -m unittest discover -s tests -p "test_*.py" -v
```

结果（2026-02-25）：

- `Ran 320 tests in 6.756s`
- `OK`

---

## 47. 2026-02-25 未完成项续改（批量异常收敛：trace/merge/mshelp/pdf_md）

### 47.1 本轮改动

- `converter/failure_trace_utils.py`
  - 收敛 3 处通用异常捕获（`failed_copy_path` 目录推断、`makedirs`、失败日志写盘）。

- `converter/merge_index_doc.py`
  - 收敛 2 处通用异常捕获（页边距设置分支、主生成流程）。

- `converter/mshelp_topics.py`
  - 收敛 HTML 解析分支中的通用异常捕获（1 处）。

- `converter/pdf_markdown_export.py`
  - 收敛 3 处通用异常捕获（`extract_text`、`compute_md5`、`build_short_id`）。

### 47.2 本轮测试更新

- 更新 `tests/test_converter_failure_trace_utils_module.py`
  - 新增 `test_failure_trace_utils_module_has_no_bare_except_exception`
- 更新 `tests/test_converter_merge_index_doc_module.py`
  - 新增 `test_merge_index_doc_module_has_no_bare_except_exception`
- 更新 `tests/test_converter_mshelp_topics_module.py`
  - 新增 `test_mshelp_topics_module_has_no_bare_except_exception`
- 更新 `tests/test_converter_pdf_markdown_export_module.py`
  - 新增 `test_pdf_markdown_export_module_has_no_bare_except_exception`

### 47.3 回归结果

执行命令：

```bash
python -m unittest discover -s tests -p "test_*.py" -v
```

结果（2026-02-25）：

- `Ran 324 tests in 6.187s`
- `OK`

---

## 48. 2026-02-25 未完成项续改（批量异常收敛：trace/merge/mshelp/pdf_md 第二轮）

### 48.1 本轮改动

- `converter/failure_trace_utils.py`
  - 收敛 3 处通用异常捕获（路径推断、目录创建、JSON 写盘）。

- `converter/merge_index_doc.py`
  - 收敛 2 处通用异常捕获（页面设置分支、主流程异常）。

- `converter/mshelp_topics.py`
  - 收敛 BS4 解析分支通用异常捕获（1 处）。

- `converter/pdf_markdown_export.py`
  - 收敛 3 处通用异常捕获（`extract_text`、`compute_md5`、`build_short_id`）。

### 48.2 本轮测试更新

- 更新 `tests/test_converter_failure_trace_utils_module.py`
  - 增加静态约束：`test_failure_trace_utils_module_has_no_bare_except_exception`
- 更新 `tests/test_converter_merge_index_doc_module.py`
  - 增加静态约束：`test_merge_index_doc_module_has_no_bare_except_exception`
- 更新 `tests/test_converter_mshelp_topics_module.py`
  - 增加静态约束：`test_mshelp_topics_module_has_no_bare_except_exception`
- 更新 `tests/test_converter_pdf_markdown_export_module.py`
  - 增加静态约束：`test_pdf_markdown_export_module_has_no_bare_except_exception`

### 48.3 回归结果

执行命令：

```bash
python -m unittest discover -s tests -p "test_*.py" -v
```

结果（2026-02-25）：

- `Ran 324 tests in 6.112s`
- `OK`

---

## 49. 2026-02-25 未完成项续改（批量异常收敛：batch/chromadb/excel/convert/corpus）

### 49.1 本轮改动

- `converter/batch_parallel.py`
  - 收敛 1 处通用异常捕获（并行 worker `future.result()` 失败分支）。

- `converter/batch_sequential.py`
  - 收敛 1 处通用异常捕获（单文件转换主流程失败分支）。

- `converter/checkpoint_runtime.py`
  - 收敛 checkpoint 加载分支通用异常捕获，细化为 `JSONDecodeError + I/O/类型运行时异常`。

- `converter/chromadb_export.py`
  - 收敛 ChromaDB upsert 主流程通用异常捕获（1 处）。

- `converter/chromadb_docs.py`
  - 收敛 Markdown 读取分支通用异常捕获（1 处）。

- `converter/excel_content_scan.py`
  - 收敛工作表扫描内外两层通用异常捕获（2 处）。

- `converter/excel_json_export.py`
  - 收敛 5 处通用异常捕获：
    - 正整数配置读取兜底
    - 相对路径兜底
    - 主解析流程失败
    - `wb_values.close()` 兜底
    - `wb_formula.close()` 兜底

- `converter/excel_json_batch.py`
  - 收敛批量导出循环中的通用异常捕获（1 处）。

- `converter/convert_thread.py`
  - 收敛 Word/Excel/PPT 各分支与关闭/退出兜底中的通用异常捕获（9 处）。

- `converter/corpus_manifest.py`
  - 收敛 LLM Hub 目录创建、文件复制、大小读取、manifest/README 写盘、hub 构建包装层通用异常捕获（7 处）。

### 49.2 本轮测试更新

- 更新 `tests/test_converter_batch_parallel_module.py`
  - 新增 `test_batch_parallel_module_has_no_bare_except_exception`
- 更新 `tests/test_converter_batch_sequential_module.py`
  - 新增 `test_batch_sequential_module_has_no_bare_except_exception`
- 更新 `tests/test_converter_checkpoint_runtime_module.py`
  - 新增 `test_checkpoint_runtime_module_has_no_bare_except_exception`
- 更新 `tests/test_converter_chromadb_export_module.py`
  - 新增 `test_chromadb_export_module_has_no_bare_except_exception`
- 更新 `tests/test_converter_chromadb_docs_module.py`
  - 新增 `test_chromadb_docs_module_has_no_bare_except_exception`
- 更新 `tests/test_converter_excel_content_scan_module.py`
  - 新增 `test_excel_content_scan_module_has_no_bare_except_exception`
- 更新 `tests/test_converter_excel_json_export_module.py`
  - 新增 `test_excel_json_export_module_has_no_bare_except_exception`
- 更新 `tests/test_converter_excel_json_batch_module.py`
  - 新增 `test_excel_json_batch_module_has_no_bare_except_exception`
- 更新 `tests/test_converter_convert_thread_module.py`
  - 新增 `test_convert_thread_module_has_no_bare_except_exception`
- 更新 `tests/test_converter_corpus_manifest_module.py`
  - 新增 `test_corpus_manifest_module_has_no_bare_except_exception`

### 49.3 回归结果

执行命令：

```bash
python -m unittest discover -s tests -p "test_*.py" -v
```

结果（2026-02-25）：

- `Ran 338 tests in 6.203s`
- `OK`

### 49.4 未完成项状态

- 当前 `except Exception` 总量（`converter + office_converter.py`）：`34`
- 相比上一轮统计 `74`，本轮净减少 `40`
- 剩余集中在：`office_converter.py`、`merge_pdfs.py`、`mshelp_merge.py`、`collect_index.py`、`safe_exec.py`、`process_single.py` 等模块。

---

## 50. 2026-02-25 未完成项续改（批量异常收敛：collect/incremental/mshelp/process/safe_exec）

### 50.1 本轮改动

- `converter/collect_index.py`
  - 收敛 2 处通用异常捕获（复制失败、列宽计算兜底）。

- `converter/incremental_filters.py`
  - 收敛全局 MD5 去重中的通用异常捕获（1 处）。

- `converter/incremental_scan.py`
  - 收敛重命名检测阶段哈希计算兜底通用异常捕获（1 处）。

- `converter/mac_convert.py`
  - 收敛 LibreOffice 调用失败通用异常捕获（1 处）。

- `converter/merge_markdown.py`
  - 收敛 Markdown 读取失败通用异常捕获（1 处）。

- `converter/mshelp_merge.py`
  - 收敛 4 处通用异常捕获（配置解析、内容读取、DOCX/PDF 回调导出）。

- `converter/process_single.py`
  - 收敛清理分支通用异常捕获（1 处）。
  - 同时将内部超时/未生成 PDF 的主动抛错由 `Exception` 改为 `RuntimeError`（行为保持：仍会进入失败路径）。

- `converter/safe_exec.py`
  - 移除 `except Exception`，改为可组合的 `retryable_exceptions`（基础异常 + 可选 `com_error_cls`）。
  - `program stopped` 与最终 COM 错误改为 `RuntimeError`，语义不变。

- `converter/update_package_index.py`
  - 收敛 XLSX 写盘包装层通用异常捕获（1 处）。

### 50.2 本轮测试更新

- 更新 `tests/test_converter_collect_index_module.py`
  - 新增 `test_collect_index_module_has_no_bare_except_exception`
- 更新 `tests/test_converter_incremental_filters_module.py`
  - 新增 `test_incremental_filters_module_has_no_bare_except_exception`
- 更新 `tests/test_converter_incremental_scan_module.py`
  - 新增 `test_incremental_scan_module_has_no_bare_except_exception`
- 更新 `tests/test_converter_mac_convert_module.py`
  - 新增 `test_mac_convert_module_has_no_bare_except_exception`
- 更新 `tests/test_converter_merge_markdown_module.py`
  - 新增 `test_merge_markdown_module_has_no_bare_except_exception`
- 更新 `tests/test_converter_mshelp_merge_module.py`
  - 新增 `test_mshelp_merge_module_has_no_bare_except_exception`
- 更新 `tests/test_converter_process_single_module.py`
  - 新增 `test_process_single_module_has_no_bare_except_exception`
- 更新 `tests/test_converter_safe_exec_module.py`
  - 新增 `test_safe_exec_module_has_no_bare_except_exception`
- 更新 `tests/test_converter_update_package_index_module.py`
  - 新增 `test_update_package_index_module_has_no_bare_except_exception`

### 50.3 回归结果

执行命令：

```bash
python -m unittest discover -s tests -p "test_*.py" -v
```

结果（2026-02-25）：

- `Ran 347 tests in 5.926s`
- `OK`

### 50.4 未完成项状态

- 当前 `except Exception` 总量（`converter + office_converter.py`）：`21`
- 相比第 49 轮统计 `34`，本轮净减少 `13`
- 当前仅剩：
  - `office_converter.py`：13 处
  - `converter/merge_pdfs.py`：8 处

---

## 51. 2026-02-25 未完成项续改（收尾：merge_pdfs + office_converter）

### 51.1 本轮改动

- `converter/merge_pdfs.py`
  - 收敛全部 8 处通用异常捕获：
    - Word 启动失败
    - `relpath` 兜底
    - 单 PDF 读取/追加失败
    - merged PDF MD5 兜底
    - map 写出失败
    - 单任务外层失败
    - Word 退出兜底
    - Excel merge list 保存失败

- `office_converter.py`
  - 收敛全部 13 处通用异常捕获：
    - 控制台 `reconfigure`
    - 可选依赖导入（`chromadb/bs4/docx/reportlab/markitdown`）改为 `ImportError`
    - signal 注册兜底
    - `_add_perf_seconds` 数值转换兜底
    - `_kill_process_by_name` 兜底
    - `scan_pdf_content` 兜底
    - `_extract_sheet_pivot_tables` 位置字段兜底
    - `_convert_source_to_markdown_text` 文件读取兜底
    - CLI 入口默认配置生成兜底

- 语法修正
  - 本轮中途一次回归暴露 `office_converter.py` 中 `_extract_sheet_pivot_tables` 的 `except` 缩进错误；
  - 已修复并通过 `python -m py_compile office_converter.py` 与后续全量回归验证。

### 51.2 本轮测试更新

- 更新 `tests/test_converter_merge_pdfs_module.py`
  - 新增 `test_merge_pdfs_module_has_no_bare_except_exception`
- 新增 `tests/test_office_converter_module.py`
  - 新增 `test_office_converter_module_has_no_bare_except_exception`

### 51.3 回归结果

执行命令：

```bash
python -m unittest discover -s tests -p "test_*.py" -v
```

结果（2026-02-25）：

- `Ran 349 tests in 5.924s`
- `OK`

### 51.4 收尾状态

- 当前 `except Exception` 总量（`converter + office_converter.py`）：`0`
- 本轮后，拆分异常收敛阶段在目标范围内已清零完成。

---

## 52. 2026-02-25 未完成项续改（交互选择逻辑再拆分一轮）

### 52.1 未完成项确认

- 异常收敛主线已完成（`except Exception` 清零），但 `office_converter.py` 中仍有一组交互式选择方法属于“可继续薄委托”的剩余拆分空间：
  - `ask_for_subfolder`
  - `select_run_mode`
  - `select_collect_mode`
  - `select_merge_mode`
  - `select_content_strategy`
  - `select_engine_mode`

### 52.2 本轮改动

- 新增 `converter/interactive_choices.py`
  - 提取上述 6 个 CLI 交互选择函数，改为纯函数 + 依赖注入（`input/print/readable`）。

- 更新 `office_converter.py`
  - 上述 6 个方法改为薄委托，内部仅做状态回写（`self.run_mode/self.collect_mode/...`）。
  - 新增对应 `*_impl` 导入。

### 52.3 本轮测试更新

- 新增 `tests/test_converter_interactive_choices_module.py`
  - 核心行为测试（6 个函数）
  - `OfficeConverter` 委托测试（6 个方法）
  - 静态约束：`test_interactive_choices_module_has_no_bare_except_exception`

### 52.4 回归结果

执行命令：

```bash
python -m unittest discover -s tests -p "test_*.py" -v
```

结果（2026-02-25）：

- `Ran 352 tests in 5.771s`
- `OK`

### 52.5 当前状态

- `converter + office_converter.py` 中 `except Exception`：`0`
- 交互式选择逻辑已进一步从 `office_converter.py` 下沉到 `converter/`，拆分完成度继续提升。

---

## 53. 2026-02-25 未完成项续改（运行时方法批量下沉）

### 53.1 未完成项确认

- 在第 52 轮后，`office_converter.py` 仍有若干可继续薄委托的方法（路径计算、PDF 内容扫描、错误摘要、MSHelp-only 流程、运行摘要输出）。
- 本轮按“尽可能一次多拆分”的要求，对这组方法进行成组下沉，减少单文件职责密度。

### 53.2 本轮改动

- 新增 `converter/target_path.py`
  - 提取 `get_target_path`。

- 新增 `converter/pdf_content_scan.py`
  - 提取 `scan_pdf_content`。

- 新增 `converter/error_summary.py`
  - 提取 `get_error_summary_for_display`。

- 新增 `converter/mshelp_workflow.py`
  - 提取 `_run_mshelp_only` 的主流程。

- 新增 `converter/runtime_summary.py`
  - 提取 `print_runtime_summary`。

- 更新 `office_converter.py`
  - 对以上 5 组逻辑改为薄委托调用 `*_impl`。

### 53.3 本轮测试更新

- 新增 `tests/test_converter_runtime_extracts_module.py`
  - 覆盖 5 个新模块核心行为测试；
  - 覆盖 `OfficeConverter` 对应方法委托测试；
  - 新模块静态约束：无裸 `except Exception`。

### 53.4 回归结果

执行命令：

```bash
python -m unittest discover -s tests -p "test_*.py" -v
```

结果（2026-02-25）：

- `Ran 358 tests in 5.948s`
- `OK`

### 53.5 当前状态

- `converter + office_converter.py` 中 `except Exception`：`0`
- `office_converter.py` 可继续拆分项已进一步减少，当前剩余更多为轻量委托或初始化聚合逻辑。

---

## 54. 2026-02-25 未完成项续改（初始化状态与生命周期批量下沉）

### 54.1 本轮改动

- 新增 `converter/bootstrap_state.py`
  - 提取：
    - `build_default_perf_metrics`
    - `initialize_runtime_state`
    - `initialize_output_tracking_state`
    - `initialize_error_tracking_state`
    - `register_signal_handlers`

- 新增 `converter/runtime_lifecycle.py`
  - 提取：
    - `cleanup_all_processes`
    - `close_office_apps`
    - `on_office_file_processed`
    - `check_and_handle_running_processes`

- 更新 `office_converter.py`
  - `__init__` 中运行时初始化逻辑改为调用 `initialize_*_impl`。
  - `_reset_perf_metrics` 改为调用 `build_default_perf_metrics_impl`。
  - `cleanup_all_processes / close_office_apps / _on_office_file_processed / check_and_handle_running_processes` 改为薄委托。

### 54.2 回归中的问题与修复

- 在该批次回归中暴露出 `signal_handler` 缺失导致的初始化异常（`AttributeError`）。
- 已在 `office_converter.py` 新增 `signal_handler(self, signum, _frame)`，语义为记录 warning 并设置 `self.is_running = False`。

### 54.3 本轮测试更新

- 新增 `tests/test_converter_bootstrap_lifecycle_module.py`
  - 覆盖 `bootstrap_state/runtime_lifecycle` 核心行为；
  - 覆盖 `OfficeConverter` 委托行为；
  - 校验新模块无裸 `except Exception`。

### 54.4 回归结果

执行命令：

```bash
python -m unittest discover -s tests -p "test_*.py" -v
```

结果（2026-02-25）：

- `Ran 362 tests in 5.847s`
- `OK`

---

## 55. 2026-02-25 未完成项续改（路径配置与性能计时批量下沉）

### 55.1 本轮改动

- 更新 `converter/path_config.py`
  - 新增 `init_paths_from_config`：
    - 统一处理 `temp_sandbox_root`（绝对/相对路径）；
    - 统一生成并创建 `temp_sandbox / failed_dir / merge_output_dir`。
  - 新增 `save_config`：
    - 统一配置落盘逻辑，保留原有错误日志语义。

- 更新 `converter/perf_summary.py`
  - 新增 `add_perf_seconds`：
    - 提取 `_add_perf_seconds` 的合法性处理（键存在、数值转换、负值过滤）。

- 更新 `office_converter.py`
  - `_init_paths_from_config` 改为委托 `init_paths_from_config_impl` 并回写实例路径字段；
  - `save_config` 改为委托 `save_config_impl`；
  - `_add_perf_seconds` 改为委托 `add_perf_seconds_impl`。

### 55.2 本轮测试更新

- 更新 `tests/test_converter_path_config_module.py`
  - 新增 `init_paths_from_config` 核心行为测试；
  - 新增 `save_config` 核心行为测试；
  - 新增 `OfficeConverter._init_paths_from_config/save_config` 委托测试。

- 更新 `tests/test_converter_perf_summary_module.py`
  - 新增 `add_perf_seconds` 核心行为测试；
  - 新增 `OfficeConverter._add_perf_seconds` 委托测试。

### 55.3 回归结果

执行命令：

```bash
python -m unittest discover -s tests -p "test_*.py" -v
```

结果（2026-02-25）：

- `Ran 367 tests in 5.745s`
- `OK`

### 55.4 当前状态

- 本轮继续完成“尽可能一次多拆分内容”，`office_converter.py` 的初始化路径与性能累加逻辑已进一步下沉。
- `converter + office_converter.py` 中裸 `except Exception` 仍保持 `0`。

---

## 56. 2026-02-25 未完成项续改（merge/prompt/excel/markdown 再下沉一轮）

### 56.1 本轮改动

- 更新 `converter/merge_candidates.py`
  - 新增 `resolve_merge_scan_context`，提取 merge 扫描根目录与排除目录解析逻辑。
- 更新 `office_converter.py`
  - `_scan_merge_candidates_by_ext` 改为委托 `resolve_merge_scan_context + scan_candidates_by_ext`。

- 更新 `converter/excel_chart_extract.py`
  - 新增 `extract_sheet_pivot_tables`。
- 更新 `office_converter.py`
  - `_extract_sheet_pivot_tables` 改为薄委托。

- 更新 `converter/prompt_wrapper.py`
  - 新增 `collect_prompt_ready_candidates`，提取 Prompt_Ready 输入候选聚合逻辑。
- 更新 `office_converter.py`
  - `_write_prompt_ready` 改为调用 `collect_prompt_ready_candidates` 后委托写出。

- 新增 `converter/markdown_source_reader.py`
  - 提取 `convert_source_to_markdown_text`（MarkItDown 优先 + 文件读取兜底）。
- 更新 `office_converter.py`
  - `_convert_source_to_markdown_text` 改为薄委托到新模块。

### 56.2 本轮测试更新

- 更新 `tests/test_converter_merge_candidates_module.py`
  - 新增 `resolve_merge_scan_context` 核心行为覆盖。
- 更新 `tests/test_converter_excel_chart_extract_module.py`
  - 新增 `extract_sheet_pivot_tables` 核心行为与委托覆盖。
- 重写并更新 `tests/test_converter_prompt_wrapper_module.py`
  - 增加 `collect_prompt_ready_candidates` 覆盖；
  - 强化 `_write_prompt_ready` 委托参数断言。
- 新增 `tests/test_converter_markdown_source_reader_module.py`
  - 覆盖 `convert_source_to_markdown_text` 核心行为；
  - 覆盖 `OfficeConverter` 委托；
  - 静态约束：新模块无裸 `except Exception`。

### 56.3 回归结果

执行命令：

```bash
python -m unittest discover -s tests -p "test_*.py" -v
```

结果（2026-02-25）：

- `Ran 371 tests in 6.198s`
- `OK`

### 56.4 当前状态

- 本轮后 `office_converter.py` 中 merge 扫描上下文、pivot 提取、prompt 候选拼接、source->markdown 读取已继续下沉。
- `converter + office_converter.py` 中裸 `except Exception` 仍保持 `0`。

---

## 57. 2026-02-25 未完成项续改（CLI 向导与显示辅助再拆分一轮）

### 57.1 本轮改动

- 新增 `converter/cli_wizard_flow.py`
  - 提取 `cli_wizard` 主流程编排为 `run_cli_wizard`（交互步骤、模式分支、进程检查与路径初始化顺序）。

- 新增 `converter/display_helpers.py`
  - 提取：
    - `print_welcome`
    - `print_step_title`

- 更新 `converter/runtime_lifecycle.py`
  - 新增 `kill_current_app`，提取引擎-进程名映射与复用策略判断。

- 更新 `office_converter.py`
  - `cli_wizard` 改为委托 `run_cli_wizard_impl`；
  - `print_welcome/print_step_title` 改为委托 `print_*_impl`；
  - `_kill_current_app` 改为委托 `kill_current_app_impl`。

### 57.2 本轮测试更新

- 新增 `tests/test_converter_cli_display_module.py`
  - 覆盖 `cli_wizard_flow` 核心流程；
  - 覆盖 `display_helpers` 核心行为；
  - 覆盖 `OfficeConverter` 新委托路径；
  - 静态约束：新模块无裸 `except Exception`。

- 更新 `tests/test_converter_bootstrap_lifecycle_module.py`
  - 增加 `kill_current_app` 核心行为覆盖；
  - 增加 `_kill_current_app` 委托覆盖。

### 57.3 回归中的问题与修复

- 本轮中途出现一次 `converter/runtime_lifecycle.py` 缩进错误（`IndentationError`），已修复后重新执行定向与全量回归。

### 57.4 回归结果

执行命令：

```bash
python -m unittest discover -s tests -p "test_*.py" -v
```

结果（2026-02-25）：

- `Ran 375 tests in 5.962s`
- `OK`

### 57.5 当前状态

- `office_converter.py` 的 CLI 向导流程、显示输出与当前应用清理策略已进一步下沉。
- `converter + office_converter.py` 中裸 `except Exception` 仍保持 `0`。

---

## 58. 2026-02-25 未完成项续改（状态回写逻辑再下沉一轮）

### 58.1 本轮改动

- 更新 `converter/config_defaults.py`
  - 新增 `apply_config_defaults_for_converter`，提取 `_apply_config_defaults` 中 runtime 字段回写逻辑。

- 更新 `converter/mshelp_records.py`
  - 新增 `append_mshelp_record`，提取 `_append_mshelp_record` 的 record 组装与追加逻辑。

- 更新 `converter/incremental_scan.py`
  - 新增 `apply_incremental_filter_for_converter`，提取 `_apply_incremental_filter` 中上下文回写逻辑。

- 更新 `converter/update_package_export.py`
  - 新增 `generate_update_package_for_converter`，提取 `_generate_update_package` 中产物路径与状态回写逻辑。

- 更新 `converter/bootstrap_state.py`
  - 新增 `handle_stop_signal`，提取 signal 到运行态标记变更逻辑。

- 更新 `office_converter.py`
  - `signal_handler` 改为委托 `handle_stop_signal_impl`；
  - `_apply_config_defaults` 改为委托 `apply_config_defaults_for_converter`；
  - `_append_mshelp_record` 改为委托 `append_mshelp_record`；
  - `_apply_incremental_filter` 改为委托 `apply_incremental_filter_for_converter`；
  - `_generate_update_package` 改为委托 `generate_update_package_for_converter`；
  - 新增 `_incremental_log_info/_update_package_log_info` 作为模块化日志注入入口。

### 58.2 本轮测试更新

- 更新 `tests/test_converter_config_defaults_module.py`
  - 新增 `_apply_config_defaults` 委托测试。
- 更新 `tests/test_converter_mshelp_records_module.py`
  - 新增 `append_mshelp_record` 核心行为覆盖。
- 更新 `tests/test_converter_incremental_scan_module.py`
  - 增加 `apply_incremental_filter_for_converter` 对照覆盖。
- 更新 `tests/test_converter_update_package_export_module.py`
  - 调整 `_generate_update_package` 委托测试为新 helper 路径；
  - 新增 `generate_update_package_for_converter` 状态回写测试。
- 更新 `tests/test_converter_bootstrap_lifecycle_module.py`
  - 新增 `handle_stop_signal` 覆盖；
  - 新增 `signal_handler` 委托覆盖。

### 58.3 回归中的问题与修复

- 本轮中途出现 `bootstrap_state.py` 的一次 `IndentationError`（`register_signal_handlers` 的 `try` 块缩进错误），已修复后重跑定向与全量回归。

### 58.4 回归结果

执行命令：

```bash
python -m unittest discover -s tests -p "test_*.py" -v
```

结果（2026-02-25）：

- `Ran 377 tests in 5.785s`
- `OK`

### 58.5 当前状态

- `office_converter.py` 中状态回写型逻辑再次减少，进一步向 `converter/` 模块集中。
- `converter + office_converter.py` 中裸 `except Exception` 仍保持 `0`。

---

## 59. 2026-02-25 未完成项续改（输出回写逻辑再下沉一轮）

### 59.1 本轮改动

- 更新 `converter/index_runtime.py`
  - 新增 `write_conversion_index_workbook_for_converter`，提取转换索引工作簿写出后的 `converter.convert_index_path` 回写逻辑。

- 更新 `converter/markdown_quality_report.py`
  - 新增 `write_markdown_quality_report_for_converter`，提取质量报告路径与产物列表回写逻辑。

- 更新 `converter/records_json_export.py`
  - 新增 `write_records_json_exports_for_converter`，提取 JSON 产物列表回写逻辑。

- 更新 `converter/prompt_wrapper.py`
  - 新增 `write_prompt_ready_for_converter`，提取候选聚合 + Prompt_Ready 输出状态回写逻辑。

- 更新 `office_converter.py`
  - `_write_conversion_index_workbook` 改为委托 `write_conversion_index_workbook_for_converter`；
  - `_write_markdown_quality_report` 改为委托 `write_markdown_quality_report_for_converter`；
  - `_write_records_json_exports` 改为委托 `write_records_json_exports_for_converter`；
  - `_write_prompt_ready` 改为委托 `write_prompt_ready_for_converter`；
  - 清理对应不再使用的旧导入。

### 59.2 本轮测试更新

- 更新 `tests/test_converter_index_runtime_module.py`
  - 新增 `write_conversion_index_workbook_for_converter` 核心行为覆盖；
  - 调整 `OfficeConverter` 委托测试到新 wrapper 路径。

- 更新 `tests/test_converter_markdown_quality_report_module.py`
  - 新增 `write_markdown_quality_report_for_converter` 覆盖；
  - 调整委托测试断言到新 wrapper 参数。

- 更新 `tests/test_converter_records_json_export_module.py`
  - 新增 `write_records_json_exports_for_converter` 覆盖；
  - 调整委托测试断言到新 wrapper 参数。

- 更新 `tests/test_converter_prompt_wrapper_module.py`
  - 新增 `write_prompt_ready_for_converter` 覆盖；
  - 调整委托测试为新 wrapper 路径。

### 59.3 回归结果

执行命令：

```bash
python -m unittest discover -s tests -p "test_*.py" -v
```

结果（2026-02-25）：

- `Ran 377 tests in 6.122s`
- `OK`

### 59.4 当前状态

- `office_converter.py` 中“输出后回写状态”方法再次收敛为薄委托。
- `converter + office_converter.py` 中裸 `except Exception` 仍保持 `0`。

---

## 60. 2026-02-25 未完成项续改（failed/process/merge-map 再下沉一轮）

### 60.1 本轮改动

- 更新 `converter/failure_report.py`
  - 新增 `export_failed_files_report_for_converter`，提取 `output_dir` 解析与 `failed_report_path` 回写逻辑。

- 更新 `converter/runtime_lifecycle.py`
  - 新增 `check_and_handle_running_processes_for_converter`，提取 `reuse_process` 回写逻辑。

- 更新 `converter/index_runtime.py`
  - 新增 `write_merge_map_for_converter`，提取 `_write_merge_map` 的“空记录短路 + 委托写出”逻辑。

- 更新 `office_converter.py`
  - `check_and_handle_running_processes` 改为委托 `check_and_handle_running_processes_for_converter`；
  - `export_failed_files_report` 改为委托 `export_failed_files_report_for_converter`；
  - `_write_merge_map` 改为委托 `write_merge_map_for_converter`；
  - 清理对应旧 `*_impl` 导入。

### 60.2 本轮测试更新

- 更新 `tests/test_converter_failure_report_module.py`
  - 新增 `export_failed_files_report_for_converter` 核心行为覆盖；
  - 调整 `OfficeConverter` 委托测试到新 wrapper。

- 更新 `tests/test_converter_bootstrap_lifecycle_module.py`
  - 新增 `check_and_handle_running_processes_for_converter` 覆盖；
  - 调整 `OfficeConverter.check_and_handle_running_processes` 委托测试到新 wrapper。

- 更新 `tests/test_converter_index_runtime_module.py`
  - 新增 `write_merge_map_for_converter` 覆盖；
  - 调整 `OfficeConverter._write_merge_map` 委托测试到新 wrapper。

### 60.3 回归结果

执行命令：

```bash
python -m unittest discover -s tests -p "test_*.py" -v
```

结果（2026-02-25）：

- `Ran 377 tests in 6.163s`
- `OK`

### 60.4 当前状态

- `office_converter.py` 的 failed report / process handling / merge-map 路径进一步收敛为薄委托。
- `converter + office_converter.py` 中裸 `except Exception` 仍保持 `0`。

---

## 61. 2026-02-25 未完成项一次性收口（剩余 6 个方法全部拆分）

### 61.1 本轮改动

- 更新 `converter/bootstrap_state.py`
  - 新增 `initialize_converter_for_runtime`，提取 `OfficeConverter.__init__` 的初始化编排逻辑（runtime state / output state / signal / load config / init paths / error state）。

- 更新 `converter/path_config.py`
  - 新增 `init_paths_from_config_for_converter`，提取 `_init_paths_from_config` 的状态回写逻辑。

- 更新 `converter/process_ops.py`
  - 新增 `kill_process_by_name_for_converter`，提取 `_kill_process_by_name` 的异常兜底逻辑。

- 更新 `converter/mshelp_merge.py`
  - 新增 `merge_mshelp_markdowns_for_converter`，提取 `_merge_mshelp_markdowns` 的参数装配与 `mshelp_merge_seconds` 计时回写。

- 更新 `converter/batch_parallel.py`
  - 新增 `convert_single_file_threadsafe_for_converter`，提取 `_convert_single_file_threadsafe` 的 COM 初始化/释放与单文件处理调用。

- 更新 `converter/merge_candidates.py`
  - 新增 `scan_merge_candidates_by_ext_for_converter`，提取 `_scan_merge_candidates_by_ext` 的扫描上下文构建 + 扫描调用。

- 更新 `office_converter.py`
  - `__init__` 改为委托 `initialize_converter_for_runtime`；
  - `_init_paths_from_config` 改为委托 `init_paths_from_config_for_converter`；
  - `_kill_process_by_name` 改为委托 `kill_process_by_name_for_converter`；
  - `_merge_mshelp_markdowns` 改为委托 `merge_mshelp_markdowns_for_converter`；
  - `_convert_single_file_threadsafe` 改为委托 `convert_single_file_threadsafe_for_converter`；
  - `_scan_merge_candidates_by_ext` 改为委托 `scan_merge_candidates_by_ext_for_converter`；
  - 清理对应旧导入。

### 61.2 本轮测试更新

- 更新 `tests/test_converter_bootstrap_lifecycle_module.py`
  - 新增 `initialize_converter_for_runtime` 核心行为覆盖；
  - 新增 `OfficeConverter.__init__` 委托测试。

- 更新 `tests/test_converter_path_config_module.py`
  - 新增 `init_paths_from_config_for_converter` 核心覆盖；
  - 调整 `_init_paths_from_config` 委托测试目标到新 wrapper。

- 更新 `tests/test_converter_process_ops_module.py`
  - 新增 `kill_process_by_name_for_converter` 覆盖；
  - 调整 `_kill_process_by_name` 委托测试目标到新 wrapper。

- 更新 `tests/test_converter_mshelp_merge_module.py`
  - 新增 `merge_mshelp_markdowns_for_converter` 覆盖；
  - 调整 `_merge_mshelp_markdowns` 委托测试目标到新 wrapper。

- 更新 `tests/test_converter_batch_parallel_module.py`
  - 新增 `convert_single_file_threadsafe_for_converter` 核心覆盖；
  - 新增 `_convert_single_file_threadsafe` 委托测试。

- 更新 `tests/test_converter_merge_candidates_module.py`
  - 新增 `_scan_merge_candidates_by_ext` 对 wrapper 的委托测试。

### 61.3 回归结果

执行命令：

```bash
python -m unittest discover -s tests -p "test_*.py" -v
```

结果（2026-02-25）：

- `Ran 381 tests in 8.897s`
- `OK`

### 61.4 当前状态

- 本轮前识别的剩余 6 个待拆分方法已全部拆分完成。
- `office_converter.py` 相关路径均已收敛为薄委托调用。
- `converter + office_converter.py` 中裸 `except Exception` 仍保持 `0`。
