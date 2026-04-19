# Changelog

All notable changes to this project are documented in this file.

## [v5.21.1] - 2026-04-18
### Fixed
- **任务向导**：将「归集子模式」与「复制布局」从第 2 步移到第 3 步（路径与增量），避免选「归集/索引」后进入路径页时看不到扁平化选项；确认页摘要中补充归集子模式与复制布局。

---

## [v5.21.0] - 2026-04-18
### Added
- **Collect 复制布局**：在「复制 + 生成索引」（`copy_and_index`）子模式下，新增配置 `collect_copy_layout`：`preserve_tree`（默认，保持源目录相对结构）与 `flat`（全部复制到目标根目录；同名文件依次命名为 `name__1.ext`、`name__2.ext`…）。运行参数 Tab 与任务向导提供单选项；`index_only` 不受影响。

### Tests
- 更新配置校验/默认配置/加载相关单测；新增扁平化目标路径分配的单元测试。

---

## [v5.20.0 · Post-release UI 整理] - 2026-04-18
### Added
- **向导"去重策略"行**：任务向导第 3 步新增去重小节。Collect 模式显示只读说明（SHA256 内容去重 + Duplicates 表）；Convert / 合并模式显示「按类型全局 MD5 去重」复选框与提示。
- **任务编辑复用向导**：编辑任务直接调用 `_open_task_wizard(task=?)`，窗口标题切为「编辑任务」，预填所有字段（名称/描述/多源/目标/增量/输出/合并/收集策略/扩展名/去重）。删除旧的 `_open_task_edit_form` 与 `_edit_task_binding_in_dialog`（净 -347 行）。
- **定时一览**：工具菜单新增「定时一览」窗口，集中查看 `tasks/schedules.json` 中所有计划。
- **定时任务补全**：支持多频率、待触发队列、列表可见性、立即触发 / 删除操作。
- **扩展名 chip 编辑器**：共享 chip UI，向导和共用配置一致。
- **全局默认子分页 / 顶部面包屑**：顶部面包屑点击跳回任务中心。

### Changed
- **任务中心强制化**：移除 classic / task 模式切换，`OfficeGUI` 始终以任务模式运行；清理 `app_mode_*`、`btn_go_task_center`、「仅当前配置」toggle 相关 i18n 键与死代码。
- **dark 模式**：Checkbutton `command` 在 var 翻转后触发，不再二次反转；无 ttkbootstrap 时回退 cosmo。
- **启动构建容错**：`_build_task_tab_content` 末尾的 `_refresh_task_list_ui` 套 try/except，避免单点异常打爆后续 tab 构建链。

### Fixed
- **新建任务向导白屏**：修复 `self.config` 被误用（它是 tk 控件方法而非配置 dict），改用 `_load_config_for_write()`。
- **配置中心白屏**：`task_manager._iter_task_ids_from_disk` 排除 `schedules.json` 等非任务文件，避免被当作任务触发 `KeyError: 'id'` 导致级联失败。

### Documentation
- README（根 + 2026）、使用说明书 统一到 v5.20.0。
- 过时 plans 迁入 `docs/archive/`。

---
## [v5.20.0] - 2026-02-28
### Added
- **保存到任务**：任务 Tab 新增「保存到任务」按钮，一键将当前界面配置绑定写回所选任务（等价于编辑任务并选择「绑定当前配置」），无需再打开编辑弹窗。
- **批量运行**：任务列表支持多选（Ctrl/Shift+点击），新增「批量运行」按钮，按选择顺序将任务入队并依次执行；停止时清空队列。
- **定时运行**：新增「定时运行」按钮与计划逻辑，可为任务设置每日运行时间（HH:MM），计划持久化在 `tasks/schedules.json`，应用启动后每 60 秒检查并在到点时自动运行该任务（进程内调度，需程序保持打开）。

### Changed
- 任务列表 Treeview 的 `selectmode` 由 `browse` 改为 `extended`，以支持多选。
- 执行逻辑：抽出 `_run_single_task(task_id, resume)`，供单次运行、批量队列与定时调度共用；批量时在 worker 结束后调用 `_maybe_run_next_queued_task` 启动下一任务。

### Documentation
- AGENTS.md 新增「版本与变更记录」强制规则：交付新功能或重要修复须更新 `office_converter.py` 的 `__version__`、`CHANGELOG.md` 及 AGENTS.md 的变更摘要。

---
## [Post v5.19.1] - 2026-02-24
### Added
- **任务列表配置范围过滤**：新增“仅当前配置”开关，任务列表可只显示与当前活动配置/配置档绑定的任务。
- **任务筛选一致性**：状态筛选候选项改为基于当前可见任务集合生成，避免出现筛选项与列表不一致。
- **默认配置字段对齐**：`create_default_config()` 明确包含 `ui.task_current_config_only = true`。

### Fixed
- **合并 Map CSV 写入失败**：E2E 运行中发现 `merge_pdfs` 写入的 `map_records` 包含字段 `merged_pdf_md5`，但 `index_runtime.write_merge_map` 的 CSV fieldnames 未包含该列，导致 `DictWriter` 报错 "dict contains fields not in fieldnames"。已在 `converter/index_runtime.py` 的 `write_merge_map` 的 fieldnames 中增加 `merged_pdf_md5`，与 Excel MergeIndex 表及 `merge_pdfs` 写入逻辑一致。
- **Word 转 PDF（WPS/混合环境）**：当 `Documents.Open` 返回的对象无 `ExportAsFixedFormat` 或 `SaveAs`（如 WPS 返回的 "Open" 接口）时，先尝试 `doc.SaveAs(..., FileFormat=wdFormatPDF)`，再尝试 `app.ActiveDocument.ExportAsFixedFormat` / `app.ActiveDocument.SaveAs`，减少部分 .doc/.docx 转换失败。
- **Excel 转 PDF Open 失败**：MS Excel 路径下 `Workbooks.Open` 带完整参数失败（如 COM -2147352567）时，回退为仅路径的 `Workbooks.Open(file_source)` 重试，提高部分 .xls/.xlsx 可转换率。
- **合并阶段单 PDF 损坏导致整轮失败**：合并时若某个 PDF 读取失败（如 PyPDF 报 "Stream has ended unexpectedly"、无效文件头等），不再导致整次 run 崩溃；`merge_pdfs` 中对该 PDF 打日志并跳过，继续合并其余文件；索引 PDF 读取失败时同样跳过并继续。
- **Markdown 导出阶段**：`pdf_markdown_export` 与 `pdf_content_scan` 中 PDF 读取若抛出异常（含 PyPDF 流错误），捕获后返回 None/跳过，避免单文件损坏导致整轮 E2E 失败。

### Changed
- **任务模式配置持久化**：`ui.task_current_config_only` 已接入读取、保存、配置合成与脏状态比较逻辑。
- **任务模式交互稳定性**：任务列表刷新、选择、写入细节等流程的非致命异常上报覆盖范围扩大，避免 UI 静默失败。

### Tests
- 新增/补充测试：
  - `tests/test_task_list_filter_sort.py`（含当前配置范围过滤）
  - `tests/test_default_config_schema.py`（默认 schema 字段校验）
  - `tests/test_nonfatal_ui_error_reporting.py`
  - `tests/test_gui_run_mode_state_behavior.py`
- 2026-02-24 全量验证：`python -m unittest discover -s tests -p "test_*.py" -v` -> `Ran 71 tests ... OK`

### Documentation
- 更新 `README.md`、`docs/test-reports/README.md`、`docs/test-reports/TEST_REPORT_SUMMARY.md`、`docs/notes/使用说明书.md`。
- 新增 `docs/plans/2026-02-24-office-converter-split-plan.md`（`office_converter.py` 拆分实施计划）。

---
## [v5.19.1] - 2026-02-15
### Fixed
- **GUI 滚动页面渲染修复**：修复 Canvas 内嵌窗口初始宽度为 0 导致所有 Tab 内容不可见的问题
  - 添加 `<Map>` / `<Expose>` 事件绑定，Canvas 可见时自动同步宽度
  - 添加启动后多次延迟刷新和 Tab 切换刷新机制
  - 兼容不同 Python 版本和 IDE 运行环境（VSCode / 终端）
- **PIL 导入冲突修复**：启动时清理 `sys.path` 中第三方污染路径（如 pipecad），避免加载错误的 Pillow

---

## [v5.19.0] - 2026-02-15
### Added
- **并发转换功能**：
  - 新增多线程并发转换支持，使用 ThreadPoolExecutor 实现
  - 可配置并发工作线程数（1-16，默认 4）
  - 预期可提升大批量文件转换速度 3-4 倍
  - 配置项：`enable_parallel_conversion`（开关）、`parallel_workers`（线程数）
- **断点续传功能**：
  - 转换过程中定期保存进度断点（默认每 10 个文件）
  - 程序中断后重启可自动从断点恢复
  - 可配置自动恢复或手动确认恢复
  - 断点文件存储在 `<target>/_AI/checkpoints/` 目录
  - 配置项：`enable_checkpoint`（开关）、`checkpoint_auto_resume`（自动恢复）、`parallel_checkpoint_interval`（保存间隔）
- **GUI 控件**：
  - 新增并发转换开关和并发数设置
  - 新增断点续传开关和自动恢复选项
  - 中英文双语提示文案

### Changed
- `run()` 方法根据 `enable_parallel_conversion` 配置自动选择串行或并发模式
- 串行模式 `run_batch()` 现也支持断点续传功能

### Technical Notes
- COM 对象线程安全：每个并发线程独立调用 `CoInitialize()` 和 `CoUninitialize()`
- 断点文件使用 source_folder 的 MD5 哈希作为标识，确保唯一性
- 并发模式下使用线程锁保护共享统计数据

---

## [v5.18.0] - 2026-02-15
### Added
- **失败文件异常处理增强**：
  - 新增错误类型自动分类：权限不足、文件被占用、文件损坏、COM错误、超时、磁盘空间不足、密码保护等
  - 每种错误类型提供针对性的处理建议
  - 转换失败时自动记录详细错误信息（错误类型、是否可重试、是否需人工处理）
- **失败文件报告导出**：
  - 自动生成 JSON 格式详细报告 (`failed_files_report_YYYYMMDD_HHMMSS.json`)
  - 自动生成可读文本报告 (`failed_files_report_YYYYMMDD_HHMMSS.txt`)
  - 报告包含：统计摘要、按错误类型分组的文件列表、处理建议
- **GUI 失败文件摘要显示**：运行结束后在日志面板显示失败文件统计和处理建议
- **统计信息增强**：新增按错误类型分类的计数（permission_denied、file_locked、file_corrupted 等）

### Changed
- 错误日志更详细：记录错误类型而不仅仅是原始异常信息
- 失败文件复制到 `_FAILED_FILES` 目录时保留原有功能

### Documentation
- 使用说明书新增「失败文件处理指南」章节，包含错误类型说明和处理建议

---

## [v5.17.0] - 2026-02-13
### Added
- **Google Drive 上传**：在「成果文件」Tab 中可将 `_LLM_UPLOAD` 目录一键上传到用户自己的 Google Drive。支持 OAuth 桌面流程、可选目标文件夹 ID、上传后更新 `llm_upload_manifest.json` 的 gdrive 区段；client_secrets 与 token 不入库、不写日志，详见 `docs/plans/Google_Drive_上传_实现规划.md`。
- Merge output filename pattern: configurable `merge_filename_pattern` in config and GUI (placeholders: `{category}`, `{timestamp}`, `{date}`, `{time}`, `{idx}`). Default: `Merged_{category}_{timestamp}_{idx}`.
- Build script `build_exe.py`: exe output name now includes version (e.g. `ZhiWei_v5.17.0.exe`), version read from `office_converter.py`; **packaging now clears `dist` and `build` before building**.
- **README.md** for GitHub: project intro, features, install, run, build, and doc links.

### Changed
- Multi-folder processing: converter now scans and processes **all** source folders in both task mode and classic mode. Added `_get_source_roots()` and `_get_source_root_for_path()`; convert, merge, collect, and MSHelp flows use multiple source roots when `source_folders` is set.
- Classic mode: each run step explicitly sets `source_folders = [step["source"]]` so one folder per step; task mode passes full `source_folders` and processes all in one run.

### Fixed
- Multi-select source folders were not actually processing files under each folder: only the first folder was used in task mode; converter now iterates over all `source_folders` for scanning and relpath resolution.

### Documentation & Release
- **Product rename**: 品牌名 **知喂 (ZhiWei)**，副标题「知识投喂工具」/「Office 转知识库」；exe 与脚本使用 `ZhiWei`。
- GUI: English mode removed; interface is Chinese-only.
- All user-facing docs updated (使用说明书、打包说明、CHANGELOG); AI handover doc and README aligned for public release.

---

## [v5.16.0] - 2026-02-13
### Added
- GUI: Multiple source folder selection with parent folder scanning feature for both classic mode and task wizard.
- GUI: Source folder multi-select dialog now supports two methods:
  - Method 1 (Recommended): Select parent folder and scan all subfolders at once
  - Method 2: Manual individual folder addition
- GUI: Added "Select All", "Deselect All", and "Invert Selection" controls for quick folder selection.
- GUI: Multi-select dialogs with improved window sizes to ensure all controls are fully visible.
- GUI: Task wizard source folder selection now supports the same multi-select functionality as classic mode.

## [Unreleased] - 2026-02-11
### Added
- GUI: New independent tab `成果文件` (Output Files) for managing LLM upload hub, upload manifest, and dedup settings. Always visible regardless of run mode.
- LLM Upload: Configurable upload manifest generation – text readme (`README_UPLOAD_LIST.txt`) and JSON manifest (`llm_upload_manifest.json`) can be independently toggled.
- LLM Upload: Configurable merge dedup – when enabled (default), individual source files already contained in merged documents are excluded from upload.
- LLM Upload: Flatten `_LLM_UPLOAD` by default (`llm_delivery_flatten = true`), all files go to root directory with clean original filenames instead of nested subdirectories.
- LLM Upload: Content-only whitelist filter – only `markdown_export`, `mshelp_merged_markdown`, `excel_structured_json`, and optionally merged/converted PDFs are collected; metadata files (manifests, quality reports, records JSON, ChromaDB) are excluded.
- MSHelp Merge: Unified merge size parameter – MSHelp merged markdown now uses `max_merge_size_mb` (default 80 MB) for package splitting, removing separate `mshelp_merge_max_docs` / `mshelp_merge_max_chars` parameters.
- GUI: Single-level main Notebook with 6 tabs (`模式与路径` / `转换选项` / `合并 / 梳理` / `MSHelp` / `快速定位` / `高级设置`), replacing the previous 2-level "运行中心 / 配置管理" structure.
- GUI: Merged sparse tabs so that merge + collect ("梳理") options share one page, reducing empty pages and tab switching.
- GUI: NotebookLM Locator restored as an independent tab (`快速定位`), always enabled regardless of run mode, for convenient source tracing at any time.
- GUI: Two-column layouts for `转换选项` and `高级设置` tabs to better utilize horizontal space on wide screens.
- GUI: Global spacing compression (labelframe padding, section gaps, help-row spacing, sub-option indentation) to reduce vertical scrolling while keeping visual hierarchy.
- GUI: Window size presets with screen-height-aware defaults (`1280x860` for standard displays, `1360x920` for ≥1080p) and stricter minimum size (`1000x700`).
- GUI: Remember main window size/position and state; persist under `ui.window_geometry` / `ui.window_state` in `config.json` and restore on next launch.
- Corpus manifest export: auto-generate `corpus.json` with artifact metadata, conversion records, merge records, and summary.
- GUI artifact summary output after each step.
- AI export toggles in GUI/runtime config: `enable_markdown`, `enable_excel_json`.
- Markdown export from converted PDFs into `_AI/Markdown/`.
- Markdown cleanup and structure enhancement: repeated header/footer/page-number stripping and heading/paragraph normalization.
- Markdown quality report export into `_AI/MarkdownQuality/` with sampling and aggregate stats.
- Records JSON export for conversion/merge index into `_AI/Records/`.
- Excel structured JSON export into `_AI/ExcelJSON/` with semantic enrichment (header detection, records preview, column profiles, formula sample, sheet links, merged ranges).
- Excel JSON deep semantic metadata: workbook defined names, sheet charts, and sheet pivot table descriptors.
- Incremental sync v1:
  - `FileRegistry` JSON ledger.
  - Added/Modified/Renamed/Unchanged/Deleted scan summary.
  - Source priority rule: skip same-dir same-stem PDF when Office source exists.
  - Same-type global MD5 dedup (runtime).
  - GUI toggles for incremental and dedup options.
- Incremental sync v2:
  - `Update_Package` generation for incremental runs (incremental PDFs + index JSON/CSV/XLSX + incremental manifest).
  - GUI toggle for enabling/disabling update package generation.
  - Rename detection in incremental scan (`renamed`), including registry/update-package trace fields.
  - Registry key strategy normalized to source-relative forward-slash paths (with backward compatibility for legacy absolute-key ledgers).
- ChromaDB adapter (base):
  - Optional export switch `enable_chromadb_export`.
  - Markdown chunk collection and ChromaDB `PersistentClient` upsert flow.
  - JSONL fallback output plus export manifest under `_AI/ChromaDB/`.
  - GUI toggle and artifact-summary integration.
- MSHelp API docs mode (`run_mode=mshelp_only`):
  - Scan `MSHelpViewer` folders and process `.cab` help packages.
  - CAB to Markdown conversion pipeline with source trace records.
  - MSHelp index outputs (`MSHelp_Index_*.json/.csv`).
  - Merged MSHelp package outputs (`MSHelp_API_Merged_*.md`, optional `.docx/.pdf`).
- Project handover package for next AI:
  - New technical handover document with architecture, state, risks, and next-step implementation guidance.
  - Unified documentation updates for upcoming V1.1 requirements.

### Changed
- Config defaults expanded and normalized in loader/default-config generation.
- Conversion batch now returns structured per-file results for downstream bookkeeping.
- Incremental registry persisted at end of conversion stage.
- GUI mouse wheel handling changed to global stable binding with canvas-subtree filtering (better Windows behavior).
- Rewrote `docs/INCREMENTAL_SYNC_DESIGN.md` to a clean UTF-8 version and aligned it with current incremental implementation (config keys, paths, states, and outputs).
- Refined Run Center layout for Convert mode: split runtime options into clearer blocks (`Convert`, `Filters & Strategy`, `AI Export`, `Incremental / Dedup`) to reduce configuration mixing.
- Refined Config Center layout for Convert mode: added dedicated tabs for `AI Defaults` and `Incremental Defaults`.
- Refined Config Center grouping depth: split AI defaults into `Text & Manifest` + `Structured Data & Vector`, and split incremental defaults into `Scan & Rename` + `Dedup & Package`.
- Refined `Shared Defaults` grouping: split into dedicated blocks for config path, process strategy, and log output.
- Refined `Merge Defaults` grouping: split into `Merge Behavior` and `Merge Output`, and unified merge-enable label through i18n.
- Added a Config Center hint label clarifying these values are persisted defaults written to `config.json`.
- Refined `Rules & Keywords` grouping: split into `Exclude Rules` and `Keyword Strategy` defaults.
- Added per-section reset controls in Config Center (`Reset This Section`) for Shared/Convert/AI/Incremental/Merge/UI/Rules defaults, with explicit "not saved yet" status feedback.
- Added Config Center dirty-state indicator (saved/unsaved) that flips to unsaved after default-value edits and resets to saved after `Save All`.
- Improved dirty-state accuracy: after both `Save All` and `Save Current Mode`, status is recomputed by UI-vs-config snapshot diff instead of naive scope-based clearing.
- Added per-section dirty markers in Config Center tabs (tab title suffix `*`) based on snapshot diff, so unsaved changes are visible by section.
- Added Config Center unsaved-section summary line (lists dirty sections by name), synchronized with tab `*` markers and snapshot-diff state.
- Added `Go to Unsaved` action in Config Center to jump directly to the first dirty config tab.
- Added clickable dirty-section chips in Config Center summary so each unsaved section can be opened directly.
- Added `Save Unsaved Sections` action in Config Center to persist only dirty sections instead of full config writes.
- Added `Save This Section` action in each Config Center section to persist that section only.
- Improved `Save Unsaved Sections` feedback: now reports saved section names and shows a clear "no unsaved sections" notice when applicable.
- Added `Revert Unsaved Sections` action in Config Center to roll back dirty sections to current `config.json` baseline, with section-level feedback.
- Added dynamic dirty-count labels for Config Center actions (`Save Unsaved Sections` / `Revert Unsaved Sections`).
- Improved `Revert Unsaved Sections` safety: added confirmation dialog with section scope and refreshed baseline from `config.json` before rollback.
- Added configurable revert-confirmation toggle in UI defaults (`confirm_revert_dirty`), persisted in `config.json`.
- Updated project docs to align with the next requested user goals:
  - Centralized LLM upload output folder (single ingestion path).
  - Adjustable sandbox location and low-disk safety strategy for large incremental runs.

### Fixed
- Repaired broken/mis-encoded string literals in `office_converter.py` that could cause runtime/compile failures.
- Normalized high-frequency conversion status labels and conflict statuses to readable values (`success/overwrite/conflict_saved`, `skip_*`) in main conversion flow.
- Added console stream fallback (`errors=\"replace\"`) in runtime modules to prevent `UnicodeEncodeError` on legacy Windows code pages.
- Repaired the corrupted `ask_retry_failed_files` block that caused `SyntaxError` on startup.
- Fixed `_compute_md5` binding (`@staticmethod`) so source MD5 fields are reliably populated in incremental indexes/manifests.
- Cleaned merge/collect CLI user-facing messages and index-sheet headers to readable text (removed mojibake output paths/prompts).
- Normalized remaining user-facing `print`/`logging`/`raise` messages in `office_converter.py` to avoid Windows mojibake in runtime output.
- Cleaned mojibake comments/docstrings in key merge/batch sections for maintainability (no logic changes).
- Fixed malformed multiline f-strings in `office_gui.py` start/stop workflow logs that caused `SyntaxError` on Windows.
- Cleaned remaining mojibake comments/messages in `office_converter.py` (path/content/COM/core headers, Excel keyword hit log, index page setup comment, and index font name).
- Refined GUI wheel routing on Windows: scrollable child controls (`Listbox/Text/Treeview/Canvas`) now keep native wheel behavior, avoiding page-canvas wheel hijacking.
- Fixed mojibake icon on path-row open button and replaced it with a stable ASCII symbol (`>`).
- Cleaned remaining mojibake comments/docstrings in `office_gui.py` key sections (`footer`, `UI state sync`, `config read/write`).

## [v5.15.7] - 2026-01-08
### Fixed
- PDF merge compatibility for `pypdf` versions without `indirect_ref` (fallback to `indirect_reference`).

## [v5.15.6] - 2026-01-08
### Fixed
- Merge output reliability: post-write validation and output path reporting.
- UX: after merge, open `_MERGED` directory when outputs exist.

## [v5.15.5] - 2026-01-08
### Changed
- Index page hotspots: all index entries are clickable across multiple index pages.

## [v5.15.4] - 2026-01-08
### Fixed
- Removed duplicated legacy merge logic causing incorrect Excel output and double processing.
- Merge list Excel now correctly writes one row per component file.

## [v5.15.3] - 2026-01-08
### Fixed
- Added missing `traceback` import for merge error logging.

## [v5.15.2] - 2026-01-08
### Added
- Clickable links on generated PDF index pages.
- Enhanced merge Excel listing (output filename + source filename per row).

### Changed
- Refactored merge task generation for stability.
- Improved index filename truncation to avoid layout overflow.

## [v5.15.1] - 2026-01-08
### Added
- Date-based file filtering (before/after).
- Merge options: index page generation and Excel list export.
- UI controls for new options.


