# Changelog

All notable changes to this project are documented in this file.

## [v5.17.0] - 2026-02-13
### Added
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
