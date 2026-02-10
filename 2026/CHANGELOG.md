# Changelog

All notable changes to this project will be documented in this file.

## [v5.15.7] - 2026-01-08
### Fixed
- **PDF Merge Compatibility**: Fixed `AttributeError: 'PageObject' object has no attribute 'indirect_ref'` by adding fallback support for older `pypdf` versions (using `indirect_reference` or skipping links safely).

## [v5.15.6] - 2026-01-08
### Fixed
- **Merge Output**: Added post-write validation and printed output paths to avoid “silent no-output” situations.
- **UX**: After merge, opens the `_MERGED` directory when outputs exist.

## [v5.15.5] - 2026-01-08
### Changed
- **PDF Index Hotspots**: Made all index entries clickable across multiple index pages.

## [v5.15.4] - 2026-01-08
### Fixed
- **Merge Logic**: Completely removed duplicated legacy merge logic that caused incorrect Excel output and double processing.
- **Excel Listing**: Verified that the merge list now correctly outputs one row per component file.

## [v5.15.3] - 2026-01-08
### Fixed
- **Runtime Error**: Fixed missing `traceback` import causing crash when logging merge errors.

## [v5.15.2] - 2026-01-08
### Added
- **PDF Index Hotspots**: The generated index page now includes clickable links jumping to the corresponding document start page.
- **Enhanced Excel Listing**: The merge list Excel now records one row per component file (Output Filename, Source Filename) for easier tracking.

### Changed
- **Merge Logic**: Refactored merge task generation for better stability.
- **Index Layout**: Improved filename truncation on index page to prevent layout overflow.

## [v5.15.1] - 2026-01-08
### Added
- Date-based file filtering (before/after specific date).
- Merge options: Index page generation and Excel file listing.
- UI controls for new features.
# v5.16.0

- **GUI 双 Tab 拆分**: 主界面重构为“运行中心 / 配置中心”，将高频运行操作与低频配置管理解耦，默认进入运行中心。
- **定位模块归位**: NotebookLM 快速定位模块固定在“运行中心”，保留 Everything/Listary 联动。
- **配置页强调保存**: 在“配置中心”增加显式保存入口，同时保留原保存逻辑与配置兼容性。
- **NotebookLM 溯源定位**: 合并输出新增 `*.map.csv`/`*.map.json`，包含页码范围、源文件路径、MD5、短ID。
- **短ID书签**: 合并书签可写入 `[ID:XXXXXXXX] 文件名`，便于 NotebookLM/人工引用时快速回查。
- **定位 CLI**: 新增 `locate_source.py`，支持 `merged+page` 与 `short_id` 双路径反查，并支持 `--json` 输出。
- **检索增强**: 新增 `search_adapter.py`，支持 Everything（`es.exe`）查询联动与 Listary 查询词生成。
- **GUI 快速定位面板**: 新增 NotebookLM 定位区（刷新 map、定位、打开文件/目录、Everything 搜索、复制 Listary 查询）。
