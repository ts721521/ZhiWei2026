# 增量同步与增量包设计（v2）

> 文档状态：已按当前实现校正（2026-02-10）  
> 对应代码：`office_converter.py` 中 `FileRegistry`、`_apply_incremental_filter`、`_flush_incremental_registry`、`_generate_update_package`

## 1. 目标与范围

### 1.1 目标
- 只处理新增/修改文件，跳过未变化文件。
- 对“重命名”进行检测并在索引中保留追溯关系。
- 生成可直接投喂 AI 的增量包（PDF 子集 + 索引 + manifest）。

### 1.2 生效范围
- 仅在 `run_mode` 为 `convert_only` 或 `convert_then_merge` 时参与转换链路。
- `merge_only` 与 `collect_only` 不触发增量扫描流程。

## 2. 配置项（与代码一致）

```json
{
  "enable_incremental_mode": false,
  "incremental_verify_hash": false,
  "incremental_reprocess_renamed": false,
  "incremental_registry_path": "",
  "source_priority_skip_same_name_pdf": false,
  "global_md5_dedup": false,
  "enable_update_package": true,
  "update_package_root": ""
}
```

配置说明：
- `enable_incremental_mode`：开启增量模式主开关。
- `incremental_verify_hash`：增量比较时是否启用 SHA256 深度校验。
- `incremental_reprocess_renamed`：检测到重命名后是否重新转换。
- `incremental_registry_path`：账本路径；留空时使用默认路径。
- `source_priority_skip_same_name_pdf`：同目录同名时优先 Office，跳过 PDF。
- `global_md5_dedup`：同类型文件（word/excel/powerpoint/pdf）按 MD5 去重。
- `enable_update_package`：是否生成增量包。
- `update_package_root`：增量包输出根目录；留空走默认路径。

## 3. 路径与默认输出

### 3.1 增量账本
- 默认路径：`<target_folder>/_AI/registry/incremental_registry.json`
- 账本键策略：
- 使用相对 `source_folder` 的正斜杠路径（`/`）。
- 在 Windows 下 key 统一转小写（兼容大小写不敏感文件系统）。

### 3.2 增量包
- 默认根目录：`<target_folder>/_AI/Update_Package/`
- 单次输出目录：`Update_Package_YYYYMMDD_HHMMSS/`
- 目录内容：
- `incremental_manifest.json`
- `incremental_index.json`
- `incremental_index.csv`
- `incremental_index.xlsx`（仅当 `openpyxl` 可用）
- `PDF/`（本次成功转换并纳入增量包的 PDF 子集）

## 4. 流程设计

执行顺序（与 `run()` 实际顺序一致）：
1. 扫描候选文件（扩展名 + 目录过滤 + 日期过滤）。
2. `source priority`：同目录同 stem 且存在 Office 时跳过 PDF。
3. 增量扫描：输出 Added/Modified/Renamed/Unchanged/Deleted 统计。
4. 全局 MD5 去重：同类型且同 MD5 的文件跳过。
5. 执行转换批处理（仅对保留文件）。
6. 刷新账本 `incremental_registry.json`。
7. 生成增量包（如开启）。

## 5. 状态判定规则

### 5.1 基础状态
- `added`：账本无记录。
- `modified`：同路径下尺寸/mtime（或哈希）变化。
- `unchanged`：与账本一致，跳过转换。
- `deleted`：账本存在但本次扫描不存在。

### 5.2 重命名检测（`renamed`）

检测入口：
- 仅在“本次 `added`”与“历史 `deleted`”之间匹配。

匹配条件（顺序）：
1. 同扩展名 + 同大小，且 SHA256 一致（最优先）。
2. 同扩展名 + 同大小，且 `mtime_ns` 一致。
3. 仅剩唯一候选时，使用 `ext+size` 唯一匹配兜底。

匹配后行为：
- 当前文件状态从 `added` 改为 `renamed`。
- 记录 `renamed_from`、`rename_match_type`。
- 若 `incremental_reprocess_renamed=false`，该文件不进入转换队列，结果记为 `renamed_detected`。

## 6. 账本结构（核心字段）

`entries[key]` 典型字段：
- `source_path`
- `ext`
- `source_size`
- `source_mtime`
- `source_mtime_ns`
- `source_hash_sha256`（当启用哈希校验或重命名检测触发计算时写入）
- `change_state`
- `renamed_from`
- `rename_match_type`
- `last_seen_at`
- `last_run_mode`
- `last_status`
- `last_error`
- `last_processed_at`
- `last_output_pdf`
- `last_output_pdf_md5`

顶层结构包含：
- `version`
- `updated_at`
- `key_strategy`
- `entry_count`
- `entries`
- `last_run`（扫描统计汇总）

## 7. 增量包索引字段

`incremental_index.json/csv/xlsx` 字段：
- `seq`
- `change_state`
- `process_status`
- `source_file`
- `source_path`
- `source_md5`
- `source_sha256`
- `renamed_from`
- `rename_match_type`
- `packaged_pdf`
- `packaged_pdf_path`
- `packaged_pdf_md5`
- `note`

说明：
- 仅 `process_status=success` 且目标 PDF 存在时才复制到 `PDF/`。
- `renamed_detected` 且未重转时不会产生 `packaged_pdf`。

## 8. 设计边界与已知行为

- 同名优先规则仅作用于“同目录 + 同 stem”场景，不跨目录误杀文件。
- 全局 MD5 去重按“同类型桶”执行，不做跨类型去重（例如 docx 与 pdf 不互判）。
- 删除（`deleted`）只体现在账本和 manifest 统计，不会删除既有输出文件。
- 重命名未重转时默认只记账和索引，不自动复用/移动旧 PDF 产物。

## 9. 验收清单

- 开启增量模式后，`unchanged` 文件不再重复转换。
- 重命名可在扫描统计中体现为 `renamed`，并写入 `renamed_from`。
- 每次运行结束后账本可落盘并可被下次运行复用。
- 开启增量包时，目录结构和索引文件完整可追溯。

---

## 10. Addendum (2026-02-11): LLM Delivery Hub + Sandbox Capacity Guard
This addendum defines two required enhancements for the next iteration:

### A) LLM Delivery Hub
- Build one consolidated upload folder per run.
- Keep source mapping via `llm_upload_manifest.json`.
- Do not replace existing `_AI/*` outputs; hub is an additional convenience layer.

### B) Sandbox Capacity Guard
- Before long conversion starts, check free disk space for sandbox root.
- Enforce configurable policy:
  - `block`: stop run
  - `confirm`: ask user
  - `warn`: continue with warning

### Incremental-specific note
For large incremental updates (10k+ files), guard should run before scan->convert stage to avoid mid-run disk exhaustion.
