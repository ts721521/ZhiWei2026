# NotebookLM Preset Configs

These presets are for common NotebookLM ingestion scenarios.

## How To Use

1. Copy one preset JSON to your runtime config (or load directly in GUI).
2. Change `source_folder/source_folders` and `target_folder` to your environment.
3. Keep traceability keys enabled unless you explicitly do not need source backtracking.

## Required Scenarios

### 1) Merged MD Only
- File: `configs/presets/notebooklm/config.presets.notebooklm_merged_md_only.json`
- Use when NotebookLM input is markdown-first and you want fewer files.
- Output intent:
  - merged markdown on
  - PDF off
  - independent outputs off

### 2) Merged PDF Only
- File: `configs/presets/notebooklm/config.presets.notebooklm_merged_pdf_only.json`
- Use when your downstream reviewers prefer original pagination/layout from PDF.
- Output intent:
  - merged PDF on
  - markdown off
  - independent outputs off

### 3) Merged PDF + Merged MD
- File: `configs/presets/notebooklm/config.presets.notebooklm_merged_pdf_md.json`
- Use when both retrieval quality (MD) and visual fidelity (PDF) are needed.
- Output intent:
  - merged PDF on
  - merged markdown on
  - independent outputs off

### 4) Word PDF+MD 合并与图片溯源（推荐用于 LLM 理解）
- File: `configs/presets/notebooklm/config.presets.notebooklm_word_pdf_md_trace.json`
- **用途**：PPT 转 PDF、Word 转 MD+PDF、Excel 转 MD；Word/PPT 的 PDF 合并；合并 PDF 带书签与 map；**Word 内图片在合并 PDF 中的位置可追溯，便于 LLM 理解**。
- 关键开关：`enable_merge_map`、`bookmark_with_short_id`、`enable_markdown_image_manifest` 均为 true；与 merged_pdf_md 类似，但强调图片与合并 PDF 页码的对应关系（image manifest）。
- 使用前请设置 `source_folder`/`source_folders` 和 `target_folder`。

## 预设配置说明一览

| 名称 | 文件 | 配置说明（一句话） | 备注摘要 |
|------|------|-------------------|----------|
| merged_md_only | config.presets.notebooklm_merged_md_only.json | 仅合并 Markdown，不上传 PDF，减少文件数量。 | 使用前设置路径；保持溯源开关。 |
| merged_pdf_only | config.presets.notebooklm_merged_pdf_only.json | 仅合并 PDF，保留版式与页码，适合原样排版审阅。 | 使用前设置路径；审阅依赖 PDF 页码时选用。 |
| merged_pdf_md | config.presets.notebooklm_merged_pdf_md.json | 合并 PDF + 合并 Markdown，兼顾检索质量与版式保真。 | 使用前设置路径；需同时出 PDF 与 MD 时推荐。 |
| word_pdf_md_trace | config.presets.notebooklm_word_pdf_md_trace.json | PPT→PDF、Word→MD+PDF、Excel→MD，Word 的 PDF 合并，图片在合并 PDF 中位置可追溯便于 LLM 理解。 | 使用前设置路径；保留 merge map 与 image manifest。 |
| fast_md_only | config.presets.notebooklm_fast_md_only.json | 极速仅 MD 模式，不走 Office/WPS COM，依赖 markitdown。 | 需安装 markitdown；使用前设置路径。 |
| incremental_weekly_sync | config.presets.notebooklm_incremental_weekly_sync.json | 增量周同步，只处理新增/变更，减少全量重跑。 | 需配置 incremental_registry_path；使用前设置路径。 |
| traceability_audit | config.presets.notebooklm_traceability_audit.json | 溯源审计：独立输出、不合并，便于排查与证据留存。 | 用于排查源文件映射与审计；使用前设置路径。 |

更多配置总览与流程图见 [CONFIG_REFERENCE.md](CONFIG_REFERENCE.md)。

### 5) Fast MD Only (No Office/WPS COM path)
- File: `configs/presets/notebooklm/config.presets.notebooklm_fast_md_only.json`
- Use when you want maximum speed and direct markdown outputs.
- Notes:
  - requires `markitdown` runtime availability
  - no merged artifacts by default

### 6) Incremental Weekly Sync
- File: `configs/presets/notebooklm/config.presets.notebooklm_incremental_weekly_sync.json`
- Use for recurring sync jobs to reduce full reprocessing cost.
- Notes:
  - incremental registry enabled
  - merged PDF+MD enabled

### 7) Traceability Audit
- File: `configs/presets/notebooklm/config.presets.notebooklm_traceability_audit.json`
- Use for troubleshooting backtracking and evidence review.
- Notes:
  - independent outputs enabled
  - merge off
  - extra index/trace logs on

## Traceability (Reverse Location) Checklist

For stable source backtracking, keep these enabled:

- `enable_traceability_anchor_and_map = true`
- `enable_merge_map = true`
- `bookmark_with_short_id = true`
- `short_id_prefix` set (default `ZW-`)
- `enable_corpus_manifest = true`
- `enable_upload_json_manifest = true`

Key artifacts to check after run:

- `trace_map.xlsx`
- `corpus.json`
- merged file front sections and `source_short_id` markers
- `_LLM_UPLOAD/llm_upload_manifest.json`

## Practical Size Control For NotebookLM

- `markdown_max_size_mb` is the unified markdown size cap (FastMD single outputs + Knowledge Bundle + merged markdown outputs).
- If `markdown_max_size_mb` is not set, system falls back to `max_merge_size_mb`.
- Keep `max_merge_size_mb` / `markdown_max_size_mb` conservative (e.g. `100~120`) to reduce single-file risk.
- Prefer merged-only upload sets for large corpora.
- If very large corpora are expected, split by category/date and run multiple jobs.
