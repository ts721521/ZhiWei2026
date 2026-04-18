# Config Layout

This directory centralizes all non-runtime configuration files to keep `2026/` root clean.

## Structure

- `templates/`
  - Baseline templates for manual bootstrap.
  - Example: `templates/config.example.json`
- `scenarios/`
  - Runnable scenario configs for scripts and smoke/E2E jobs.
  - Example: `scenarios/notebooklm/config.notebooklm_test.json`
- `presets/`
  - User-facing preset packs for common workflows.
  - Example: `presets/notebooklm/config.presets.notebooklm_merged_pdf_md.json`
- `diagnostics/`
  - Temporary or targeted troubleshooting configs.
  - Keep only active probes; archive stale ones.
- `archive/`
  - Deprecated/unused configs kept only for traceability.

## Rules

1. Runtime config stays as `2026/config.json` (generated/edited by app).
2. New reusable config files must be created under one of the folders above, not in `2026/` root.
3. New docs and scripts must reference paths under `configs/`.
4. For backward compatibility, scripts may read legacy root paths, then migrate to `configs/` on write.

## Metadata Convention (`_meta`)

To keep preset/scenario files self-descriptive and maintainable, each reusable config should include a top-level `_meta` object.

Recommended fields:

- `config_kind`: `preset` or `scenario`
- `preset_group` / `scenario_group`: domain group such as `notebooklm`
- `preset_name` / `scenario_name`: stable identifier
- `description`: one-line purpose
- `last_updated`: date in `YYYY-MM-DD`
- `notes`: list of operational hints

Example:

```json
{
  "_meta": {
    "config_kind": "preset",
    "preset_group": "notebooklm",
    "preset_name": "merged_pdf_md",
    "description": "NotebookLM merged PDF + merged markdown for balanced quality/fidelity.",
    "last_updated": "2026-02-27",
    "notes": [
      "Set source_folder/source_folders and target_folder before run.",
      "Recommended default when both PDF and markdown are needed."
    ]
  }
}
```

## 配置说明与界面展示

- **`_meta.description`** 与 **`_meta.notes`** 会显示在软件「加载配置」对话框的**备注**列中；内置预设与场景在列表中不再显示为固定的 "Built-in config"，而是从各 JSON 的 `_meta` 读取后展示，便于用户直接看到配置用途。
- 完整配置说明、预设一览表及「按场景选预设」「数据流」流程图见：**[docs/design/CONFIG_REFERENCE.md](../docs/design/CONFIG_REFERENCE.md)**；NotebookLM 预设的详细说明见 **[docs/design/CONFIG_PRESETS_NOTEBOOKLM.md](../docs/design/CONFIG_PRESETS_NOTEBOOKLM.md)**。

