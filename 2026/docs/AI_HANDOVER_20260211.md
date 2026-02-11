# AI Handover (2026-02-11)

## 1. Project Snapshot
- Repo: `GPTVersion/2026`
- Core runtime:
  - `office_converter.py` (pipeline, conversion, merge, AI exports, incremental, MSHelp)
  - `office_gui.py` (Tk/ttkbootstrap UI, run/config centers)
  - `ui_translations.py` (zh/en i18n)
- Current version marker in code: `v5.15.6` (code), docs/changelog updated to current workstream.

## 2. Implemented Capability Baseline
### 2.1 Existing run modes
- `convert_only`
- `merge_only`
- `convert_then_merge`
- `collect_only`
- `mshelp_only` (new independent main feature)

### 2.2 MSHelp feature status (already implemented)
- Scan `MSHelpViewer` folders under source root.
- Convert CAB/MSHC help payload into readable Markdown.
- Export MSHelp source index (`json/csv`).
- Build merged packages (`md`, optional `docx/pdf`).
- GUI includes dedicated MSHelp mode + runtime/config parameters.

### 2.3 AI artifacts current output pattern (important)
Current outputs are distributed in multiple folders:
- `_AI/Markdown/`
- `_AI/MarkdownQuality/`
- `_AI/ExcelJSON/`
- `_AI/Records/`
- `_AI/ChromaDB/`
- `_AI/MSHelp/` and `_AI/MSHelp/Merged/`
- `_AI/Update_Package/`

This is the key pain point from user for next iteration.

## 3. User-confirmed Next Requirements (priority)
### 3.1 Requirement A: Single-folder LLM upload output
User wants all files intended for NotebookLM/LLM upload to be centralized in ONE folder, not scattered.

Expected objective:
- One direct upload directory per run (or stable single entry).
- Keep source traceability and manifest.
- Do not break existing raw `_AI/*` outputs.

### 3.2 Requirement B: Sandbox path + disk-space safety
User is concerned about large-scale runs (10k files) exhausting disk.

Current state:
- Sandbox root is already configurable via `temp_sandbox_root`.

Gap:
- No pre-run free-space check.
- No low-space policy/guardrail in UI + runtime.

## 4. Recommended V1.1 Design (for next AI)
### 4.1 New config keys
- `enable_llm_delivery_hub` (bool, default `true`)
- `llm_delivery_root` (string, default `<target>/_LLM_UPLOAD`)
- `llm_delivery_flatten` (bool, default `false`)  
  - `false`: categorized subfolders under one root
  - `true`: flatten to single-level with collision-safe names
- `llm_delivery_include_pdf` (bool, default `false`)  
  - avoid giant upload bundles by default
- `sandbox_min_free_gb` (int, default `10`)
- `sandbox_low_space_policy` (enum: `block|confirm|warn`, default `block`)

### 4.2 Backend additions (`office_converter.py`)
Add postprocess phase functions:
- `_collect_llm_delivery_candidates()`
- `_sync_llm_delivery_hub()`
- `_write_llm_upload_manifest()`
- `_check_sandbox_free_space_or_raise()`

Suggested integration points:
- Before batch start: sandbox free-space precheck.
- After all artifacts generated: build `_LLM_UPLOAD` hub and manifest.
- Add hub files into `corpus.json` artifacts.

### 4.3 GUI additions (`office_gui.py`)
- Runtime + config controls for above keys.
- Show sandbox current free space and low-space policy.
- Add artifact summary lines:
  - `LLM hub path`
  - `LLM deliverable file count`

### 4.4 i18n additions (`ui_translations.py`)
Add keys for:
- group title, labels, toggles, policy options
- validation messages for `sandbox_min_free_gb`
- low-space warnings/confirm dialog texts

## 5. Acceptance Criteria (must-pass)
1. User can upload from one folder only (`_LLM_UPLOAD`).
2. `llm_upload_manifest.json` exists and includes source->delivery mapping.
3. Legacy `_AI/*` directories still generated (backward compatibility).
4. Sandbox low-space condition is detected before long run starts.
5. Behavior follows policy (`block/confirm/warn`) and is visible in GUI logs.

## 6. Validation Plan
### 6.1 Functional
- Run each mode once and verify hub composition.
- Verify filename collision handling in flattened mode.
- Verify manifest correctness for at least 20 files.

### 6.2 Large-scale
- Simulate 10k file scan with low free-space environment.
- Confirm precheck blocks/warns as configured.

### 6.3 Regression
- `python -m py_compile office_converter.py office_gui.py ui_translations.py`
- sanity run for existing convert+merge workflow.

## 7. Risks and Mitigation
- Risk: duplicate storage overhead when copying files to hub.
  - Mitigation: allow hardlink/symlink mode in future; start with copy for compatibility.
- Risk: oversized LLM upload package.
  - Mitigation: default exclude PDF; optional include.
- Risk: mode-specific outputs inconsistent.
  - Mitigation: central candidate collector by artifact kind + whitelist.

## 8. Suggested Implementation Order
1. Add config keys + defaults + GUI inputs.
2. Add sandbox free-space precheck and policy handling.
3. Add LLM hub sync + manifest generation.
4. Wire artifact summary and corpus integration.
5. Update docs and run validation checklist.

## 9. Files to Touch Next
- `office_converter.py`
- `office_gui.py`
- `ui_translations.py`
- `docs/PRODUCT_REQUIREMENTS.md`
- `docs/TASK_LIST.md`
- `使用说明书.md`
- `CHANGELOG.md`

## 10. Handover Notes
- Keep V1 compatibility as hard constraint.
- Do not remove existing `_AI` outputs; add hub as overlay layer.
- If free-space precheck is uncertain on non-Windows, implement conservative warn fallback.
