# Unified Output Controls Implementation Plan

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task.

**Goal:** Unify output controls so all run modes support independent PDF/MD format toggles and merged/independent strategy toggles, plus the new "Merge & Convert" sub-functions.

**Architecture:** Add output decision primitives in `office_converter.py`, then route convert/merge flows through these primitives. Extend GUI runtime config to expose explicit global output controls and merge sub-function choice. Update user manual with scenario-specific flowcharts and config recommendations.

**Tech Stack:** Python (tkinter/ttkbootstrap), existing converter pipeline, Markdown docs.

---

### Task 1: Add failing tests for output decision logic

**Files:**
- Create: `tests/test_output_controls.py`
- Modify: none

1. Add tests for conversion output plan combinations:
- convert_only + PDF independent
- convert_only + MD independent
- convert_then_merge + merged PDF only
- convert_then_merge + merged MD only
- both formats disabled
2. Run tests and verify failure (missing function/class API).

### Task 2: Implement output decision primitives and defaults

**Files:**
- Modify: `office_converter.py`

1. Add config defaults:
- `output_enable_pdf`
- `output_enable_md`
- `output_enable_merged`
- `output_enable_independent`
- `merge_convert_submode`
2. Implement static/helper output plan function for convert flow.
3. Wire helper into `process_single_file` and `run_batch`.

### Task 3: Implement merge markdown + merge-and-convert sub-functions

**Files:**
- Modify: `office_converter.py`

1. Add merge submode constants and handlers:
- `merge_only`
- `pdf_to_md`
2. Add markdown merge pipeline for `.md` files (all-in-one/category split naming aligned with current merge outputs).
3. Extend merge-mode runtime branch:
- merge-only path: merge PDF and/or MD based on output toggles.
- pdf-to-md path: convert PDF->MD, optional merge outputs by toggles.
4. Add missing-MD behavior:
- prompt continue/exit in interactive mode.
- return continue by default in non-interactive mode (GUI precheck will enforce explicit choice).

### Task 4: Add GUI controls and persistence

**Files:**
- Modify: `office_gui.py`
- Modify: `ui_translations.py`

1. Add prominent global output panel in run-mode area:
- `PDF`
- `MD`
- `Merged Output`
- `Independent Output`
2. Rename displayed mode label to `Merge & Convert`.
3. Add merge sub-function selector:
- `Merge Existing (PDF/MD)`
- `PDF -> MD`
4. Persist/load new config keys in runtime snapshot + save/load paths.
5. Add GUI preflight prompt for "MD merge requested but no .md files": Continue/Exit.

### Task 5: Update docs with flowcharts and per-scenario guidance

**Files:**
- Modify: `使用说明书.md`

1. Update mode names and output control section.
2. Add flowchart for each usage scenario:
- Convert only
- Merge & Convert / Merge Existing
- Merge & Convert / PDF->MD
- Convert+Merge
- Collect
- MSHelp
3. Add recommended config templates per scenario.

### Task 6: Verify

**Files:**
- Modify: none

1. Run unit tests.
2. Run `python -m py_compile office_converter.py office_gui.py ui_translations.py`.
3. Report changed files and key behavior mapping.
