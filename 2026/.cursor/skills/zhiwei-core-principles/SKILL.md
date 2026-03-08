---
name: zhiwei-core-principles
description: Core principles for ZhiWei project - NotebookLM corpus preparation. This skill is ALWAYS relevant for any development task. The project goal is to convert and merge files for NotebookLM, NOT to copy files.
---

# ZhiWei Core Principles

These principles must be followed for ALL development in this project.

## Project Mission

**ZhiWei (知喂) is a corpus preparation tool for NotebookLM.**

The core workflow is:
1. Convert Office documents (Word/Excel/PPT) to PDF/Markdown
2. Merge converted files into volumes suitable for NotebookLM
3. Output to `_LLM_UPLOAD` directory for upload

**IMPORTANT: The program CONVERTS and MERGES files. It does NOT simply copy files.**

## NotebookLM Limits

| Limit | Value | Action if exceeded |
|-------|-------|-------------------|
| Single file size | <= 100MB | Adjust merge strategy, split volumes |
| Sources per notebook | <= 300 (free tier) | Merge more files, reduce output count |
| Supported formats | PDF, DOCX, TXT, Markdown, Google Docs/Sheets/Slides | Output in these formats |

## _LLM_UPLOAD Directory Rules

1. **MUST be cleared before each run** - No old files should remain
2. **Only contains current run output** - Fresh merged/converted files only
3. **Files must be properly merged** - Not just copied from source

## Development Principles

1. **Convert + Merge, not Copy** - Every file in `_LLM_UPLOAD` should be processed
2. **Respect size limits** - Merged files should not exceed 100MB
3. **Respect count limits** - Output files should not exceed 300 for free tier
4. **Fresh output each run** - Clear old data before generating new

## When Making Code Changes

Ask yourself:
- Does this change support the convert+merge workflow?
- Will the output comply with NotebookLM limits?
- Is `_LLM_UPLOAD` handled correctly (cleared before run)?

## Key Files

| File | Purpose |
|------|---------|
| `converter/corpus_manifest.py` | `_LLM_UPLOAD` directory management |
| `converter/merge_pdfs.py` | PDF merging logic |
| `converter/merge_markdown.py` | Markdown merging logic |
| `office_converter.py` | Main conversion workflow |