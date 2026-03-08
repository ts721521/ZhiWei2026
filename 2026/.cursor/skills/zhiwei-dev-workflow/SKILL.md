---
name: zhiwei-dev-workflow
description: ZhiWei project development workflow including code changes, documentation updates, and version management. Use when user asks to develop new features, fix bugs, or manage project workflow.
---

# ZhiWei Development Workflow

Follow this workflow for all development tasks in the ZhiWei project.

## Pre-Development

1. Read `AGENTS.md` for current rules and conventions
2. Check `docs/TASK_LIST.md` for pending tasks
3. Review `docs/plans/` for relevant design documents

## Development Steps

### 1. Code Changes

Make changes to relevant files in:
- `converter/` - Core conversion logic
- `gui/mixins/` - GUI components
- `tests/` - Test files

### 2. Required Documentation Updates

When modifying code paths, update these files:

| Changed | Must Update |
|---------|-------------|
| Core modules | `docs/plans/2026-02-24-office-converter-split-handover.md` |
| Tests | `docs/test-reports/TEST_REPORT_SUMMARY.md` |
| Paths/Directories | `AGENTS.md` |

### 3. Version Updates

For new features or important fixes:

1. Update version in `office_converter.py`:
```python
__version__ = "X.Y.Z"
```

2. Add CHANGELOG entry in `CHANGELOG.md`:
```markdown
## [vX.Y.Z] - YYYY-MM-DD

### Added
- Description of new feature

### Fixed
- Description of bug fix
```

### 4. Test Execution

Run tests before completing:
```bash
python -m unittest discover -s tests -p "test_*.py" -v
```

### 5. Handover Documentation

Before ending the session, update:
- `docs/AI_交接文档_下一阶段开发.md` or
- `docs/test-reports/TEST_REPORT_SUMMARY.md`

Include:
- Changes made this round
- Test results
- Unfinished items and risks

## Project Entry Points

```bash
# GUI
python office_gui.py

# CLI
python office_converter.py --source "path" --target "path" --run-mode convert_then_merge

# Help
python office_converter.py --help
```

## Key Files

| File | Purpose |
|------|---------|
| `office_converter.py` | Core converter, version `__version__` |
| `office_gui.py` | GUI entry point |
| `task_manager.py` | Task storage |
| `AGENTS.md` | Collaboration rules |