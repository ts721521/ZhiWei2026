---
name: zhiwei-doc-sync
description: Verify and enforce documentation synchronization for code changes in ZhiWei project. Use when user asks to check doc sync, verify documentation, run check_doc_sync.py, or ensure documentation is up to date.
---

# ZhiWei Documentation Sync

Ensure code changes are properly synchronized with documentation.

## Automatic Check

Run the doc sync validation script:

```bash
cd d:\GitHub\ZhiWei2026\2026
python scripts/check_doc_sync.py --staged
# or for specific files:
python scripts/check_doc_sync.py --changed file1.py file2.py
```

## Mandatory Documentation Updates

When changing these code paths, MUST update:

| Code Path | Must Update |
|-----------|-------------|
| `office_converter.py` | `docs/test-reports/TEST_REPORT_SUMMARY.md` |
| `office_gui.py` | `docs/plans/2026-02-24-office-converter-split-handover.md` |
| `task_manager.py` | `docs/plans/2026-02-24-office-converter-split-handover.md` |
| `converter/**` | Both handover and test summary |
| `gui/**` | Both handover and test summary |
| `tests/**` | `docs/test-reports/TEST_REPORT_SUMMARY.md` |

## Path/Directory Changes

If changing paths or directory structure, MUST also update `AGENTS.md`.

## Documentation Files

### Primary Documentation
- `AGENTS.md` - Collaboration rules
- `docs/AI_交接文档_下一阶段开发.md` - Project overview and handover
- `docs/plans/2026-02-24-office-converter-split-handover.md` - Current code status
- `docs/test-reports/TEST_REPORT_SUMMARY.md` - Test results

### Supporting Documentation
- `docs/plans/` - Design documents
- `docs/test-reports/` - Test reports
- `CHANGELOG.md` - Version history

## Checklist Before Commit

- [ ] Code changes complete
- [ ] Tests run and pass
- [ ] Doc sync check passes: `python scripts/check_doc_sync.py --staged`
- [ ] Updated handover or test summary if required
- [ ] Updated AGENTS.md if paths changed
- [ ] Updated CHANGELOG.md if new version

## Common Issues

| Issue | Solution |
|-------|----------|
| Check fails on modified file | Add/update corresponding doc file |
| Check fails on new file | Check if new module needs documentation |
| AGENTS.md update needed | Update section 3 "强制文档同步规则" |