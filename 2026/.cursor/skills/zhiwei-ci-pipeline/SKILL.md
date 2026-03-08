---
name: zhiwei-ci-pipeline
description: ZhiWei CI/CD pipeline management including quality gates, doc sync checks, and test automation. Use when user asks to check CI status, fix CI failures, run quality gates, or configure automated checks.
---

# ZhiWei CI/CD Pipeline

Manage CI/CD workflows for the ZhiWei project.

## CI Configuration

**Workflow file**: `.github/workflows/quality-gate.yml`

**Triggers**:
- Pull requests
- Push to any branch

## Quality Gate Steps

### 1. Document Sync Check

Verifies code changes are properly documented:

```bash
cd d:\GitHub\ZhiWei2026\2026
python scripts/check_doc_sync.py --staged
# or for specific files:
python scripts/check_doc_sync.py --changed file1.py file2.py
```

### 2. Unit Tests

```bash
cd d:\GitHub\ZhiWei2026\2026
python -m unittest discover -s tests -p "test_*.py" -v
```

## Common CI Failures

| Issue | Fix |
|-------|-----|
| Doc sync failed | Update `docs/plans/2026-02-24-office-converter-split-handover.md` or `docs/test-reports/TEST_REPORT_SUMMARY.md` |
| Tests failed | Fix failing tests, ensure all tests pass |
| Import errors | Check Python path, ensure running from `2026/` directory |

## CI Monitoring

- Check CI status in GitHub Actions tab
- Use `ci-watcher` subagent to monitor CI runs
- Review failed logs in GitHub Actions workflow runs

## Pre-commit Hook

Install local pre-commit hook for local verification:

```bash
cd d:\GitHub\ZhiWei2026\2026
python scripts/install_git_hook.py
```

This runs `check_doc_sync.py` before each commit.

## Related Scripts

| Script | Purpose |
|--------|---------|
| `scripts/check_doc_sync.py` | Verify doc sync compliance |
| `scripts/install_git_hook.py` | Install pre-commit hook |