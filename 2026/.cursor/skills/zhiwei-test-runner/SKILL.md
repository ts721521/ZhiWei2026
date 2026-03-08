---
name: zhiwei-test-runner
description: Run ZhiWei project tests including unit tests and NotebookLM E2E tests. Use when user asks to run tests, execute test suite, verify changes, or needs test reports.
---

# ZhiWei Test Runner

Run and manage tests for the ZhiWei (知喂) Office converter project.

## Test Commands

### Unit Tests

```bash
cd d:\GitHub\ZhiWei2026\2026
python -m unittest discover -s tests -p "test_*.py" -v
```

**Location**: `tests/` directory

### NotebookLM E2E Test

```bash
cd d:\GitHub\ZhiWei2026\2026
python scripts/run_notebooklm_e2e.py
```

**Repair prompt if failed**: `docs/test-reports/notebooklm_e2e_repair_prompt.txt`

## Test Workflow

1. Run unit tests first
2. If unit tests pass, run E2E tests for relevant features
3. Record results in `docs/test-reports/TEST_REPORT_SUMMARY.md`

## Documentation Updates

After running tests, update the test report with:
- Command executed
- Number of tests (Ran N tests)
- Result (OK / FAILED)
- Summary of changes this round

## Common Issues

- **Import errors**: Ensure running from `2026/` directory
- **Config errors**: Check `config.json` exists or use `configs/templates/config.example.json`
- **E2E failures**: Review `docs/test-reports/notebooklm_e2e_repair_prompt.txt` for guidance