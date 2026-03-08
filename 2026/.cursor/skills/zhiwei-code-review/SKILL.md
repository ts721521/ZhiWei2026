---
name: zhiwei-code-review
description: Review ZhiWei project code for quality, security, and best practices. Use when user asks to review code, perform code review, check for bugs, or analyze code changes.
---

# ZhiWei Code Review

Perform comprehensive code reviews for the ZhiWei project.

## Review Checklist

### 1. Correctness
- [ ] Logic is correct and handles edge cases
- [ ] No off-by-one errors
- [ ] Error handling is appropriate

### 2. Security
- [ ] No hardcoded credentials or secrets
- [ ] File paths are validated
- [ ] No SQL injection vulnerabilities
- [ ] No command injection risks

### 3. Code Quality
- [ ] Follows project naming conventions (snake_case for Python)
- [ ] Functions are appropriately sized and focused
- [ ] No code duplication
- [ ] Comments explain non-obvious logic

### 4. Project Standards
- [ ] Follows AGENTS.md rules
- [ ] Documentation updated if paths/structures changed
- [ ] Version updated if new feature/fix

### 5. Testing
- [ ] Tests cover the changes
- [ ] Test file naming follows convention (`test_*.py`)

## Review Focus Areas

### converter/ modules
- Check error handling in `convert_thread.py`, `run_workflow.py`
- Verify config validation in `config_validation.py`
- Check resource cleanup (file handles, threads)

### gui/ modules
- Verify event handlers are properly bound
- Check for memory leaks in callbacks
- Validate config read/write operations

### Tests
- Ensure tests are independent
- Check for proper mocking
- Verify test coverage for edge cases

## Feedback Format

Use this format for review feedback:

- 🔴 **Critical**: Must fix before merge
- 🟡 **Warning**: Should consider fixing
- 🟢 **Suggestion**: Optional improvement
- 📝 **Note**: Informational

## Files to Review

Key files requiring careful review:
- `office_converter.py` - Main entry point
- `office_gui.py` - GUI entry point
- `task_manager.py` - Task persistence
- `converter/convert_thread.py` - Concurrent processing
- `converter/run_workflow.py` - Main workflow logic
- `gui/mixins/*.py` - GUI components