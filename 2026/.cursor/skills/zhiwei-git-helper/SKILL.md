---
name: zhiwei-git-helper
description: Git workflow helper for ZhiWei project including commit, branch, and PR management. Use when user asks to commit changes, create branch, manage git workflow, or create pull request.
---

# ZhiWei Git Helper

Manage Git operations for the ZhiWei project.

## Commit Workflow

### 1. Check Status
```bash
cd d:\GitHub\ZhiWei2026
git status
```

### 2. Review Changes
```bash
git diff
git diff --staged
```

### 3. Stage Files
```bash
# Stage specific files
git add file1.py file2.py

# Stage all changes
git add -A
```

### 4. Commit

Follow commit message format:
```
<type>(<scope>): <description>

[optional body]
```

Types: `feat`, `fix`, `docs`, `test`, `refactor`, `chore`

Example:
```
feat(converter): add concurrent PDF processing

Implement ThreadPoolExecutor for parallel PDF conversion
- Add max_workers config option
- Handle thread-safe file operations
```

## Branch Management

### Create New Branch
```bash
git checkout -b feature/your-feature-name
git checkout -b fix/bug-description
```

### Switch Branch
```bash
git checkout main
git checkout your-branch
```

## Pull Request

### Create PR using gh CLI
```bash
# Push branch
git push -u origin HEAD

# Create PR
gh pr create --title "Feature description" --body "## Summary
- Added feature X
- Fixed issue Y

## Test plan
- [ ] Run unit tests"
```

## Pre-commit Hook

Install for local validation:

```bash
cd d:\GitHub\ZhiWei2026\2026
python scripts/install_git_hook.py
```

This runs doc sync check before each commit.

## Ignore Files

These are already ignored, do NOT commit:
- `config.json` - Runtime config
- `2026/tasks/*.json` - Task data
- `2026/config_profiles/*.json` - User profiles
- `__pycache__/`
- `*.pyc`
- `.env`

## Common Commands

| Task | Command |
|------|---------|
| View history | `git log --oneline -20` |
| Undo staged | `git reset HEAD file` |
| Undo changes | `git checkout -- file` |
| View branch | `git branch -a` |