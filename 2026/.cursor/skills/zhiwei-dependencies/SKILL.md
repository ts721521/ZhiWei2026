---
name: zhiwei-dependencies
description: Manage Python dependencies for ZhiWei project. Use when user asks about dependencies, requirements, install packages, or fix import errors.
---

# ZhiWei Dependencies Management

Manage and troubleshoot Python dependencies.

## Core Dependencies

**File:** `requirements.txt`

Core packages:
- `ttkbootstrap` - Modern Tkinter theme
- `pypdf` - PDF processing
- `openpyxl` - Excel handling
- `python-docx` - Word documents
- `pillow` - Image processing

## Optional Dependencies

| Package | Feature | Install Command |
|---------|---------|-----------------|
| `google-api-python-client` | Google Drive upload | `pip install google-api-python-client google-auth-httplib2 google-auth-oauthlib` |
| `chromadb` | Vector database export | `pip install chromadb` |
| `beautifulsoup4` | HTML parsing | `pip install beautifulsoup4` |
| `markitdown` | Markdown conversion | `pip install markitdown` |

## Dependency Commands

```bash
# Install all requirements
cd d:\GitHub\ZhiWei2026\2026
pip install -r requirements.txt

# Install Google Drive dependencies (GUI button available)
pip install google-api-python-client google-auth-httplib2 google-auth-oauthlib

# Check installed packages
pip list

# Update a package
pip install --upgrade package_name
```

## Import Error Handling

The code uses try/except for optional imports:

```python
try:
    import chromadb
    HAS_CHROMADB = True
except ImportError:
    HAS_CHROMADB = False
```

**When adding new optional dependencies:**
1. Use try/except ImportError pattern
2. Set HAS_* flag
3. Provide fallback or skip functionality
4. Document in requirements.txt as optional

## Testing Dependencies

Before running tests, ensure:
```bash
pip install -r requirements.txt
pip install pytest  # if using pytest
```

## Common Issues

| Error | Solution |
|-------|----------|
| `ModuleNotFoundError: No module named 'xxx'` | `pip install xxx` |
| `ImportError: cannot import name` | Check version compatibility |
| `Permission denied` | Use `--user` flag or venv |