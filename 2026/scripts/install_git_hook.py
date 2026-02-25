#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Install repository pre-commit hook for doc sync checks."""

from __future__ import annotations

import shutil
from pathlib import Path


ROOT = Path(__file__).resolve().parent.parent
SRC = ROOT / ".githooks" / "pre-commit"


def _find_git_dir(start: Path) -> Path | None:
    cur = start
    while True:
        cand = cur / ".git"
        if cand.exists():
            return cand
        if cur.parent == cur:
            return None
        cur = cur.parent


def main() -> int:
    if not SRC.exists():
        print(f"ERROR: missing hook template: {SRC}")
        return 1
    git_dir = _find_git_dir(ROOT)
    if git_dir is None:
        print("ERROR: .git directory not found; run inside a git repository")
        return 1
    dst = git_dir / "hooks" / "pre-commit"
    dst.parent.mkdir(parents=True, exist_ok=True)
    shutil.copyfile(SRC, dst)
    try:
        current = dst.stat().st_mode
        dst.chmod(current | 0o111)
    except Exception:
        # On some Windows environments chmod execute bit may be ignored.
        pass
    print(f"Installed pre-commit hook: {dst}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
