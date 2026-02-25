#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Check whether code changes are synchronized with required docs."""

from __future__ import annotations

import argparse
import shutil
import subprocess
import sys
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parent.parent
REQUIRED_DOCS = {
    "docs/plans/2026-02-24-office-converter-split-handover.md",
    "docs/test-reports/TEST_REPORT_SUMMARY.md",
}
ARCH_CHANGE_HINTS = (
    "office_gui.py",
    "office_converter.py",
    "gui/",
    "converter/",
)
CODE_CHANGE_HINTS = (
    "office_gui.py",
    "office_converter.py",
    "task_manager.py",
    "converter/",
    "gui/",
    "tests/",
)


def _norm(path: str) -> str:
    return path.replace("\\", "/").lstrip("./")


def _is_code_change(path: str) -> bool:
    p = _norm(path)
    return any(p == h or p.startswith(h) for h in CODE_CHANGE_HINTS)


def _is_arch_change(path: str) -> bool:
    p = _norm(path)
    return any(p == h or p.startswith(h) for h in ARCH_CHANGE_HINTS)


def _git_changed_files(staged: bool) -> list[str]:
    if shutil.which("git") is None:
        raise RuntimeError("git not found; please pass --changed explicitly")
    cmd = ["git", "diff", "--name-only"]
    if staged:
        cmd.append("--cached")
    res = subprocess.run(
        cmd,
        cwd=REPO_ROOT,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        check=False,
    )
    if res.returncode != 0:
        raise RuntimeError(res.stderr.strip() or "git diff failed")
    return [line.strip() for line in res.stdout.splitlines() if line.strip()]


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--changed", nargs="*", default=None, help="changed file paths")
    parser.add_argument("--staged", action="store_true", help="read changed files from git staged diff")
    args = parser.parse_args()

    if args.changed is not None and args.staged:
        print("ERROR: use either --changed or --staged", file=sys.stderr)
        return 2

    try:
        if args.changed is not None:
            changed = [_norm(p) for p in args.changed]
        elif args.staged:
            changed = [_norm(p) for p in _git_changed_files(staged=True)]
        else:
            changed = [_norm(p) for p in _git_changed_files(staged=False)]
    except RuntimeError as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 2

    if not changed:
        print("OK: no changed files detected")
        return 0

    changed_set = set(changed)
    has_code_change = any(_is_code_change(p) for p in changed)
    has_arch_change = any(_is_arch_change(p) for p in changed)

    missing: list[str] = []
    if has_code_change:
        for doc in sorted(REQUIRED_DOCS):
            if doc not in changed_set:
                missing.append(doc)

    if has_arch_change and "AGENTS.md" not in changed_set:
        missing.append("AGENTS.md")

    if missing:
        print("ERROR: documentation sync check failed")
        for m in missing:
            print(f"  - missing update: {m}")
        print("Hint: include required docs in this change set.")
        return 1

    print("OK: documentation sync check passed")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
