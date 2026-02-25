# -*- coding: utf-8 -*-
"""Minimal MarkItDown smoke script for packaging validation."""

import argparse
import json
import os
import sys

from converter.markitdown_pack_probe import run_markitdown_probe


def main(argv=None):
    parser = argparse.ArgumentParser(description="MarkItDown smoke probe")
    parser.add_argument("--input", required=True, help="Source Office file path")
    parser.add_argument(
        "--output",
        default=os.path.join("_tmp", "markitdown_probe_output.md"),
        help="Output markdown path",
    )
    args = parser.parse_args(argv)

    result = run_markitdown_probe(args.input, args.output)
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0 if result.get("status") == "ok" else 1


if __name__ == "__main__":
    sys.exit(main())
