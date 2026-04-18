# -*- coding: utf-8 -*-
"""MSHelp 独立运行入口。

GUI 已剥离 MSHelp 模式，但底层 converter/mshelp_*.py 仍在维护，
可通过本脚本直接运行 MSHelpViewer 扫描 + CAB 转换 + 索引/合并流程：

    python tools/mshelp_run.py --config config.json
    python tools/mshelp_run.py --config config_profiles/myprofile.json --non-interactive

脚本会强制把 run_mode 设为 mshelp_only，其他参数（源/目标/输出开关等）
仍读取传入的 config.json，与 GUI 保存的配置保持一致。
"""

from __future__ import annotations

import argparse
import os
import sys


def _ensure_project_root_on_path():
    here = os.path.dirname(os.path.abspath(__file__))
    project_root = os.path.dirname(here)
    if project_root not in sys.path:
        sys.path.insert(0, project_root)


def main(argv=None):
    _ensure_project_root_on_path()

    from office_converter import (  # noqa: E402  延迟导入，确保 sys.path 已就位
        MODE_MSHELP_ONLY,
        OfficeConverter,
    )

    parser = argparse.ArgumentParser(
        description="Run MSHelp-only workflow (UI 已剥离的独立 CLI 入口)",
    )
    parser.add_argument(
        "--config",
        default="config.json",
        help="配置文件路径，默认读取项目根目录的 config.json",
    )
    parser.add_argument(
        "--non-interactive",
        action="store_true",
        help="跳过所有交互确认（适合脚本/计划任务调用）",
    )
    args = parser.parse_args(argv)

    config_path = os.path.abspath(args.config)
    if not os.path.exists(config_path):
        print(f"[ERROR] config not found: {config_path}", file=sys.stderr)
        return 2

    converter = OfficeConverter(
        config_path=config_path,
        interactive=not args.non_interactive,
    )
    converter.run_mode = MODE_MSHELP_ONLY
    return 0 if converter.run() else 1


if __name__ == "__main__":
    raise SystemExit(main())
