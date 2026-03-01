#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
诊断「转换范围」：仅执行扫描与三道过滤（source_priority / incremental / global_md5_dedup），
不进行实际转换。用于排查「很多文件未转换」时，是扫描/过滤阶段就排除了，还是转换阶段失败。

用法（在 2026 目录下执行）：
  python scripts/diagnose_convert_scope.py [config_path]

若不传 config_path，默认使用 2026/config.json。
"""
from __future__ import print_function

import os
import sys

_SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
_REPO_ROOT = os.path.dirname(_SCRIPT_DIR)
_DEFAULT_CONFIG = os.path.join(_REPO_ROOT, "config.json")


def _main():
    config_path = sys.argv[1] if len(sys.argv) > 1 else _DEFAULT_CONFIG
    if not os.path.isfile(config_path):
        print("未找到配置文件:", config_path)
        print("用法: python scripts/diagnose_convert_scope.py [config_path]")
        return 1

    # 保证从 2026 目录可导入 office_converter 与 converter
    if _REPO_ROOT not in sys.path:
        sys.path.insert(0, _REPO_ROOT)
    os.chdir(_REPO_ROOT)

    print("=== 转换范围诊断（仅扫描 + 过滤，不转换）===\n")
    print("配置:", config_path)
    print()

    try:
        from office_converter import OfficeConverter
    except ImportError as e:
        print("无法导入 office_converter，请在本仓库 2026 目录下执行: python scripts/diagnose_convert_scope.py")
        print("错误:", e)
        return 1

    converter = OfficeConverter(config_path=config_path, interactive=False)

    # 仅当运行模式为转换类时才做扫描与过滤
    from converter.constants import MODE_CONVERT_ONLY, MODE_CONVERT_THEN_MERGE

    if converter.run_mode not in (MODE_CONVERT_ONLY, MODE_CONVERT_THEN_MERGE):
        print("当前 run_mode 非转换模式，跳过扫描。run_mode:", converter.run_mode)
        return 0

    # 1) 扫描候选
    files = converter._scan_convert_candidates()
    if files is None:
        files = []
    scan_count = len(files)
    print("1. 扫描候选数 (scan candidate file count):", scan_count)
    if scan_count == 0:
        print("   可能原因: 源目录不存在/无权限、allowed_extensions 为空或过窄、excluded_folders 排除过多。")
        print("   请检查 config 中 source_folder/source_folders、allowed_extensions、excluded_folders。")
        return 0

    # 2) source_priority
    files, source_priority_skips = converter._apply_source_priority_filter(files)
    print("2. 同名优先 Office 后 (after source_priority): 待处理=%s, 跳过=%s" % (len(files), len(source_priority_skips)))

    # 3) incremental
    files, incremental_ctx = converter._apply_incremental_filter(files)
    inc = incremental_ctx or {}
    print(
        "3. 增量过滤后 (after incremental): 待处理=%s | added=%s, modified=%s, unchanged=%s, renamed=%s, deleted=%s"
        % (
            len(files),
            inc.get("added_count", 0),
            inc.get("modified_count", 0),
            inc.get("unchanged_count", 0),
            inc.get("renamed_count", 0),
            inc.get("deleted_count", 0),
        )
    )

    # 4) global_md5_dedup
    files, dedup_skips = converter._apply_global_md5_dedup(files)
    print("4. 全局 MD5 去重后 (after global_md5_dedup): 待处理=%s, 跳过=%s" % (len(files), len(dedup_skips)))

    final_total = len(files)
    print()
    print("最终待转换数量 (total to process):", final_total)
    if final_total == 0 and scan_count > 0:
        print("   所有候选均在过滤阶段被跳过，未进入转换。请根据上面各步跳过数检查配置（增量/同名优先/去重）。")
    elif final_total > 0:
        print("   上述数量将参与实际转换；转换阶段还可能因空文件/策略/内容跳过或失败，请结合运行日志与失败报告查看。")
    return 0


if __name__ == "__main__":
    sys.exit(_main())
