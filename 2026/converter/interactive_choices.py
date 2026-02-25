# -*- coding: utf-8 -*-
"""Interactive CLI selection helpers extracted from office_converter.py."""

import os


def ask_for_subfolder(
    config,
    *,
    print_step_title_fn,
    print_fn=print,
    input_fn=input,
    abspath_fn=os.path.abspath,
    join_fn=os.path.join,
):
    print_step_title_fn("Step 2/4: Optional Output Subfolder")
    print_fn("You can create a subfolder for this run under target directory.")
    print_fn("-" * 60)
    sub = input_fn("Subfolder name (Enter to skip): ").strip()
    if sub:
        for char in '<>:"/\\|?*':
            sub = sub.replace(char, "")
        config["target_folder"] = abspath_fn(join_fn(config["target_folder"], sub))
        print_fn(f"--> Run target folder: {config['target_folder']}")
    else:
        print_fn("--> Keep original target folder from config.")


def select_run_mode(
    *,
    print_step_title_fn,
    get_readable_run_mode_fn,
    mode_convert_only,
    mode_merge_only,
    mode_convert_then_merge,
    mode_collect_only,
    mode_mshelp_only,
    print_fn=print,
    input_fn=input,
):
    print_step_title_fn("Step 3/4: Select Run Mode")
    print_fn("  [1] Convert only")
    print_fn("  [2] Merge & convert")
    print_fn("  [3] Convert then merge (recommended)")
    print_fn("  [4] Collect / deduplicate only")
    print_fn("  [5] MSHelp API docs (CAB->MD, merged package)")
    print_fn("-" * 60)
    choice = input_fn("Choose (1/2/3/4/5, default 3): ").strip()
    if choice == "1":
        run_mode = mode_convert_only
    elif choice == "2":
        run_mode = mode_merge_only
    elif choice == "4":
        run_mode = mode_collect_only
    elif choice == "5":
        run_mode = mode_mshelp_only
    else:
        run_mode = mode_convert_then_merge
    print_fn(f"--> Run mode: {get_readable_run_mode_fn(run_mode)} ({run_mode})")
    return run_mode


def select_collect_mode(
    *,
    print_step_title_fn,
    get_readable_collect_mode_fn,
    collect_mode_copy_and_index,
    collect_mode_index_only,
    print_fn=print,
    input_fn=input,
):
    print_step_title_fn("Select Collect Sub-Mode")
    print_fn("  [1] Deduplicate + copy + Excel index")
    print_fn("  [2] Generate Excel index only (no copy)")
    print_fn("-" * 60)
    choice = input_fn("Choose (1/2, default 1): ").strip()
    if choice == "2":
        collect_mode = collect_mode_index_only
    else:
        collect_mode = collect_mode_copy_and_index
    print_fn(f"--> Collect mode: {get_readable_collect_mode_fn(collect_mode)} ({collect_mode})")
    return collect_mode


def select_merge_mode(
    config,
    *,
    print_step_title_fn,
    get_readable_merge_mode_fn,
    merge_mode_category,
    merge_mode_all_in_one,
    print_fn=print,
    input_fn=input,
):
    if not config.get("enable_merge", True):
        return merge_mode_category

    cfg_mode = config.get("merge_mode", merge_mode_category)
    if cfg_mode in (merge_mode_all_in_one, merge_mode_category):
        print_fn(
            f"--> Merge mode from config: {get_readable_merge_mode_fn(cfg_mode)} ({cfg_mode})"
        )
        return cfg_mode

    print_step_title_fn("Select Merge Mode")
    print_fn("  [1] Category split (Price/Word/Excel/PPT/PDF)")
    print_fn("  [2] All in one PDF")
    print_fn("-" * 60)
    choice = input_fn("Choose (1/2, default 1): ").strip()
    if choice == "2":
        merge_mode = merge_mode_all_in_one
    else:
        merge_mode = merge_mode_category
    print_fn(f"--> Merge mode: {get_readable_merge_mode_fn(merge_mode)} ({merge_mode})")
    return merge_mode


def select_content_strategy(
    price_keywords,
    *,
    print_step_title_fn,
    get_readable_content_strategy_fn,
    strategy_standard,
    strategy_smart_tag,
    strategy_price_only,
    print_fn=print,
    input_fn=input,
):
    print_step_title_fn("Step 4/4: Select Content Strategy")
    print_fn("  [1] Standard classification")
    print_fn("  [2] Smart tag (price keyword hit)")
    print_fn("  [3] Price only")
    print_fn("-" * 60)
    print_fn(f"Current keywords: {price_keywords}")
    choice = input_fn("Choose (1/2/3, default 1): ").strip()
    if choice == "2":
        content_strategy = strategy_smart_tag
    elif choice == "3":
        content_strategy = strategy_price_only
    else:
        content_strategy = strategy_standard
    print_fn(
        f"--> Strategy: {get_readable_content_strategy_fn(content_strategy)} ({content_strategy})\n"
    )
    return content_strategy


def select_engine_mode(
    config,
    *,
    print_step_title_fn,
    get_readable_engine_type_fn,
    engine_ask,
    engine_wps,
    engine_ms,
    print_fn=print,
    input_fn=input,
):
    default = config.get("default_engine", engine_ask)
    if default == engine_wps:
        print_fn("--> [auto] engine: WPS Office")
        return engine_wps
    if default == engine_ms:
        print_fn("--> [auto] engine: Microsoft Office")
        return engine_ms

    print_step_title_fn("Select Office Engine")
    print_fn("  [1] WPS Office")
    print_fn("  [2] Microsoft Office")
    print_fn("-" * 60)
    while True:
        choice = input_fn("Choose (1/2, default 1): ").strip()
        if choice in ("", "1"):
            engine_type = engine_wps
            break
        if choice == "2":
            engine_type = engine_ms
            break
        print_fn("Invalid input. Please enter 1 or 2.")
    print_fn(f"--> Selected: {get_readable_engine_type_fn(engine_type)} ({engine_type})\n")
    return engine_type
