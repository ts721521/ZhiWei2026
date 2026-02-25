# -*- coding: utf-8 -*-
"""Runtime summary rendering helper extracted from office_converter.py."""


def print_runtime_summary(
    *,
    config,
    run_mode,
    merge_mode,
    content_strategy,
    mode_merge_only,
    get_output_pref_fn,
    get_merge_convert_submode_fn,
    should_reuse_office_app_fn,
    get_office_restart_every_fn,
    print_fn=print,
):
    print_fn("\n" + "=" * 60)
    print_fn(" Runtime Summary")
    print_fn("=" * 60)
    print_fn(f"  source_folder : {config.get('source_folder', '')}")
    print_fn(f"  target_folder : {config.get('target_folder', '')}")
    print_fn(f"  run_mode      : {run_mode}")
    print_fn(f"  merge_mode    : {merge_mode}")
    pref = get_output_pref_fn()
    print_fn(
        f"  output        : pdf={pref['pdf']} md={pref['md']} merged={pref['merged']} independent={pref['independent']}"
    )
    if run_mode == mode_merge_only:
        print_fn(f"  merge_submode : {get_merge_convert_submode_fn()}")
    print_fn(f"  strategy      : {content_strategy}")
    print_fn(f"  reuse_office  : {should_reuse_office_app_fn()}")
    print_fn(f"  restart_every : {get_office_restart_every_fn()}")
    print_fn("=" * 60)
