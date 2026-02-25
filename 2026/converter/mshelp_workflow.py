# -*- coding: utf-8 -*-
"""MSHelp-only run workflow extracted from office_converter.py."""


def run_mshelp_only(
    *,
    stats,
    scan_mshelp_cab_candidates_fn,
    run_batch_fn,
    write_mshelp_index_files_fn,
    merge_mshelp_markdowns_fn,
    add_perf_seconds_fn,
    perf_counter_fn,
    log_info_fn,
    print_fn=print,
):
    log_info_fn("scanning MSHelpViewer folders...")
    scan_start = perf_counter_fn()
    mshelp_dirs, cab_files = scan_mshelp_cab_candidates_fn()
    add_perf_seconds_fn("scan_seconds", perf_counter_fn() - scan_start)

    log_info_fn("MSHelpViewer folder count: %s", len(mshelp_dirs))
    log_info_fn("MSHelp CAB candidate count: %s", len(cab_files))

    stats["total"] = len(cab_files)
    results = []
    if cab_files:
        log_info_fn("start processing MSHelp CAB files: %s", len(cab_files))
        batch_start = perf_counter_fn()
        results.extend(run_batch_fn(cab_files))
        add_perf_seconds_fn("batch_seconds", perf_counter_fn() - batch_start)
    else:
        print_fn("\n[INFO] No MSHelp CAB files found under source folder.")

    mshelp_index_outputs = write_mshelp_index_files_fn()
    mshelp_merged_outputs = merge_mshelp_markdowns_fn()
    return results, mshelp_dirs, mshelp_index_outputs, mshelp_merged_outputs
