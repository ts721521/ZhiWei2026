# -*- coding: utf-8 -*-
"""Sandbox disk-space guard extracted from office_converter.py."""


def check_sandbox_free_space_or_raise(
    config,
    *,
    exists_fn,
    splitdrive_fn,
    getcwd_fn,
    disk_usage_fn,
    log_info_fn,
    log_warning_fn,
    print_fn=print,
):
    if not config.get("enable_sandbox", True):
        return

    threshold_gb = config.get("sandbox_min_free_gb", 10)
    try:
        threshold_gb = float(threshold_gb)
    except (TypeError, ValueError):
        threshold_gb = 0
    if threshold_gb <= 0:
        return

    sandbox_root = config.get("temp_sandbox_root") or config.get("target_folder", "")
    if not sandbox_root:
        return

    probe_path = sandbox_root
    if not exists_fn(probe_path):
        drive, _ = splitdrive_fn(probe_path)
        if drive:
            probe_path = drive + "\\"
        else:
            probe_path = getcwd_fn()

    try:
        usage = disk_usage_fn(probe_path)
    except (OSError, RuntimeError, TypeError, ValueError) as e:
        log_warning_fn(f"disk usage check failed for {probe_path}: {e}")
        return

    free_gb = usage.free / (1024 * 1024 * 1024)
    policy = (config.get("sandbox_low_space_policy") or "block").lower()

    msg = (
        f"Sandbox free space check: path={probe_path}, "
        f"free={free_gb:.2f} GB, threshold={threshold_gb:.2f} GB, policy={policy}"
    )
    log_info_fn(msg)

    if free_gb >= threshold_gb:
        return

    warn_text = (
        f"[WARN] Sandbox free space is below threshold: "
        f"{free_gb:.2f} GB < {threshold_gb:.2f} GB (policy={policy})"
    )
    print_fn("\n" + warn_text)
    log_warning_fn(warn_text)

    if policy == "block":
        raise RuntimeError(
            "Sandbox free space below configured minimum; run blocked by policy."
        )
