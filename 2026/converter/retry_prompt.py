# -*- coding: utf-8 -*-
"""Retry prompt helper extracted from office_converter.py."""


def ask_retry_failed_files(
    failed_count,
    *,
    error_records,
    timeout,
    has_msvcrt,
    msvcrt_module,
    time_module,
    input_fn=input,
    print_fn=print,
):
    print_fn("\n" + "=" * 60)
    print_fn(f"[WARN] {failed_count} files failed (including timeout).")
    if error_records:
        print_fn("Failed examples (up to 10):")
        for p in error_records[:10]:
            print_fn("  -", p)
    print_fn("-" * 60)
    print_fn("Retry failed files?")
    print_fn("  Enter Y + Enter -> retry")
    print_fn("  Enter N + Enter -> no retry")
    print_fn(f"  If no input in {timeout}s, default is no retry.")
    print_fn("=" * 60)

    if not has_msvcrt:
        ans = input_fn("Input [Y/N] and press Enter: ").strip().lower()
        return ans == "y"

    buf = ""
    start = time_module.time()
    last_shown = None

    while True:
        elapsed = time_module.time() - start
        remain = int(timeout - elapsed)
        if remain < 0:
            print_fn("\n[INFO] timeout reached, default no retry.")
            return False

        if last_shown != remain:
            print_fn(f"\rInput [Y/N] within {remain:2d}s: {buf}", end="", flush=True)
            last_shown = remain

        if msvcrt_module.kbhit():
            ch = msvcrt_module.getwch()
            if ch in ("\r", "\n"):
                ans = buf.strip().lower()
                print_fn()
                if ans == "y":
                    print_fn("[SELECT] retry failed files.\n")
                    return True
                print_fn("[SELECT] do not retry failed files.\n")
                return False
            if ch == "\b":
                buf = buf[:-1]
            else:
                buf += ch
        time_module.sleep(0.1)
