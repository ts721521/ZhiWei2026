# -*- coding: utf-8 -*-
"""Performance summary helpers extracted from office_converter.py."""


def build_perf_summary(perf_metrics, stats):
    m = perf_metrics or {}
    stats = stats or {}
    lines = [
        "",
        "=== 性能统计 ===",
        f"扫描耗时: {m.get('scan_seconds', 0.0):.2f}s",
        f"转换主流程耗时: {m.get('batch_seconds', 0.0):.2f}s",
        f"  - Office/PDF核心耗时: {m.get('convert_core_seconds', 0.0):.2f}s",
        f"  - 等待PDF落盘耗时: {m.get('pdf_wait_seconds', 0.0):.2f}s",
        f"  - Markdown导出耗时: {m.get('markdown_seconds', 0.0):.2f}s",
        f"  - MSHelp合并耗时: {m.get('mshelp_merge_seconds', 0.0):.2f}s",
        f"合并耗时: {m.get('merge_seconds', 0.0):.2f}s",
        f"后处理耗时: {m.get('postprocess_seconds', 0.0):.2f}s",
        f"总耗时: {m.get('total_seconds', 0.0):.2f}s",
    ]
    success_count = stats.get("success", 0) or 0
    if success_count > 0:
        avg = m.get("total_seconds", 0.0) / success_count
        lines.append(f"平均每成功文件耗时: {avg:.2f}s")
        if m.get("total_seconds", 0.0) > 0:
            lines.append(
                f"吞吐率(成功文件/分钟): {success_count / m.get('total_seconds', 1.0) * 60:.2f}"
            )
    return "\n".join(lines) + "\n"
