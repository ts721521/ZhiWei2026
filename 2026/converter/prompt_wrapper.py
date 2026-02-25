# -*- coding: utf-8 -*-
"""Prompt wrapper generation helper."""

import os
from datetime import datetime


_TEMPLATES = {
    "new_solution": "请基于参考资料输出新的技术方案，给出架构、实施步骤、风险与验收标准。",
    "tech_clarification": "请基于参考资料输出历史技术澄清总结，按问题-结论-证据组织。",
    "device_extract": "请基于参考资料提取设备清单，输出结构化表格并标注关键参数来源。",
}


def collect_prompt_ready_candidates(
    generated_fast_md_outputs,
    generated_merge_markdown_outputs,
    generated_markdown_outputs,
):
    candidates = []
    candidates.extend(generated_fast_md_outputs or [])
    candidates.extend(generated_merge_markdown_outputs or [])
    candidates.extend(generated_markdown_outputs or [])
    return candidates


def write_prompt_ready(
    *,
    config,
    target_folder,
    candidate_markdown_paths,
    now_fn=datetime.now,
):
    if not config.get("enable_prompt_wrapper", False):
        return None
    if not target_folder:
        return None

    existing = [p for p in candidate_markdown_paths or [] if p and os.path.exists(p)]
    if not existing:
        return None

    template_type = str(config.get("prompt_template_type", "new_solution") or "new_solution")
    template_text = _TEMPLATES.get(template_type, _TEMPLATES["new_solution"])
    out_path = os.path.join(target_folder, "Prompt_Ready.txt")

    with open(out_path, "w", encoding="utf-8") as out:
        out.write("【系统指令】\n")
        out.write(template_text + "\n")
        out.write("请在此处输入你的具体需求。\n\n")
        out.write(f"generated_at: {now_fn().isoformat(timespec='seconds')}\n")
        out.write(f"template_type: {template_type}\n\n")
        out.write("【参考资料开始】\n\n")
        for idx, path in enumerate(existing, 1):
            out.write(f"### [{idx}] {os.path.basename(path)}\n\n")
            with open(path, "r", encoding="utf-8", errors="ignore") as src:
                while True:
                    chunk = src.read(64 * 1024)
                    if not chunk:
                        break
                    out.write(chunk)
            out.write("\n\n")
        out.write("【参考资料结束】\n")
    return out_path


def write_prompt_ready_for_converter(converter, *, now_fn=datetime.now, log_info_fn=None):
    target_root = converter.config.get("target_folder", "")
    candidates = collect_prompt_ready_candidates(
        getattr(converter, "generated_fast_md_outputs", []) or [],
        getattr(converter, "generated_merge_markdown_outputs", []) or [],
        getattr(converter, "generated_markdown_outputs", []) or [],
    )
    out_path = write_prompt_ready(
        config=converter.config,
        target_folder=target_root,
        candidate_markdown_paths=candidates,
        now_fn=now_fn,
    )
    if out_path:
        converter.prompt_ready_path = out_path
        converter.generated_prompt_outputs = [out_path]
        if callable(log_info_fn):
            log_info_fn(f"Prompt_Ready generated: {out_path}")
    return out_path
