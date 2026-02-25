# -*- coding: utf-8 -*-
"""Runtime wrapper for AI path helpers extracted from office_converter.py."""


def build_ai_output_path_from_source_for_converter(
    converter,
    source_path,
    sub_dir,
    ext,
    *,
    build_ai_output_path_from_source_fn,
):
    return build_ai_output_path_from_source_fn(
        source_path,
        sub_dir,
        ext,
        converter.config.get("target_folder", ""),
        source_root_resolver=converter._get_source_root_for_path,
    )
