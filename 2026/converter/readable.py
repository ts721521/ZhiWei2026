# -*- coding: utf-8 -*-
"""Human-readable label helpers extracted from office_converter.py."""

from converter.constants import (
    COLLECT_MODE_COPY_AND_INDEX,
    COLLECT_MODE_INDEX_ONLY,
    ENGINE_MS,
    ENGINE_WPS,
    MERGE_MODE_ALL_IN_ONE,
    MERGE_MODE_CATEGORY,
    MODE_COLLECT_ONLY,
    MODE_CONVERT_ONLY,
    MODE_CONVERT_THEN_MERGE,
    MODE_MERGE_ONLY,
    MODE_MSHELP_ONLY,
    STRATEGY_PRICE_ONLY,
    STRATEGY_SMART_TAG,
    STRATEGY_STANDARD,
)


def readable_run_mode(run_mode):
    m = {
        MODE_CONVERT_ONLY: "convert_only",
        MODE_MERGE_ONLY: "merge_only",
        MODE_CONVERT_THEN_MERGE: "convert_then_merge",
        MODE_COLLECT_ONLY: "collect_only",
        MODE_MSHELP_ONLY: "mshelp_only",
    }
    return m.get(run_mode, run_mode)


def readable_collect_mode(collect_mode):
    m = {
        COLLECT_MODE_COPY_AND_INDEX: "copy_and_index",
        COLLECT_MODE_INDEX_ONLY: "index_only",
    }
    return m.get(collect_mode, collect_mode)


def readable_content_strategy(content_strategy):
    m = {
        STRATEGY_STANDARD: "standard",
        STRATEGY_SMART_TAG: "smart_tag",
        STRATEGY_PRICE_ONLY: "price_only",
    }
    return m.get(content_strategy, content_strategy)


def readable_engine_type(engine_type):
    m = {
        ENGINE_WPS: "WPS Office",
        ENGINE_MS: "Microsoft Office",
        None: "not_used",
    }
    return m.get(engine_type, str(engine_type))


def readable_merge_mode(merge_mode):
    m = {
        MERGE_MODE_CATEGORY: "category_split",
        MERGE_MODE_ALL_IN_ONE: "all_in_one",
    }
    return m.get(merge_mode, merge_mode)
