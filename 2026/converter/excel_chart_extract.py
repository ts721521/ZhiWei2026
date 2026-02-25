# -*- coding: utf-8 -*-
"""Excel chart extraction helpers extracted from office_converter.py."""


def extract_sheet_charts(
    ws_formula,
    *,
    extract_chart_title_text_fn,
    stringify_chart_anchor_fn,
    series_ref_limit=50,
):
    charts = []
    if ws_formula is None:
        return charts
    for idx, chart in enumerate(getattr(ws_formula, "_charts", []) or [], 1):
        refs = []
        for series in getattr(chart, "series", []) or []:
            for attr in ("val", "xVal", "yVal", "cat", "bubbleSize"):
                ref_obj = getattr(series, attr, None)
                if ref_obj is None:
                    continue
                num_ref = getattr(ref_obj, "numRef", None)
                str_ref = getattr(ref_obj, "strRef", None)
                for ref in (num_ref, str_ref):
                    formula_ref = getattr(ref, "f", None) if ref is not None else None
                    if formula_ref:
                        refs.append(str(formula_ref))
        dedup_refs = sorted(set(refs))
        charts.append(
            {
                "index_1based": idx,
                "chart_type": chart.__class__.__name__,
                "title": extract_chart_title_text_fn(chart),
                "anchor": stringify_chart_anchor_fn(getattr(chart, "anchor", None)),
                "series_ref_count": len(dedup_refs),
                "series_refs_truncated": len(dedup_refs) > series_ref_limit,
                "series_refs": dedup_refs[:series_ref_limit],
            }
        )
    return charts


def extract_sheet_pivot_tables(ws_formula):
    pivots = []
    if ws_formula is None:
        return pivots
    for idx, pivot in enumerate(getattr(ws_formula, "_pivots", []) or [], 1):
        loc_ref = ""
        try:
            location = getattr(pivot, "location", None)
            loc_ref = str(getattr(location, "ref", "") or "")
        except (AttributeError, RuntimeError, TypeError, ValueError):
            loc_ref = ""
        pivots.append(
            {
                "index_1based": idx,
                "name": str(getattr(pivot, "name", "") or ""),
                "cache_id": getattr(pivot, "cacheId", None),
                "location_ref": loc_ref,
            }
        )
    return pivots
