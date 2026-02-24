# -*- coding: utf-8 -*-
"""Excel JSON export helper functions extracted from office_converter.py."""

import re
from datetime import date as dt_date
from datetime import datetime
from datetime import time as dt_time


def json_safe_value(value):
    if value is None:
        return None
    if isinstance(value, (str, int, float, bool)):
        return value
    if isinstance(value, (datetime, dt_date, dt_time)):
        try:
            return value.isoformat()
        except Exception:
            return str(value)
    return str(value)


def is_empty_json_cell(value):
    if value is None:
        return True
    return isinstance(value, str) and value.strip() == ""


def is_effectively_empty_row(row_values):
    if not row_values:
        return True
    return all(is_empty_json_cell(v) for v in row_values)


def looks_like_header_row(row_values):
    if not row_values:
        return False
    non_empty = [
        v for v in row_values if not (v is None or (isinstance(v, str) and v.strip() == ""))
    ]
    if not non_empty:
        return False
    string_like = [v for v in non_empty if isinstance(v, str)]
    threshold = 1 if len(non_empty) == 1 else max(2, (len(non_empty) + 1) // 2)
    return len(string_like) >= threshold


def normalize_header_row(header_raw, width):
    names = []
    seen = {}
    for idx in range(width):
        base = ""
        if idx < len(header_raw):
            hv = header_raw[idx]
            if hv is not None:
                base = str(hv).strip()
        if not base:
            base = f"col_{idx + 1}"

        if base not in seen:
            seen[base] = 1
            names.append(base)
        else:
            seen[base] += 1
            names.append(f"{base}_{seen[base]}")
    return names


def detect_json_value_type(value):
    if value is None:
        return "null"
    if isinstance(value, bool):
        return "boolean"
    if isinstance(value, int) and not isinstance(value, bool):
        return "integer"
    if isinstance(value, float):
        return "number"
    if isinstance(value, datetime):
        return "datetime"
    if isinstance(value, dt_date):
        return "date"
    if isinstance(value, dt_time):
        return "time"
    if isinstance(value, str):
        return "string"
    return "other"


def build_column_profiles(header, raw_rows, sample_limit):
    if not header:
        return []
    profiles = []
    for idx, name in enumerate(header):
        non_null = 0
        type_counts = {}
        sample_values = []
        for row in raw_rows[:sample_limit]:
            v = row[idx] if idx < len(row) else None
            if v is None:
                continue
            if isinstance(v, str) and v.strip() == "":
                continue
            non_null += 1
            t = detect_json_value_type(v)
            type_counts[t] = type_counts.get(t, 0) + 1
            if len(sample_values) < 3:
                sample_values.append(json_safe_value(v))
        inferred_type = "null"
        if type_counts:
            inferred_type = sorted(type_counts.items(), key=lambda kv: (-kv[1], kv[0]))[0][0]
        profiles.append(
            {
                "index_1based": idx + 1,
                "name": name,
                "non_null_count": non_null,
                "inferred_type": inferred_type,
                "type_counts": type_counts,
                "sample_values": sample_values,
            }
        )
    return profiles


def col_index_to_label(col_index_1based):
    n = int(col_index_1based)
    if n <= 0:
        return "A"
    chars = []
    while n > 0:
        n, rem = divmod(n - 1, 26)
        chars.append(chr(ord("A") + rem))
    return "".join(reversed(chars))


def extract_formula_sheet_refs(formula_text, current_sheet_name):
    if not formula_text or not isinstance(formula_text, str):
        return set()

    refs = set()
    for m in re.finditer(r"(?:'((?:[^']|'')+)'|([A-Za-z0-9_.\[\]]+))!", formula_text):
        cand = m.group(1) if m.group(1) is not None else m.group(2)
        cand = (cand or "").replace("''", "'").strip()
        if not cand:
            continue
        if "]" in cand:
            cand = cand.split("]", 1)[1].strip()
        if cand and cand != current_sheet_name:
            refs.add(cand)
    return refs
