# -*- coding: utf-8 -*-
"""Locate source files by merged PDF page or short id."""

from __future__ import annotations

import argparse
import glob
import json
import os
import sys
from dataclasses import asdict, dataclass
from typing import Optional

from search_adapter import EverythingAdapter, build_listary_query
from converter.traceability import normalize_short_id_for_match

EXIT_OK = 0
EXIT_NOT_FOUND = 2
EXIT_INVALID_INPUT = 3
EXIT_MAP_MISSING = 4
EXIT_SEARCH_ERROR = 5


@dataclass
class LocatorRecord:
    merge_batch_id: str
    merged_pdf_name: str
    merged_pdf_path: str
    source_index: int
    source_filename: str
    source_abspath: str
    source_relpath: str
    source_md5: str
    source_short_id: str
    start_page_1based: int
    end_page_1based: int
    page_count: int
    bookmark_title: str


@dataclass
class LocateResult:
    status: str
    record: Optional[LocatorRecord]
    alternatives: list[LocatorRecord]
    error_code: int


def _to_record(d: dict) -> LocatorRecord:
    return LocatorRecord(
        merge_batch_id=str(d.get("merge_batch_id", "")),
        merged_pdf_name=str(d.get("merged_pdf_name", "")),
        merged_pdf_path=str(d.get("merged_pdf_path", "")),
        source_index=int(d.get("source_index", 0)),
        source_filename=str(d.get("source_filename", "")),
        source_abspath=str(d.get("source_abspath", "")),
        source_relpath=str(d.get("source_relpath", "")),
        source_md5=str(d.get("source_md5", "")),
        source_short_id=str(d.get("source_short_id", "")),
        start_page_1based=int(d.get("start_page_1based", 0)),
        end_page_1based=int(d.get("end_page_1based", 0)),
        page_count=int(d.get("page_count", 0)),
        bookmark_title=str(d.get("bookmark_title", "")),
    )


def _load_json_records(map_json_path: str) -> list[LocatorRecord]:
    with open(map_json_path, "r", encoding="utf-8") as f:
        data = json.load(f)
    records = data.get("records", []) if isinstance(data, dict) else []
    return [_to_record(r) for r in records]


def _find_map_json_by_merged_name(merged_name: str, map_dir: str) -> Optional[str]:
    merged_basename = os.path.basename(merged_name)
    merged_stem, _ = os.path.splitext(merged_basename)
    exact = os.path.join(map_dir, f"{merged_stem}.map.json")
    if os.path.exists(exact):
        return exact

    candidates = glob.glob(os.path.join(map_dir, "*.map.json"))
    hit = []
    for p in candidates:
        try:
            with open(p, "r", encoding="utf-8") as f:
                payload = json.load(f)
            records = payload.get("records", [])
            if not records:
                continue
            if os.path.basename(records[0].get("merged_pdf_name", "")) == merged_basename:
                hit.append(p)
        except Exception:
            continue

    if not hit:
        return None
    hit.sort(key=lambda x: os.path.getmtime(x), reverse=True)
    return hit[0]


def locate_by_page(merged_name: str, page: int, map_dir: str) -> LocateResult:
    if page <= 0:
        return LocateResult("invalid", None, [], EXIT_INVALID_INPUT)
    map_json = _find_map_json_by_merged_name(merged_name, map_dir)
    if not map_json:
        return LocateResult("map_missing", None, [], EXIT_MAP_MISSING)

    records = _load_json_records(map_json)
    if not records:
        return LocateResult("not_found", None, [], EXIT_NOT_FOUND)

    alternatives: list[LocatorRecord] = []
    best_lower = None
    best_upper = None

    for r in records:
        if r.start_page_1based <= page <= r.end_page_1based:
            return LocateResult("ok", r, [], EXIT_OK)
        if r.end_page_1based < page and (best_lower is None or r.end_page_1based > best_lower.end_page_1based):
            best_lower = r
        if r.start_page_1based > page and (best_upper is None or r.start_page_1based < best_upper.start_page_1based):
            best_upper = r

    if best_lower:
        alternatives.append(best_lower)
    if best_upper:
        alternatives.append(best_upper)
    return LocateResult("not_found", None, alternatives, EXIT_NOT_FOUND)


def locate_by_short_id(short_id: str, map_dir: str) -> LocateResult:
    query_raw = short_id.strip().upper()
    if not query_raw:
        return LocateResult("invalid", None, [], EXIT_INVALID_INPUT)
    short_id = normalize_short_id_for_match(query_raw)

    all_maps = glob.glob(os.path.join(map_dir, "*.map.json"))
    if not all_maps:
        return LocateResult("map_missing", None, [], EXIT_MAP_MISSING)

    matches: list[LocatorRecord] = []
    for m in all_maps:
        try:
            records = _load_json_records(m)
        except Exception:
            continue
        for r in records:
            sid = normalize_short_id_for_match(r.source_short_id)
            if sid == short_id or r.source_short_id.strip().upper() == query_raw:
                matches.append(r)

    if not matches:
        return LocateResult("not_found", None, [], EXIT_NOT_FOUND)
    if len(matches) == 1:
        return LocateResult("ok", matches[0], [], EXIT_OK)
    return LocateResult("ambiguous", None, matches, EXIT_NOT_FOUND)


def _print_result(result: LocateResult, as_json: bool):
    if as_json:
        payload = {
            "status": result.status,
            "error_code": result.error_code,
            "record": asdict(result.record) if result.record else None,
            "alternatives": [asdict(x) for x in result.alternatives],
        }
        print(json.dumps(payload, ensure_ascii=False, indent=2))
        return

    if result.record:
        r = result.record
        print(f"状态: {result.status}")
        print(f"文件: {r.source_filename}")
        print(f"路径: {r.source_abspath}")
        print(f"短ID: {r.source_short_id}")
        print(f"MD5: {r.source_md5}")
        print(f"页码范围: {r.start_page_1based}-{r.end_page_1based}")
        print(f"来源合并文件: {r.merged_pdf_name}")
    else:
        print(f"状态: {result.status}")
        if result.alternatives:
            print("候选记录:")
            for x in result.alternatives[:10]:
                print(
                    f"- {x.source_filename} | 页码 {x.start_page_1based}-{x.end_page_1based} | ID {x.source_short_id}"
                )


def main() -> int:
    parser = argparse.ArgumentParser(description="Locate source document from merge map")
    parser.add_argument("--merged", help="Merged PDF filename")
    parser.add_argument("--page", type=int, help="1-based page number in merged PDF")
    parser.add_argument("--id", dest="short_id", help="Short ID generated from md5")
    parser.add_argument("--map-dir", default=".", help="Directory containing *.map.json")
    parser.add_argument("--list", action="store_true", help="List all records for a merged file")
    parser.add_argument("--json", action="store_true", help="Output JSON")
    parser.add_argument("--everything-open", action="store_true", help="Open search query via Everything")
    parser.add_argument("--everything-es-path", default="", help="Custom path to es.exe")
    args = parser.parse_args()

    map_dir = os.path.abspath(args.map_dir)
    if not os.path.isdir(map_dir):
        print(f"map目录不存在: {map_dir}")
        return EXIT_MAP_MISSING

    result: LocateResult
    if args.list:
        if not args.merged:
            print("--list 需要配合 --merged")
            return EXIT_INVALID_INPUT
        map_json = _find_map_json_by_merged_name(args.merged, map_dir)
        if not map_json:
            print("未找到对应 map 文件")
            return EXIT_MAP_MISSING
        records = _load_json_records(map_json)
        payload = {
            "status": "ok",
            "error_code": 0,
            "record": None,
            "alternatives": [asdict(r) for r in records],
        }
        print(json.dumps(payload, ensure_ascii=False, indent=2) if args.json else f"共 {len(records)} 条记录")
        if not args.json:
            for r in records:
                print(f"- {r.source_filename} | 页码 {r.start_page_1based}-{r.end_page_1based} | ID {r.source_short_id}")
        return EXIT_OK

    if args.short_id:
        result = locate_by_short_id(args.short_id, map_dir)
    else:
        if not args.merged or not args.page:
            print("请提供 --merged + --page，或提供 --id")
            return EXIT_INVALID_INPUT
        result = locate_by_page(args.merged, args.page, map_dir)

    _print_result(result, as_json=args.json)

    if args.everything_open and result.record:
        adapter = EverythingAdapter(es_path=args.everything_es_path)
        directory = os.path.dirname(result.record.source_abspath)
        search_ret = adapter.run_query(result.record.source_filename, directory)
        if not search_ret.ok:
            if args.json:
                print(json.dumps({"everything": asdict(search_ret)}, ensure_ascii=False))
            else:
                print(f"Everything 调用失败: {search_ret.stderr}")
                print(build_listary_query(
                    result.record.source_short_id,
                    result.record.source_md5,
                    result.record.source_filename,
                    result.record.source_abspath,
                ))
            return EXIT_SEARCH_ERROR

    return result.error_code


if __name__ == "__main__":
    sys.exit(main())
