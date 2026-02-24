# -*- coding: utf-8 -*-
"""MSHelp topic parsing helper extracted from office_converter.py."""

import os


def parse_mshelp_topics(
    html_root,
    *,
    find_files_recursive_fn,
    has_bs4=False,
    beautifulsoup_cls=None,
    meta_content_by_names_fn=None,
):
    html_files = find_files_recursive_fn(html_root, (".htm", ".html"))
    if not html_files:
        return []

    topics = {}
    for fpath in html_files:
        rel_path = os.path.relpath(fpath, html_root).replace("\\", "/")
        topic_id = rel_path
        parent_id = ""
        title = os.path.splitext(os.path.basename(fpath))[0]

        if has_bs4 and beautifulsoup_cls is not None:
            try:
                with open(fpath, "rb") as f:
                    raw = f.read()
                soup = beautifulsoup_cls(raw, "html.parser")
                meta_id = (
                    meta_content_by_names_fn(soup, ["Microsoft.Help.Id"])
                    if callable(meta_content_by_names_fn)
                    else ""
                )
                meta_parent = (
                    meta_content_by_names_fn(soup, ["Microsoft.Help.TocParent"])
                    if callable(meta_content_by_names_fn)
                    else ""
                )
                meta_title = (
                    meta_content_by_names_fn(soup, ["Title"])
                    if callable(meta_content_by_names_fn)
                    else ""
                )
                title_tag = soup.find("title")
                if meta_id:
                    topic_id = meta_id
                if meta_parent:
                    parent_id = meta_parent
                if meta_title:
                    title = meta_title
                elif title_tag and title_tag.get_text(strip=True):
                    title = title_tag.get_text(strip=True)
            except Exception:
                pass

        topics[topic_id] = {
            "id": topic_id,
            "parent": parent_id,
            "title": title or os.path.basename(fpath),
            "file": fpath,
            "children": [],
        }

    if not topics:
        return []

    roots = []
    for tid, topic in topics.items():
        pid = str(topic.get("parent", "") or "").strip()
        if not pid or pid in ("-1", tid) or pid not in topics:
            roots.append(tid)
        else:
            topics[pid]["children"].append(tid)

    for topic in topics.values():
        topic["children"].sort(
            key=lambda cid: (topics[cid].get("title", ""), topics[cid].get("file", ""))
        )
    roots.sort(key=lambda rid: (topics[rid].get("title", ""), topics[rid].get("file", "")))

    ordered = []
    visited = set()

    def walk(topic_id):
        if topic_id in visited or topic_id not in topics:
            return
        visited.add(topic_id)
        ordered.append(topics[topic_id])
        for child_id in topics[topic_id].get("children", []):
            walk(child_id)

    for rid in roots:
        walk(rid)
    for tid in sorted(topics.keys()):
        walk(tid)

    return ordered
