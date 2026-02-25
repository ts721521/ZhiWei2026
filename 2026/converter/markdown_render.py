# -*- coding: utf-8 -*-
"""HTML -> Markdown rendering helper extracted from office_converter.py."""

import re


def table_to_markdown_lines(table_tag, *, normalize_md_line_fn):
    rows = []
    for tr in table_tag.find_all("tr"):
        cells = tr.find_all(["th", "td"])
        if not cells:
            continue
        row = [normalize_md_line_fn(c.get_text(" ", strip=True)) for c in cells]
        rows.append(row)
    if not rows:
        return []

    width = max(len(r) for r in rows)
    norm_rows = []
    for r in rows:
        norm_rows.append(r + [""] * (width - len(r)))

    header = [c.replace("|", "\\|") for c in norm_rows[0]]
    lines = [
        "| " + " | ".join(header) + " |",
        "| " + " | ".join(["---"] * width) + " |",
    ]
    for r in norm_rows[1:]:
        escaped = [c.replace("|", "\\|") for c in r]
        lines.append("| " + " | ".join(escaped) + " |")
    return lines


def render_html_to_markdown(
    html_path,
    *,
    has_bs4,
    beautifulsoup_cls=None,
    normalize_md_line_fn=None,
    table_to_markdown_lines_fn=None,
):
    if normalize_md_line_fn is None:
        normalize_md_line_fn = lambda s: str(s or "").strip()
    if table_to_markdown_lines_fn is None:
        table_to_markdown_lines_fn = lambda _table: []

    if not has_bs4:
        with open(html_path, "r", encoding="utf-8", errors="ignore") as f:
            raw = f.read()
        text = re.sub(r"<[^>]+>", " ", raw)
        text = re.sub(r"\s+", " ", text).strip()
        return text

    if beautifulsoup_cls is None:
        raise RuntimeError("BeautifulSoup class is required when has_bs4=True")

    with open(html_path, "rb") as f:
        soup = beautifulsoup_cls(f.read(), "html.parser")

    for tag in soup(["script", "style", "noscript", "svg"]):
        tag.decompose()

    body = soup.body if soup.body else soup
    lines = []

    def append_para(text):
        t = normalize_md_line_fn(text)
        if t:
            lines.append(t)
            lines.append("")

    def render(node):
        name = getattr(node, "name", None)
        if not name:
            return
        name = str(name).lower()
        if name in ("script", "style", "noscript", "svg"):
            return

        if name in ("h1", "h2", "h3", "h4", "h5", "h6"):
            level = int(name[1])
            title = normalize_md_line_fn(node.get_text(" ", strip=True))
            if title:
                lines.append("#" * level + " " + title)
                lines.append("")
            return

        if name == "pre":
            code = node.get_text("\n", strip=False).strip("\n")
            if code:
                lines.append("```text")
                lines.append(code)
                lines.append("```")
                lines.append("")
            return

        if name in ("ul", "ol"):
            lis = node.find_all("li", recursive=False)
            if lis:
                for idx, li in enumerate(lis, 1):
                    item_text = normalize_md_line_fn(li.get_text(" ", strip=True))
                    if not item_text:
                        continue
                    prefix = f"{idx}. " if name == "ol" else "- "
                    lines.append(prefix + item_text)
                lines.append("")
            return

        if name == "table":
            table_lines = table_to_markdown_lines_fn(node)
            if table_lines:
                lines.extend(table_lines)
                lines.append("")
            return

        if name in ("p", "blockquote"):
            append_para(node.get_text(" ", strip=True))
            return

        if name in (
            "article",
            "section",
            "main",
            "body",
            "div",
            "header",
            "footer",
            "aside",
            "nav",
        ):
            for child in node.children:
                if getattr(child, "name", None):
                    render(child)
            return

        child_tags = [c for c in node.children if getattr(c, "name", None)]
        if child_tags:
            for child in child_tags:
                render(child)
            return
        append_para(node.get_text(" ", strip=True))

    for child in body.children:
        if getattr(child, "name", None):
            render(child)

    compact = []
    blank = False
    for line in lines:
        if str(line).strip():
            compact.append(line.rstrip())
            blank = False
        else:
            if not blank:
                compact.append("")
            blank = True
    return "\n".join(compact).strip()
