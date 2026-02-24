# -*- coding: utf-8 -*-
"""Markdown -> DOCX export helper extracted from office_converter.py."""


def export_markdown_to_docx(
    md_path,
    out_docx,
    *,
    has_pydocx,
    document_cls=None,
    re_module=None,
):
    if not has_pydocx or document_cls is None:
        raise RuntimeError("python-docx is not installed.")

    if re_module is None:
        import re as re_module  # noqa: PLC0415

    with open(md_path, "r", encoding="utf-8", errors="ignore") as f:
        lines = f.read().splitlines()

    doc = document_cls()
    in_code = False
    for raw in lines:
        line = str(raw or "")
        if line.strip().startswith("```"):
            in_code = not in_code
            continue
        if in_code:
            doc.add_paragraph(line)
            continue
        s = line.strip()
        if not s:
            doc.add_paragraph("")
            continue
        if s.startswith("### "):
            doc.add_heading(s[4:], level=3)
        elif s.startswith("## "):
            doc.add_heading(s[3:], level=2)
        elif s.startswith("# "):
            doc.add_heading(s[2:], level=1)
        elif s.startswith("- "):
            doc.add_paragraph(s[2:], style="List Bullet")
        elif re_module.match(r"^\d+\.\s+", s):
            text = re_module.sub(r"^\d+\.\s+", "", s)
            doc.add_paragraph(text, style="List Number")
        else:
            doc.add_paragraph(s)
    doc.save(out_docx)
