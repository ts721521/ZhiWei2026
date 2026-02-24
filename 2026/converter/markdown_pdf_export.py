# -*- coding: utf-8 -*-
"""Markdown -> PDF export helper extracted from office_converter.py."""


def export_markdown_to_pdf(
    md_path,
    out_pdf,
    *,
    has_reportlab,
    canvas_cls=None,
    page_size=None,
    wrap_plain_text_for_pdf_fn=None,
):
    if not has_reportlab or canvas_cls is None or page_size is None:
        raise RuntimeError("reportlab is not installed.")

    if wrap_plain_text_for_pdf_fn is None:
        wrap_plain_text_for_pdf_fn = lambda text, width=100: [text]

    with open(md_path, "r", encoding="utf-8", errors="ignore") as f:
        lines = f.read().splitlines()

    c = canvas_cls(out_pdf, pagesize=page_size)
    page_w, page_h = page_size
    x = 36
    y = page_h - 36
    line_h = 12

    for raw in lines:
        text = str(raw or "")
        wrapped = wrap_plain_text_for_pdf_fn(text, width=100)
        for w in wrapped:
            if y <= 36:
                c.showPage()
                y = page_h - 36
            c.drawString(x, y, w)
            y -= line_h
    c.save()
