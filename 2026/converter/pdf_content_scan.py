# -*- coding: utf-8 -*-
"""PDF content scanning helper extracted from office_converter.py."""


def scan_pdf_content(
    pdf_path,
    *,
    price_keywords,
    has_pypdf,
    pdf_reader_cls,
    max_pages=5,
):
    if not has_pypdf:
        return False
    try:
        reader = pdf_reader_cls(pdf_path)
        page_count = min(len(reader.pages), max_pages)
        for i in range(page_count):
            text = reader.pages[i].extract_text()
            if not text:
                continue
            for kw in price_keywords:
                if kw in text:
                    return True
    except (OSError, RuntimeError, TypeError, ValueError, AttributeError):
        pass
    return False
