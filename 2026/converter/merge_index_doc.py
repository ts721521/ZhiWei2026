# -*- coding: utf-8 -*-
"""Merge index Word->PDF helper extracted from office_converter.py."""

import os


def create_index_doc_and_convert(
    word_app,
    file_list,
    title,
    *,
    temp_sandbox,
    uuid4_hex_fn,
    log_error_fn,
):
    try:
        doc = word_app.Documents.Add()
        word_app.Visible = False

        try:
            doc.PageSetup.PaperSize = 7
            doc.PageSetup.TopMargin = 72
            doc.PageSetup.BottomMargin = 72
            doc.PageSetup.LeftMargin = 72
            doc.PageSetup.RightMargin = 72
        except (AttributeError, RuntimeError, TypeError, ValueError):
            pass

        selection = word_app.Selection
        lines_per_page = 32

        def write_header():
            selection.ParagraphFormat.Alignment = 1
            selection.Font.Name = "Microsoft YaHei"
            selection.Font.Size = 16
            selection.Font.Bold = True
            selection.ParagraphFormat.LineSpacingRule = 0
            selection.TypeText(title + "\n")

            selection.ParagraphFormat.Alignment = 0
            selection.Font.Size = 10.5
            selection.Font.Bold = False
            selection.TypeText("\n")

            selection.ParagraphFormat.LineSpacingRule = 4
            selection.ParagraphFormat.LineSpacing = 20

        write_header()

        for i, fname in enumerate(file_list, 1):
            if i > 1 and (i - 1) % lines_per_page == 0:
                selection.InsertBreak(7)
                write_header()
            if len(fname) > 45:
                fname = fname[:42] + "..."
            selection.TypeText(f"{i}. {fname}\n")

        temp_pdf = os.path.join(temp_sandbox, f"index_{uuid4_hex_fn()}.pdf")

        doc.ExportAsFixedFormat(
            OutputFileName=temp_pdf,
            ExportFormat=17,
            OpenAfterExport=False,
            OptimizeFor=0,
            CreateBookmarks=1,
            DocStructureTags=True,
        )
        doc.Close(SaveChanges=0)
        return temp_pdf
    except (AttributeError, RuntimeError, OSError, TypeError, ValueError) as e:
        log_error_fn(f"failed to generate merge index page: {e}")
        return None
