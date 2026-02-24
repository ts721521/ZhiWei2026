import unittest

from office_converter import OfficeConverter


class ConverterMarkdownTextUtilsSplitTests(unittest.TestCase):
    def test_markdown_text_utils_core_behaviors(self):
        from converter.markdown_text_utils import (
            clean_markdown_page_lines,
            collect_margin_candidates,
            looks_like_heading_line,
            looks_like_page_number_line,
            normalize_extracted_text,
            normalize_margin_line,
            render_markdown_blocks,
        )

        self.assertEqual(
            normalize_extracted_text("a\r\n\r\nb\r\n\r\n\r\nc"),
            "a\n\nb\n\nc",
        )
        self.assertEqual(normalize_margin_line("  Header - 1  "), "header1")
        self.assertTrue(looks_like_page_number_line("Page 2 of 9"))
        self.assertTrue(looks_like_page_number_line("第 2 页 / 共 9 页"))
        self.assertFalse(looks_like_page_number_line("正文内容"))

        pages = [
            "My Header\nline a\nPage 1",
            "My Header\nline b\nPage 2",
            "My Header\nline c\nPage 3",
        ]
        hk, fk = collect_margin_candidates(pages)
        self.assertIn(normalize_margin_line("My Header"), hk)
        self.assertEqual(len(fk), 0)

        kept, stats = clean_markdown_page_lines(
            "My Header\n正文\nPage 4",
            hk,
            fk,
        )
        self.assertEqual(kept, ["正文"])
        self.assertEqual(stats["removed_header_lines"], 1)
        self.assertEqual(stats["removed_page_number_lines"], 1)

        self.assertTrue(looks_like_heading_line("1.2 Title"))
        self.assertTrue(looks_like_heading_line("Chapter 3"))
        self.assertTrue(looks_like_heading_line("SYSTEM OVERVIEW"))
        self.assertFalse(looks_like_heading_line("this is a regular sentence"))

        blocks, heading_count = render_markdown_blocks(
            ["1. Intro", "hello", "world", "", "line one-", "line two"]
        )
        self.assertEqual(heading_count, 1)
        self.assertEqual(blocks[0], "### 1. Intro")
        self.assertEqual(blocks[1], "hello world")
        self.assertEqual(blocks[2], "line oneline two")

    def test_office_converter_markdown_methods_delegate_to_module(self):
        from converter.markdown_text_utils import (
            clean_markdown_page_lines,
            collect_margin_candidates,
            looks_like_heading_line,
            looks_like_page_number_line,
            normalize_extracted_text,
            normalize_margin_line,
            render_markdown_blocks,
        )

        self.assertEqual(
            OfficeConverter._normalize_extracted_text("a\r\nb"),
            normalize_extracted_text("a\r\nb"),
        )
        self.assertEqual(
            OfficeConverter._normalize_margin_line(" A - 1 "),
            normalize_margin_line(" A - 1 "),
        )
        self.assertEqual(
            OfficeConverter._looks_like_page_number_line("Page 1"),
            looks_like_page_number_line("Page 1"),
        )
        self.assertEqual(
            OfficeConverter._collect_margin_candidates(["H\nx\n1", "H\ny\n2"]),
            collect_margin_candidates(["H\nx\n1", "H\ny\n2"]),
        )
        self.assertEqual(
            OfficeConverter._clean_markdown_page_lines("H\nx\n1", {"h"}, set()),
            clean_markdown_page_lines("H\nx\n1", {"h"}, set()),
        )
        self.assertEqual(
            OfficeConverter._looks_like_heading_line("1. A"),
            looks_like_heading_line("1. A"),
        )
        self.assertEqual(
            OfficeConverter._render_markdown_blocks(["1. A", "b"]),
            render_markdown_blocks(["1. A", "b"]),
        )


if __name__ == "__main__":
    unittest.main()
