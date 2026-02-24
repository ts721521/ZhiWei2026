import unittest

from pypdf import PdfWriter


class PypdfWriterApiTests(unittest.TestCase):
    def test_writer_creates_blank_pages(self):
        writer = PdfWriter()
        writer.add_blank_page(width=595, height=842)
        writer.add_blank_page(width=595, height=842)
        self.assertEqual(2, len(writer.pages))

    def test_writer_exposes_link_capable_api(self):
        # pypdf>=4 removed add_link; modern APIs include add_annotation/add_uri.
        writer = PdfWriter()
        has_modern_link_api = hasattr(writer, "add_annotation") and hasattr(
            writer, "add_uri"
        )
        self.assertTrue(has_modern_link_api)


if __name__ == "__main__":
    unittest.main()
