import unittest

from pypdf import PdfWriter


class PypdfWriterDirTests(unittest.TestCase):
    def test_writer_add_methods_include_expected_subset(self):
        add_methods = {x for x in dir(PdfWriter) if x.startswith("add")}
        self.assertIn("add_blank_page", add_methods)
        self.assertIn("add_annotation", add_methods)
        self.assertIn("add_uri", add_methods)


if __name__ == "__main__":
    unittest.main()
