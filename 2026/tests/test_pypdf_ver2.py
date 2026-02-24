import unittest

import pypdf
import pypdf.annotations


class PypdfAnnotationsImportTests(unittest.TestCase):
    def test_annotations_module_imports(self):
        self.assertTrue(hasattr(pypdf, "annotations"))
        self.assertTrue(hasattr(pypdf.annotations, "Link"))
        self.assertTrue(hasattr(pypdf.annotations, "Text"))


if __name__ == "__main__":
    unittest.main()
