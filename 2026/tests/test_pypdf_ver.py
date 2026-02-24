import unittest

import pypdf
from pypdf.annotations import FreeText


class PypdfVersionTests(unittest.TestCase):
    def test_version_string_present(self):
        self.assertIsInstance(pypdf.__version__, str)
        self.assertTrue(pypdf.__version__.strip())

    def test_annotations_module_exposes_freetext(self):
        self.assertTrue(callable(FreeText))


if __name__ == "__main__":
    unittest.main()
