import os
import unittest

from office_converter import OfficeConverter


class ConverterChromaDbUtilsSplitTests(unittest.TestCase):
    def test_chromadb_utils_core_behaviors(self):
        from converter.chromadb_utils import (
            chunk_text_for_vector,
            resolve_chromadb_persist_dir,
            sanitize_chromadb_collection_name,
        )

        self.assertEqual(sanitize_chromadb_collection_name("  A B  "), "A_B")
        self.assertEqual(sanitize_chromadb_collection_name(".."), "office_corpus")
        self.assertTrue(len(sanitize_chromadb_collection_name("x" * 100)) <= 63)

        cfg = {"target_folder": r"C:\tmp\target", "chromadb_persist_dir": "db_rel"}
        self.assertEqual(
            resolve_chromadb_persist_dir(cfg),
            os.path.abspath(os.path.join(cfg["target_folder"], "db_rel")),
        )
        self.assertEqual(
            resolve_chromadb_persist_dir({"target_folder": r"C:\tmp\target"}),
            os.path.join(r"C:\tmp\target", "_AI", "ChromaDB", "db"),
        )

        chunks = chunk_text_for_vector("a" * 550, max_chars=200, overlap=50)
        self.assertGreaterEqual(len(chunks), 3)
        self.assertTrue(all(isinstance(c, str) and c for c in chunks))

    def test_office_converter_chromadb_methods_delegate_to_module(self):
        from converter.chromadb_utils import (
            chunk_text_for_vector,
            resolve_chromadb_persist_dir,
            sanitize_chromadb_collection_name,
        )

        dummy = OfficeConverter.__new__(OfficeConverter)
        dummy.config = {"target_folder": r"C:\tmp\target", "chromadb_persist_dir": "db_rel"}

        self.assertEqual(
            OfficeConverter._sanitize_chromadb_collection_name(" A B "),
            sanitize_chromadb_collection_name(" A B "),
        )
        self.assertEqual(
            dummy._resolve_chromadb_persist_dir(),
            resolve_chromadb_persist_dir(dummy.config),
        )
        self.assertEqual(
            OfficeConverter._chunk_text_for_vector("abcdef", 3, 1),
            chunk_text_for_vector("abcdef", 3, 1),
        )


if __name__ == "__main__":
    unittest.main()
