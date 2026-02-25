import os
import tempfile
import unittest
from pathlib import Path

from office_converter import OfficeConverter


class ConverterChromaDbDocsSplitTests(unittest.TestCase):
    def test_chromadb_docs_module_has_no_bare_except_exception(self):
        module_text = Path("converter/chromadb_docs.py").read_text(encoding="utf-8")
        self.assertNotIn("except Exception", module_text)

    def test_chromadb_docs_core_behaviors(self):
        from converter.chromadb_docs import collect_chromadb_documents

        root = tempfile.mkdtemp(prefix="chromadb_docs_")
        md = os.path.join(root, "a.md")
        with open(md, "w", encoding="utf-8") as f:
            f.write("hello world")

        docs = collect_chromadb_documents(
            generated_markdown_outputs=[md, md, os.path.join(root, "missing.md")],
            markdown_quality_records=[
                {"markdown_path": md, "source_pdf": r"C:\x\a.pdf"},
            ],
            config={"chromadb_max_chars_per_chunk": 20, "chromadb_chunk_overlap": 5},
            chunk_text_for_vector_fn=lambda text, max_chars, overlap: ["c1", "c2"],
        )
        self.assertEqual(len(docs), 2)
        self.assertTrue(docs[0]["id"].startswith("md_"))
        self.assertEqual(docs[0]["metadata"]["source_markdown_path"], os.path.abspath(md))
        self.assertEqual(docs[0]["metadata"]["source_pdf_path"], r"C:\x\a.pdf")
        self.assertEqual(docs[0]["metadata"]["chunk_count"], 2)

        try:
            os.remove(md)
            os.rmdir(root)
        except Exception:
            pass

    def test_office_converter_collect_chromadb_documents_delegates_to_module(self):
        from converter.chromadb_docs import collect_chromadb_documents

        root = tempfile.mkdtemp(prefix="chromadb_docs_delegate_")
        md = os.path.join(root, "a.md")
        with open(md, "w", encoding="utf-8") as f:
            f.write("hello")

        dummy = OfficeConverter.__new__(OfficeConverter)
        dummy.generated_markdown_outputs = [md]
        dummy.markdown_quality_records = [{"markdown_path": md, "source_pdf": ""}]
        dummy.config = {"chromadb_max_chars_per_chunk": 10, "chromadb_chunk_overlap": 0}
        dummy._chunk_text_for_vector = lambda text, max_chars, overlap: ["x"]

        expected = collect_chromadb_documents(
            generated_markdown_outputs=dummy.generated_markdown_outputs,
            markdown_quality_records=dummy.markdown_quality_records,
            config=dummy.config,
            chunk_text_for_vector_fn=dummy._chunk_text_for_vector,
        )
        actual = dummy._collect_chromadb_documents()
        self.assertEqual(actual, expected)

        try:
            os.remove(md)
            os.rmdir(root)
        except Exception:
            pass


if __name__ == "__main__":
    unittest.main()
