import os
import tempfile
import unittest

from office_converter import OfficeConverter


class ConverterUpdatePackageIndexSplitTests(unittest.TestCase):
    def test_update_package_index_core_behaviors(self):
        from converter.update_package_index import write_update_package_index_xlsx

        root = tempfile.mkdtemp(prefix="update_pkg_idx_")
        xlsx = os.path.join(root, "idx.xlsx")
        links = []

        out = write_update_package_index_xlsx(
            xlsx,
            [
                {
                    "seq": 1,
                    "change_state": "modified",
                    "process_status": "success",
                    "source_file": "a.docx",
                    "source_path": r"C:\x\a.docx",
                    "renamed_from": r"C:\x\old.docx",
                    "packaged_pdf_path": r"C:\x\a.pdf",
                }
            ],
            has_openpyxl=False,
            make_file_hyperlink_fn=lambda p: links.append(p) or p,
        )
        self.assertIsNone(out)
        self.assertEqual(links, [])

        import office_converter as oc

        out2 = write_update_package_index_xlsx(
            xlsx,
            [
                {
                    "seq": 1,
                    "change_state": "modified",
                    "process_status": "success",
                    "source_file": "a.docx",
                    "source_path": r"C:\x\a.docx",
                    "renamed_from": r"C:\x\old.docx",
                    "packaged_pdf_path": r"C:\x\a.pdf",
                }
            ],
            has_openpyxl=oc.HAS_OPENPYXL,
            workbook_cls=getattr(oc, "Workbook", None),
            font_cls=getattr(oc, "Font", None),
            make_file_hyperlink_fn=lambda p: links.append(p) or p,
        )
        if oc.HAS_OPENPYXL:
            self.assertEqual(out2, xlsx)
            self.assertTrue(os.path.exists(xlsx))
            self.assertGreaterEqual(len(links), 3)
        else:
            self.assertIsNone(out2)

        try:
            if os.path.exists(xlsx):
                os.remove(xlsx)
            os.rmdir(root)
        except Exception:
            pass

    def test_office_converter_update_package_index_method_delegates_to_module(self):
        from converter.update_package_index import write_update_package_index_xlsx
        import office_converter as oc

        root = tempfile.mkdtemp(prefix="update_pkg_idx_delegate_")
        xlsx = os.path.join(root, "idx.xlsx")
        dummy = OfficeConverter.__new__(OfficeConverter)
        dummy._make_file_hyperlink = lambda p: p

        records = [{"seq": 1, "source_path": r"C:\x\a.docx"}]
        expected = write_update_package_index_xlsx(
            xlsx,
            records,
            has_openpyxl=oc.HAS_OPENPYXL,
            workbook_cls=getattr(oc, "Workbook", None),
            font_cls=getattr(oc, "Font", None),
            make_file_hyperlink_fn=dummy._make_file_hyperlink,
        )
        actual = dummy._write_update_package_index_xlsx(xlsx, records)
        self.assertEqual(actual, expected)

        try:
            if os.path.exists(xlsx):
                os.remove(xlsx)
            os.rmdir(root)
        except Exception:
            pass


if __name__ == "__main__":
    unittest.main()
