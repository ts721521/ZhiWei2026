import unittest

from office_converter import OfficeConverter


class _Cell:
    def __init__(self):
        self.hyperlink = None
        self.style = None


class _WS:
    def __init__(self):
        self.rows = []
        self._cells = {}

    def append(self, row):
        self.rows.append(row)

    def cell(self, row, column):
        key = (row, column)
        if key not in self._cells:
            self._cells[key] = _Cell()
        return self._cells[key]


class ConverterIndexSheetsSplitTests(unittest.TestCase):
    def test_index_sheets_core_behaviors(self):
        from converter.index_sheets import (
            write_conversion_index_sheet,
            write_merge_index_sheet,
        )

        ws1 = _WS()
        write_conversion_index_sheet(
            ws1,
            [
                {
                    "source_filename": "a.docx",
                    "source_abspath": "C:\\a.docx",
                    "source_md5": "m1",
                    "pdf_filename": "a.pdf",
                    "pdf_abspath": "C:\\a.pdf",
                    "pdf_md5": "m2",
                    "status": "success",
                }
            ],
            style_header_row_fn=lambda _ws: None,
            auto_fit_sheet_fn=lambda _ws: None,
            make_file_hyperlink_fn=lambda p: f"link://{p}",
        )
        self.assertEqual(len(ws1.rows), 2)
        self.assertEqual(ws1.cell(2, 3).hyperlink, "link://C:\\a.docx")
        self.assertEqual(ws1.cell(2, 6).hyperlink, "link://C:\\a.pdf")

        ws2 = _WS()
        write_merge_index_sheet(
            ws2,
            [
                {
                    "merged_pdf_name": "m.pdf",
                    "merged_pdf_path": "C:\\m.pdf",
                    "merged_pdf_md5": "x",
                    "source_index": 1,
                    "source_filename": "s.pdf",
                    "source_abspath": "C:\\s.pdf",
                    "source_md5": "y",
                    "source_short_id": "id",
                    "start_page_1based": 1,
                    "end_page_1based": 3,
                    "page_count": 3,
                }
            ],
            style_header_row_fn=lambda _ws: None,
            auto_fit_sheet_fn=lambda _ws: None,
            make_file_hyperlink_fn=lambda p: f"link://{p}",
        )
        self.assertEqual(len(ws2.rows), 2)
        self.assertEqual(ws2.cell(2, 3).hyperlink, "link://C:\\m.pdf")
        self.assertEqual(ws2.cell(2, 7).hyperlink, "link://C:\\s.pdf")

    def test_office_converter_index_sheet_methods_delegate_to_module(self):
        import office_converter as oc

        ori_conv = oc.write_conversion_index_sheet_impl
        ori_merge = oc.write_merge_index_sheet_impl
        try:
            seen = {}

            def _fake_conv(ws, records, **kwargs):
                seen["conv"] = (ws, records, kwargs)
                return "c"

            def _fake_merge(ws, records, **kwargs):
                seen["merge"] = (ws, records, kwargs)
                return "m"

            oc.write_conversion_index_sheet_impl = _fake_conv
            oc.write_merge_index_sheet_impl = _fake_merge

            dummy = OfficeConverter.__new__(OfficeConverter)
            dummy._style_header_row = lambda _ws: None
            dummy._auto_fit_sheet = lambda _ws: None
            dummy._make_file_hyperlink = lambda p: p

            self.assertEqual(dummy._write_conversion_index_sheet("ws1", [{"a": 1}]), "c")
            self.assertEqual(dummy._write_merge_index_sheet("ws2", [{"b": 2}]), "m")
            self.assertEqual(seen["conv"][0], "ws1")
            self.assertEqual(seen["merge"][0], "ws2")
        finally:
            oc.write_conversion_index_sheet_impl = ori_conv
            oc.write_merge_index_sheet_impl = ori_merge


if __name__ == "__main__":
    unittest.main()
