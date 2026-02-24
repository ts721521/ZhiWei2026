import os
import tempfile
import unittest
from datetime import datetime

from office_converter import OfficeConverter


class ConverterMshelpMergeSplitTests(unittest.TestCase):
    def test_mshelp_merge_core_behaviors(self):
        from converter.mshelp_merge import merge_mshelp_markdowns

        root = tempfile.mkdtemp(prefix="mshelp_merge_")
        md = os.path.join(root, "a.md")
        with open(md, "w", encoding="utf-8") as f:
            f.write("content-a")

        records = [
            {
                "source_cab": os.path.join(root, "a.cab"),
                "mshelpviewer_dir": os.path.join(root, "MSHelpViewer"),
                "markdown_path": md,
                "topic_count": 2,
            }
        ]
        fixed_now = datetime(2026, 2, 24, 12, 34, 56)
        generated = []
        outputs = merge_mshelp_markdowns(
            records,
            {
                "enable_mshelp_merge_output": True,
                "target_folder": root,
                "source_folder": root,
                "enable_mshelp_output_docx": False,
                "enable_mshelp_output_pdf": False,
            },
            generated_outputs=generated,
            now_fn=lambda: fixed_now,
        )
        self.assertEqual(1, len(outputs))
        self.assertEqual(outputs, generated)
        self.assertTrue(outputs[0].endswith("MSHelp_API_Merged_20260224_123456_001.md"))
        self.assertTrue(os.path.exists(outputs[0]))
        with open(outputs[0], "r", encoding="utf-8") as f:
            text = f.read()
        self.assertIn("MSHelp API Merged Package 1/1", text)
        self.assertIn("content-a", text)

        for p in [outputs[0], md]:
            try:
                os.remove(p)
            except Exception:
                pass
        for d in [
            os.path.join(root, "_AI", "MSHelp", "Merged"),
            os.path.join(root, "_AI", "MSHelp"),
            os.path.join(root, "_AI"),
            root,
        ]:
            try:
                os.rmdir(d)
            except Exception:
                pass

    def test_mshelp_merge_with_docx_pdf_callbacks(self):
        from converter.mshelp_merge import merge_mshelp_markdowns

        root = tempfile.mkdtemp(prefix="mshelp_merge_cb_")
        md = os.path.join(root, "a.md")
        with open(md, "w", encoding="utf-8") as f:
            f.write("content-a")

        def mk(path):
            with open(path, "w", encoding="utf-8") as ff:
                ff.write("x")

        outputs = merge_mshelp_markdowns(
            [
                {
                    "source_cab": "a.cab",
                    "mshelpviewer_dir": "d",
                    "markdown_path": md,
                    "topic_count": 1,
                }
            ],
            {
                "enable_mshelp_merge_output": True,
                "target_folder": root,
                "source_folder": root,
                "enable_mshelp_output_docx": True,
                "enable_mshelp_output_pdf": True,
            },
            export_markdown_to_docx_fn=lambda inp, out: mk(out),
            export_markdown_to_pdf_fn=lambda inp, out: mk(out),
            now_fn=lambda: datetime(2026, 2, 24, 12, 35, 0),
        )
        self.assertEqual(3, len(outputs))
        self.assertTrue(any(p.endswith(".docx") for p in outputs))
        self.assertTrue(any(p.endswith(".pdf") for p in outputs))
        self.assertTrue(all(os.path.exists(p) for p in outputs))

        for p in outputs + [md]:
            try:
                os.remove(p)
            except Exception:
                pass
        for d in [
            os.path.join(root, "_AI", "MSHelp", "Merged"),
            os.path.join(root, "_AI", "MSHelp"),
            os.path.join(root, "_AI"),
            root,
        ]:
            try:
                os.rmdir(d)
            except Exception:
                pass

    def test_office_converter_merge_mshelp_markdowns_delegates_to_module(self):
        import office_converter as oc

        original = oc.merge_mshelp_markdowns_impl
        try:
            seen = {}

            def _fake(records, config, **kwargs):
                seen["records"] = records
                seen["config"] = config
                return ["x"]

            oc.merge_mshelp_markdowns_impl = _fake
            dummy = OfficeConverter.__new__(OfficeConverter)
            dummy.mshelp_records = [{"markdown_path": "a.md"}]
            dummy.config = {"enable_mshelp_merge_output": True}
            dummy.generated_mshelp_outputs = []
            dummy._export_markdown_to_docx = lambda i, o: None
            dummy._export_markdown_to_pdf = lambda i, o: None
            dummy._add_perf_seconds = lambda k, v: None

            out = dummy._merge_mshelp_markdowns()
            self.assertEqual(out, ["x"])
            self.assertEqual(seen["records"], dummy.mshelp_records)
        finally:
            oc.merge_mshelp_markdowns_impl = original


if __name__ == "__main__":
    unittest.main()
