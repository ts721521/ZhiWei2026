import unittest

from office_converter import OfficeConverter


class ConverterMergeModePipelineSplitTests(unittest.TestCase):
    def test_merge_mode_pipeline_core_behaviors_pdf_to_md(self):
        from converter.constants import MERGE_CONVERT_SUBMODE_PDF_TO_MD
        from converter.merge_mode_pipeline import run_merge_mode_pipeline

        class Dummy:
            config = {"enable_merge": True}
            generated_markdown_outputs = []

            def _get_output_pref(self):
                return {"pdf": False, "md": True, "merged": True, "independent": True}

            def _get_merge_convert_submode(self):
                return MERGE_CONVERT_SUBMODE_PDF_TO_MD

            def _scan_merge_candidates_by_ext(self, ext):
                if ext == ".pdf":
                    return [r"C:\x\a.pdf"]
                if ext == ".md":
                    return [r"C:\x\a.md"]
                return []

            def _export_pdf_markdown(self, path, source_path_hint=None):
                return r"C:\x\a.md"

            def _confirm_continue_missing_md_merge(self):
                return True

            def merge_pdfs(self):
                return []

            def merge_markdowns(self, candidates=None):
                self.last_candidates = list(candidates or [])
                return ["merged.md"]

        dummy = Dummy()
        batch_results = []
        out = run_merge_mode_pipeline(dummy, batch_results)
        self.assertEqual(out, ["merged.md"])
        self.assertEqual(len(batch_results), 1)
        self.assertEqual(batch_results[0]["detail"], "pdf_to_md")
        self.assertEqual(dummy.last_candidates, [r"C:\x\a.md"])

    def test_office_converter_run_merge_mode_pipeline_delegates_to_module(self):
        import office_converter as oc

        original = oc.run_merge_mode_pipeline_impl
        try:
            seen = {}

            def _fake(converter, batch_results):
                seen["converter"] = converter
                seen["batch_results"] = batch_results
                return ["ok.pdf"]

            oc.run_merge_mode_pipeline_impl = _fake
            dummy = OfficeConverter.__new__(OfficeConverter)
            rows = []
            out = dummy._run_merge_mode_pipeline(rows)
            self.assertEqual(out, ["ok.pdf"])
            self.assertIs(seen["converter"], dummy)
            self.assertIs(seen["batch_results"], rows)
        finally:
            oc.run_merge_mode_pipeline_impl = original


if __name__ == "__main__":
    unittest.main()
