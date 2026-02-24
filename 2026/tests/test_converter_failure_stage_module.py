import unittest

from office_converter import OfficeConverter


class ConverterFailureStageSplitTests(unittest.TestCase):
    def test_failure_stage_module_exists(self):
        from converter.failure_stage import (
            get_failure_output_expectation,
            infer_failure_stage,
            sanitize_failure_log_stem,
        )

        self.assertEqual(sanitize_failure_log_stem("  a/b:c  "), "a_b_c")
        self.assertEqual(sanitize_failure_log_stem("..."), "failed_file")

        plan = get_failure_output_expectation(
            "convert_then_merge",
            {},
            lambda run_mode, cfg: {"need_final_pdf": 1, "need_markdown": 0},
        )
        self.assertEqual(plan, {"need_final_pdf": True, "need_markdown": False})

        fallback = get_failure_output_expectation(
            "convert_then_merge",
            {},
            lambda run_mode, cfg: (_ for _ in ()).throw(RuntimeError("x")),
        )
        self.assertEqual(
            fallback,
            {"need_final_pdf": None, "need_markdown": None},
        )

        self.assertEqual(
            infer_failure_stage("a.docx", context={"phase": "scan", "scan_scope": "access"}),
            "scan_access",
        )
        self.assertEqual(
            infer_failure_stage("a.docx", raw_error="markdown export failed for x"),
            "markdown_export",
        )
        self.assertEqual(
            infer_failure_stage("a.cab", cab_extensions=[".cab"]),
            "cab_to_markdown",
        )
        self.assertEqual(
            infer_failure_stage(
                "a.pdf",
                expected_outputs_getter=lambda: {
                    "need_markdown": True,
                    "need_final_pdf": False,
                },
            ),
            "pdf_to_markdown",
        )
        self.assertEqual(
            infer_failure_stage("a.pdf", expected_outputs_getter=lambda: {}),
            "pdf_pipeline",
        )

    def test_office_converter_failure_stage_methods_delegate_to_module(self):
        from converter.failure_stage import (
            get_failure_output_expectation,
            infer_failure_stage,
            sanitize_failure_log_stem,
        )

        dummy = OfficeConverter.__new__(OfficeConverter)
        dummy.run_mode = "convert_then_merge"
        dummy.config = {"allowed_extensions": {"cab": [".cab"]}}
        dummy.compute_convert_output_plan = (
            lambda run_mode, cfg: {"need_final_pdf": False, "need_markdown": True}
        )

        self.assertEqual(
            OfficeConverter._sanitize_failure_log_stem("a/b"),
            sanitize_failure_log_stem("a/b"),
        )
        self.assertEqual(
            dummy._get_failure_output_expectation(),
            get_failure_output_expectation(
                dummy.run_mode, dummy.config, dummy.compute_convert_output_plan
            ),
        )
        self.assertEqual(
            dummy._infer_failure_stage("x.pdf"),
            infer_failure_stage(
                "x.pdf",
                cab_extensions=dummy.config.get("allowed_extensions", {}).get("cab", []),
                expected_outputs_getter=dummy._get_failure_output_expectation,
            ),
        )


if __name__ == "__main__":
    unittest.main()
