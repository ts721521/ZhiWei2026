import unittest

from office_converter import OfficeConverter


class ConverterAiPathsRuntimeSplitTests(unittest.TestCase):
    def test_ai_paths_runtime_core_behaviors(self):
        from converter.ai_paths_runtime import build_ai_output_path_from_source_for_converter

        class Dummy:
            def __init__(self):
                self.config = {"target_folder": "C:/target"}
                self._get_source_root_for_path = lambda _p: "C:/src"

        seen = {}

        def _fake(source_path, sub_dir, ext, target_root, source_root_resolver=None):
            seen["args"] = (source_path, sub_dir, ext, target_root)
            seen["resolver"] = source_root_resolver
            return "x"

        out = build_ai_output_path_from_source_for_converter(
            Dummy(),
            "C:/src/a.xlsx",
            "ExcelJSON",
            ".json",
            build_ai_output_path_from_source_fn=_fake,
        )
        self.assertEqual(out, "x")
        self.assertEqual(seen["args"][3], "C:/target")
        self.assertTrue(callable(seen["resolver"]))

    def test_office_converter_ai_path_from_source_delegates_to_runtime_module(self):
        import office_converter as oc

        original = oc.build_ai_output_path_from_source_for_converter_impl
        try:
            seen = {}

            def _fake(converter, source_path, sub_dir, ext, **kwargs):
                seen["converter"] = converter
                seen["args"] = (source_path, sub_dir, ext)
                seen["kwargs"] = kwargs
                return "ok"

            oc.build_ai_output_path_from_source_for_converter_impl = _fake
            dummy = OfficeConverter.__new__(OfficeConverter)
            out = dummy._build_ai_output_path_from_source("a", "b", ".c")
            self.assertEqual(out, "ok")
            self.assertIs(seen["converter"], dummy)
            self.assertEqual(seen["args"], ("a", "b", ".c"))
        finally:
            oc.build_ai_output_path_from_source_for_converter_impl = original


if __name__ == "__main__":
    unittest.main()
