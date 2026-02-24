import os
import tempfile
import unittest

from office_converter import OfficeConverter


class ConverterAiPathsSplitTests(unittest.TestCase):
    def test_ai_paths_core_behaviors(self):
        from converter.ai_paths import build_ai_output_path, build_ai_output_path_from_source

        root = tempfile.mkdtemp(prefix="ai_paths_")
        target = os.path.join(root, "target")
        source = os.path.join(root, "src")
        os.makedirs(os.path.join(target, "docs"), exist_ok=True)
        os.makedirs(os.path.join(source, "A"), exist_ok=True)

        f_in_target = os.path.join(target, "docs", "x.pdf")
        f_in_source = os.path.join(source, "A", "y.docx")
        with open(f_in_target, "w", encoding="utf-8") as f:
            f.write("x")
        with open(f_in_source, "w", encoding="utf-8") as f:
            f.write("y")

        try:
            p1 = build_ai_output_path(f_in_target, "Markdown", ".md", target)
            self.assertTrue(p1.endswith(os.path.join("_AI", "Markdown", "docs", "x.md")))
            self.assertTrue(os.path.isdir(os.path.dirname(p1)))

            p2 = build_ai_output_path_from_source(
                f_in_source,
                "ExcelJSON",
                ".json",
                target,
                source_root_resolver=lambda _: source,
            )
            self.assertTrue(p2.endswith(os.path.join("_AI", "ExcelJSON", "A", "y.json")))
            self.assertTrue(os.path.isdir(os.path.dirname(p2)))
        finally:
            for p in (f_in_target, f_in_source):
                try:
                    os.remove(p)
                except Exception:
                    pass
            for d in (
                os.path.join(target, "_AI", "Markdown", "docs"),
                os.path.join(target, "_AI", "ExcelJSON", "A"),
                os.path.join(target, "docs"),
                os.path.join(source, "A"),
                target,
                source,
                root,
            ):
                try:
                    os.rmdir(d)
                except Exception:
                    pass

    def test_office_converter_ai_path_methods_delegate_to_module(self):
        from converter.ai_paths import build_ai_output_path, build_ai_output_path_from_source

        root = tempfile.mkdtemp(prefix="ai_paths_delegate_")
        target = os.path.join(root, "target")
        source = os.path.join(root, "src")
        os.makedirs(os.path.join(source, "d"), exist_ok=True)
        os.makedirs(target, exist_ok=True)
        src_file = os.path.join(source, "d", "a.docx")
        with open(src_file, "w", encoding="utf-8") as f:
            f.write("a")

        dummy = OfficeConverter.__new__(OfficeConverter)
        dummy.config = {"target_folder": target}
        dummy._get_source_root_for_path = lambda _: source

        try:
            self.assertEqual(
                dummy._build_ai_output_path(src_file, "X", ".md"),
                build_ai_output_path(src_file, "X", ".md", target),
            )
            self.assertEqual(
                dummy._build_ai_output_path_from_source(src_file, "X", ".json"),
                build_ai_output_path_from_source(
                    src_file,
                    "X",
                    ".json",
                    target,
                    source_root_resolver=dummy._get_source_root_for_path,
                ),
            )
        finally:
            try:
                os.remove(src_file)
            except Exception:
                pass
            for d in (
                os.path.join(target, "_AI", "X", "d"),
                os.path.join(target, "_AI", "X"),
                os.path.join(target, "_AI"),
                os.path.join(source, "d"),
                target,
                source,
                root,
            ):
                try:
                    os.rmdir(d)
                except Exception:
                    pass


if __name__ == "__main__":
    unittest.main()
