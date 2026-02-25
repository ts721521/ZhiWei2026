import os
import tempfile
import unittest
from pathlib import Path

from office_converter import OfficeConverter


class ConverterCabExtractSplitTests(unittest.TestCase):
    def test_cab_extract_core_behaviors(self):
        from converter.cab_extract import extract_cab_with_fallback

        root = tempfile.mkdtemp(prefix="cab_extract_")
        cab = os.path.join(root, "a.cab")
        out = os.path.join(root, "out")
        with open(cab, "wb") as f:
            f.write(b"x")

        calls = {"run": []}

        def _run(cmd, **kwargs):
            calls["run"].append(cmd)
            return None

        # win expand success path
        extract_cab_with_fallback(
            cab,
            out,
            is_win_fn=lambda: True,
            run_cmd=_run,
            find_files_recursive_fn=lambda _r, _e: ["x.html"],
            cab_7z_path="",
            get_app_path_fn=lambda: root,
            which_fn=lambda _n: "",
        )
        self.assertTrue(calls["run"])

        # fallback no 7z
        with self.assertRaises(RuntimeError):
            extract_cab_with_fallback(
                cab,
                out,
                is_win_fn=lambda: False,
                run_cmd=_run,
                find_files_recursive_fn=lambda _r, _e: [],
                cab_7z_path="",
                get_app_path_fn=lambda: root,
                which_fn=lambda _n: "",
            )

    def test_office_converter_extract_cab_delegates_to_module(self):
        import office_converter as oc

        original = oc.extract_cab_with_fallback_impl
        try:
            seen = {}

            def _fake(cab_path, extract_dir, **kwargs):
                seen["args"] = (cab_path, extract_dir)
                seen["kwargs"] = kwargs
                return "ok"

            oc.extract_cab_with_fallback_impl = _fake
            dummy = OfficeConverter.__new__(OfficeConverter)
            dummy.config = {"cab_7z_path": ""}
            dummy._find_files_recursive = lambda *_a, **_k: []

            out = dummy._extract_cab_with_fallback("a.cab", "out")
            self.assertEqual(out, "ok")
            self.assertEqual(seen["args"], ("a.cab", "out"))
        finally:
            oc.extract_cab_with_fallback_impl = original

    def test_cab_extract_module_has_no_bare_except_exception(self):
        mod_path = Path(__file__).resolve().parents[1] / "converter" / "cab_extract.py"
        text = mod_path.read_text(encoding="utf-8")
        self.assertNotIn("except Exception", text)


if __name__ == "__main__":
    unittest.main()
