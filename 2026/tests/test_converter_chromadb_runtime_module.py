import tempfile
import unittest

from office_converter import OfficeConverter


class ConverterChromaDbRuntimeSplitTests(unittest.TestCase):
    def test_chromadb_runtime_core_behaviors(self):
        from converter.chromadb_runtime import write_chromadb_export_for_converter

        class Dummy:
            def __init__(self, root):
                self.config = {"enable_chromadb_export": True, "target_folder": root}
                self.generated_chromadb_outputs = []
                self.chromadb_export_manifest_path = None
                self._collect_chromadb_documents = lambda: [{"id": "1"}]
                self._sanitize_chromadb_collection_name = lambda s: s
                self._resolve_chromadb_persist_dir = lambda: root

        root = tempfile.mkdtemp(prefix="chromadb_runtime_")
        d = Dummy(root)
        out = write_chromadb_export_for_converter(
            d,
            write_chromadb_export_fn=lambda docs, **kwargs: ("m.json", ["m.json"]),
            has_chromadb=False,
            chromadb_module=None,
            now_fn=lambda: None,
            log_info_fn=lambda _m: None,
        )
        self.assertEqual(out, "m.json")
        self.assertEqual(d.generated_chromadb_outputs, ["m.json"])
        self.assertEqual(d.chromadb_export_manifest_path, "m.json")

    def test_office_converter_write_chromadb_export_delegates_to_runtime_module(self):
        import office_converter as oc

        original = oc.write_chromadb_export_for_converter_impl
        try:
            seen = {}

            def _fake(converter, **kwargs):
                seen["converter"] = converter
                seen["kwargs"] = kwargs
                return "m.json"

            oc.write_chromadb_export_for_converter_impl = _fake
            dummy = OfficeConverter.__new__(OfficeConverter)
            out = dummy._write_chromadb_export()
            self.assertEqual(out, "m.json")
            self.assertIs(seen["converter"], dummy)
            self.assertTrue(callable(seen["kwargs"]["log_info_fn"]))
        finally:
            oc.write_chromadb_export_for_converter_impl = original


if __name__ == "__main__":
    unittest.main()
