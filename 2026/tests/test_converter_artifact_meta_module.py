import os
import tempfile
import unittest

from office_converter import OfficeConverter


class ConverterArtifactMetaSplitTests(unittest.TestCase):
    def test_artifact_meta_module_exists(self):
        from converter.artifact_meta import add_artifact, safe_file_meta

        root = tempfile.mkdtemp(prefix="artifact_meta_")
        target = os.path.join(root, "target")
        os.makedirs(target, exist_ok=True)
        path = os.path.join(target, "a.txt")
        with open(path, "w", encoding="utf-8") as f:
            f.write("hello")

        try:
            meta = safe_file_meta(path, target)
            self.assertIsInstance(meta, dict)
            self.assertEqual(meta["path_rel_to_target"], "a.txt")

            items = []
            add_artifact(items, "text", path, target)
            self.assertEqual(len(items), 1)
            self.assertEqual(items[0]["kind"], "text")
            self.assertEqual(items[0]["path_rel_to_target"], "a.txt")
        finally:
            try:
                os.remove(path)
            except Exception:
                pass
            try:
                os.rmdir(target)
            except Exception:
                pass
            try:
                os.rmdir(root)
            except Exception:
                pass

    def test_office_converter_artifact_methods_delegate_to_module(self):
        from converter.artifact_meta import add_artifact, safe_file_meta

        root = tempfile.mkdtemp(prefix="artifact_meta_delegate_")
        target = os.path.join(root, "target")
        os.makedirs(target, exist_ok=True)
        path = os.path.join(target, "b.txt")
        with open(path, "w", encoding="utf-8") as f:
            f.write("world")

        dummy = OfficeConverter.__new__(OfficeConverter)
        dummy.config = {"target_folder": target}

        try:
            self.assertEqual(dummy._safe_file_meta(path), safe_file_meta(path, target))

            list_a = []
            list_b = []
            dummy._add_artifact(list_a, "txt", path)
            add_artifact(list_b, "txt", path, target)
            self.assertEqual(list_a, list_b)
        finally:
            try:
                os.remove(path)
            except Exception:
                pass
            try:
                os.rmdir(target)
            except Exception:
                pass
            try:
                os.rmdir(root)
            except Exception:
                pass


if __name__ == "__main__":
    unittest.main()
