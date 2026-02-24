import json
import os
import tempfile
import unittest

from office_converter import OfficeConverter


class ConverterMshelpRecordsSplitTests(unittest.TestCase):
    def test_mshelp_records_core_behaviors(self):
        from converter.mshelp_records import (
            build_mshelp_record,
            write_mshelp_index_files,
        )

        root = tempfile.mkdtemp(prefix="mshelp_records_")
        cab = os.path.join(root, "MSHelpViewer", "a.cab")
        md = os.path.join(root, "out.md")
        os.makedirs(os.path.dirname(cab), exist_ok=True)
        with open(cab, "wb") as f:
            f.write(b"x")
        with open(md, "w", encoding="utf-8") as f:
            f.write("# x")

        record = build_mshelp_record(
            cab,
            md,
            3,
            folder_name="MSHelpViewer",
            get_source_root_for_path_fn=lambda p: root,
        )
        self.assertEqual(record["topic_count"], 3)
        self.assertEqual(record["status"], "success")
        self.assertTrue(record["source_cab_relpath"].endswith("MSHelpViewer\\a.cab"))

        generated = []
        outputs = write_mshelp_index_files(
            [record],
            root,
            generated_mshelp_outputs=generated,
        )
        self.assertEqual(len(outputs), 2)
        self.assertEqual(len(generated), 2)
        self.assertTrue(os.path.exists(outputs[0]))
        self.assertTrue(os.path.exists(outputs[1]))

        with open(outputs[0], "r", encoding="utf-8") as f:
            payload = json.load(f)
        self.assertEqual(payload["record_count"], 1)

        for p in [outputs[0], outputs[1], md, cab]:
            try:
                os.remove(p)
            except Exception:
                pass
        for d in [
            os.path.join(root, "_AI", "MSHelp"),
            os.path.join(root, "_AI"),
            os.path.join(root, "MSHelpViewer"),
            root,
        ]:
            try:
                os.rmdir(d)
            except Exception:
                pass

    def test_office_converter_mshelp_record_methods_delegate_to_module(self):
        from converter.mshelp_records import build_mshelp_record, write_mshelp_index_files

        root = tempfile.mkdtemp(prefix="mshelp_records_delegate_")
        cab = os.path.join(root, "MSHelpViewer", "a.cab")
        md = os.path.join(root, "out.md")
        os.makedirs(os.path.dirname(cab), exist_ok=True)
        with open(cab, "wb") as f:
            f.write(b"x")
        with open(md, "w", encoding="utf-8") as f:
            f.write("# x")

        dummy = OfficeConverter.__new__(OfficeConverter)
        dummy.config = {"mshelpviewer_folder_name": "MSHelpViewer", "target_folder": root}
        dummy.mshelp_records = []
        dummy.generated_mshelp_outputs = []
        dummy._get_source_root_for_path = lambda p: root

        expected = build_mshelp_record(
            cab,
            md,
            2,
            folder_name=dummy.config.get("mshelpviewer_folder_name", "MSHelpViewer"),
            get_source_root_for_path_fn=dummy._get_source_root_for_path,
        )
        dummy._append_mshelp_record(cab, md, 2)
        self.assertEqual(dummy.mshelp_records[-1], expected)

        expected_outputs = write_mshelp_index_files(
            dummy.mshelp_records,
            dummy.config.get("target_folder", ""),
            generated_mshelp_outputs=[],
        )
        actual_outputs = dummy._write_mshelp_index_files()
        self.assertEqual(len(actual_outputs), len(expected_outputs))
        self.assertTrue(all(os.path.exists(p) for p in actual_outputs))

        for p in actual_outputs + [md, cab]:
            try:
                os.remove(p)
            except Exception:
                pass
        for d in [
            os.path.join(root, "_AI", "MSHelp"),
            os.path.join(root, "_AI"),
            os.path.join(root, "MSHelpViewer"),
            root,
        ]:
            try:
                os.rmdir(d)
            except Exception:
                pass


if __name__ == "__main__":
    unittest.main()
