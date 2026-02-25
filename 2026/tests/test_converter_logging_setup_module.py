import os
import tempfile
import unittest
from datetime import datetime

from office_converter import OfficeConverter


class _FakeLoggerModule:
    INFO = 20

    class StreamHandler:
        def __init__(self):
            self.level = None
            self.formatter = None

        def setLevel(self, level):
            self.level = level

        def setFormatter(self, formatter):
            self.formatter = formatter

    class Formatter:
        def __init__(self, fmt):
            self.fmt = fmt

    class _Root:
        def __init__(self):
            self.handlers = []

        def addHandler(self, h):
            self.handlers.append(h)

    def __init__(self):
        self.basic_calls = []
        self.root = self._Root()

    def basicConfig(self, **kwargs):
        self.basic_calls.append(kwargs)

    def getLogger(self, _name):
        return self.root


class ConverterLoggingSetupSplitTests(unittest.TestCase):
    def test_logging_setup_core_behaviors(self):
        from converter.logging_setup import setup_logging

        root = tempfile.mkdtemp(prefix="log_setup_")
        fake = _FakeLoggerModule()
        out = setup_logging(
            config={"log_folder": root, "source_folder": "S", "target_folder": "T", "enable_sandbox": True},
            engine_type="wps",
            run_mode="convert_then_merge",
            content_strategy="standard",
            merge_mode="by_category",
            temp_sandbox="tmp",
            merge_output_dir="merged",
            app_version="x",
            get_readable_run_mode_fn=lambda: "rr",
            get_readable_content_strategy_fn=lambda: "cs",
            get_readable_merge_mode_fn=lambda: "mm",
            should_reuse_office_app_fn=lambda: True,
            get_office_restart_every_fn=lambda: 100,
            mode_convert_only="convert_only",
            mode_convert_then_merge="convert_then_merge",
            mode_merge_only="merge_only",
            now_fn=lambda: datetime(2026, 2, 24, 22, 0, 0),
            get_app_path_fn=lambda: root,
            logging_module=fake,
        )
        self.assertTrue(out.endswith(".txt"))
        self.assertTrue(os.path.exists(out))
        self.assertTrue(fake.basic_calls)

    def test_office_converter_setup_logging_delegates_to_module(self):
        import office_converter as oc

        original = oc.setup_logging_impl
        try:
            seen = {}

            def _fake(**kwargs):
                seen["kwargs"] = kwargs
                return "x.log"

            oc.setup_logging_impl = _fake
            dummy = OfficeConverter.__new__(OfficeConverter)
            dummy.config = {}
            dummy.engine_type = "wps"
            dummy.run_mode = "convert_only"
            dummy.content_strategy = "standard"
            dummy.merge_mode = "by_category"
            dummy.temp_sandbox = "tmp"
            dummy.merge_output_dir = "merged"
            dummy.get_readable_run_mode = lambda: "rr"
            dummy.get_readable_content_strategy = lambda: "cs"
            dummy.get_readable_merge_mode = lambda: "mm"
            dummy._should_reuse_office_app = lambda: True
            dummy._get_office_restart_every = lambda: 1

            dummy.setup_logging()
            self.assertEqual(dummy.log_path, "x.log")
            self.assertEqual(seen["kwargs"]["engine_type"], "wps")
        finally:
            oc.setup_logging_impl = original


if __name__ == "__main__":
    unittest.main()
