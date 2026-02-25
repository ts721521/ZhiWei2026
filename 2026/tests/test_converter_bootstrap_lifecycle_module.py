import unittest
from pathlib import Path

from office_converter import OfficeConverter


class ConverterBootstrapLifecycleSplitTests(unittest.TestCase):
    def test_bootstrap_state_core_behaviors(self):
        from converter.bootstrap_state import (
            build_default_perf_metrics,
            handle_stop_signal,
            initialize_converter_for_runtime,
            initialize_error_tracking_state,
            initialize_output_tracking_state,
            initialize_runtime_state,
            register_signal_handlers,
        )

        class Dummy:
            pass

        d = Dummy()
        initialize_runtime_state(
            d,
            config_path="cfg.json",
            interactive=True,
            mode_convert_then_merge="convert_then_merge",
            collect_mode_copy_and_index="copy_and_index",
            merge_mode_category="category",
            strategy_standard="standard",
        )
        self.assertEqual(d.config_path, "cfg.json")
        self.assertEqual(d.run_mode, "convert_then_merge")

        initialize_output_tracking_state(d)
        self.assertEqual(d.generated_pdfs, [])
        self.assertIn("scan_seconds", d.perf_metrics)

        initialize_error_tracking_state(d)
        self.assertEqual(d.stats["total"], 0)
        self.assertEqual(d.error_records, [])

        calls = []

        class Sig:
            SIGINT = 2
            SIGTERM = 15

            @staticmethod
            def signal(sig, fn):
                calls.append((sig, fn))

        register_signal_handlers(
            signal_module=Sig,
            current_thread_fn=lambda: "t",
            main_thread_fn=lambda: "t",
            signal_handler_fn=lambda *_a: None,
        )
        self.assertEqual(len(calls), 2)
        marker = {"is_running": True, "logs": []}
        handle_stop_signal(
            2,
            set_running_fn=lambda v: marker.__setitem__("is_running", v),
            log_warning_fn=lambda m: marker["logs"].append(m),
        )
        self.assertFalse(marker["is_running"])
        self.assertTrue(marker["logs"])
        self.assertIn("merge_seconds", build_default_perf_metrics())

        class Dummy2:
            def __init__(self):
                self.called = []

            def signal_handler(self, *_a):
                return None

            def load_config(self, path):
                self.called.append(("load", path))
                self.config = {"target_folder": "x"}

            def _init_paths_from_config(self):
                self.called.append(("paths", None))

        d2 = Dummy2()
        initialize_converter_for_runtime(
            d2,
            config_path="cfg2.json",
            interactive=False,
            mode_convert_then_merge="convert_then_merge",
            collect_mode_copy_and_index="copy_and_index",
            merge_mode_category="category",
            strategy_standard="standard",
            signal_module=Sig,
            current_thread_fn=lambda: "t",
            main_thread_fn=lambda: "t",
        )
        self.assertEqual(d2.config_path, "cfg2.json")
        self.assertEqual([("load", "cfg2.json"), ("paths", None)], d2.called)

    def test_runtime_lifecycle_core_behaviors(self):
        from converter.runtime_lifecycle import (
            check_and_handle_running_processes,
            check_and_handle_running_processes_for_converter,
            cleanup_all_processes,
            close_office_apps,
            kill_current_app,
            on_office_file_processed,
        )

        killed = []
        cleanup_all_processes(
            "wps",
            process_names_for_engine_fn=lambda _e: ["a", "b"],
            kill_process_by_name_fn=lambda n: killed.append(n),
        )
        self.assertEqual(killed, ["a", "b"])

        called = {"c": 0}
        close_office_apps(
            reuse_process=False,
            run_mode="convert_only",
            mode_merge_only="merge_only",
            mode_collect_only="collect_only",
            cleanup_all_processes_fn=lambda: called.__setitem__("c", called["c"] + 1),
        )
        self.assertEqual(called["c"], 1)

        killed_apps = []
        kill_current_app(
            "word",
            reuse_process=False,
            force=False,
            engine_type="wps",
            engine_wps="wps",
            engine_ms="ms",
            kill_process_by_name_fn=lambda n: killed_apps.append(n),
        )
        self.assertEqual(["wps"], killed_apps)

        logs = []
        cnt = on_office_file_processed(
            ".docx",
            should_reuse_office_app_fn=lambda: True,
            reuse_process=False,
            get_office_restart_every_fn=lambda: 2,
            get_app_type_for_ext_fn=lambda _e: "word",
            office_file_counter=1,
            kill_current_app_fn=lambda _t, force=False: logs.append(force),
            log_info_fn=lambda m: logs.append(m),
        )
        self.assertEqual(cnt, 2)
        self.assertTrue(any(isinstance(x, str) for x in logs))

        reuse = check_and_handle_running_processes(
            run_mode="convert_only",
            config={"kill_process_mode": "ask"},
            interactive=True,
            resolve_process_handling_fn=lambda **_k: {
                "skip": False,
                "cleanup_all": True,
                "reuse_process": True,
            },
            cleanup_all_processes_fn=lambda: logs.append("cleanup"),
        )
        self.assertTrue(reuse)
        self.assertIn("cleanup", logs)

        class _Dummy:
            run_mode = "convert_only"
            config = {"kill_process_mode": "ask"}
            interactive = True
            reuse_process = False

            @staticmethod
            def cleanup_all_processes():
                logs.append("cleanup2")

        r2 = check_and_handle_running_processes_for_converter(
            _Dummy,
            resolve_process_handling_fn=lambda **_k: {
                "skip": False,
                "cleanup_all": False,
                "reuse_process": True,
            },
        )
        self.assertTrue(r2)
        self.assertTrue(_Dummy.reuse_process)

    def test_office_converter_methods_delegate_to_new_modules(self):
        import office_converter as oc

        originals = (
            oc.build_default_perf_metrics_impl,
            oc.cleanup_all_processes_impl,
            oc.close_office_apps_impl,
            oc.kill_current_app_impl,
            oc.on_office_file_processed_impl,
            oc.check_and_handle_running_processes_for_converter,
            oc.handle_stop_signal_impl,
            oc.initialize_converter_for_runtime,
        )
        try:
            oc.build_default_perf_metrics_impl = lambda: {"scan_seconds": 1.0}
            oc.cleanup_all_processes_impl = lambda *_a, **_k: "cleanup"
            oc.close_office_apps_impl = lambda **_k: "close"
            oc.kill_current_app_impl = lambda *_a, **_k: "kill"
            oc.on_office_file_processed_impl = lambda *_a, **_k: 7
            oc.check_and_handle_running_processes_for_converter = lambda *_a, **_k: True
            oc.handle_stop_signal_impl = lambda *_a, **_k: "sig"
            oc.initialize_converter_for_runtime = lambda *_a, **_k: None

            dummy = OfficeConverter.__new__(OfficeConverter)
            dummy.engine_type = "wps"
            dummy.reuse_process = False
            dummy.run_mode = "convert_only"
            dummy.config = {"kill_process_mode": "ask"}
            dummy.interactive = True
            dummy._office_file_counter = 0
            dummy._should_reuse_office_app = lambda: True
            dummy._get_office_restart_every = lambda: 1
            dummy._get_app_type_for_ext = lambda _e: "word"

            dummy._reset_perf_metrics()
            self.assertEqual(dummy.perf_metrics["scan_seconds"], 1.0)
            self.assertEqual(dummy.cleanup_all_processes(), "cleanup")
            self.assertEqual(dummy.close_office_apps(), "close")
            self.assertEqual(dummy._kill_current_app("word"), "kill")
            self.assertEqual(dummy.signal_handler(2, None), "sig")
            dummy._on_office_file_processed(".docx")
            self.assertEqual(dummy._office_file_counter, 7)
            self.assertTrue(dummy.check_and_handle_running_processes())
        finally:
            (
                oc.build_default_perf_metrics_impl,
                oc.cleanup_all_processes_impl,
                oc.close_office_apps_impl,
                oc.kill_current_app_impl,
                oc.on_office_file_processed_impl,
                oc.check_and_handle_running_processes_for_converter,
                oc.handle_stop_signal_impl,
                oc.initialize_converter_for_runtime,
            ) = originals

    def test_office_converter_init_delegates_to_bootstrap_wrapper(self):
        import office_converter as oc

        original = oc.initialize_converter_for_runtime
        try:
            seen = {}

            def _fake(converter, **kwargs):
                seen["converter"] = converter
                seen["kwargs"] = kwargs
                converter.config = {}

            oc.initialize_converter_for_runtime = _fake
            conv = OfficeConverter("cfg.json", interactive=True)
            self.assertIs(seen["converter"], conv)
            self.assertEqual("cfg.json", seen["kwargs"]["config_path"])
            self.assertEqual(oc.MODE_CONVERT_THEN_MERGE, seen["kwargs"]["mode_convert_then_merge"])
        finally:
            oc.initialize_converter_for_runtime = original

    def test_new_modules_have_no_bare_except_exception(self):
        for rel in (
            "converter/bootstrap_state.py",
            "converter/runtime_lifecycle.py",
        ):
            text = Path(rel).read_text(encoding="utf-8")
            self.assertNotIn("except Exception", text)


if __name__ == "__main__":
    unittest.main()
