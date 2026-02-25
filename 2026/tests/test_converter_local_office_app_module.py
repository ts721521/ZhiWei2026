import unittest
from pathlib import Path

from office_converter import OfficeConverter


class _FakePythonCom:
    def __init__(self):
        self.count = 0

    def CoInitialize(self):
        self.count += 1


class _FakeApp:
    def __init__(self):
        self.Visible = True
        self.DisplayAlerts = True
        self.AskToUpdateLinks = True


class _FakeClient:
    def __init__(self, app, fail_dispatch=False):
        self.app = app
        self.fail_dispatch = fail_dispatch
        self.dispatch_called = 0
        self.dispatchex_called = 0

    def Dispatch(self, prog_id):
        self.dispatch_called += 1
        self.last_prog_id = prog_id
        if self.fail_dispatch:
            raise RuntimeError("x")
        return self.app

    def DispatchEx(self, prog_id):
        self.dispatchex_called += 1
        self.last_prog_id = prog_id
        return self.app


class ConverterLocalOfficeAppSplitTests(unittest.TestCase):
    def test_local_office_app_core_behaviors(self):
        from converter.local_office_app import get_local_app

        pycom = _FakePythonCom()
        app = _FakeApp()
        client = _FakeClient(app, fail_dispatch=True)

        out = get_local_app(
            app_type="excel",
            engine_type="ms",
            has_win32=True,
            engine_wps="wps",
            engine_ms="ms",
            pythoncom_module=pycom,
            win32_client=client,
        )
        self.assertIs(out, app)
        self.assertEqual(pycom.count, 1)
        self.assertEqual(client.dispatch_called, 1)
        self.assertEqual(client.dispatchex_called, 1)
        self.assertFalse(app.Visible)
        self.assertFalse(app.DisplayAlerts)
        self.assertFalse(app.AskToUpdateLinks)

    def test_office_converter_get_local_app_delegates_to_module(self):
        import office_converter as oc

        original = oc.get_local_app_impl
        try:
            seen = {}

            def _fake(**kwargs):
                seen["kwargs"] = kwargs
                return "app"

            oc.get_local_app_impl = _fake
            dummy = OfficeConverter.__new__(OfficeConverter)
            dummy.engine_type = "wps"
            out = dummy._get_local_app("word")
            self.assertEqual(out, "app")
            self.assertEqual(seen["kwargs"]["app_type"], "word")
            self.assertEqual(seen["kwargs"]["engine_type"], "wps")
        finally:
            oc.get_local_app_impl = original

    def test_local_office_app_module_has_no_bare_except_exception(self):
        mod_path = Path(__file__).resolve().parents[1] / "converter" / "local_office_app.py"
        text = mod_path.read_text(encoding="utf-8")
        self.assertNotIn("except Exception", text)


if __name__ == "__main__":
    unittest.main()
