import unittest

from office_converter import OfficeConverter


class ConverterSafeExecSplitTests(unittest.TestCase):
    def test_safe_exec_core_behaviors(self):
        from converter.safe_exec import safe_exec

        self.assertEqual(safe_exec(lambda x: x + 1, 1, retries=0), 2)

        calls = {"n": 0}

        def flaky():
            calls["n"] += 1
            if calls["n"] < 3:
                raise RuntimeError("tmp")
            return "ok"

        sleep_calls = []
        out = safe_exec(flaky, retries=3, sleep_fn=lambda s: sleep_calls.append(s))
        self.assertEqual(out, "ok")
        self.assertEqual(calls["n"], 3)
        self.assertEqual(sleep_calls, [1, 1])

        class _ComErr(Exception):
            def __init__(self, hresult):
                super().__init__(f"err:{hresult}")
                self.hresult = hresult

        busy_calls = {"n": 0}
        sleeps = []

        def busy_then_ok():
            busy_calls["n"] += 1
            if busy_calls["n"] < 3:
                raise _ComErr(99)
            return "done"

        out = safe_exec(
            busy_then_ok,
            retries=3,
            sleep_fn=lambda s: sleeps.append(s),
            randint_fn=lambda a, b: 2,
            com_error_cls=_ComErr,
            rpc_server_busy_code=99,
        )
        self.assertEqual(out, "done")
        self.assertEqual(sleeps, [2, 2])

        def always_com_error():
            raise _ComErr(123)

        with self.assertRaisesRegex(Exception, r"COM error \(123\)"):
            safe_exec(
                always_com_error,
                retries=1,
                sleep_fn=lambda s: None,
                com_error_cls=_ComErr,
                rpc_server_busy_code=99,
            )

    def test_office_converter_safe_exec_delegates_to_module(self):
        dummy = OfficeConverter.__new__(OfficeConverter)
        dummy.is_running = True
        self.assertEqual(dummy._safe_exec(lambda x: x * 2, 3, retries=0), 6)

        dummy.is_running = False
        with self.assertRaisesRegex(Exception, "program stopped"):
            dummy._safe_exec(lambda: "x", retries=0)


if __name__ == "__main__":
    unittest.main()
