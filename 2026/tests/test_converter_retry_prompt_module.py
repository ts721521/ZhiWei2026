import unittest

from office_converter import OfficeConverter


class _FakeMsvcrt:
    def __init__(self, chars):
        self._chars = list(chars)

    def kbhit(self):
        return bool(self._chars)

    def getwch(self):
        if self._chars:
            return self._chars.pop(0)
        return ""


class _FakeTime:
    def __init__(self):
        self.t = 0.0

    def time(self):
        self.t += 0.2
        return self.t

    def sleep(self, _s):
        self.t += 0.1


class ConverterRetryPromptSplitTests(unittest.TestCase):
    def test_retry_prompt_core_behaviors(self):
        from converter.retry_prompt import ask_retry_failed_files

        out = ask_retry_failed_files(
            2,
            error_records=["a", "b"],
            timeout=1,
            has_msvcrt=False,
            msvcrt_module=None,
            time_module=_FakeTime(),
            input_fn=lambda _p: "y",
            print_fn=lambda *_a, **_k: None,
        )
        self.assertTrue(out)

        out2 = ask_retry_failed_files(
            1,
            error_records=[],
            timeout=2,
            has_msvcrt=True,
            msvcrt_module=_FakeMsvcrt(["n", "\r"]),
            time_module=_FakeTime(),
            input_fn=lambda _p: "n",
            print_fn=lambda *_a, **_k: None,
        )
        self.assertFalse(out2)

    def test_office_converter_retry_prompt_delegates_to_module(self):
        import office_converter as oc

        original = oc.ask_retry_failed_files_impl
        try:
            seen = {}

            def _fake(failed_count, **kwargs):
                seen["failed_count"] = failed_count
                seen["kwargs"] = kwargs
                return True

            oc.ask_retry_failed_files_impl = _fake
            dummy = OfficeConverter.__new__(OfficeConverter)
            dummy.error_records = ["x"]

            out = dummy.ask_retry_failed_files(3, timeout=5)
            self.assertTrue(out)
            self.assertEqual(seen["failed_count"], 3)
            self.assertEqual(seen["kwargs"]["timeout"], 5)
        finally:
            oc.ask_retry_failed_files_impl = original


if __name__ == "__main__":
    unittest.main()
