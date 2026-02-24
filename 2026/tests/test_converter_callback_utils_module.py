import unittest

from office_converter import OfficeConverter


class ConverterCallbackUtilsSplitTests(unittest.TestCase):
    def test_callback_utils_core_behaviors(self):
        from converter.callback_utils import emit_file_done, emit_file_plan

        planned = []
        done = []

        emit_file_plan(lambda items: planned.extend(items), ["a", "b"])
        self.assertEqual(planned, ["a", "b"])

        emit_file_done(lambda rec: done.append(rec), {"x": 1})
        self.assertEqual(done, [{"x": 1}])

        # invalid callback / invalid record should be ignored
        emit_file_plan(None, ["x"])
        emit_file_done(lambda rec: done.append(rec), "not-dict")
        self.assertEqual(len(done), 1)

    def test_office_converter_callback_methods_delegate_to_module(self):
        dummy = OfficeConverter.__new__(OfficeConverter)
        planned = []
        done = []
        dummy.file_plan_callback = lambda items: planned.extend(items)
        dummy.file_done_callback = lambda rec: done.append(rec)

        dummy._emit_file_plan(["x", "y"])
        dummy._emit_file_done({"k": "v"})

        self.assertEqual(planned, ["x", "y"])
        self.assertEqual(done, [{"k": "v"}])


if __name__ == "__main__":
    unittest.main()
