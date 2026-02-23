import ast
import unittest
from pathlib import Path


class OfficeGUIEntrypointTests(unittest.TestCase):
    def test_office_gui_has_main_entrypoint_with_mainloop(self):
        source_path = Path(__file__).resolve().parent.parent / "office_gui.py"
        source = source_path.read_text(encoding="utf-8")
        tree = ast.parse(source)

        main_guard = None
        for node in tree.body:
            if isinstance(node, ast.If) and self._is_main_guard(node.test):
                main_guard = node
                break

        self.assertIsNotNone(
            main_guard,
            "office_gui.py must define `if __name__ == \"__main__\":` entrypoint",
        )

        has_mainloop_call = any(
            isinstance(sub_node, ast.Call)
            and isinstance(sub_node.func, ast.Attribute)
            and sub_node.func.attr == "mainloop"
            for sub_node in ast.walk(main_guard)
        )
        self.assertTrue(
            has_mainloop_call,
            "office_gui.py entrypoint must call mainloop() to keep the GUI window alive",
        )

    def test_office_gui_initializes_hover_tip_class(self):
        source_path = Path(__file__).resolve().parent.parent / "office_gui.py"
        source = source_path.read_text(encoding="utf-8")
        self.assertIn(
            "self._hover_tip_cls = HoverTip",
            source,
            "OfficeGUI.__init__ must assign _hover_tip_cls for TooltipMixin",
        )

    @staticmethod
    def _is_main_guard(test_node: ast.AST) -> bool:
        if not isinstance(test_node, ast.Compare):
            return False
        if len(test_node.ops) != 1 or len(test_node.comparators) != 1:
            return False
        if not isinstance(test_node.ops[0], ast.Eq):
            return False
        if not isinstance(test_node.left, ast.Name) or test_node.left.id != "__name__":
            return False
        comparator = test_node.comparators[0]
        return isinstance(comparator, ast.Constant) and comparator.value == "__main__"


if __name__ == "__main__":
    unittest.main()
