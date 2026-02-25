import os
import tempfile
import unittest
from pathlib import Path

from office_converter import OfficeConverter


class ConverterMshelpTopicsSplitTests(unittest.TestCase):
    def test_mshelp_topics_core_behaviors_without_bs4(self):
        from converter.mshelp_topics import parse_mshelp_topics

        root = tempfile.mkdtemp(prefix="mshelp_topics_")
        p1 = os.path.join(root, "z.html")
        p2 = os.path.join(root, "a.htm")
        for p in (p1, p2):
            with open(p, "w", encoding="utf-8") as f:
                f.write("<html><head><title>T</title></head><body>x</body></html>")

        topics = parse_mshelp_topics(
            root,
            find_files_recursive_fn=lambda d, exts: [p1, p2],
            has_bs4=False,
        )
        self.assertEqual(len(topics), 2)
        self.assertEqual(topics[0]["id"], "a.htm")
        self.assertEqual(topics[1]["id"], "z.html")
        self.assertEqual(topics[0]["children"], [])

        for p in (p1, p2):
            try:
                os.remove(p)
            except Exception:
                pass
        try:
            os.rmdir(root)
        except Exception:
            pass

    def test_office_converter_parse_mshelp_topics_delegates_to_module(self):
        from converter.mshelp_topics import parse_mshelp_topics
        import office_converter as oc

        root = tempfile.mkdtemp(prefix="mshelp_topics_delegate_")
        p = os.path.join(root, "a.html")
        with open(p, "w", encoding="utf-8") as f:
            f.write("<html><head><title>A</title></head><body>x</body></html>")

        dummy = OfficeConverter.__new__(OfficeConverter)
        dummy._find_files_recursive = lambda d, exts: [p]
        dummy._meta_content_by_names = lambda soup, names: ""

        expected = parse_mshelp_topics(
            root,
            find_files_recursive_fn=dummy._find_files_recursive,
            has_bs4=oc.HAS_BS4,
            beautifulsoup_cls=getattr(oc, "BeautifulSoup", None) if oc.HAS_BS4 else None,
            meta_content_by_names_fn=dummy._meta_content_by_names,
        )
        actual = dummy._parse_mshelp_topics(root)
        self.assertEqual(actual, expected)

        try:
            os.remove(p)
            os.rmdir(root)
        except Exception:
            pass

    def test_mshelp_topics_module_has_no_bare_except_exception(self):
        mod_path = Path(__file__).resolve().parents[1] / "converter" / "mshelp_topics.py"
        text = mod_path.read_text(encoding="utf-8")
        self.assertNotIn("except Exception", text)


if __name__ == "__main__":
    unittest.main()
