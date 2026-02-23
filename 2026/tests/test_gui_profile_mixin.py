import unittest

from gui_profile_mixin import ProfileManagementMixin


class _DummyProfileMixin(ProfileManagementMixin):
    pass


class ProfileManagementMixinTests(unittest.TestCase):
    def test_sanitize_profile_stem_replaces_invalid_chars(self):
        mixin = _DummyProfileMixin()
        out = mixin._sanitize_profile_stem('prod:/\\*?"<>| config.')
        self.assertEqual("prod_ config", out)


if __name__ == "__main__":
    unittest.main()
