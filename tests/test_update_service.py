import unittest

from update_service import ReleaseInfo, is_newer_version, version_key


class VersionComparisonTests(unittest.TestCase):
    def test_newer_minor_version(self):
        self.assertTrue(is_newer_version("0.2.0", "0.1.9"))

    def test_v_prefix_is_ignored(self):
        self.assertFalse(is_newer_version("v1.2.3", "1.2.3"))

    def test_stable_is_newer_than_prerelease(self):
        self.assertGreater(version_key("1.0.0"), version_key("1.0.0-beta"))

    def test_release_version_uses_tag(self):
        release = ReleaseInfo("v2.3.4", "Release", "https://example.com", "", "")
        self.assertEqual(release.version, "2.3.4")


if __name__ == "__main__":
    unittest.main()
