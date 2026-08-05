import unittest
import json
from unittest.mock import patch
from urllib.error import HTTPError

from update_service import ReleaseInfo, fetch_latest_release, is_newer_version, version_key


class FakeRedirectResponse:
    def __enter__(self):
        return self

    def __exit__(self, *_args):
        return False

    @staticmethod
    def geturl():
        return "https://gitee.com/zhang-jiaxin654/qunzhong-toupiao/releases"

    @staticmethod
    def read():
        return b'<a href="/zhang-jiaxin654/qunzhong-toupiao/releases/tag/v0.3.4">v0.3.4</a>'


class FakeApiResponse:
    def __enter__(self):
        return self

    def __exit__(self, *_args):
        return False

    @staticmethod
    def read():
        return json.dumps(
            {
                "tag_name": "v0.3.4",
                "name": "群众投票 v0.3.4",
                "body": "更新源迁移到 Gitee",
                "created_at": "2026-08-05T12:00:00+08:00",
            }
        ).encode("utf-8")


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

    @patch("update_service.urlopen", return_value=FakeApiResponse())
    def test_gitee_api_constructs_release_page_url(self, _mocked_urlopen):
        release = fetch_latest_release(timeout=1)
        self.assertEqual(release.version, "0.3.4")
        self.assertEqual(
            release.html_url,
            "https://gitee.com/zhang-jiaxin654/qunzhong-toupiao/releases/tag/v0.3.4",
        )
        self.assertEqual(release.published_at, "2026-08-05T12:00:00+08:00")

    @patch(
        "update_service.urlopen",
        side_effect=[
            HTTPError("https://gitee.com/api/v5/test", 403, "rate limited", {}, None),
            FakeRedirectResponse(),
        ],
    )
    def test_rate_limit_falls_back_to_latest_release_page(self, mocked_urlopen):
        release = fetch_latest_release(timeout=1)
        self.assertEqual(release.version, "0.3.4")
        self.assertEqual(mocked_urlopen.call_count, 2)


if __name__ == "__main__":
    unittest.main()
