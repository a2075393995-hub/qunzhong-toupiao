from __future__ import annotations

import json
import tempfile
import unittest
from pathlib import Path

from template_profiles import (
    load_template_profile,
    mapping_area_counts,
    rebuild_pending_pair_index,
    save_template_profile,
    template_profile_key,
)
from vote_core import blank_mapping, configured_option_keys


class TemplateProfileTests(unittest.TestCase):
    def setUp(self):
        self.temporary_directory = tempfile.TemporaryDirectory()
        self.root = Path(self.temporary_directory.name)
        self.store = self.root / "profiles.json"

    def tearDown(self):
        self.temporary_directory.cleanup()

    def template(self, name: str, content: bytes = b"template") -> Path:
        path = self.root / name
        path.write_bytes(content)
        return path

    @staticmethod
    def marked_mapping():
        mapping = blank_mapping()
        mapping["filenamePrefix"] = "已保存-"
        mapping["options"]["结果1:同意"] = {
            "label": "同意",
            "pairs": [
                {
                    "judgment": {"table": 0, "row": 1, "col": 0},
                    "mark": {"table": 0, "row": 1, "col": 1},
                    "markStyle": {"horizontal": "center", "vertical": "middle", "fontSize": 12},
                },
                {"judgment": {"table": 0, "row": 2, "col": 0}},
            ],
        }
        return mapping

    def test_saved_profile_restores_all_option_areas(self):
        template = self.template("vote.docx")
        mapping = self.marked_mapping()
        self.assertTrue(save_template_profile(template, mapping, self.store))
        restored = load_template_profile(template, self.store)
        self.assertEqual(configured_option_keys(mapping), configured_option_keys(restored))
        self.assertEqual(mapping["options"]["结果1:同意"]["pairs"], restored["options"]["结果1:同意"]["pairs"])
        self.assertEqual((2, 1), mapping_area_counts(restored))

    def test_same_content_template_copy_uses_same_profile(self):
        first = self.template("first.docx", b"same-template")
        copy = self.template("copy.docx", b"same-template")
        mapping = self.marked_mapping()
        save_template_profile(first, mapping, self.store)
        self.assertEqual(template_profile_key(first), template_profile_key(copy))
        self.assertEqual(mapping, load_template_profile(copy, self.store))

    def test_changed_template_does_not_receive_old_coordinates(self):
        template = self.template("vote.docx", b"version-one")
        original_key = template_profile_key(template)
        save_template_profile(template, self.marked_mapping(), self.store, profile_key=original_key)
        template.write_bytes(b"version-two")
        save_template_profile(template, self.marked_mapping(), self.store, profile_key=original_key)
        self.assertIsNone(load_template_profile(template, self.store))

    def test_corrupted_json_is_ignored_and_can_be_replaced(self):
        template = self.template("vote.docx")
        self.store.write_text("{not valid json", encoding="utf-8")
        self.assertIsNone(load_template_profile(template, self.store))
        self.assertTrue(save_template_profile(template, self.marked_mapping(), self.store))
        self.assertIsInstance(json.loads(self.store.read_text(encoding="utf-8")), dict)

    def test_multiple_template_profiles_do_not_overwrite_each_other(self):
        first = self.template("first.docx", b"first")
        second = self.template("second.docx", b"second")
        first_mapping = self.marked_mapping()
        second_mapping = blank_mapping()
        second_mapping["fieldTargets"]["name"] = {"table": 1, "row": 2, "col": 3}
        save_template_profile(first, first_mapping, self.store)
        save_template_profile(second, second_mapping, self.store)
        self.assertEqual(first_mapping, load_template_profile(first, self.store))
        self.assertEqual(second_mapping, load_template_profile(second, self.store))

    def test_reselecting_same_template_restores_latest_session_mapping(self):
        template = self.template("vote.docx")
        current_mapping = self.marked_mapping()
        save_template_profile(template, current_mapping, self.store)
        reselected_mapping = load_template_profile(template, self.store)
        self.assertEqual(current_mapping, reselected_mapping)

    def test_pending_pair_index_is_rebuilt_after_restore(self):
        mapping = self.marked_mapping()
        self.assertEqual({"结果1:同意": 1}, rebuild_pending_pair_index(mapping))


if __name__ == "__main__":
    unittest.main()
