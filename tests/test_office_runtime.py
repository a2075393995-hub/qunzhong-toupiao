from __future__ import annotations

import os
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import office_runtime


class OfficeRuntimeTests(unittest.TestCase):
    def test_configured_runtime_accepts_office_root(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            soffice = Path(temp_dir) / "program" / "soffice.exe"
            soffice.parent.mkdir(parents=True)
            soffice.touch()
            with patch.dict(os.environ, {office_runtime.OFFICE_RUNTIME_ENV: temp_dir}, clear=False):
                self.assertEqual(office_runtime.find_bundled_soffice(), soffice)

    def test_bundled_runtime_is_preferred_over_system_runtime(self):
        bundled = Path("C:/portable/program/soffice.exe")
        system = Path("C:/installed/program/soffice.exe")
        with patch("office_runtime.find_bundled_soffice", return_value=bundled), patch(
            "office_runtime.find_system_soffice", return_value=system
        ):
            self.assertEqual(office_runtime.find_soffice(), bundled)

    def test_force_bundled_flag(self):
        with patch.dict(os.environ, {office_runtime.FORCE_BUNDLED_OFFICE_ENV: "1"}, clear=False):
            self.assertTrue(office_runtime.force_bundled_office())


if __name__ == "__main__":
    unittest.main()
