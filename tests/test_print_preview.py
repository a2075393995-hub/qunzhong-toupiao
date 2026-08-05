from __future__ import annotations

import tempfile
import subprocess
import unittest
from pathlib import Path
from unittest.mock import Mock, patch

import print_preview


class PrintPreviewBackendTests(unittest.TestCase):
    def setUp(self):
        self.temporary_directory = tempfile.TemporaryDirectory()
        self.root = Path(self.temporary_directory.name).resolve()
        self.source = (self.root / "template.docx").resolve()
        self.target = (self.root / "preview.pdf").resolve()
        self.source.write_bytes(b"docx")

    def tearDown(self):
        self.temporary_directory.cleanup()

    @staticmethod
    def successful_export(name, calls):
        def export(_source, target, *args):
            calls.append((name, *args))
            target.write_bytes(b"%PDF-test")
            return target

        return export

    def test_word_is_preferred_over_wps_and_bundled_engine(self):
        calls = []
        with patch("print_preview.force_bundled_office", return_value=False), patch(
            "print_preview.com_progid_registered", side_effect=lambda progid: progid == print_preview.WORD_PROGID
        ), patch("print_preview.find_wps_progid", return_value="KWPS.Application"), patch(
            "print_preview.export_with_word", side_effect=self.successful_export("word", calls)
        ), patch("print_preview.export_with_wps", side_effect=self.successful_export("wps", calls)), patch(
            "print_preview.export_with_bundled_engine", side_effect=self.successful_export("bundled", calls)
        ):
            result = print_preview.docx_to_pdf(self.source, self.target)
        self.assertEqual(self.target, result)
        self.assertEqual([("word",)], calls)

    def test_wps_is_preferred_when_word_is_unavailable(self):
        calls = []
        with patch("print_preview.force_bundled_office", return_value=False), patch(
            "print_preview.com_progid_registered", return_value=False
        ), patch("print_preview.find_wps_progid", return_value="KWPS.Application"), patch(
            "print_preview.export_with_wps", side_effect=self.successful_export("wps", calls)
        ), patch(
            "print_preview.export_with_bundled_engine", side_effect=self.successful_export("bundled", calls)
        ):
            result = print_preview.docx_to_pdf(self.source, self.target)
        self.assertEqual(self.target, result)
        self.assertEqual([("wps", "KWPS.Application")], calls)

    def test_word_failure_continues_with_wps(self):
        calls = []

        def failed_word(_source, target):
            calls.append(("word",))
            target.write_bytes(b"partial")
            raise RuntimeError("Word export failed")

        with patch("print_preview.force_bundled_office", return_value=False), patch(
            "print_preview.com_progid_registered", return_value=True
        ), patch("print_preview.find_wps_progid", return_value="KWPS.Application"), patch(
            "print_preview.export_with_word", side_effect=failed_word
        ), patch(
            "print_preview.export_with_wps", side_effect=self.successful_export("wps", calls)
        ), patch("print_preview.export_with_bundled_engine") as bundled:
            result = print_preview.docx_to_pdf(self.source, self.target)
        self.assertEqual(self.target, result)
        self.assertEqual([("word",), ("wps", "KWPS.Application")], calls)
        self.assertEqual(b"%PDF-test", self.target.read_bytes())
        bundled.assert_not_called()

    def test_wps_failure_falls_back_to_bundled_engine(self):
        calls = []

        def failed_wps(_source, target, progid):
            calls.append(("wps", progid))
            target.write_bytes(b"partial")
            raise RuntimeError("WPS export failed")

        with patch("print_preview.force_bundled_office", return_value=False), patch(
            "print_preview.com_progid_registered", return_value=False
        ), patch("print_preview.find_wps_progid", return_value="KWPS.Application"), patch(
            "print_preview.export_with_wps", side_effect=failed_wps
        ), patch(
            "print_preview.export_with_bundled_engine", side_effect=self.successful_export("bundled", calls)
        ):
            result = print_preview.docx_to_pdf(self.source, self.target)
        self.assertEqual(self.target, result)
        self.assertEqual([("wps", "KWPS.Application"), ("bundled",)], calls)
        self.assertEqual(b"%PDF-test", self.target.read_bytes())

    def test_force_bundled_engine_skips_word_and_wps(self):
        calls = []
        with patch("print_preview.force_bundled_office", return_value=True), patch(
            "print_preview.export_with_bundled_engine", side_effect=self.successful_export("bundled", calls)
        ), patch("print_preview.export_with_word") as word, patch("print_preview.export_with_wps") as wps:
            result = print_preview.docx_to_pdf(self.source, self.target)
        self.assertEqual(self.target, result)
        self.assertEqual([("bundled",)], calls)
        word.assert_not_called()
        wps.assert_not_called()

    def test_missing_wps_registration_does_not_block_bundled_engine(self):
        calls = []
        with patch("print_preview.force_bundled_office", return_value=False), patch(
            "print_preview.com_progid_registered", return_value=False
        ), patch("print_preview.find_wps_progid", return_value=None), patch(
            "print_preview.export_with_bundled_engine", side_effect=self.successful_export("bundled", calls)
        ):
            result = print_preview.docx_to_pdf(self.source, self.target)
        self.assertEqual(self.target, result)
        self.assertEqual([("bundled",)], calls)

    def test_com_timeout_terminates_helper_process_tree(self):
        process = Mock()
        process.wait.side_effect = subprocess.TimeoutExpired("office", 120)
        process.pid = 1234
        with patch("print_preview.subprocess.Popen", return_value=process), patch(
            "print_preview._terminate_process_tree"
        ) as terminate:
            with self.assertRaisesRegex(RuntimeError, "转换超时"):
                print_preview._export_with_com(self.source, self.target, "KWPS.Application")
        terminate.assert_called_once_with(process)


if __name__ == "__main__":
    unittest.main()
