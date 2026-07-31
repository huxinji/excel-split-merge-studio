from __future__ import annotations

import unittest

from excel_studio.services.header_detector import detect_header


class HeaderDetectorTests(unittest.TestCase):
    def test_detects_header_after_report_title(self) -> None:
        rows = [
            ["Quarterly report", None, None],
            ["ID", "Name", "Date"],
            [1, "A", "2026-01-01"],
            [2, "B", "2026-01-02"],
        ]
        result = detect_header(rows)
        self.assertEqual(result.row_number, 2)
        self.assertGreaterEqual(result.confidence, 0.8)

    def test_empty_input_has_zero_confidence(self) -> None:
        result = detect_header([])
        self.assertEqual(result.row_number, 1)
        self.assertEqual(result.confidence, 0.0)
