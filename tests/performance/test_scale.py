from __future__ import annotations

import tempfile
import time
import unittest
from pathlib import Path

import pandas as pd
from openpyxl import Workbook

from excel_studio.models.task import ScanOptions
from excel_studio.services.file_scanner import scan_files
from excel_studio.services.split_engine import fixed_row_groups


class ScaleTests(unittest.TestCase):
    def test_million_row_partitioning_is_complete(self) -> None:
        frame = pd.DataFrame({"ID": range(1_000_000)})
        started = time.monotonic()
        groups = fixed_row_groups(frame, 100_000)
        elapsed = time.monotonic() - started
        self.assertEqual(len(groups), 10)
        self.assertEqual(sum(len(group) for group in groups), 1_000_000)
        self.assertLess(elapsed, 20.0)

    def test_scans_ten_and_one_hundred_workbooks(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            seed = root / "seed.xlsx"
            workbook = Workbook()
            workbook.active.append(["ID", "Value"])
            workbook.active.append([1, "A"])
            workbook.save(seed)
            payload = seed.read_bytes()
            seed.unlink()
            paths = []
            for index in range(100):
                path = root / f"book-{index:03d}.xlsx"
                path.write_bytes(payload)
                paths.append(path)
            ten = scan_files(paths[:10], ScanOptions())
            hundred = scan_files(paths, ScanOptions())
            self.assertEqual(len(ten), 10)
            self.assertEqual(len(hundred), 100)
            self.assertTrue(all(record.status == "ready" for record in hundred))


if __name__ == "__main__":
    unittest.main()
