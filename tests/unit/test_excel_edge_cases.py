from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook

from excel_studio.models.task import HeaderMode, TableStructure
from excel_studio.services.table_reader import read_sheet_frame
from excel_studio.services.workbook_inspector import build_file_record


class ExcelEdgeCaseTests(unittest.TestCase):
    def test_multiline_header_formula_and_xlsm_are_read(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            path = Path(temporary) / "macro-like.xlsm"
            workbook = Workbook()
            sheet = workbook.active
            sheet.title = "Data"
            sheet.append(["Customer", "Financial"])
            sheet.append(["ID", "Amount"])
            sheet.append([1, "=1+2"])
            workbook.save(path)
            record = build_file_record(path)
            self.assertEqual(record.status, "ready")
            frame = read_sheet_frame(
                path,
                "Data",
                TableStructure(
                    header_mode=HeaderMode.MULTI_ROW,
                    header_row=1,
                    header_end_row=2,
                    data_start_row=3,
                ),
            )
            self.assertEqual(list(frame.columns), ["Customer_ID", "Financial_Amount"])
            self.assertEqual(frame.iloc[0, 1], "=1+2")


if __name__ == "__main__":
    unittest.main()
