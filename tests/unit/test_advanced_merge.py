from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

import pandas as pd
from openpyxl import Workbook

from excel_studio.models.advanced import (
    AdvancedMergeTaskConfig,
    CleaningOptions,
    DuplicateKeyPolicy,
    JoinType,
)
from excel_studio.models.operations import MergeMode, OutputOptions
from excel_studio.models.task import HeaderMode, TableStructure
from excel_studio.services.advanced_merge_engine import (
    clean_frame,
    execute_advanced_merge,
    horizontal_concat,
    key_join,
    normalize_field_name,
    suggest_field_mapping,
)


class AdvancedMergeTests(unittest.TestCase):
    def test_horizontal_concat_pads_rows_and_renames_duplicates(self) -> None:
        first = pd.DataFrame({"ID": [1, 2], "Value": [10, 20]})
        second = pd.DataFrame({"ID": [1], "Status": ["ok"]})
        merged = horizontal_concat([first, second])
        self.assertEqual(list(merged.columns), ["ID", "Value", "ID_Source2", "Status"])
        self.assertEqual(len(merged), 2)
        self.assertTrue(pd.isna(merged.loc[1, "Status"]))

    def test_left_join_reports_unmatched_rows(self) -> None:
        left = pd.DataFrame({"ID": [1, 2], "Name": ["A", "B"]})
        right = pd.DataFrame({"Code": [1, 3], "Amount": [10, 30]})
        merged, stats = key_join(left, right, ["ID"], ["Code"], JoinType.LEFT)
        self.assertEqual(len(merged), 2)
        self.assertEqual(stats["matched"], 1)
        self.assertEqual(stats["left_only"], 1)

    def test_full_outer_join_keeps_both_sides(self) -> None:
        left = pd.DataFrame({"ID": [1, 2]})
        right = pd.DataFrame({"ID": [2, 3], "Amount": [20, 30]})
        merged, stats = key_join(left, right, ["ID"], ["ID"], JoinType.OUTER)
        self.assertEqual(len(merged), 3)
        self.assertEqual(stats["right_only"], 1)

    def test_duplicate_policy_rejects_or_keeps_first(self) -> None:
        left = pd.DataFrame({"ID": [1], "Name": ["A"]})
        right = pd.DataFrame({"ID": [1, 1], "Amount": [10, 20]})
        with self.assertRaises(ValueError):
            key_join(
                left,
                right,
                ["ID"],
                ["ID"],
                JoinType.LEFT,
                DuplicateKeyPolicy.REJECT,
            )
        merged, stats = key_join(
            left,
            right,
            ["ID"],
            ["ID"],
            JoinType.LEFT,
            DuplicateKeyPolicy.FIRST,
        )
        self.assertEqual(len(merged), 1)
        self.assertEqual(merged.loc[0, "Amount"], 10)
        self.assertEqual(stats["right_duplicates"], 2)

    def test_cleaning_and_mapping_suggestions_are_explicit(self) -> None:
        frame = pd.DataFrame({"Name": ["  A   B ", "  A   B "]})
        cleaned = clean_frame(frame, CleaningOptions(drop_duplicate_rows=True))
        self.assertEqual(cleaned.loc[0, "Name"], "A B")
        self.assertEqual(len(cleaned), 1)
        self.assertEqual(normalize_field_name("Ｎａｍｅ  （ID）"), "name (id)")
        suggestions = suggest_field_mapping(["Customer Name"], ["CustomerName"], 0.7)
        self.assertEqual(suggestions[0][:2], ("Customer Name", "CustomerName"))

    def test_execute_key_join_writes_output(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            master = root / "master.xlsx"
            related = root / "related.xlsx"
            for path, rows in (
                (master, [["ID", "Name"], [1, "A"], [2, "B"]]),
                (related, [["Code", "Amount"], [1, 10], [3, 30]]),
            ):
                workbook = Workbook()
                sheet = workbook.active
                for row in rows:
                    sheet.append(row)
                workbook.save(path)
            result = execute_advanced_merge(
                AdvancedMergeTaskConfig(
                    input_files=[master, related],
                    structure=TableStructure(
                        header_mode=HeaderMode.ROW_NUMBER, header_row=1, data_start_row=2
                    ),
                    mode=MergeMode.KEY_JOIN,
                    output=OutputOptions(
                        directory=root / "output", naming_template="joined_{row_count}"
                    ),
                    join_type=JoinType.LEFT,
                    left_keys=["ID"],
                    right_keys=["Code"],
                )
            )
            self.assertFalse(result.errors)
            self.assertEqual(len(result.output_files), 1)
            frame = pd.read_excel(result.output_files[0])
            self.assertEqual(len(frame), 2)
            self.assertEqual(frame.loc[0, "Amount"], 10)


if __name__ == "__main__":
    unittest.main()
