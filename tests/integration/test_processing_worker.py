from __future__ import annotations

import os
import tempfile
import time
import unittest
from pathlib import Path

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from openpyxl import Workbook  # noqa: E402
from PySide6.QtCore import QThread  # noqa: E402
from PySide6.QtWidgets import QApplication  # noqa: E402

from excel_studio.models.operations import (  # noqa: E402
    OutputOptions,
    SplitMode,
    SplitOptions,
    SplitTaskConfig,
)
from excel_studio.models.task import HeaderMode, TableStructure  # noqa: E402
from excel_studio.services.history_service import HistoryStore  # noqa: E402
from excel_studio.workers.task_worker import ProcessingWorker  # noqa: E402


class ProcessingWorkerIntegrationTests(unittest.TestCase):
    def test_worker_runs_engine_and_reports_off_ui_thread(self) -> None:
        app = QApplication.instance() or QApplication([])
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            source = root / "source.xlsx"
            workbook = Workbook()
            sheet = workbook.active
            sheet.append(["ID", "Branch"])
            sheet.append([1, "A"])
            sheet.append([2, "B"])
            workbook.save(source)
            config = SplitTaskConfig(
                input_files=[source],
                structure=TableStructure(
                    header_mode=HeaderMode.ROW_NUMBER, header_row=1, data_start_row=2
                ),
                split=SplitOptions(mode=SplitMode.BY_FIELD, fields=["Branch"]),
                output=OutputOptions(directory=root / "output", naming_template="{split_value}"),
            )
            thread = QThread()
            worker = ProcessingWorker(config, HistoryStore(root / "config"))
            worker.moveToThread(thread)
            summaries = []
            failures = []
            worker.completed.connect(summaries.append)
            worker.failed.connect(failures.append)
            worker.completed.connect(thread.quit)
            worker.failed.connect(thread.quit)
            thread.started.connect(worker.run)
            thread.start()
            deadline = time.monotonic() + 10
            while thread.isRunning() and time.monotonic() < deadline:
                app.processEvents()
                time.sleep(0.01)
            thread.wait(1000)
            self.assertFalse(failures)
            self.assertEqual(len(summaries), 1)
            self.assertEqual(len(summaries[0].result.output_files), 2)
            self.assertTrue((summaries[0].report_directory / "task-report.json").exists())


if __name__ == "__main__":
    unittest.main()
