from __future__ import annotations

from pathlib import Path

from PySide6.QtCore import QObject, Signal, Slot

from excel_studio.models.task import ScanOptions
from excel_studio.models.workbook import FileRecord, PreviewResult
from excel_studio.services.file_scanner import scan_files
from excel_studio.services.workbook_inspector import preview_workbook


class ScanWorker(QObject):
    progress = Signal(int, int, str)
    completed = Signal(object)
    failed = Signal(str)

    def __init__(self, paths: list[Path], options: ScanOptions) -> None:
        super().__init__()
        self._paths = paths
        self._options = options

    @Slot()
    def run(self) -> None:
        try:
            records: list[FileRecord] = scan_files(
                self._paths,
                self._options,
                lambda current, total, name: self.progress.emit(current, total, name),
            )
            self.completed.emit(records)
        except Exception as error:
            self.failed.emit(str(error))


class PreviewWorker(QObject):
    completed = Signal(object)
    failed = Signal(str)

    def __init__(
        self,
        path: Path,
        sheet_name: str | None,
        header_mode: str,
        header_row: int,
        header_end_row: int,
        preview_rows: int,
    ) -> None:
        super().__init__()
        self._path = path
        self._sheet_name = sheet_name
        self._header_mode = header_mode
        self._header_row = header_row
        self._header_end_row = header_end_row
        self._preview_rows = preview_rows

    @Slot()
    def run(self) -> None:
        try:
            result: PreviewResult = preview_workbook(
                self._path,
                self._sheet_name,
                self._header_mode,
                self._header_row,
                self._header_end_row,
                self._preview_rows,
            )
            self.completed.emit(result)
        except Exception as error:
            self.failed.emit(str(error))
