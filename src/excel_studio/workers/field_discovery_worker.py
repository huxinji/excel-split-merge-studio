from __future__ import annotations

from PySide6.QtCore import QObject, Signal, Slot

from excel_studio.models.field_catalog import FieldCatalog, FieldDiscoveryRequest
from excel_studio.services.field_discovery import discover_fields


class FieldDiscoveryWorker(QObject):
    completed = Signal(object)
    failed = Signal(str)

    def __init__(
        self,
        requests: list[FieldDiscoveryRequest],
        header_mode: str,
        header_row: int,
        header_end_row: int,
        preview_rows: int,
    ) -> None:
        super().__init__()
        self._requests = requests
        self._header_mode = header_mode
        self._header_row = header_row
        self._header_end_row = header_end_row
        self._preview_rows = preview_rows

    @Slot()
    def run(self) -> None:
        try:
            catalog: FieldCatalog = discover_fields(
                self._requests,
                self._header_mode,
                self._header_row,
                self._header_end_row,
                self._preview_rows,
            )
            self.completed.emit(catalog)
        except Exception as error:
            self.failed.emit(str(error))
