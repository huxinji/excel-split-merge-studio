from __future__ import annotations

from pathlib import Path

from PySide6.QtCore import QObject, Signal, Slot

from excel_studio.models.operations import BlankValuePolicy
from excel_studio.models.task import TableStructure
from excel_studio.services.hierarchy_split import (
    HierarchyCatalog,
    build_hierarchy_catalog,
)
from excel_studio.services.table_reader import read_sheet_frame


class HierarchyValuesWorker(QObject):
    completed = Signal(int, object)
    failed = Signal(int, str)

    def __init__(
        self,
        request_id: int,
        path: Path,
        sheet_name: str,
        structure: TableStructure,
        hierarchy_fields: list[str],
        hierarchy_filters: dict[str, str | list[str]],
        split_field: str,
        empty_label: str,
        blank_policy: BlankValuePolicy,
    ) -> None:
        super().__init__()
        self._request_id = request_id
        self._path = path
        self._sheet_name = sheet_name
        self._structure = structure
        self._hierarchy_fields = hierarchy_fields
        self._hierarchy_filters = hierarchy_filters
        self._split_field = split_field
        self._empty_label = empty_label
        self._blank_policy = blank_policy

    @Slot()
    def run(self) -> None:
        try:
            frame = read_sheet_frame(self._path, self._sheet_name, self._structure)
            catalog: HierarchyCatalog = build_hierarchy_catalog(
                frame,
                self._hierarchy_fields,
                self._hierarchy_filters,
                self._split_field,
                self._empty_label,
                self._blank_policy,
            )
            self.completed.emit(self._request_id, catalog)
        except Exception as error:
            self.failed.emit(self._request_id, str(error))
