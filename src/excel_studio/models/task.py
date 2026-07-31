from __future__ import annotations

from dataclasses import dataclass, field
from enum import StrEnum
from pathlib import Path


class Operation(StrEnum):
    SPLIT = "split"
    MERGE = "merge"


class HeaderMode(StrEnum):
    AUTO = "auto"
    ROW_NUMBER = "row_number"
    MULTI_ROW = "multi_row"
    NONE = "none"


@dataclass(slots=True)
class TableStructure:
    header_mode: HeaderMode = HeaderMode.AUTO
    header_row: int = 1
    header_end_row: int = 1
    data_start_row: int = 2
    data_end_row: int | None = None
    data_start_column: int = 1
    data_end_column: int | None = None
    skip_bottom_rows: int = 0
    drop_empty_rows: bool = True
    drop_empty_columns: bool = True


@dataclass(slots=True)
class ScanOptions:
    scan_subfolders: bool = False
    ignore_hidden: bool = True
    ignore_temp_files: bool = True
    name_contains: str = ""
    name_excludes: str = ""
    include_text_files: bool = False


@dataclass(slots=True)
class TaskDraft:
    operation: Operation
    input_paths: list[Path] = field(default_factory=list)
    selected_sheets: dict[str, list[str]] = field(default_factory=dict)
    structure: TableStructure = field(default_factory=TableStructure)
    output_directory: Path | None = None
