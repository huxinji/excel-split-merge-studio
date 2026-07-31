from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any


@dataclass(slots=True)
class SheetInfo:
    name: str
    state: str = "visible"
    max_row: int = 0
    max_column: int = 0
    is_empty: bool = True
    header_row: int | None = None
    header_confidence: float = 0.0


@dataclass(slots=True)
class FileRecord:
    path: Path
    size_bytes: int
    modified_at: datetime
    extension: str
    sheets: list[SheetInfo] = field(default_factory=list)
    status: str = "ready"
    warning: str = ""

    @property
    def name(self) -> str:
        return self.path.name


@dataclass(slots=True)
class HeaderDetectionResult:
    row_number: int
    confidence: float
    reasons: list[str] = field(default_factory=list)


@dataclass(slots=True)
class PreviewResult:
    path: Path
    sheet_name: str
    columns: list[str]
    rows: list[list[Any]]
    total_rows: int
    total_columns: int
    header: HeaderDetectionResult
    warnings: list[str] = field(default_factory=list)
