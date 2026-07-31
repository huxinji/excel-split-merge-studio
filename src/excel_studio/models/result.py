from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path


@dataclass(slots=True)
class ReconciliationResult:
    input_rows: int
    output_rows: int
    excluded_rows: int = 0
    warning: str = ""

    @property
    def is_balanced(self) -> bool:
        return self.input_rows == self.output_rows + self.excluded_rows


@dataclass(slots=True)
class TaskResult:
    output_files: list[Path] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    errors: list[str] = field(default_factory=list)
    reconciliation: ReconciliationResult | None = None
