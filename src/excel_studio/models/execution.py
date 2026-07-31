from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path

from excel_studio.models.result import TaskResult


@dataclass(slots=True)
class TaskProgress:
    overall_percent: int = 0
    current_percent: int = 0
    step: str = ""
    current_file: str = ""
    current_sheet: str = ""
    completed_files: int = 0
    total_files: int = 0
    processed_rows: int = 0
    total_rows: int = 0
    success_count: int = 0
    skipped_count: int = 0
    warning_count: int = 0
    failure_count: int = 0
    elapsed_seconds: float = 0.0
    eta_seconds: float | None = None
    rows_per_second: float = 0.0


@dataclass(slots=True)
class TaskExecutionSummary:
    task_id: str
    operation: str
    status: str
    started_at: datetime
    finished_at: datetime
    output_directory: Path
    result: TaskResult
    report_directory: Path | None = None
    log_messages: list[str] = field(default_factory=list)

    @property
    def duration_seconds(self) -> float:
        return max(0.0, (self.finished_at - self.started_at).total_seconds())
