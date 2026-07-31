from __future__ import annotations

import json
from pathlib import Path
from typing import Any

from excel_studio.models.execution import TaskExecutionSummary


class HistoryStore:
    def __init__(self, config_directory: Path, limit: int = 10) -> None:
        self.path = config_directory / "task-history.json"
        self.limit = limit

    def load(self) -> list[dict[str, Any]]:
        if not self.path.exists():
            return []
        try:
            value = json.loads(self.path.read_text(encoding="utf-8"))
            return value if isinstance(value, list) else []
        except (OSError, ValueError):
            return []

    def add(self, summary: TaskExecutionSummary) -> None:
        records = self.load()
        records.insert(
            0,
            {
                "task_id": summary.task_id,
                "operation": summary.operation,
                "status": summary.status,
                "started_at": summary.started_at.isoformat(),
                "finished_at": summary.finished_at.isoformat(),
                "duration_seconds": summary.duration_seconds,
                "output_directory": str(summary.output_directory),
                "report_directory": str(summary.report_directory or ""),
                "output_files": [str(path) for path in summary.result.output_files],
                "warnings": len(summary.result.warnings),
                "errors": len(summary.result.errors),
            },
        )
        self.path.parent.mkdir(parents=True, exist_ok=True)
        temporary = self.path.with_suffix(".tmp")
        temporary.write_text(
            json.dumps(records[: self.limit], ensure_ascii=False, indent=2), encoding="utf-8"
        )
        temporary.replace(self.path)
