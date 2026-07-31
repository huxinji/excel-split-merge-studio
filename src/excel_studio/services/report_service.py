from __future__ import annotations

import json
from dataclasses import asdict, is_dataclass
from datetime import datetime
from enum import Enum
from pathlib import Path
from typing import Any

import pandas as pd

from excel_studio.models.execution import TaskExecutionSummary


def to_jsonable(value: Any) -> Any:
    if is_dataclass(value):
        return {key: to_jsonable(item) for key, item in asdict(value).items()}
    if isinstance(value, Path):
        return str(value)
    if isinstance(value, datetime):
        return value.isoformat()
    if isinstance(value, Enum):
        return value.value
    if isinstance(value, dict):
        return {str(key): to_jsonable(item) for key, item in value.items()}
    if isinstance(value, (list, tuple, set)):
        return [to_jsonable(item) for item in value]
    return value


def write_task_reports(summary: TaskExecutionSummary, config: Any) -> Path:
    report_directory = summary.output_directory / "reports" / summary.task_id
    report_directory.mkdir(parents=True, exist_ok=True)
    reconciliation = summary.result.reconciliation
    payload = {
        "task_id": summary.task_id,
        "operation": summary.operation,
        "status": summary.status,
        "started_at": summary.started_at.isoformat(),
        "finished_at": summary.finished_at.isoformat(),
        "duration_seconds": summary.duration_seconds,
        "configuration": to_jsonable(config),
        "output_files": [str(path) for path in summary.result.output_files],
        "warnings": summary.result.warnings,
        "errors": summary.result.errors,
        "reconciliation": to_jsonable(reconciliation),
    }
    json_path = report_directory / "task-report.json"
    json_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    log_path = report_directory / "task.log"
    log_path.write_text("\n".join(summary.log_messages) + "\n", encoding="utf-8")
    excel_path = report_directory / "task-report.xlsx"
    summary_rows = [
        ("Task ID", summary.task_id),
        ("Operation", summary.operation),
        ("Status", summary.status),
        ("Started", summary.started_at.isoformat()),
        ("Finished", summary.finished_at.isoformat()),
        ("Duration seconds", summary.duration_seconds),
        ("Output files", len(summary.result.output_files)),
        ("Warnings", len(summary.result.warnings)),
        ("Errors", len(summary.result.errors)),
    ]
    with pd.ExcelWriter(excel_path, engine="xlsxwriter") as writer:
        pd.DataFrame(summary_rows, columns=["Metric", "Value"]).to_excel(
            writer, sheet_name="Summary", index=False
        )
        input_files = payload["configuration"].get("input_files", [])
        pd.DataFrame({"Path": input_files}).to_excel(writer, sheet_name="Input Files", index=False)
        pd.DataFrame({"Path": payload["output_files"]}).to_excel(
            writer, sheet_name="Output Files", index=False
        )
        pd.DataFrame({"Warning": summary.result.warnings}).to_excel(
            writer, sheet_name="Warnings", index=False
        )
        pd.DataFrame({"Error": summary.result.errors}).to_excel(
            writer, sheet_name="Errors", index=False
        )
        mapping = payload["configuration"].get("field_mapping", {})
        pd.DataFrame(
            [{"Source": source, "Target": target} for source, target in mapping.items()]
        ).to_excel(writer, sheet_name="Field Mapping", index=False)
        if reconciliation is None:
            reconciliation_rows: list[dict[str, Any]] = []
        else:
            reconciliation_rows = [
                {
                    "Input Rows": reconciliation.input_rows,
                    "Output Rows": reconciliation.output_rows,
                    "Excluded Rows": reconciliation.excluded_rows,
                    "Balanced": reconciliation.is_balanced,
                    "Warning": reconciliation.warning,
                }
            ]
        pd.DataFrame(reconciliation_rows).to_excel(
            writer, sheet_name="Row Reconciliation", index=False
        )
    return report_directory
