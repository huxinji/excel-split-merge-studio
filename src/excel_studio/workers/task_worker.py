from __future__ import annotations

import threading
import time
import uuid
from datetime import datetime
from typing import Any

from PySide6.QtCore import QObject, Signal, Slot

from excel_studio.models.advanced import AdvancedMergeTaskConfig
from excel_studio.models.execution import TaskExecutionSummary, TaskProgress
from excel_studio.models.operations import MergeTaskConfig, SplitTaskConfig
from excel_studio.models.result import TaskResult
from excel_studio.services.history_service import HistoryStore
from excel_studio.services.pro_engine import (
    EngineHooks,
    execute_pro_advanced_merge,
    execute_pro_merge,
    execute_pro_split,
)
from excel_studio.services.report_service import write_task_reports


class TaskCancelled(BaseException):
    """Cooperative cancellation signal that is not swallowed by per-file error handling."""


class TaskControl:
    def __init__(self) -> None:
        self._condition = threading.Condition()
        self._paused = False
        self._cancelled = False

    @property
    def paused(self) -> bool:
        with self._condition:
            return self._paused

    @property
    def cancelled(self) -> bool:
        with self._condition:
            return self._cancelled

    def pause(self) -> None:
        with self._condition:
            self._paused = True

    def resume(self) -> None:
        with self._condition:
            self._paused = False
            self._condition.notify_all()

    def cancel(self) -> None:
        with self._condition:
            self._cancelled = True
            self._paused = False
            self._condition.notify_all()

    def checkpoint(self) -> None:
        with self._condition:
            while self._paused and not self._cancelled:
                self._condition.wait(timeout=0.25)
            if self._cancelled:
                raise TaskCancelled("Task cancelled at a safe checkpoint")


class ProcessingWorker(QObject):
    progress = Signal(object)
    log = Signal(str)
    completed = Signal(object)
    cancelled = Signal(object)
    failed = Signal(str)

    def __init__(self, config: Any, history_store: HistoryStore) -> None:
        super().__init__()
        self.config = config
        self.history_store = history_store
        self.control = TaskControl()
        self.task_id = datetime.now().strftime("%Y%m%d-%H%M%S-") + uuid.uuid4().hex[:6]
        self._logs: list[str] = []
        self._started_monotonic = 0.0

    def pause(self) -> None:
        self.control.pause()
        self._write_log("Task paused; waiting for the next safe checkpoint")

    def resume(self) -> None:
        self.control.resume()
        self._write_log("Task resumed")

    def cancel(self) -> None:
        self.control.cancel()
        self._write_log("Cancellation requested")

    def _write_log(self, message: str) -> None:
        entry = f"{datetime.now().isoformat(timespec='seconds')} | {message}"
        self._logs.append(entry)
        self.log.emit(entry)

    def _emit_progress(self, percent: int, step: str, current_file: str = "") -> None:
        elapsed = max(0.0, time.monotonic() - self._started_monotonic)
        total_files = len(getattr(self.config, "input_files", []))
        completed_files = (
            total_files
            if percent >= 80
            else min(total_files, int(total_files * max(0, percent - 15) / 65))
        )
        self.progress.emit(
            TaskProgress(
                overall_percent=max(0, min(100, percent)),
                current_percent=max(0, min(100, percent)),
                step=step,
                current_file=current_file,
                completed_files=completed_files,
                total_files=total_files,
                elapsed_seconds=elapsed,
                eta_seconds=(elapsed * (100 - percent) / percent) if percent else None,
            )
        )

    def _summary(
        self,
        operation: str,
        status: str,
        started_at: datetime,
        result: TaskResult,
    ) -> TaskExecutionSummary:
        return TaskExecutionSummary(
            task_id=self.task_id,
            operation=operation,
            status=status,
            started_at=started_at,
            finished_at=datetime.now(),
            output_directory=self.config.output.directory,
            result=result,
            log_messages=list(self._logs),
        )

    @Slot()
    def run(self) -> None:
        started_at = datetime.now()
        self._started_monotonic = time.monotonic()
        operation = "split" if isinstance(self.config, SplitTaskConfig) else "merge"
        hooks = EngineHooks(self.control, self._write_log, self._emit_progress)
        try:
            self._write_log(f"Task {self.task_id} started ({operation})")
            self._emit_progress(5, "Validating configuration")
            self.control.checkpoint()
            first_file = str(self.config.input_files[0]) if self.config.input_files else ""
            self._emit_progress(15, "Reading workbooks", first_file)
            if isinstance(self.config, SplitTaskConfig):
                result = execute_pro_split(self.config, hooks)
            elif isinstance(self.config, AdvancedMergeTaskConfig):
                result = execute_pro_advanced_merge(self.config, hooks)
            elif isinstance(self.config, MergeTaskConfig):
                result = execute_pro_merge(self.config, hooks)
            else:
                raise TypeError(f"Unsupported task configuration: {type(self.config).__name__}")
            self._emit_progress(82, "Reconciling rows")
            self.control.checkpoint()
            status = "success"
            if result.errors and result.output_files:
                status = "partial"
            elif result.errors:
                status = "failed"
            summary = self._summary(operation, status, started_at, result)
            self._emit_progress(92, "Writing task reports")
            self._write_log(
                f"Engine finished with {len(result.output_files)} outputs, "
                f"{len(result.warnings)} warnings, {len(result.errors)} errors"
            )
            try:
                summary.report_directory = write_task_reports(summary, self.config)
            except Exception as error:
                result.warnings.append(f"Report generation failed: {error}")
                self._write_log(f"Report generation failed: {error}")
            summary.finished_at = datetime.now()
            summary.log_messages = list(self._logs)
            self.history_store.add(summary)
            self._emit_progress(100, "Completed")
            self._write_log(f"Task finished with status {status}")
            self.completed.emit(summary)
        except TaskCancelled:
            summary = self._summary(operation, "cancelled", started_at, TaskResult())
            self._write_log("Task cancelled safely")
            try:
                summary.report_directory = write_task_reports(summary, self.config)
            except Exception as error:
                self._write_log(f"Cancellation report failed: {error}")
            summary.log_messages = list(self._logs)
            self.history_store.add(summary)
            self.cancelled.emit(summary)
        except Exception as error:
            self._write_log(f"Fatal task error: {error}")
            self.failed.emit(str(error))
