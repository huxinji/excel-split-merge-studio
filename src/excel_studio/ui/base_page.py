from __future__ import annotations

from collections.abc import Callable
from pathlib import Path
from typing import Any

from PySide6.QtCore import QObject, Qt, QThread, Signal
from PySide6.QtGui import QDragEnterEvent, QDropEvent
from PySide6.QtWidgets import (
    QFrame,
    QHBoxLayout,
    QScrollArea,
    QSplitter,
    QVBoxLayout,
    QWidget,
)

from excel_studio.config.i18n import TranslationManager
from excel_studio.config.settings import AppSettings
from excel_studio.models.workbook import PreviewResult
from excel_studio.ui.widgets import LogPanel, PreviewPanel
from excel_studio.workers.scan_worker import PreviewWorker


class OperationPage(QWidget):
    configuration_changed = Signal()
    paths_dropped = Signal(object)

    def __init__(self, i18n: TranslationManager, settings: AppSettings) -> None:
        super().__init__()
        self.i18n = i18n
        self.settings = settings
        self.current_preview: PreviewResult | None = None
        self.preview_path: Path | None = None
        self.preview_sheet: str | None = None
        self._threads: set[QThread] = set()
        self._workers: set[QObject] = set()
        self.setAcceptDrops(True)

        root = QHBoxLayout(self)
        root.setContentsMargins(14, 12, 14, 10)
        root.setSpacing(0)
        self.splitter = QSplitter()
        self.splitter.setChildrenCollapsible(False)
        root.addWidget(self.splitter)

        self.left_scroll = QScrollArea()
        self.left_scroll.setWidgetResizable(True)
        self.left_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.left_content = QWidget()
        self.left_layout = QVBoxLayout(self.left_content)
        self.left_layout.setContentsMargins(2, 2, 8, 12)
        self.left_layout.setSpacing(12)
        self.left_scroll.setWidget(self.left_content)
        self.splitter.addWidget(self.left_scroll)

        self.right_panel = QFrame()
        self.right_panel.setMinimumWidth(365)
        self.right_panel.setMaximumWidth(520)
        right_layout = QVBoxLayout(self.right_panel)
        right_layout.setContentsMargins(0, 2, 2, 2)
        right_layout.setSpacing(10)
        self.preview_panel = PreviewPanel(i18n)
        self.log_panel = LogPanel(i18n)
        self.preview_panel.refresh_requested.connect(self.refresh_preview)
        right_layout.addWidget(self.preview_panel, 1)
        right_layout.addWidget(self.log_panel, 0)
        self.splitter.addWidget(self.right_panel)
        self.splitter.setStretchFactor(0, 5)
        self.splitter.setStretchFactor(1, 3)
        self.splitter.setSizes([820, 420])

    def add_card(self, card: QWidget) -> None:
        self.left_layout.addWidget(card)

    def finish_cards(self) -> None:
        self.left_layout.addStretch(1)

    def append_log(self, message: str) -> None:
        self.log_panel.append(message)

    def _run_worker(
        self,
        worker: Any,
        completed: Callable[[Any], None],
        failed: Callable[[str], None],
    ) -> None:
        thread = QThread(self)
        self._threads.add(thread)
        self._workers.add(worker)
        worker.moveToThread(thread)
        thread.started.connect(worker.run)
        worker.completed.connect(completed)
        worker.completed.connect(thread.quit)
        worker.failed.connect(failed)
        worker.failed.connect(thread.quit)
        thread.finished.connect(worker.deleteLater)
        thread.finished.connect(thread.deleteLater)
        thread.finished.connect(lambda current=thread: self._threads.discard(current))
        thread.finished.connect(lambda current=worker: self._workers.discard(current))
        thread.start()

    def set_preview_target(self, path: Path | None, sheet_name: str | None) -> None:
        self.preview_path = path
        self.preview_sheet = sheet_name

    def refresh_preview(self) -> None:
        if self.preview_path is None:
            self.append_log(self.i18n.text("log.no_preview_source"))
            return
        self.preview_panel.set_busy(True)
        worker = PreviewWorker(
            self.preview_path,
            self.preview_sheet,
            "row_number",
            self.header_row_value(),
            self.header_row_value(),
            self.settings.preview_rows,
        )
        self._run_worker(worker, self._preview_completed, self._preview_failed)

    def _preview_completed(self, result: PreviewResult) -> None:
        self.current_preview = result
        self.preview_panel.set_busy(False)
        self.preview_panel.set_preview(result)
        self.on_preview_ready(result)
        self.append_log(
            self.i18n.text(
                "log.preview_ready",
                file=result.path.name,
                sheet=result.sheet_name,
                rows=result.total_rows,
            )
        )

    def _preview_failed(self, message: str) -> None:
        self.preview_panel.set_busy(False)
        self.append_log(self.i18n.text("log.preview_failed", error=message))

    def on_preview_ready(self, result: PreviewResult) -> None:
        del result

    def header_row_value(self) -> int:
        return 1

    def retranslate_ui(self) -> None:
        self.preview_panel.retranslate_ui()
        self.log_panel.retranslate_ui()

    def dragEnterEvent(self, event: QDragEnterEvent) -> None:  # noqa: N802
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dropEvent(self, event: QDropEvent) -> None:  # noqa: N802
        paths = [Path(url.toLocalFile()) for url in event.mimeData().urls() if url.isLocalFile()]
        if paths:
            self.paths_dropped.emit(paths)
            event.acceptProposedAction()
