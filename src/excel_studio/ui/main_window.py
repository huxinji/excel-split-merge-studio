from __future__ import annotations

from pathlib import Path

from PySide6.QtCore import QThread, QUrl
from PySide6.QtGui import QCloseEvent, QDesktopServices
from PySide6.QtWidgets import (
    QApplication,
    QComboBox,
    QFileDialog,
    QFrame,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QProgressBar,
    QPushButton,
    QStackedWidget,
    QVBoxLayout,
    QWidget,
)

from excel_studio.config.constants import VERSION
from excel_studio.config.i18n import TranslationManager
from excel_studio.config.settings import AppSettings, SettingsStore
from excel_studio.models.execution import TaskExecutionSummary, TaskProgress
from excel_studio.services.history_service import HistoryStore
from excel_studio.ui.merge_page import MergePage
from excel_studio.ui.split_page import SplitPage
from excel_studio.ui.theme import product_stylesheet
from excel_studio.workers.task_worker import ProcessingWorker


class MainWindow(QMainWindow):
    def __init__(
        self,
        settings: AppSettings,
        settings_store: SettingsStore,
        i18n: TranslationManager,
    ) -> None:
        super().__init__()
        self.settings = settings
        self.settings_store = settings_store
        self.i18n = i18n
        self.history_store = HistoryStore(settings_store.path.parent)
        self.task_thread: QThread | None = None
        self.task_worker: ProcessingWorker | None = None
        self._output_user_changed = False
        self._active_operation = "split"

        self.setObjectName("ProductWindow")
        self.setStyleSheet(product_stylesheet())
        self._build_ui()
        self._connect_signals()
        self.retranslate_ui()
        self._apply_responsive_geometry()
        self.switch_workspace("split")

    def _build_ui(self) -> None:
        root = QWidget()
        root.setObjectName("AppRoot")
        layout = QVBoxLayout(root)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)
        self.setCentralWidget(root)

        self.header = self._build_header()
        layout.addWidget(self.header)
        self.page_stack = QStackedWidget()
        self.split_page = SplitPage(self.i18n, self.settings)
        self.merge_page = MergePage(self.i18n, self.settings)
        self.page_stack.addWidget(self.split_page)
        self.page_stack.addWidget(self.merge_page)
        layout.addWidget(self.page_stack, 1)
        self.footer = self._build_footer()
        layout.addWidget(self.footer)

    def _build_header(self) -> QFrame:
        header = QFrame()
        header.setObjectName("ProductHeader")
        layout = QHBoxLayout(header)
        layout.setContentsMargins(20, 13, 20, 13)
        layout.setSpacing(12)

        title_box = QVBoxLayout()
        title_box.setSpacing(2)
        self.title_label = QLabel()
        self.title_label.setObjectName("ProductTitle")
        self.subtitle_label = QLabel()
        self.subtitle_label.setObjectName("ProductSubtitle")
        title_box.addWidget(self.title_label)
        title_box.addWidget(self.subtitle_label)
        layout.addLayout(title_box)
        self.version_label = QLabel(f"v{VERSION}")
        self.version_label.setObjectName("VersionBadge")
        layout.addWidget(self.version_label)
        layout.addStretch(1)

        switch = QFrame()
        switch.setObjectName("WorkspaceSwitch")
        switch_layout = QHBoxLayout(switch)
        switch_layout.setContentsMargins(3, 3, 3, 3)
        switch_layout.setSpacing(3)
        self.split_workspace_button = QPushButton()
        self.merge_workspace_button = QPushButton()
        for button in (self.split_workspace_button, self.merge_workspace_button):
            button.setObjectName("WorkspaceButton")
            button.setCheckable(True)
            switch_layout.addWidget(button)
        layout.addWidget(switch)

        self.language_combo = QComboBox()
        self.language_combo.setObjectName("LanguageCombo")
        self.language_combo.addItem("简体中文", "zh_CN")
        self.language_combo.addItem("English", "en_US")
        layout.addWidget(self.language_combo)
        return header

    def _build_footer(self) -> QFrame:
        footer = QFrame()
        footer.setObjectName("FooterBar")
        layout = QHBoxLayout(footer)
        layout.setContentsMargins(16, 10, 16, 10)
        layout.setSpacing(10)

        left = QVBoxLayout()
        left.setSpacing(6)
        output_row = QHBoxLayout()
        self.output_label = QLabel()
        self.output_directory_edit = QLineEdit()
        self.output_directory_edit.setClearButtonEnabled(True)
        self.output_browse_button = QPushButton()
        self.output_browse_button.setObjectName("SecondaryButton")
        self.open_output_button = QPushButton()
        self.open_output_button.setObjectName("SuccessButton")
        self.open_output_button.setEnabled(False)
        output_row.addWidget(self.output_label)
        output_row.addWidget(self.output_directory_edit, 1)
        output_row.addWidget(self.output_browse_button)
        output_row.addWidget(self.open_output_button)
        left.addLayout(output_row)

        progress_row = QHBoxLayout()
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        self.progress_bar.setTextVisible(False)
        self.status_label = QLabel()
        self.progress_detail_label = QLabel()
        progress_row.addWidget(self.progress_bar, 1)
        progress_row.addWidget(self.status_label)
        progress_row.addWidget(self.progress_detail_label)
        left.addLayout(progress_row)
        layout.addLayout(left, 1)

        self.pause_button = QPushButton()
        self.pause_button.setObjectName("SecondaryButton")
        self.pause_button.setEnabled(False)
        self.cancel_button = QPushButton()
        self.cancel_button.setObjectName("DangerButton")
        self.cancel_button.setEnabled(False)
        self.start_button = QPushButton()
        self.start_button.setObjectName("PrimaryButton")
        layout.addWidget(self.pause_button)
        layout.addWidget(self.cancel_button)
        layout.addWidget(self.start_button)
        return footer

    def _connect_signals(self) -> None:
        self.split_workspace_button.clicked.connect(lambda: self.switch_workspace("split"))
        self.merge_workspace_button.clicked.connect(lambda: self.switch_workspace("merge"))
        self.language_combo.currentIndexChanged.connect(self._language_changed)
        self.output_browse_button.clicked.connect(self.choose_output_directory)
        self.open_output_button.clicked.connect(self.open_output_directory)
        self.output_directory_edit.textEdited.connect(self._output_edited)
        self.output_directory_edit.textChanged.connect(lambda _text: self._update_readiness())
        self.start_button.clicked.connect(self.start_task)
        self.pause_button.clicked.connect(self.toggle_pause)
        self.cancel_button.clicked.connect(self.cancel_task)
        self.split_page.configuration_changed.connect(self._update_readiness)
        self.merge_page.configuration_changed.connect(self._update_readiness)
        self.split_page.suggested_output_directory.connect(self._accept_suggested_output)
        self.i18n.language_changed.connect(lambda _language: self.retranslate_ui())

    def _apply_responsive_geometry(self) -> None:
        application = QApplication.instance()
        screen = self.screen() or (application.primaryScreen() if application else None)
        if screen is None:
            self.resize(1280, 800)
            return
        available = screen.availableGeometry()
        minimum_width = min(1040, max(900, available.width() - 120))
        minimum_height = min(640, max(570, available.height() - 120))
        self.setMinimumSize(minimum_width, minimum_height)
        width = min(1440, max(minimum_width, int(available.width() * 0.86)))
        height = min(900, max(minimum_height, int(available.height() * 0.84)))
        self.resize(width, height)
        self.move(
            available.x() + (available.width() - width) // 2,
            available.y() + (available.height() - height) // 2,
        )

    def switch_workspace(self, operation: str) -> None:
        self._active_operation = "merge" if operation == "merge" else "split"
        merge = self._active_operation == "merge"
        self.page_stack.setCurrentWidget(self.merge_page if merge else self.split_page)
        self.split_workspace_button.setChecked(not merge)
        self.merge_workspace_button.setChecked(merge)
        self.start_button.setText(
            self.i18n.text("actions.start_merge" if merge else "actions.start_split")
        )
        self._update_readiness()

    def active_page(self) -> SplitPage | MergePage:
        return self.merge_page if self._active_operation == "merge" else self.split_page

    def _language_changed(self, _index: int) -> None:
        language = str(self.language_combo.currentData())
        self.settings.language = language
        self.settings_store.save(self.settings)
        self.i18n.set_language(language)

    def _output_edited(self, _text: str) -> None:
        self._output_user_changed = True

    def _accept_suggested_output(self, path: str) -> None:
        if not self._output_user_changed or not self.output_directory_edit.text().strip():
            self.output_directory_edit.setText(path)

    def choose_output_directory(self) -> None:
        selected = QFileDialog.getExistingDirectory(
            self,
            self.i18n.text("dialog.select_output"),
            self.output_directory_edit.text().strip(),
        )
        if selected:
            self._output_user_changed = True
            self.output_directory_edit.setText(selected)

    def open_output_directory(self) -> None:
        path = Path(self.output_directory_edit.text().strip())
        if path.exists():
            QDesktopServices.openUrl(QUrl.fromLocalFile(str(path)))

    def _update_readiness(self) -> None:
        if self.task_worker is not None:
            self.start_button.setEnabled(False)
            return
        page = self.active_page()
        ready = page.validation_error() is None and bool(self.output_directory_edit.text().strip())
        self.start_button.setEnabled(ready)
        self.start_button.setToolTip(
            "" if ready else (page.validation_error() or self.i18n.text("validation.output"))
        )

    def start_task(self) -> None:
        output_text = self.output_directory_edit.text().strip()
        if not output_text:
            QMessageBox.warning(
                self,
                self.i18n.text("dialog.warning"),
                self.i18n.text("validation.output"),
            )
            return
        try:
            config = self.active_page().build_config(Path(output_text))
        except ValueError as error:
            QMessageBox.warning(
                self,
                self.i18n.text("dialog.warning"),
                str(error),
            )
            return
        answer = QMessageBox.question(
            self,
            self.i18n.text("confirm.title"),
            self.i18n.text(
                "confirm.body",
                operation=self.i18n.text(f"workspace.{self._active_operation}"),
                files=len(config.input_files),
                output=output_text,
            ),
        )
        if answer != QMessageBox.StandardButton.Yes:
            return
        self._set_running(True)
        self.progress_bar.setValue(0)
        self.status_label.setText(self.i18n.text("status.starting"))
        self.active_page().append_log(self.i18n.text("log.task_queued"))
        self.task_thread = QThread(self)
        self.task_worker = ProcessingWorker(config, self.history_store)
        self.task_worker.moveToThread(self.task_thread)
        self.task_thread.started.connect(self.task_worker.run)
        self.task_worker.progress.connect(self._task_progress)
        self.task_worker.log.connect(self.active_page().append_log)
        self.task_worker.completed.connect(self._task_completed)
        self.task_worker.cancelled.connect(self._task_cancelled)
        self.task_worker.failed.connect(self._task_failed)
        self.task_worker.completed.connect(self.task_thread.quit)
        self.task_worker.cancelled.connect(self.task_thread.quit)
        self.task_worker.failed.connect(self.task_thread.quit)
        self.task_thread.finished.connect(self._task_thread_finished)
        self.task_thread.start()

    def _set_running(self, running: bool) -> None:
        self.page_stack.setEnabled(not running)
        self.split_workspace_button.setEnabled(not running)
        self.merge_workspace_button.setEnabled(not running)
        self.language_combo.setEnabled(not running)
        self.output_directory_edit.setEnabled(not running)
        self.output_browse_button.setEnabled(not running)
        self.start_button.setEnabled(not running)
        self.pause_button.setEnabled(running)
        self.cancel_button.setEnabled(running)

    def _task_progress(self, progress: TaskProgress) -> None:
        self.progress_bar.setValue(progress.overall_percent)
        self.status_label.setText(progress.step)
        details = []
        if progress.current_file:
            details.append(progress.current_file)
        if progress.total_files:
            details.append(f"{progress.completed_files}/{progress.total_files}")
        self.progress_detail_label.setText(" · ".join(details))

    def _task_completed(self, summary: TaskExecutionSummary) -> None:
        self.progress_bar.setValue(100)
        self.status_label.setText(self.i18n.text("status.completed"))
        self.progress_detail_label.setText(
            self.i18n.text(
                "status.output_count",
                count=len(summary.result.output_files),
            )
        )
        self.active_page().append_log(
            self.i18n.text(
                "log.task_completed",
                outputs=len(summary.result.output_files),
                warnings=len(summary.result.warnings),
                errors=len(summary.result.errors),
            )
        )
        for warning in summary.result.warnings:
            self.active_page().append_log(self.i18n.text("log.warning", message=warning))
        for error in summary.result.errors:
            self.active_page().append_log(self.i18n.text("log.error", message=error))
        self.open_output_button.setEnabled(bool(summary.result.output_files))

    def _task_cancelled(self, _summary: TaskExecutionSummary) -> None:
        self.status_label.setText(self.i18n.text("status.cancelled"))
        self.active_page().append_log(self.i18n.text("log.task_cancelled"))

    def _task_failed(self, message: str) -> None:
        self.status_label.setText(self.i18n.text("status.failed"))
        self.active_page().append_log(self.i18n.text("log.task_failed", error=message))
        QMessageBox.critical(
            self,
            self.i18n.text("dialog.error"),
            self.i18n.text("log.task_failed", error=message),
        )

    def _task_thread_finished(self) -> None:
        if self.task_worker is not None:
            self.task_worker.deleteLater()
        if self.task_thread is not None:
            self.task_thread.deleteLater()
        self.task_worker = None
        self.task_thread = None
        self.pause_button.setProperty("paused", False)
        self._set_running(False)
        self._update_readiness()

    def toggle_pause(self) -> None:
        if self.task_worker is None:
            return
        paused = bool(self.pause_button.property("paused"))
        if paused:
            self.task_worker.resume()
        else:
            self.task_worker.pause()
        self.pause_button.setProperty("paused", not paused)
        self.pause_button.setText(
            self.i18n.text("actions.resume" if not paused else "actions.pause")
        )

    def cancel_task(self) -> None:
        if self.task_worker is not None:
            self.task_worker.cancel()
            self.cancel_button.setEnabled(False)
            self.status_label.setText(self.i18n.text("status.cancelling"))

    def retranslate_ui(self) -> None:
        self.setWindowTitle(f"{self.i18n.text('app.title')} {VERSION}")
        self.title_label.setText(self.i18n.text("app.title"))
        self.subtitle_label.setText(self.i18n.text("header.subtitle"))
        self.split_workspace_button.setText(self.i18n.text("workspace.split"))
        self.merge_workspace_button.setText(self.i18n.text("workspace.merge"))
        index = self.language_combo.findData(self.i18n.language)
        self.language_combo.blockSignals(True)
        self.language_combo.setCurrentIndex(max(0, index))
        self.language_combo.blockSignals(False)
        self.output_label.setText(self.i18n.text("footer.output"))
        self.output_browse_button.setText(self.i18n.text("footer.choose"))
        self.open_output_button.setText(self.i18n.text("footer.open"))
        self.pause_button.setText(self.i18n.text("actions.pause"))
        self.cancel_button.setText(self.i18n.text("actions.cancel"))
        self.start_button.setText(
            self.i18n.text(
                "actions.start_merge"
                if self._active_operation == "merge"
                else "actions.start_split"
            )
        )
        if self.task_worker is None:
            self.status_label.setText(self.i18n.text("status.ready"))
        self.split_page.retranslate_ui()
        self.merge_page.retranslate_ui()
        self._update_readiness()

    def closeEvent(self, event: QCloseEvent) -> None:  # noqa: N802
        if self.task_worker is not None:
            QMessageBox.information(
                self,
                self.i18n.text("dialog.information"),
                self.i18n.text("close.task_running"),
            )
            event.ignore()
            return
        event.accept()
