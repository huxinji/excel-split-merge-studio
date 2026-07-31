from __future__ import annotations

from collections import Counter
from pathlib import Path

from PySide6.QtCore import Qt, Signal
from PySide6.QtWidgets import (
    QAbstractItemView,
    QCheckBox,
    QComboBox,
    QFileDialog,
    QFormLayout,
    QFrame,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QPushButton,
    QSpinBox,
    QStackedWidget,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from excel_studio.config.i18n import TranslationManager
from excel_studio.config.settings import AppSettings
from excel_studio.models.operations import (
    BlankValuePolicy,
    ExistingFilePolicy,
    OutputFormat,
    OutputMode,
    OutputOptions,
    SplitDistribution,
    SplitMode,
    SplitOptions,
    SplitTaskConfig,
)
from excel_studio.models.task import HeaderMode, ScanOptions, TableStructure
from excel_studio.models.workbook import FileRecord, PreviewResult
from excel_studio.ui.base_page import OperationPage
from excel_studio.ui.widgets import SegmentedSelector, StepCard
from excel_studio.workers.scan_worker import ScanWorker


class SplitPage(OperationPage):
    suggested_output_directory = Signal(str)

    def __init__(self, i18n: TranslationManager, settings: AppSettings) -> None:
        super().__init__(i18n, settings)
        self.record: FileRecord | None = None
        self._updating_sheets = False
        self._build_source_card()
        self._build_sheet_card()
        self._build_rule_card()
        self._build_output_card()
        self.finish_cards()
        self.paths_dropped.connect(self._paths_dropped)
        self.retranslate_ui()
        self.mode_selector.set_current(SplitMode.BY_FIELD.value)
        self.sheet_scope_selector.set_current("visible")

    def _build_source_card(self) -> None:
        self.source_card = StepCard(1, "blue")
        form = QFormLayout()
        form.setContentsMargins(0, 2, 0, 0)
        form.setHorizontalSpacing(12)
        form.setVerticalSpacing(8)

        file_row = QHBoxLayout()
        self.input_path_edit = QLineEdit()
        self.input_path_edit.setReadOnly(True)
        self.browse_button = QPushButton()
        self.browse_button.setObjectName("PrimaryButton")
        self.browse_button.clicked.connect(self.choose_file)
        file_row.addWidget(self.input_path_edit, 1)
        file_row.addWidget(self.browse_button)
        self.file_label = QLabel()
        form.addRow(self.file_label, file_row)

        header_row = QHBoxLayout()
        self.header_spin = QSpinBox()
        self.header_spin.setRange(1, 100_000)
        self.header_spin.setValue(1)
        self.reload_button = QPushButton()
        self.reload_button.setObjectName("SecondaryButton")
        self.reload_button.clicked.connect(self._reload)
        header_row.addWidget(self.header_spin)
        header_row.addStretch(1)
        header_row.addWidget(self.reload_button)
        self.header_label = QLabel()
        form.addRow(self.header_label, header_row)
        self.header_hint = QLabel()
        self.header_hint.setObjectName("Muted")
        self.header_hint.setWordWrap(True)
        form.addRow("", self.header_hint)
        self.source_card.layout.addLayout(form)
        self.add_card(self.source_card)
        self.header_spin.valueChanged.connect(lambda _value: self.configuration_changed.emit())

    def _build_sheet_card(self) -> None:
        self.sheet_card = StepCard(2, "cyan")
        self.sheet_scope_selector = SegmentedSelector(("all", "visible", "single", "multiple"))
        self.sheet_scope_selector.changed.connect(self._sheet_scope_changed)
        self.sheet_card.add(self.sheet_scope_selector)
        self.sheet_table = QTableWidget(0, 5)
        self.sheet_table.setAlternatingRowColors(True)
        self.sheet_table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.sheet_table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.sheet_table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.sheet_table.verticalHeader().setVisible(False)
        self.sheet_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        self.sheet_table.setMinimumHeight(150)
        self.sheet_table.itemChanged.connect(self._sheet_item_changed)
        self.sheet_table.currentCellChanged.connect(self._sheet_current_changed)
        self.sheet_card.add(self.sheet_table)
        self.sheet_hint = QLabel()
        self.sheet_hint.setObjectName("Muted")
        self.sheet_hint.setWordWrap(True)
        self.sheet_card.add(self.sheet_hint)
        self.add_card(self.sheet_card)

    def _build_rule_card(self) -> None:
        self.rule_card = StepCard(3, "violet")
        self.mode_selector = SegmentedSelector(
            (
                SplitMode.BY_FIELD.value,
                SplitMode.FIXED_ROWS.value,
                SplitMode.BY_PARTS.value,
                SplitMode.BY_SHEET.value,
            )
        )
        self.mode_selector.changed.connect(self._mode_changed)
        self.rule_card.add(self.mode_selector)

        self.mode_stack = QStackedWidget()
        self.field_panel = self._build_field_panel()
        self.rows_panel = self._build_rows_panel()
        self.parts_panel = self._build_parts_panel()
        self.sheet_panel = self._build_sheet_panel()
        for panel in (self.field_panel, self.rows_panel, self.parts_panel, self.sheet_panel):
            self.mode_stack.addWidget(panel)
        self.rule_card.add(self.mode_stack)
        self.add_card(self.rule_card)

    def _build_field_panel(self) -> QWidget:
        panel = QFrame()
        panel.setObjectName("PrimaryRulePanel")
        layout = QVBoxLayout(panel)
        layout.setContentsMargins(14, 13, 14, 14)
        layout.setSpacing(9)
        self.field_primary_title = QLabel()
        self.field_primary_title.setObjectName("PrimaryRuleTitle")
        layout.addWidget(self.field_primary_title)

        row = QHBoxLayout()
        self.field_label = QLabel()
        self.field_label.setMinimumWidth(115)
        self.field_combo = QComboBox()
        self.field_combo.setMinimumWidth(280)
        self.field_combo.setEnabled(False)
        self.refresh_fields_button = QPushButton()
        self.refresh_fields_button.setObjectName("SecondaryButton")
        self.refresh_fields_button.clicked.connect(self.refresh_preview)
        row.addWidget(self.field_label)
        row.addWidget(self.field_combo, 1)
        row.addWidget(self.refresh_fields_button)
        layout.addLayout(row)

        self.field_hint = QLabel()
        self.field_hint.setObjectName("Muted")
        self.field_hint.setWordWrap(True)
        layout.addWidget(self.field_hint)

        blank_row = QHBoxLayout()
        self.blank_policy_label = QLabel()
        self.blank_policy_combo = QComboBox()
        self.blank_policy_combo.addItem("", BlankValuePolicy.GROUP.value)
        self.blank_policy_combo.addItem("", BlankValuePolicy.SKIP.value)
        blank_row.addWidget(self.blank_policy_label)
        blank_row.addWidget(self.blank_policy_combo, 1)
        layout.addLayout(blank_row)

        self.field_value_summary = QLabel()
        self.field_value_summary.setObjectName("Muted")
        layout.addWidget(self.field_value_summary)
        self.field_value_table = QTableWidget(0, 2)
        self.field_value_table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.field_value_table.setMaximumHeight(150)
        self.field_value_table.verticalHeader().setVisible(False)
        self.field_value_table.horizontalHeader().setSectionResizeMode(
            0, QHeaderView.ResizeMode.Stretch
        )
        self.field_value_table.horizontalHeader().setSectionResizeMode(
            1, QHeaderView.ResizeMode.ResizeToContents
        )
        layout.addWidget(self.field_value_table)
        self.field_combo.currentIndexChanged.connect(self._update_field_values)
        self.blank_policy_combo.currentIndexChanged.connect(
            lambda _index: self.configuration_changed.emit()
        )
        return panel

    def _build_rows_panel(self) -> QWidget:
        panel = QFrame()
        layout = QFormLayout(panel)
        layout.setContentsMargins(8, 8, 8, 8)
        self.rows_label = QLabel()
        self.rows_spin = QSpinBox()
        self.rows_spin.setRange(1, 10_000_000)
        self.rows_spin.setValue(1000)
        layout.addRow(self.rows_label, self.rows_spin)
        self.distribution_label = QLabel()
        self.distribution_combo = QComboBox()
        self.distribution_combo.addItem("", SplitDistribution.FIXED_LIMIT.value)
        self.distribution_combo.addItem("", SplitDistribution.BALANCED.value)
        layout.addRow(self.distribution_label, self.distribution_combo)
        self.rows_hint = QLabel()
        self.rows_hint.setObjectName("Muted")
        self.rows_hint.setWordWrap(True)
        layout.addRow("", self.rows_hint)
        self.rows_spin.valueChanged.connect(lambda _value: self.configuration_changed.emit())
        return panel

    def _build_parts_panel(self) -> QWidget:
        panel = QFrame()
        layout = QFormLayout(panel)
        layout.setContentsMargins(8, 8, 8, 8)
        self.parts_label = QLabel()
        self.parts_spin = QSpinBox()
        self.parts_spin.setRange(2, 10_000)
        self.parts_spin.setValue(5)
        layout.addRow(self.parts_label, self.parts_spin)
        self.parts_hint = QLabel()
        self.parts_hint.setObjectName("Muted")
        self.parts_hint.setWordWrap(True)
        layout.addRow("", self.parts_hint)
        self.parts_spin.valueChanged.connect(lambda _value: self.configuration_changed.emit())
        return panel

    def _build_sheet_panel(self) -> QWidget:
        panel = QFrame()
        layout = QVBoxLayout(panel)
        layout.setContentsMargins(8, 8, 8, 8)
        self.sheet_mode_hint = QLabel()
        self.sheet_mode_hint.setObjectName("Muted")
        self.sheet_mode_hint.setWordWrap(True)
        layout.addWidget(self.sheet_mode_hint)
        return panel

    def _build_output_card(self) -> None:
        self.output_card = StepCard(4, "amber")
        form = QFormLayout()
        form.setContentsMargins(0, 2, 0, 0)
        form.setHorizontalSpacing(12)
        form.setVerticalSpacing(7)
        self.output_labels = [QLabel() for _ in range(7)]

        self.output_mode_combo = QComboBox()
        self.output_mode_combo.addItem("", OutputMode.SEPARATE_FILES.value)
        self.output_mode_combo.addItem("", OutputMode.SINGLE_WORKBOOK.value)
        form.addRow(self.output_labels[0], self.output_mode_combo)
        self.output_format_combo = QComboBox()
        self.output_format_combo.addItem("", OutputFormat.XLSX.value)
        self.output_format_combo.addItem("", OutputFormat.CSV.value)
        form.addRow(self.output_labels[1], self.output_format_combo)
        self.workbook_name_edit = QLineEdit("split_results")
        form.addRow(self.output_labels[2], self.workbook_name_edit)
        self.naming_template_edit = QLineEdit("{original_name}_{split_value}")
        form.addRow(self.output_labels[3], self.naming_template_edit)
        name_row = QHBoxLayout()
        self.prefix_edit = QLineEdit()
        self.suffix_edit = QLineEdit()
        name_row.addWidget(self.prefix_edit)
        name_row.addWidget(self.suffix_edit)
        form.addRow(self.output_labels[4], name_row)
        self.existing_policy_combo = QComboBox()
        for policy in ExistingFilePolicy:
            self.existing_policy_combo.addItem("", policy.value)
        form.addRow(self.output_labels[5], self.existing_policy_combo)

        options = QVBoxLayout()
        self.preserve_format_check = QCheckBox()
        self.preserve_format_check.setChecked(True)
        self.style_output_check = QCheckBox()
        self.style_output_check.setChecked(True)
        self.create_report_check = QCheckBox()
        self.create_report_check.setChecked(True)
        self.open_when_done_check = QCheckBox()
        self.open_when_done_check.setChecked(True)
        for check in (
            self.preserve_format_check,
            self.style_output_check,
            self.create_report_check,
            self.open_when_done_check,
        ):
            options.addWidget(check)
            check.toggled.connect(lambda _checked: self.configuration_changed.emit())
        form.addRow(self.output_labels[6], options)
        self.output_card.layout.addLayout(form)
        self.output_hint = QLabel()
        self.output_hint.setObjectName("Muted")
        self.output_hint.setWordWrap(True)
        self.output_card.add(self.output_hint)
        self.add_card(self.output_card)
        for combo in (self.output_mode_combo, self.output_format_combo):
            combo.currentIndexChanged.connect(self._sync_output_constraints)
        for edit in (
            self.workbook_name_edit,
            self.naming_template_edit,
            self.prefix_edit,
            self.suffix_edit,
        ):
            edit.textChanged.connect(lambda _text: self.configuration_changed.emit())

    def choose_file(self) -> None:
        path, _selected = QFileDialog.getOpenFileName(
            self,
            self.i18n.text("dialog.select_source"),
            "",
            self.i18n.text("dialog.file_filter"),
        )
        if path:
            self.load_path(Path(path))

    def _paths_dropped(self, paths: object) -> None:
        candidates = [path for path in paths if isinstance(path, Path) and path.is_file()]
        if candidates:
            self.load_path(candidates[0])

    def load_path(self, path: Path) -> None:
        self.input_path_edit.setText(str(path))
        self.browse_button.setEnabled(False)
        self.append_log(self.i18n.text("log.scanning", file=path.name))
        worker = ScanWorker(
            [path],
            ScanOptions(include_text_files=True, ignore_hidden=True, ignore_temp_files=True),
        )
        worker.progress.connect(
            lambda current, total, name: self.append_log(
                self.i18n.text("log.scan_progress", current=current, total=total, file=name)
            )
        )
        self._run_worker(worker, self._scan_completed, self._scan_failed)

    def _reload(self) -> None:
        if self.record is not None:
            self.load_path(self.record.path)

    def _scan_completed(self, records: list[FileRecord]) -> None:
        self.browse_button.setEnabled(True)
        if not records:
            self.record = None
            self.append_log(self.i18n.text("log.no_supported_files"))
            return
        self.record = records[0]
        if self.record.status == "error":
            self.append_log(self.i18n.text("log.scan_failed", error=self.record.warning))
            return
        self._populate_sheets()
        self.suggested_output_directory.emit(str(self.record.path.parent / "Split_Output"))
        self.append_log(
            self.i18n.text(
                "log.source_loaded", file=self.record.name, sheets=len(self.record.sheets)
            )
        )
        self.configuration_changed.emit()

    def _scan_failed(self, message: str) -> None:
        self.browse_button.setEnabled(True)
        self.append_log(self.i18n.text("log.scan_failed", error=message))

    def _populate_sheets(self) -> None:
        self._updating_sheets = True
        try:
            sheets = self.record.sheets if self.record is not None else []
            self.sheet_table.setRowCount(len(sheets))
            for row, sheet in enumerate(sheets):
                check = QTableWidgetItem()
                check.setFlags(
                    Qt.ItemFlag.ItemIsUserCheckable
                    | Qt.ItemFlag.ItemIsEnabled
                    | Qt.ItemFlag.ItemIsSelectable
                )
                check.setCheckState(
                    Qt.CheckState.Checked if sheet.state == "visible" else Qt.CheckState.Unchecked
                )
                self.sheet_table.setItem(row, 0, check)
                self.sheet_table.setItem(row, 1, QTableWidgetItem(sheet.name))
                self.sheet_table.setItem(row, 2, QTableWidgetItem(sheet.state))
                self.sheet_table.setItem(row, 3, QTableWidgetItem(f"{sheet.max_row:,}"))
                self.sheet_table.setItem(row, 4, QTableWidgetItem(str(sheet.max_column)))
            if sheets:
                self.sheet_table.setCurrentCell(0, 1)
                self.set_preview_target(self.record.path, sheets[0].name)
                self.refresh_preview()
        finally:
            self._updating_sheets = False
        self._sheet_scope_changed(self.sheet_scope_selector.current())

    def _sheet_scope_changed(self, scope: str) -> None:
        if self._updating_sheets or self.record is None:
            return
        self._updating_sheets = True
        try:
            current = max(0, self.sheet_table.currentRow())
            for row, sheet in enumerate(self.record.sheets):
                check = self.sheet_table.item(row, 0)
                if check is None:
                    continue
                if scope == "all":
                    checked = True
                elif scope == "visible":
                    checked = sheet.state == "visible"
                elif scope == "single":
                    checked = row == current
                else:
                    checked = check.checkState() == Qt.CheckState.Checked
                check.setCheckState(Qt.CheckState.Checked if checked else Qt.CheckState.Unchecked)
        finally:
            self._updating_sheets = False
        self.configuration_changed.emit()

    def _sheet_item_changed(self, _item: QTableWidgetItem) -> None:
        if not self._updating_sheets:
            self.sheet_scope_selector.button("multiple").setChecked(True)
            self.configuration_changed.emit()

    def _sheet_current_changed(
        self, current_row: int, _current_column: int, _previous_row: int, _previous_column: int
    ) -> None:
        if self.record is None or current_row < 0 or current_row >= len(self.record.sheets):
            return
        sheet = self.record.sheets[current_row]
        self.set_preview_target(self.record.path, sheet.name)
        if self.sheet_scope_selector.current() == "single":
            self._sheet_scope_changed("single")
        self.refresh_preview()

    def on_preview_ready(self, result: PreviewResult) -> None:
        previous = self.field_combo.currentText()
        valid_fields: list[str] = []
        for index, column in enumerate(result.columns):
            name = str(column).strip()
            if not name or name.casefold().startswith("unnamed"):
                continue
            if any(
                index < len(row) and row[index] is not None and str(row[index]).strip()
                for row in result.rows
            ):
                valid_fields.append(name)
        self.field_combo.blockSignals(True)
        self.field_combo.clear()
        self.field_combo.addItems(valid_fields)
        if previous in valid_fields:
            self.field_combo.setCurrentText(previous)
        self.field_combo.setEnabled(bool(valid_fields))
        self.field_combo.blockSignals(False)
        self._update_field_values()
        self.configuration_changed.emit()

    def _update_field_values(self) -> None:
        self.field_value_table.clearContents()
        if self.current_preview is None or self.field_combo.currentIndex() < 0:
            self.field_value_table.setRowCount(0)
            self.field_value_summary.setText(self.i18n.text("split.field.no_values"))
            self.configuration_changed.emit()
            return
        field = self.field_combo.currentText()
        try:
            column_index = self.current_preview.columns.index(field)
        except ValueError:
            return
        values = [
            self.i18n.text("split.field.blank_value")
            if column_index >= len(row)
            or row[column_index] is None
            or not str(row[column_index]).strip()
            else str(row[column_index]).strip()
            for row in self.current_preview.rows
        ]
        counts = Counter(values)
        shown = counts.most_common(12)
        self.field_value_table.setRowCount(len(shown))
        for row, (value, count) in enumerate(shown):
            self.field_value_table.setItem(row, 0, QTableWidgetItem(value))
            self.field_value_table.setItem(row, 1, QTableWidgetItem(str(count)))
        self.field_value_summary.setText(
            self.i18n.text(
                "split.field.value_summary",
                count=len(counts),
                rows=len(values),
            )
        )
        self.configuration_changed.emit()

    def _mode_changed(self, value: str) -> None:
        mapping = {
            SplitMode.BY_FIELD.value: 0,
            SplitMode.FIXED_ROWS.value: 1,
            SplitMode.BY_PARTS.value: 2,
            SplitMode.BY_SHEET.value: 3,
        }
        self.mode_stack.setCurrentIndex(mapping.get(value, 0))
        if value == SplitMode.BY_SHEET.value:
            self.output_mode_combo.setCurrentIndex(
                self.output_mode_combo.findData(OutputMode.SEPARATE_FILES.value)
            )
            self.output_format_combo.setCurrentIndex(
                self.output_format_combo.findData(OutputFormat.XLSX.value)
            )
        self._sync_output_constraints()
        self.configuration_changed.emit()

    def _sync_output_constraints(self) -> None:
        sheet_mode = self.current_mode() == SplitMode.BY_SHEET
        workbook_mode = self.output_mode_combo.currentData() == OutputMode.SINGLE_WORKBOOK.value
        if sheet_mode:
            self.output_mode_combo.setCurrentIndex(
                self.output_mode_combo.findData(OutputMode.SEPARATE_FILES.value)
            )
            self.output_format_combo.setCurrentIndex(
                self.output_format_combo.findData(OutputFormat.XLSX.value)
            )
        elif workbook_mode:
            self.output_format_combo.setCurrentIndex(
                self.output_format_combo.findData(OutputFormat.XLSX.value)
            )
        self.output_mode_combo.setEnabled(not sheet_mode)
        self.output_format_combo.setEnabled(not sheet_mode and not workbook_mode)
        self.workbook_name_edit.setEnabled(workbook_mode and not sheet_mode)
        self.preserve_format_check.setEnabled(
            self.output_format_combo.currentData() == OutputFormat.XLSX.value
        )
        self.configuration_changed.emit()

    def current_mode(self) -> SplitMode:
        value = self.mode_selector.current() or SplitMode.BY_FIELD.value
        return SplitMode(value)

    def selected_sheet_names(self) -> list[str]:
        names: list[str] = []
        for row in range(self.sheet_table.rowCount()):
            check = self.sheet_table.item(row, 0)
            name = self.sheet_table.item(row, 1)
            if (
                check is not None
                and name is not None
                and check.checkState() == Qt.CheckState.Checked
            ):
                names.append(name.text())
        return names

    def header_row_value(self) -> int:
        return self.header_spin.value()

    def validation_error(self) -> str | None:
        if self.record is None:
            return self.i18n.text("validation.select_source")
        if not self.selected_sheet_names():
            return self.i18n.text("validation.select_sheet")
        if self.current_mode() == SplitMode.BY_FIELD and not self.field_combo.currentText():
            return self.i18n.text("validation.select_split_field")
        return None

    def build_config(self, output_directory: Path) -> SplitTaskConfig:
        error = self.validation_error()
        if error is not None or self.record is None:
            raise ValueError(error or self.i18n.text("validation.select_source"))
        mode = self.current_mode()
        output_mode = OutputMode(str(self.output_mode_combo.currentData()))
        output_format = OutputFormat(str(self.output_format_combo.currentData()))
        if mode == SplitMode.BY_SHEET:
            output_mode = OutputMode.SEPARATE_FILES
            output_format = OutputFormat.XLSX
        return SplitTaskConfig(
            input_files=[self.record.path],
            structure=TableStructure(
                header_mode=HeaderMode.ROW_NUMBER,
                header_row=self.header_spin.value(),
                header_end_row=self.header_spin.value(),
                data_start_row=self.header_spin.value() + 1,
            ),
            split=SplitOptions(
                mode=mode,
                fields=[self.field_combo.currentText()] if mode == SplitMode.BY_FIELD else [],
                rows_per_file=self.rows_spin.value(),
                parts=self.parts_spin.value(),
                distribution=SplitDistribution(str(self.distribution_combo.currentData())),
                blank_value_policy=BlankValuePolicy(str(self.blank_policy_combo.currentData())),
                selected_sheets={self.record.path: self.selected_sheet_names()},
            ),
            output=OutputOptions(
                directory=output_directory,
                naming_template=self.naming_template_edit.text().strip()
                or "{original_name}_{split_value}",
                prefix=self.prefix_edit.text(),
                suffix=self.suffix_edit.text(),
                add_source_file=False,
                add_source_sheet=False,
                output_format=output_format,
                output_mode=output_mode,
                existing_file_policy=ExistingFilePolicy(
                    str(self.existing_policy_combo.currentData())
                ),
                workbook_name=self.workbook_name_edit.text().strip() or "split_results",
                preserve_format=self.preserve_format_check.isChecked(),
                style_output=self.style_output_check.isChecked(),
                create_report_sheet=self.create_report_check.isChecked(),
                open_when_done=self.open_when_done_check.isChecked(),
            ),
        )

    @staticmethod
    def _set_combo_texts(combo: QComboBox, labels: dict[str, str]) -> None:
        for index in range(combo.count()):
            value = str(combo.itemData(index))
            combo.setItemText(index, labels[value])

    def retranslate_ui(self) -> None:
        super().retranslate_ui()
        self.source_card.title_label.setText(self.i18n.text("split.step.source"))
        self.source_card.description_label.setText(self.i18n.text("split.step.source_hint"))
        self.file_label.setText(self.i18n.text("split.source.file"))
        self.browse_button.setText(self.i18n.text("split.source.browse"))
        self.header_label.setText(self.i18n.text("common.header_row"))
        self.reload_button.setText(self.i18n.text("common.reload"))
        self.header_hint.setText(self.i18n.text("split.source.header_hint"))

        self.sheet_card.title_label.setText(self.i18n.text("split.step.sheets"))
        self.sheet_card.description_label.setText(self.i18n.text("split.step.sheets_hint"))
        for value in ("all", "visible", "single", "multiple"):
            self.sheet_scope_selector.button(value).setText(
                self.i18n.text(f"split.sheet_scope.{value}")
            )
        self.sheet_table.setHorizontalHeaderLabels(
            [
                self.i18n.text("table.select"),
                self.i18n.text("table.sheet"),
                self.i18n.text("table.state"),
                self.i18n.text("table.rows"),
                self.i18n.text("table.columns"),
            ]
        )
        self.sheet_hint.setText(self.i18n.text("split.sheet_scope.hint"))

        self.rule_card.title_label.setText(self.i18n.text("split.step.rules"))
        self.rule_card.description_label.setText(self.i18n.text("split.step.rules_hint"))
        for mode in SplitMode:
            self.mode_selector.button(mode.value).setText(
                self.i18n.text(f"split.mode.{mode.value}")
            )
        self.field_primary_title.setText(self.i18n.text("split.field.primary_title"))
        self.field_label.setText(self.i18n.text("split.field.select"))
        self.refresh_fields_button.setText(self.i18n.text("split.field.refresh"))
        self.field_hint.setText(self.i18n.text("split.field.hint"))
        self.blank_policy_label.setText(self.i18n.text("split.field.blank_policy"))
        self._set_combo_texts(
            self.blank_policy_combo,
            {
                BlankValuePolicy.GROUP.value: self.i18n.text("split.blank.group"),
                BlankValuePolicy.SKIP.value: self.i18n.text("split.blank.skip"),
            },
        )
        self.field_value_table.setHorizontalHeaderLabels(
            [
                self.i18n.text("split.field.value"),
                self.i18n.text("split.field.preview_count"),
            ]
        )
        self.rows_label.setText(self.i18n.text("split.rows.count"))
        self.distribution_label.setText(self.i18n.text("split.rows.distribution"))
        self._set_combo_texts(
            self.distribution_combo,
            {
                SplitDistribution.FIXED_LIMIT.value: self.i18n.text("split.distribution.fixed"),
                SplitDistribution.BALANCED.value: self.i18n.text("split.distribution.balanced"),
            },
        )
        self.rows_hint.setText(self.i18n.text("split.rows.hint"))
        self.parts_label.setText(self.i18n.text("split.parts.count"))
        self.parts_hint.setText(self.i18n.text("split.parts.hint"))
        self.sheet_mode_hint.setText(self.i18n.text("split.sheet_mode.hint"))

        self.output_card.title_label.setText(self.i18n.text("split.step.output"))
        self.output_card.description_label.setText(self.i18n.text("split.step.output_hint"))
        output_keys = (
            "split.output.mode",
            "split.output.format",
            "split.output.workbook_name",
            "split.output.naming",
            "split.output.prefix_suffix",
            "split.output.existing",
            "split.output.options",
        )
        for label, key in zip(self.output_labels, output_keys, strict=True):
            label.setText(self.i18n.text(key))
        self._set_combo_texts(
            self.output_mode_combo,
            {
                OutputMode.SEPARATE_FILES.value: self.i18n.text("split.output.separate"),
                OutputMode.SINGLE_WORKBOOK.value: self.i18n.text("split.output.single"),
            },
        )
        self._set_combo_texts(
            self.output_format_combo,
            {
                OutputFormat.XLSX.value: self.i18n.text("format.xlsx"),
                OutputFormat.CSV.value: self.i18n.text("format.csv"),
            },
        )
        self.prefix_edit.setPlaceholderText(self.i18n.text("split.output.prefix"))
        self.suffix_edit.setPlaceholderText(self.i18n.text("split.output.suffix"))
        self._set_combo_texts(
            self.existing_policy_combo,
            {
                ExistingFilePolicy.RENAME.value: self.i18n.text("existing.rename"),
                ExistingFilePolicy.OVERWRITE.value: self.i18n.text("existing.overwrite"),
                ExistingFilePolicy.SKIP.value: self.i18n.text("existing.skip"),
            },
        )
        self.preserve_format_check.setText(self.i18n.text("option.preserve_format"))
        self.style_output_check.setText(self.i18n.text("option.style_output"))
        self.create_report_check.setText(self.i18n.text("option.create_report"))
        self.open_when_done_check.setText(self.i18n.text("option.open_when_done"))
        self.output_hint.setText(self.i18n.text("split.output.hint"))
        self._update_field_values()
