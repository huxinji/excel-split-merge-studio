from __future__ import annotations

from pathlib import Path

from PySide6.QtCore import Qt
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
    QListWidget,
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
from excel_studio.models.advanced import (
    AdvancedMergeTaskConfig,
    CleaningOptions,
    DuplicateKeyPolicy,
    JoinConflictPolicy,
    JoinType,
)
from excel_studio.models.field_catalog import FieldCatalog, FieldDiscoveryRequest
from excel_studio.models.operations import (
    AggregateMethod,
    DedupeMode,
    ExistingFilePolicy,
    FieldStrategy,
    MergeMode,
    MergeOptions,
    MergeTaskConfig,
    OutputFormat,
    OutputOptions,
    SheetConflictPolicy,
)
from excel_studio.models.task import HeaderMode, ScanOptions, TableStructure
from excel_studio.models.workbook import FileRecord
from excel_studio.ui.base_page import OperationPage
from excel_studio.ui.widgets import StepCard
from excel_studio.workers.field_discovery_worker import FieldDiscoveryWorker
from excel_studio.workers.scan_worker import ScanWorker


class MergePage(OperationPage):
    def __init__(self, i18n: TranslationManager, settings: AppSettings) -> None:
        super().__init__(i18n, settings)
        self.records: list[FileRecord] = []
        self.field_catalog = FieldCatalog()
        self._building_file_table = False
        self._build_source_card()
        self._build_rule_card()
        self._build_output_card()
        self.finish_cards()
        self.paths_dropped.connect(self._paths_dropped)
        self.retranslate_ui()
        self.merge_mode_combo.setCurrentIndex(
            self.merge_mode_combo.findData(MergeMode.VERTICAL.value)
        )
        self._mode_changed()

    def _build_source_card(self) -> None:
        self.source_card = StepCard(1, "blue")
        toolbar = QHBoxLayout()
        self.add_files_button = self._button("PrimaryButton", self.add_files)
        self.add_folder_button = self._button("SecondaryButton", self.add_folder)
        self.remove_button = self._button("DangerButton", self.remove_selected)
        self.up_button = self._button("SecondaryButton", lambda: self._move_selected(-1))
        self.down_button = self._button("SecondaryButton", lambda: self._move_selected(1))
        self.clear_button = self._button("DangerButton", self.clear_files)
        self.analyze_button = self._button("SuccessButton", self.analyze_sources)
        for button in (
            self.add_files_button,
            self.add_folder_button,
            self.remove_button,
            self.up_button,
            self.down_button,
            self.clear_button,
            self.analyze_button,
        ):
            toolbar.addWidget(button)
        toolbar.addStretch(1)
        self.source_card.layout.addLayout(toolbar)

        self.file_table = QTableWidget(0, 5)
        self.file_table.setAlternatingRowColors(True)
        self.file_table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.file_table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.file_table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.file_table.verticalHeader().setVisible(False)
        self.file_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.Stretch)
        self.file_table.setMinimumHeight(160)
        self.file_table.itemChanged.connect(lambda _item: self.configuration_changed.emit())
        self.file_table.currentCellChanged.connect(self._file_current_changed)
        self.source_card.add(self.file_table)

        form = QFormLayout()
        form.setHorizontalSpacing(12)
        form.setVerticalSpacing(7)
        self.source_labels = [QLabel() for _ in range(3)]
        self.header_spin = QSpinBox()
        self.header_spin.setRange(1, 100_000)
        self.header_spin.setValue(1)
        form.addRow(self.source_labels[0], self.header_spin)
        self.sheet_rule_combo = QComboBox()
        for value in ("first", "visible", "all", "named"):
            self.sheet_rule_combo.addItem("", value)
        form.addRow(self.source_labels[1], self.sheet_rule_combo)
        self.named_sheets_edit = QLineEdit()
        form.addRow(self.source_labels[2], self.named_sheets_edit)
        self.source_card.layout.addLayout(form)
        self.source_hint = QLabel()
        self.source_hint.setObjectName("Muted")
        self.source_hint.setWordWrap(True)
        self.source_card.add(self.source_hint)
        self.add_card(self.source_card)

        self.header_spin.valueChanged.connect(self._structure_changed)
        self.sheet_rule_combo.currentIndexChanged.connect(self._sheet_rule_changed)
        self.named_sheets_edit.textChanged.connect(lambda _text: self.configuration_changed.emit())

    @staticmethod
    def _button(object_name: str, callback: object) -> QPushButton:
        button = QPushButton()
        button.setObjectName(object_name)
        button.clicked.connect(callback)
        return button

    def _build_rule_card(self) -> None:
        self.rule_card = StepCard(2, "violet")
        mode_row = QHBoxLayout()
        self.mode_label = QLabel()
        self.merge_mode_combo = QComboBox()
        for mode in (
            MergeMode.VERTICAL,
            MergeMode.KEY_JOIN,
            MergeMode.WORKBOOK,
            MergeMode.SAME_NAME,
            MergeMode.HORIZONTAL,
            MergeMode.AGGREGATE,
        ):
            self.merge_mode_combo.addItem("", mode.value)
        mode_row.addWidget(self.mode_label)
        mode_row.addWidget(self.merge_mode_combo, 1)
        self.rule_card.layout.addLayout(mode_row)

        self.mode_stack = QStackedWidget()
        self.vertical_panel = self._build_vertical_panel()
        self.join_panel = self._build_join_panel()
        self.workbook_panel = self._build_workbook_panel()
        self.same_name_panel = self._build_same_name_panel()
        self.horizontal_panel = self._build_horizontal_panel()
        self.aggregate_panel = self._build_aggregate_panel()
        for panel in (
            self.vertical_panel,
            self.join_panel,
            self.workbook_panel,
            self.same_name_panel,
            self.horizontal_panel,
            self.aggregate_panel,
        ):
            self.mode_stack.addWidget(panel)
        self.rule_card.add(self.mode_stack)
        self.add_card(self.rule_card)
        self.merge_mode_combo.currentIndexChanged.connect(self._mode_changed)

    def _build_vertical_panel(self) -> QWidget:
        panel = QWidget()
        layout = QFormLayout(panel)
        layout.setContentsMargins(6, 6, 6, 6)
        self.vertical_labels = [QLabel() for _ in range(3)]
        self.field_strategy_combo = self._enum_combo(FieldStrategy)
        layout.addRow(self.vertical_labels[0], self.field_strategy_combo)
        source_box = QVBoxLayout()
        self.add_source_file_check = QCheckBox()
        self.add_source_file_check.setChecked(True)
        self.add_source_sheet_check = QCheckBox()
        self.add_source_sheet_check.setChecked(True)
        source_box.addWidget(self.add_source_file_check)
        source_box.addWidget(self.add_source_sheet_check)
        layout.addRow(self.vertical_labels[1], source_box)
        self.dedupe_mode_combo = self._enum_combo(DedupeMode)
        layout.addRow(self.vertical_labels[2], self.dedupe_mode_combo)
        self.dedupe_fields_list = self._field_list()
        layout.addRow("", self.dedupe_fields_list)
        self.dedupe_mode_combo.currentIndexChanged.connect(self._sync_dedupe)
        return panel

    def _build_join_panel(self) -> QWidget:
        panel = QFrame()
        panel.setObjectName("PrimaryRulePanel")
        layout = QVBoxLayout(panel)
        layout.setContentsMargins(13, 12, 13, 13)
        self.join_primary_title = QLabel()
        self.join_primary_title.setObjectName("PrimaryRuleTitle")
        layout.addWidget(self.join_primary_title)
        form = QFormLayout()
        self.join_labels = [QLabel() for _ in range(9)]
        self.main_table_combo = QComboBox()
        self.lookup_table_combo = QComboBox()
        form.addRow(self.join_labels[0], self.main_table_combo)
        form.addRow(self.join_labels[1], self.lookup_table_combo)
        self.join_type_combo = self._enum_combo(JoinType)
        form.addRow(self.join_labels[2], self.join_type_combo)

        keys = QHBoxLayout()
        left_box = QVBoxLayout()
        right_box = QVBoxLayout()
        left_box.addWidget(self.join_labels[3])
        right_box.addWidget(self.join_labels[4])
        self.left_keys_list = self._field_list(125)
        self.right_keys_list = self._field_list(125)
        left_box.addWidget(self.left_keys_list)
        right_box.addWidget(self.right_keys_list)
        keys.addLayout(left_box, 1)
        keys.addLayout(right_box, 1)
        form.addRow("", keys)
        self.normalize_keys_check = QCheckBox()
        self.normalize_keys_check.setChecked(True)
        form.addRow("", self.normalize_keys_check)
        fuzzy_row = QHBoxLayout()
        self.fuzzy_check = QCheckBox()
        self.fuzzy_threshold_spin = QSpinBox()
        self.fuzzy_threshold_spin.setRange(1, 100)
        self.fuzzy_threshold_spin.setValue(85)
        fuzzy_row.addWidget(self.fuzzy_check)
        fuzzy_row.addStretch(1)
        fuzzy_row.addWidget(self.join_labels[5])
        fuzzy_row.addWidget(self.fuzzy_threshold_spin)
        form.addRow("", fuzzy_row)
        self.duplicate_policy_combo = self._enum_combo(DuplicateKeyPolicy)
        form.addRow(self.join_labels[6], self.duplicate_policy_combo)
        self.join_conflict_combo = self._enum_combo(JoinConflictPolicy)
        form.addRow(self.join_labels[7], self.join_conflict_combo)
        suffix_row = QHBoxLayout()
        self.main_suffix_edit = QLineEdit("_main")
        self.related_suffix_edit = QLineEdit("_related")
        suffix_row.addWidget(self.main_suffix_edit)
        suffix_row.addWidget(self.related_suffix_edit)
        form.addRow(self.join_labels[8], suffix_row)
        layout.addLayout(form)
        self.join_hint = QLabel()
        self.join_hint.setObjectName("Muted")
        self.join_hint.setWordWrap(True)
        layout.addWidget(self.join_hint)
        self.main_table_combo.currentIndexChanged.connect(self._populate_join_fields)
        self.lookup_table_combo.currentIndexChanged.connect(self._populate_join_fields)
        self.fuzzy_check.toggled.connect(self.fuzzy_threshold_spin.setEnabled)
        return panel

    def _build_workbook_panel(self) -> QWidget:
        panel = QWidget()
        layout = QFormLayout(panel)
        layout.setContentsMargins(6, 6, 6, 6)
        self.workbook_labels = [QLabel()]
        self.sheet_conflict_combo = self._enum_combo(SheetConflictPolicy)
        layout.addRow(self.workbook_labels[0], self.sheet_conflict_combo)
        self.workbook_preserve_check = QCheckBox()
        self.workbook_preserve_check.setChecked(True)
        layout.addRow("", self.workbook_preserve_check)
        self.workbook_hint = QLabel()
        self.workbook_hint.setObjectName("Muted")
        self.workbook_hint.setWordWrap(True)
        layout.addRow("", self.workbook_hint)
        return panel

    def _build_same_name_panel(self) -> QWidget:
        panel = QWidget()
        layout = QFormLayout(panel)
        layout.setContentsMargins(6, 6, 6, 6)
        self.same_name_labels = [QLabel(), QLabel()]
        self.target_sheet_combo = QComboBox()
        self.target_sheet_combo.setEditable(True)
        layout.addRow(self.same_name_labels[0], self.target_sheet_combo)
        self.same_name_strategy_combo = self._enum_combo(FieldStrategy)
        layout.addRow(self.same_name_labels[1], self.same_name_strategy_combo)
        self.same_name_hint = QLabel()
        self.same_name_hint.setObjectName("Muted")
        self.same_name_hint.setWordWrap(True)
        layout.addRow("", self.same_name_hint)
        return panel

    def _build_horizontal_panel(self) -> QWidget:
        panel = QWidget()
        layout = QFormLayout(panel)
        layout.setContentsMargins(6, 6, 6, 6)
        self.horizontal_labels = [QLabel()]
        self.side_prefix_check = QCheckBox()
        self.side_prefix_check.setChecked(True)
        layout.addRow("", self.side_prefix_check)
        self.side_gap_spin = QSpinBox()
        self.side_gap_spin.setRange(0, 100)
        self.side_gap_spin.setValue(1)
        layout.addRow(self.horizontal_labels[0], self.side_gap_spin)
        self.horizontal_hint = QLabel()
        self.horizontal_hint.setObjectName("Muted")
        self.horizontal_hint.setWordWrap(True)
        layout.addRow("", self.horizontal_hint)
        return panel

    def _build_aggregate_panel(self) -> QWidget:
        panel = QWidget()
        layout = QFormLayout(panel)
        layout.setContentsMargins(6, 6, 6, 6)
        self.aggregate_labels = [QLabel() for _ in range(4)]
        self.group_fields_list = self._field_list(120)
        self.value_fields_list = self._field_list(120)
        self.aggregate_method_combo = self._enum_combo(AggregateMethod)
        self.aggregate_strategy_combo = self._enum_combo(FieldStrategy)
        layout.addRow(self.aggregate_labels[0], self.group_fields_list)
        layout.addRow(self.aggregate_labels[1], self.value_fields_list)
        layout.addRow(self.aggregate_labels[2], self.aggregate_method_combo)
        layout.addRow(self.aggregate_labels[3], self.aggregate_strategy_combo)
        self.aggregate_hint = QLabel()
        self.aggregate_hint.setObjectName("Muted")
        self.aggregate_hint.setWordWrap(True)
        layout.addRow("", self.aggregate_hint)
        return panel

    @staticmethod
    def _field_list(height: int = 105) -> QListWidget:
        widget = QListWidget()
        widget.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        widget.setMaximumHeight(height)
        return widget

    @staticmethod
    def _enum_combo(enum_type: object) -> QComboBox:
        combo = QComboBox()
        for value in enum_type:
            combo.addItem("", value.value)
        return combo

    def _build_output_card(self) -> None:
        self.output_card = StepCard(3, "amber")
        form = QFormLayout()
        self.output_labels = [QLabel() for _ in range(4)]
        self.output_name_edit = QLineEdit("Combined_Result")
        form.addRow(self.output_labels[0], self.output_name_edit)
        self.output_format_combo = self._enum_combo(OutputFormat)
        form.addRow(self.output_labels[1], self.output_format_combo)
        self.existing_policy_combo = self._enum_combo(ExistingFilePolicy)
        form.addRow(self.output_labels[2], self.existing_policy_combo)
        options = QVBoxLayout()
        self.style_output_check = QCheckBox()
        self.style_output_check.setChecked(True)
        self.create_report_check = QCheckBox()
        self.create_report_check.setChecked(True)
        self.open_when_done_check = QCheckBox()
        self.open_when_done_check.setChecked(True)
        for check in (
            self.style_output_check,
            self.create_report_check,
            self.open_when_done_check,
        ):
            options.addWidget(check)
        form.addRow(self.output_labels[3], options)
        self.output_card.layout.addLayout(form)
        self.output_hint = QLabel()
        self.output_hint.setObjectName("Muted")
        self.output_hint.setWordWrap(True)
        self.output_card.add(self.output_hint)
        self.add_card(self.output_card)
        self.output_format_combo.currentIndexChanged.connect(self._sync_output)

    def add_files(self) -> None:
        paths, _selected = QFileDialog.getOpenFileNames(
            self,
            self.i18n.text("dialog.select_sources"),
            "",
            self.i18n.text("dialog.file_filter"),
        )
        if paths:
            self._scan_paths([Path(path) for path in paths], recursive=False)

    def add_folder(self) -> None:
        path = QFileDialog.getExistingDirectory(self, self.i18n.text("dialog.select_folder"))
        if path:
            self._scan_paths([Path(path)], recursive=False)

    def _paths_dropped(self, paths: object) -> None:
        candidates = [path for path in paths if isinstance(path, Path)]
        if candidates:
            self._scan_paths(candidates, recursive=False)

    def _scan_paths(self, paths: list[Path], recursive: bool) -> None:
        self.append_log(self.i18n.text("log.scanning_sources", count=len(paths)))
        worker = ScanWorker(
            paths,
            ScanOptions(
                scan_subfolders=recursive,
                include_text_files=True,
                ignore_hidden=True,
                ignore_temp_files=True,
            ),
        )
        worker.progress.connect(
            lambda current, total, name: self.append_log(
                self.i18n.text("log.scan_progress", current=current, total=total, file=name)
            )
        )
        self._run_worker(worker, self._scan_completed, self._scan_failed)

    def _scan_completed(self, records: list[FileRecord]) -> None:
        known = {record.path.resolve() for record in self.records}
        for record in records:
            if record.path.resolve() not in known:
                self.records.append(record)
                known.add(record.path.resolve())
        self._populate_file_table()
        if self.records:
            self.file_table.setCurrentCell(0, 1)
            self.analyze_sources()
        self.append_log(self.i18n.text("log.sources_loaded", count=len(self.records)))
        self.configuration_changed.emit()

    def _scan_failed(self, message: str) -> None:
        self.append_log(self.i18n.text("log.scan_failed", error=message))

    def _populate_file_table(self) -> None:
        self._building_file_table = True
        try:
            self.file_table.setRowCount(len(self.records))
            for row, record in enumerate(self.records):
                check = QTableWidgetItem()
                check.setFlags(
                    Qt.ItemFlag.ItemIsUserCheckable
                    | Qt.ItemFlag.ItemIsEnabled
                    | Qt.ItemFlag.ItemIsSelectable
                )
                check.setCheckState(
                    Qt.CheckState.Checked if record.status != "error" else Qt.CheckState.Unchecked
                )
                self.file_table.setItem(row, 0, check)
                self.file_table.setItem(row, 1, QTableWidgetItem(record.name))
                self.file_table.setItem(row, 2, QTableWidgetItem(str(record.path)))
                self.file_table.setItem(row, 3, QTableWidgetItem(str(len(record.sheets))))
                self.file_table.setItem(
                    row,
                    4,
                    QTableWidgetItem(
                        self.i18n.text(
                            "status.ready" if record.status != "error" else "status.error"
                        )
                    ),
                )
        finally:
            self._building_file_table = False
        self._refresh_table_choices()

    def selected_records(self) -> list[FileRecord]:
        selected: list[FileRecord] = []
        for row, record in enumerate(self.records):
            item = self.file_table.item(row, 0)
            if item is not None and item.checkState() == Qt.CheckState.Checked:
                selected.append(record)
        return selected

    def remove_selected(self) -> None:
        row = self.file_table.currentRow()
        if 0 <= row < len(self.records):
            self.records.pop(row)
            self._populate_file_table()
            self.analyze_sources()

    def _move_selected(self, delta: int) -> None:
        row = self.file_table.currentRow()
        target = row + delta
        if row < 0 or target < 0 or target >= len(self.records):
            return
        self.records[row], self.records[target] = self.records[target], self.records[row]
        self._populate_file_table()
        self.file_table.setCurrentCell(target, 1)
        self.analyze_sources()

    def clear_files(self) -> None:
        self.records.clear()
        self.field_catalog = FieldCatalog()
        self._populate_file_table()
        self.preview_panel.table.clear()
        self.preview_panel.table.setRowCount(0)
        self.preview_panel.table.setColumnCount(0)
        self.configuration_changed.emit()

    def _file_current_changed(
        self, current_row: int, _current_column: int, _previous_row: int, _previous_column: int
    ) -> None:
        if current_row < 0 or current_row >= len(self.records):
            return
        record = self.records[current_row]
        sheets = self._sheets_for_record(record)
        sheet = sheets[0] if sheets else None
        self.set_preview_target(record.path, sheet)
        self.refresh_preview()

    def _structure_changed(self, _value: int) -> None:
        self.configuration_changed.emit()
        if self.records:
            self.analyze_sources()

    def _sheet_rule_changed(self, _index: int) -> None:
        named = self.sheet_rule_combo.currentData() == "named"
        self.named_sheets_edit.setVisible(named)
        self.source_labels[2].setVisible(named)
        self._refresh_table_choices()
        self.analyze_sources()

    def _sheets_for_record(self, record: FileRecord) -> list[str]:
        rule = str(self.sheet_rule_combo.currentData())
        available = record.sheets
        if rule == "all":
            return [sheet.name for sheet in available]
        if rule == "visible":
            return [sheet.name for sheet in available if sheet.state == "visible"]
        if rule == "named":
            wanted = {
                value.strip().casefold()
                for value in self.named_sheets_edit.text().split(",")
                if value.strip()
            }
            return [sheet.name for sheet in available if sheet.name.casefold() in wanted]
        return [available[0].name] if available else []

    def selected_sheets(self, records: list[FileRecord] | None = None) -> dict[Path, list[str]]:
        return {
            record.path: self._sheets_for_record(record)
            for record in (records or self.selected_records())
        }

    def analyze_sources(self) -> None:
        selected = self.selected_records()
        requests = [
            FieldDiscoveryRequest(record.path, (self._sheets_for_record(record) or [None])[0])
            for record in selected
        ]
        if not requests:
            self.field_catalog = FieldCatalog()
            self._populate_fields([])
            return
        self.analyze_button.setEnabled(False)
        self.append_log(self.i18n.text("log.analyzing_fields"))
        worker = FieldDiscoveryWorker(
            requests,
            "row_number",
            self.header_spin.value(),
            self.header_spin.value(),
            self.settings.preview_rows,
        )
        self._run_worker(worker, self._analysis_completed, self._analysis_failed)

    def _analysis_completed(self, catalog: FieldCatalog) -> None:
        self.analyze_button.setEnabled(True)
        self.field_catalog = catalog
        fields: list[str] = []
        for source in catalog.sources:
            for field in source.fields:
                if field not in fields:
                    fields.append(field)
        self._populate_fields(fields)
        self.append_log(
            self.i18n.text("log.fields_ready", fields=len(fields), tables=len(catalog.sources))
        )
        for warning in catalog.warnings:
            self.append_log(warning)

    def _analysis_failed(self, message: str) -> None:
        self.analyze_button.setEnabled(True)
        self.append_log(self.i18n.text("log.analysis_failed", error=message))

    def _populate_fields(self, fields: list[str]) -> None:
        for widget in (
            self.dedupe_fields_list,
            self.group_fields_list,
            self.value_fields_list,
        ):
            widget.clear()
            widget.addItems(fields)
        self._refresh_table_choices()
        self._populate_join_fields()

    def _refresh_table_choices(self) -> None:
        selected = self.selected_records()
        current_main = self.main_table_combo.currentData()
        current_lookup = self.lookup_table_combo.currentData()
        for combo in (self.main_table_combo, self.lookup_table_combo):
            combo.blockSignals(True)
            combo.clear()
            for record in selected:
                combo.addItem(record.name, str(record.path))
            combo.blockSignals(False)
        if current_main:
            self.main_table_combo.setCurrentIndex(
                max(0, self.main_table_combo.findData(current_main))
            )
        if current_lookup:
            self.lookup_table_combo.setCurrentIndex(
                max(0, self.lookup_table_combo.findData(current_lookup))
            )
        elif self.lookup_table_combo.count() > 1:
            self.lookup_table_combo.setCurrentIndex(1)
        sheet_names: list[str] = []
        for record in selected:
            for sheet in record.sheets:
                if sheet.name not in sheet_names:
                    sheet_names.append(sheet.name)
        current_sheet = self.target_sheet_combo.currentText()
        self.target_sheet_combo.clear()
        self.target_sheet_combo.addItems(sheet_names)
        if current_sheet:
            self.target_sheet_combo.setCurrentText(current_sheet)

    def _fields_for_path(self, raw_path: object) -> list[str]:
        if not raw_path:
            return []
        wanted = Path(str(raw_path)).resolve()
        for source in self.field_catalog.sources:
            if source.path.resolve() == wanted:
                return source.fields
        return []

    def _populate_join_fields(self, _index: int = -1) -> None:
        left = self._fields_for_path(self.main_table_combo.currentData())
        right = self._fields_for_path(self.lookup_table_combo.currentData())
        self.left_keys_list.clear()
        self.right_keys_list.clear()
        self.left_keys_list.addItems(left)
        self.right_keys_list.addItems(right)
        common = [field for field in left if field in right]
        if common:
            self.left_keys_list.setCurrentRow(left.index(common[0]))
            self.left_keys_list.item(left.index(common[0])).setSelected(True)
            self.right_keys_list.setCurrentRow(right.index(common[0]))
            self.right_keys_list.item(right.index(common[0])).setSelected(True)
        self.configuration_changed.emit()

    def _mode_changed(self, _index: int = -1) -> None:
        mode = self.current_mode()
        mapping = {
            MergeMode.VERTICAL: 0,
            MergeMode.KEY_JOIN: 1,
            MergeMode.WORKBOOK: 2,
            MergeMode.SAME_NAME: 3,
            MergeMode.HORIZONTAL: 4,
            MergeMode.AGGREGATE: 5,
        }
        self.mode_stack.setCurrentIndex(mapping[mode])
        self._sync_output()
        self.configuration_changed.emit()

    def _sync_dedupe(self, _index: int = -1) -> None:
        self.dedupe_fields_list.setVisible(
            self.dedupe_mode_combo.currentData() == DedupeMode.FIELDS.value
        )

    def _sync_output(self, _index: int = -1) -> None:
        locked = self.current_mode() in {MergeMode.WORKBOOK, MergeMode.SAME_NAME}
        if locked:
            self.output_format_combo.setCurrentIndex(
                self.output_format_combo.findData(OutputFormat.XLSX.value)
            )
        self.output_format_combo.setEnabled(not locked)

    def current_mode(self) -> MergeMode:
        return MergeMode(str(self.merge_mode_combo.currentData()))

    @staticmethod
    def _selected_values(widget: QListWidget) -> list[str]:
        return [item.text() for item in widget.selectedItems()]

    def header_row_value(self) -> int:
        return self.header_spin.value()

    def validation_error(self) -> str | None:
        selected = self.selected_records()
        if len(selected) < 2:
            return self.i18n.text("validation.merge_sources")
        if self.current_mode() == MergeMode.KEY_JOIN:
            left = self._selected_values(self.left_keys_list)
            right = self._selected_values(self.right_keys_list)
            if not left or len(left) != len(right):
                return self.i18n.text("validation.join_keys")
        if self.current_mode() == MergeMode.AGGREGATE:
            if not self._selected_values(self.group_fields_list) or not self._selected_values(
                self.value_fields_list
            ):
                return self.i18n.text("validation.aggregate_fields")
        return None

    def _structure(self) -> TableStructure:
        return TableStructure(
            header_mode=HeaderMode.ROW_NUMBER,
            header_row=self.header_spin.value(),
            header_end_row=self.header_spin.value(),
            data_start_row=self.header_spin.value() + 1,
        )

    def _output(self, output_directory: Path) -> OutputOptions:
        return OutputOptions(
            directory=output_directory,
            naming_template=self.output_name_edit.text().strip() or "Combined_Result",
            add_source_file=self.add_source_file_check.isChecked(),
            add_source_sheet=self.add_source_sheet_check.isChecked(),
            output_format=OutputFormat(str(self.output_format_combo.currentData())),
            existing_file_policy=ExistingFilePolicy(str(self.existing_policy_combo.currentData())),
            workbook_name=self.output_name_edit.text().strip() or "Combined_Result",
            preserve_format=self.workbook_preserve_check.isChecked(),
            style_output=self.style_output_check.isChecked(),
            create_report_sheet=self.create_report_check.isChecked(),
            open_when_done=self.open_when_done_check.isChecked(),
        )

    def build_config(self, output_directory: Path) -> MergeTaskConfig | AdvancedMergeTaskConfig:
        error = self.validation_error()
        if error is not None:
            raise ValueError(error)
        selected = self.selected_records()
        mode = self.current_mode()
        output = self._output(output_directory)
        if mode in {MergeMode.KEY_JOIN, MergeMode.HORIZONTAL}:
            inputs = selected
            if mode == MergeMode.KEY_JOIN:
                by_path = {str(record.path): record for record in selected}
                main = by_path.get(str(self.main_table_combo.currentData()), selected[0])
                lookup = by_path.get(
                    str(self.lookup_table_combo.currentData()),
                    selected[1],
                )
                inputs = [main, lookup]
            return AdvancedMergeTaskConfig(
                input_files=[record.path for record in inputs],
                structure=self._structure(),
                mode=mode,
                output=output,
                selected_sheets=self.selected_sheets(inputs),
                join_type=JoinType(str(self.join_type_combo.currentData())),
                left_keys=self._selected_values(self.left_keys_list),
                right_keys=self._selected_values(self.right_keys_list),
                duplicate_policy=DuplicateKeyPolicy(str(self.duplicate_policy_combo.currentData())),
                normalize_keys=self.normalize_keys_check.isChecked(),
                fuzzy_match=self.fuzzy_check.isChecked(),
                fuzzy_threshold=self.fuzzy_threshold_spin.value(),
                conflict_policy=JoinConflictPolicy(str(self.join_conflict_combo.currentData())),
                main_suffix=self.main_suffix_edit.text() or "_main",
                related_suffix=self.related_suffix_edit.text() or "_related",
                side_prefix=self.side_prefix_check.isChecked(),
                side_gap=self.side_gap_spin.value(),
                cleaning=CleaningOptions(),
            )
        strategy = (
            FieldStrategy(str(self.same_name_strategy_combo.currentData()))
            if mode == MergeMode.SAME_NAME
            else FieldStrategy(str(self.field_strategy_combo.currentData()))
        )
        if mode == MergeMode.AGGREGATE:
            strategy = FieldStrategy(str(self.aggregate_strategy_combo.currentData()))
        return MergeTaskConfig(
            input_files=[record.path for record in selected],
            structure=self._structure(),
            merge=MergeOptions(
                mode=mode,
                field_strategy=strategy,
                target_sheet_name=self.target_sheet_combo.currentText().strip(),
                selected_sheets=self.selected_sheets(selected),
                add_source_file=self.add_source_file_check.isChecked(),
                add_source_sheet=self.add_source_sheet_check.isChecked(),
                dedupe_mode=DedupeMode(str(self.dedupe_mode_combo.currentData())),
                dedupe_fields=self._selected_values(self.dedupe_fields_list),
                sheet_conflict_policy=SheetConflictPolicy(
                    str(self.sheet_conflict_combo.currentData())
                ),
                preserve_format=self.workbook_preserve_check.isChecked(),
                aggregate_group_fields=self._selected_values(self.group_fields_list),
                aggregate_value_fields=self._selected_values(self.value_fields_list),
                aggregate_method=AggregateMethod(str(self.aggregate_method_combo.currentData())),
            ),
            output=output,
        )

    @staticmethod
    def _translate_combo(combo: QComboBox, translate: callable, prefix: str) -> None:
        for index in range(combo.count()):
            combo.setItemText(index, translate(f"{prefix}.{combo.itemData(index)}"))

    def retranslate_ui(self) -> None:
        super().retranslate_ui()
        self.source_card.title_label.setText(self.i18n.text("merge.step.source"))
        self.source_card.description_label.setText(self.i18n.text("merge.step.source_hint"))
        button_keys = (
            (self.add_files_button, "merge.source.add_files"),
            (self.add_folder_button, "merge.source.add_folder"),
            (self.remove_button, "merge.source.remove"),
            (self.up_button, "merge.source.up"),
            (self.down_button, "merge.source.down"),
            (self.clear_button, "merge.source.clear"),
            (self.analyze_button, "merge.source.analyze"),
        )
        for button, key in button_keys:
            button.setText(self.i18n.text(key))
        self.file_table.setHorizontalHeaderLabels(
            [
                self.i18n.text("table.select"),
                self.i18n.text("table.file"),
                self.i18n.text("table.path"),
                self.i18n.text("table.sheet_count"),
                self.i18n.text("table.status"),
            ]
        )
        for label, key in zip(
            self.source_labels,
            ("common.header_row", "merge.source.sheet_rule", "merge.source.named_sheets"),
            strict=True,
        ):
            label.setText(self.i18n.text(key))
        self._translate_combo(self.sheet_rule_combo, self.i18n.text, "merge.sheet_rule")
        self.named_sheets_edit.setPlaceholderText(self.i18n.text("merge.source.named_hint"))
        self.source_hint.setText(self.i18n.text("merge.source.hint"))

        self.rule_card.title_label.setText(self.i18n.text("merge.step.rules"))
        self.rule_card.description_label.setText(self.i18n.text("merge.step.rules_hint"))
        self.mode_label.setText(self.i18n.text("merge.mode.label"))
        self._translate_combo(self.merge_mode_combo, self.i18n.text, "merge.mode")

        for label, key in zip(
            self.vertical_labels,
            ("merge.vertical.alignment", "merge.vertical.source_fields", "merge.dedupe.mode"),
            strict=True,
        ):
            label.setText(self.i18n.text(key))
        self._translate_combo(self.field_strategy_combo, self.i18n.text, "field_strategy")
        self.add_source_file_check.setText(self.i18n.text("merge.source_field.file"))
        self.add_source_sheet_check.setText(self.i18n.text("merge.source_field.sheet"))
        self._translate_combo(self.dedupe_mode_combo, self.i18n.text, "dedupe")

        self.join_primary_title.setText(self.i18n.text("merge.join.primary_title"))
        join_keys = (
            "merge.join.main",
            "merge.join.lookup",
            "merge.join.type",
            "merge.join.left_keys",
            "merge.join.right_keys",
            "merge.join.threshold",
            "merge.join.duplicate",
            "merge.join.conflict",
            "merge.join.suffixes",
        )
        for label, key in zip(self.join_labels, join_keys, strict=True):
            label.setText(self.i18n.text(key))
        self._translate_combo(self.join_type_combo, self.i18n.text, "join_type")
        self.normalize_keys_check.setText(self.i18n.text("merge.join.normalize"))
        self.fuzzy_check.setText(self.i18n.text("merge.join.fuzzy"))
        self._translate_combo(self.duplicate_policy_combo, self.i18n.text, "duplicate")
        self._translate_combo(self.join_conflict_combo, self.i18n.text, "join_conflict")
        self.join_hint.setText(self.i18n.text("merge.join.hint"))

        self.workbook_labels[0].setText(self.i18n.text("merge.workbook.conflict"))
        self._translate_combo(self.sheet_conflict_combo, self.i18n.text, "sheet_conflict")
        self.workbook_preserve_check.setText(self.i18n.text("option.preserve_format"))
        self.workbook_hint.setText(self.i18n.text("merge.workbook.hint"))

        self.same_name_labels[0].setText(self.i18n.text("merge.same_name.sheet"))
        self.same_name_labels[1].setText(self.i18n.text("merge.vertical.alignment"))
        self._translate_combo(self.same_name_strategy_combo, self.i18n.text, "field_strategy")
        self.same_name_hint.setText(self.i18n.text("merge.same_name.hint"))

        self.side_prefix_check.setText(self.i18n.text("merge.horizontal.prefix"))
        self.horizontal_labels[0].setText(self.i18n.text("merge.horizontal.gap"))
        self.horizontal_hint.setText(self.i18n.text("merge.horizontal.hint"))

        aggregate_keys = (
            "merge.aggregate.groups",
            "merge.aggregate.values",
            "merge.aggregate.method",
            "merge.vertical.alignment",
        )
        for label, key in zip(self.aggregate_labels, aggregate_keys, strict=True):
            label.setText(self.i18n.text(key))
        self._translate_combo(self.aggregate_method_combo, self.i18n.text, "aggregate")
        self._translate_combo(self.aggregate_strategy_combo, self.i18n.text, "field_strategy")
        self.aggregate_hint.setText(self.i18n.text("merge.aggregate.hint"))

        self.output_card.title_label.setText(self.i18n.text("merge.step.output"))
        self.output_card.description_label.setText(self.i18n.text("merge.step.output_hint"))
        for label, key in zip(
            self.output_labels,
            (
                "merge.output.name",
                "merge.output.format",
                "merge.output.existing",
                "merge.output.options",
            ),
            strict=True,
        ):
            label.setText(self.i18n.text(key))
        self._translate_combo(self.output_format_combo, self.i18n.text, "format")
        self._translate_combo(self.existing_policy_combo, self.i18n.text, "existing")
        self.style_output_check.setText(self.i18n.text("option.style_output"))
        self.create_report_check.setText(self.i18n.text("option.create_report"))
        self.open_when_done_check.setText(self.i18n.text("option.open_when_done"))
        self.output_hint.setText(self.i18n.text("merge.output.hint"))
        self._populate_file_table()
        self._sync_dedupe()
        self._sheet_rule_changed(self.sheet_rule_combo.currentIndex())
