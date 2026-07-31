from __future__ import annotations

from pathlib import Path

from PySide6.QtCore import Slot
from PySide6.QtWidgets import (
    QComboBox,
    QFrame,
    QHBoxLayout,
    QLabel,
    QStackedWidget,
    QVBoxLayout,
    QWidget,
)

from excel_studio.models.operations import BlankValuePolicy, SplitMode, SplitTaskConfig
from excel_studio.models.task import HeaderMode, TableStructure
from excel_studio.models.workbook import PreviewResult
from excel_studio.services.hierarchy_split import HierarchyCatalog
from excel_studio.ui.hierarchy_value_selector import HierarchyMultiValueSelector
from excel_studio.ui.split_page import SplitPage
from excel_studio.ui.widgets import SegmentedSelector, StepCard
from excel_studio.workers.hierarchy_worker import HierarchyValuesWorker


class HierarchySplitPage(SplitPage):
    """Split page with searchable multi-value hierarchy scope selectors."""

    def _build_rule_card(self) -> None:
        self.rule_card = StepCard(3, "violet")
        self.mode_selector = SegmentedSelector(
            (
                SplitMode.BY_FIELD.value,
                SplitMode.HIERARCHY.value,
                SplitMode.FIXED_ROWS.value,
                SplitMode.BY_PARTS.value,
                SplitMode.BY_SHEET.value,
            )
        )
        self.mode_selector.changed.connect(self._mode_changed)
        self.rule_card.add(self.mode_selector)

        self.mode_stack = QStackedWidget()
        self.field_panel = self._build_field_panel()
        self.hierarchy_panel = self._build_hierarchy_panel()
        self.rows_panel = self._build_rows_panel()
        self.parts_panel = self._build_parts_panel()
        self.sheet_panel = self._build_sheet_panel()
        for panel in (
            self.field_panel,
            self.hierarchy_panel,
            self.rows_panel,
            self.parts_panel,
            self.sheet_panel,
        ):
            self.mode_stack.addWidget(panel)
        self.rule_card.add(self.mode_stack)
        self.add_card(self.rule_card)

    def _build_hierarchy_panel(self) -> QWidget:
        panel = QFrame()
        panel.setObjectName("PrimaryRulePanel")
        layout = QVBoxLayout(panel)
        layout.setContentsMargins(14, 13, 14, 14)
        layout.setSpacing(9)

        self.hierarchy_primary_title = QLabel()
        self.hierarchy_primary_title.setObjectName("PrimaryRuleTitle")
        layout.addWidget(self.hierarchy_primary_title)

        count_row = QHBoxLayout()
        self.hierarchy_count_label = QLabel()
        self.hierarchy_count_combo = QComboBox()
        self.hierarchy_count_combo.addItem("", 2)
        self.hierarchy_count_combo.addItem("", 3)
        self.hierarchy_count_combo.setCurrentIndex(0)
        count_row.addWidget(self.hierarchy_count_label)
        count_row.addWidget(self.hierarchy_count_combo)
        count_row.addStretch(1)
        layout.addLayout(count_row)

        self.hierarchy_rows: list[QFrame] = []
        self.hierarchy_level_labels: list[QLabel] = []
        self.hierarchy_field_combos: list[QComboBox] = []
        self.hierarchy_scope_labels: list[QLabel] = []
        self.hierarchy_value_combos: list[HierarchyMultiValueSelector] = []
        self.hierarchy_action_labels: list[QLabel] = []
        for level_index in range(3):
            row_frame = QFrame()
            row_layout = QHBoxLayout(row_frame)
            row_layout.setContentsMargins(0, 0, 0, 0)
            row_layout.setSpacing(8)

            level_label = QLabel()
            level_label.setMinimumWidth(82)
            field_combo = QComboBox()
            field_combo.setMinimumWidth(170)
            field_combo.setEnabled(False)
            scope_label = QLabel()
            scope_label.setMinimumWidth(72)
            value_selector = HierarchyMultiValueSelector(self.i18n)
            value_selector.setMinimumWidth(260)
            action_label = QLabel()
            action_label.setObjectName("PrimaryRuleTitle")

            row_layout.addWidget(level_label)
            row_layout.addWidget(field_combo, 1)
            row_layout.addWidget(scope_label)
            row_layout.addWidget(value_selector, 1)
            row_layout.addWidget(action_label, 1)
            layout.addWidget(row_frame)

            self.hierarchy_rows.append(row_frame)
            self.hierarchy_level_labels.append(level_label)
            self.hierarchy_field_combos.append(field_combo)
            self.hierarchy_scope_labels.append(scope_label)
            self.hierarchy_value_combos.append(value_selector)
            self.hierarchy_action_labels.append(action_label)

            field_combo.currentIndexChanged.connect(
                lambda _index, level=level_index: self._hierarchy_fields_changed(level)
            )
            value_selector.changed.connect(
                lambda level=level_index: self._hierarchy_filter_changed(level)
            )

        self.hierarchy_summary = QLabel()
        self.hierarchy_summary.setObjectName("Muted")
        self.hierarchy_summary.setWordWrap(True)
        layout.addWidget(self.hierarchy_summary)
        self.hierarchy_hint = QLabel()
        self.hierarchy_hint.setObjectName("Muted")
        self.hierarchy_hint.setWordWrap(True)
        layout.addWidget(self.hierarchy_hint)

        self._hierarchy_request_id = 0
        self._last_blank_display = ""
        self.hierarchy_count_combo.currentIndexChanged.connect(self._hierarchy_level_count_changed)
        self._update_hierarchy_structure()
        return panel

    def _mode_changed(self, value: str) -> None:
        mapping = {
            SplitMode.BY_FIELD.value: 0,
            SplitMode.HIERARCHY.value: 1,
            SplitMode.FIXED_ROWS.value: 2,
            SplitMode.BY_PARTS.value: 3,
            SplitMode.BY_SHEET.value: 4,
        }
        self.mode_stack.setCurrentIndex(mapping.get(value, 0))
        if value == SplitMode.BY_SHEET.value:
            super()._mode_changed(value)
            return
        self._sync_output_constraints()
        if value == SplitMode.HIERARCHY.value:
            self._request_hierarchy_values()
        self.configuration_changed.emit()

    def _hierarchy_level_count(self) -> int:
        value = self.hierarchy_count_combo.currentData()
        return int(value) if value is not None else 2

    def _hierarchy_fields(self) -> list[str]:
        count = self._hierarchy_level_count()
        return [combo.currentText().strip() for combo in self.hierarchy_field_combos[:count]]

    def _hierarchy_filters(
        self,
        *,
        engine_values: bool,
    ) -> dict[str, list[str]]:
        fields = self._hierarchy_fields()
        filters: dict[str, list[str]] = {}
        blank_display = self.i18n.text("split.field.blank_value")
        for index, field_name in enumerate(fields[:-1]):
            selected = self.hierarchy_value_combos[index].selected_values()
            if selected:
                filters[field_name] = [
                    "EMPTY" if engine_values and value == blank_display else value
                    for value in selected
                ]
        return filters

    def _update_hierarchy_structure(self) -> None:
        count = self._hierarchy_level_count()
        for index, row in enumerate(self.hierarchy_rows):
            visible = index < count
            row.setVisible(visible)
            if not visible:
                continue
            parent_level = index < count - 1
            self.hierarchy_scope_labels[index].setText(
                self.i18n.text(
                    "split.hierarchy.filter_value"
                    if parent_level
                    else "split.hierarchy.output_action"
                )
            )
            self.hierarchy_value_combos[index].setVisible(parent_level)
            self.hierarchy_action_labels[index].setVisible(not parent_level)
            self.hierarchy_action_labels[index].setText(
                self.i18n.text("split.hierarchy.split_here")
            )
        self.hierarchy_summary.setText(self.i18n.text("split.hierarchy.summary.pending"))

    def _hierarchy_level_count_changed(self, _index: int = -1) -> None:
        count = self._hierarchy_level_count()
        for selector in self.hierarchy_value_combos[count - 1 :]:
            selector.clear_selection(emit=False)
        self._update_hierarchy_structure()
        self._request_hierarchy_values()
        self.configuration_changed.emit()

    def _hierarchy_fields_changed(self, level: int) -> None:
        for selector in self.hierarchy_value_combos[level:]:
            selector.clear_selection(emit=False)
        self._request_hierarchy_values()
        self.configuration_changed.emit()

    def _hierarchy_filter_changed(self, level: int) -> None:
        parent_count = self._hierarchy_level_count() - 1
        for selector in self.hierarchy_value_combos[level + 1 : parent_count]:
            selector.clear_selection(emit=False)
        self._request_hierarchy_values()
        self.configuration_changed.emit()

    def _populate_hierarchy_fields(self, fields: list[str]) -> None:
        for index, combo in enumerate(self.hierarchy_field_combos):
            previous = combo.currentText()
            combo.blockSignals(True)
            combo.clear()
            combo.addItems(fields)
            if previous in fields:
                combo.setCurrentText(previous)
            elif index < len(fields):
                combo.setCurrentIndex(index)
            else:
                combo.setCurrentIndex(-1)
            combo.setEnabled(bool(fields))
            combo.blockSignals(False)
        self._update_hierarchy_structure()
        if self.current_mode() == SplitMode.HIERARCHY:
            self._request_hierarchy_values()

    def on_preview_ready(self, result: PreviewResult) -> None:
        super().on_preview_ready(result)
        fields = [self.field_combo.itemText(index) for index in range(self.field_combo.count())]
        self._populate_hierarchy_fields(fields)

    def _request_hierarchy_values(self) -> None:
        if (
            self.current_mode() != SplitMode.HIERARCHY
            or self.record is None
            or not self.preview_sheet
        ):
            return
        fields = self._hierarchy_fields()
        if (
            len(fields) != self._hierarchy_level_count()
            or any(not field for field in fields)
            or len(set(fields)) != len(fields)
        ):
            self.hierarchy_summary.setText(self.i18n.text("split.hierarchy.summary.fields"))
            return

        self._hierarchy_request_id += 1
        request_id = self._hierarchy_request_id
        self.hierarchy_summary.setText(self.i18n.text("split.hierarchy.summary.loading"))
        worker = HierarchyValuesWorker(
            request_id,
            self.record.path,
            self.preview_sheet,
            TableStructure(
                header_mode=HeaderMode.ROW_NUMBER,
                header_row=self.header_spin.value(),
                header_end_row=self.header_spin.value(),
                data_start_row=self.header_spin.value() + 1,
            ),
            fields,
            self._hierarchy_filters(engine_values=False),
            fields[-1],
            self.i18n.text("split.field.blank_value"),
            BlankValuePolicy(str(self.blank_policy_combo.currentData())),
        )
        self._run_worker(
            worker,
            self._hierarchy_values_ready,
            self._hierarchy_values_failed,
        )

    @Slot(int, object)
    def _hierarchy_values_ready(
        self,
        request_id: int,
        catalog: HierarchyCatalog,
    ) -> None:
        if request_id != self._hierarchy_request_id:
            return
        fields = self._hierarchy_fields()
        for index, field_name in enumerate(fields[:-1]):
            selector = self.hierarchy_value_combos[index]
            selector.set_values(catalog.candidate_values.get(field_name, []))

        if not catalog.complete_filters:
            self.hierarchy_summary.setText(self.i18n.text("split.hierarchy.summary.pending"))
        elif catalog.matched_rows == 0:
            self.hierarchy_summary.setText(self.i18n.text("split.hierarchy.summary.empty"))
        else:
            self.hierarchy_summary.setText(
                self.i18n.text(
                    "split.hierarchy.summary.ready",
                    rows=catalog.matched_rows,
                    field=fields[-1],
                    groups=catalog.target_count,
                )
            )
        if catalog.truncated_fields:
            self.hierarchy_summary.setText(
                self.hierarchy_summary.text()
                + " "
                + self.i18n.text("split.hierarchy.summary.truncated")
            )

    @Slot(int, str)
    def _hierarchy_values_failed(self, request_id: int, message: str) -> None:
        if request_id != self._hierarchy_request_id:
            return
        self.hierarchy_summary.setText(
            self.i18n.text("split.hierarchy.summary.failed", error=message)
        )

    def validation_error(self) -> str | None:
        error = super().validation_error()
        if error is not None:
            return error
        if self.current_mode() != SplitMode.HIERARCHY:
            return None
        fields = self._hierarchy_fields()
        count = self._hierarchy_level_count()
        if len(fields) != count or any(not field for field in fields):
            return self.i18n.text("validation.hierarchy_fields")
        if len(set(fields)) != len(fields):
            return self.i18n.text("validation.hierarchy_unique")
        for index, field_name in enumerate(fields[:-1]):
            if not self.hierarchy_value_combos[index].selected_values():
                return self.i18n.text(
                    "validation.hierarchy_filter",
                    level=index + 1,
                    field=field_name,
                )
        return None

    def build_config(self, output_directory: Path) -> SplitTaskConfig:
        config = super().build_config(output_directory)
        if config.split.mode == SplitMode.HIERARCHY:
            fields = self._hierarchy_fields()
            config.split.hierarchy_fields = fields
            config.split.hierarchy_filters = self._hierarchy_filters(engine_values=True)
            config.split.hierarchy_split_field = fields[-1]
        return config

    def retranslate_ui(self) -> None:
        previous_blank = getattr(self, "_last_blank_display", "")
        super().retranslate_ui()
        current_blank = self.i18n.text("split.field.blank_value")
        if previous_blank and previous_blank != current_blank:
            for selector in self.hierarchy_value_combos:
                selector.replace_value(previous_blank, current_blank)
        self._last_blank_display = current_blank

        self.hierarchy_primary_title.setText(self.i18n.text("split.hierarchy.primary_title"))
        self.hierarchy_count_label.setText(self.i18n.text("split.hierarchy.level_count"))
        self._set_combo_texts(
            self.hierarchy_count_combo,
            {
                "2": self.i18n.text("split.hierarchy.levels.two"),
                "3": self.i18n.text("split.hierarchy.levels.three"),
            },
        )
        for index, label in enumerate(self.hierarchy_level_labels, start=1):
            label.setText(self.i18n.text("split.hierarchy.level_field", level=index))
        for selector in self.hierarchy_value_combos:
            selector.retranslate_ui()
        self.hierarchy_hint.setText(self.i18n.text("split.hierarchy.hint"))
        self._update_hierarchy_structure()
