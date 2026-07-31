from __future__ import annotations

from collections.abc import Iterable

from PySide6.QtCore import Qt, Signal
from PySide6.QtWidgets import (
    QAbstractItemView,
    QButtonGroup,
    QFrame,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QPlainTextEdit,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from excel_studio.config.i18n import TranslationManager
from excel_studio.models.workbook import PreviewResult


class StepCard(QFrame):
    def __init__(self, step: int, accent: str = "blue") -> None:
        super().__init__()
        self.setObjectName("StepCard")
        self.setProperty("accent", accent)
        self.layout = QVBoxLayout(self)
        self.layout.setContentsMargins(16, 14, 16, 16)
        self.layout.setSpacing(11)

        heading = QHBoxLayout()
        self.number_label = QLabel(str(step))
        self.number_label.setObjectName("StepNumber")
        self.number_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.title_label = QLabel()
        self.title_label.setObjectName("CardTitle")
        self.description_label = QLabel()
        self.description_label.setObjectName("CardDescription")
        self.description_label.setWordWrap(True)

        title_box = QVBoxLayout()
        title_box.setSpacing(2)
        title_box.addWidget(self.title_label)
        title_box.addWidget(self.description_label)
        heading.addWidget(self.number_label)
        heading.addSpacing(3)
        heading.addLayout(title_box, 1)
        self.layout.addLayout(heading)

    def add(self, item: QWidget) -> None:
        self.layout.addWidget(item)


class SegmentedSelector(QFrame):
    changed = Signal(str)

    def __init__(self, values: Iterable[str]) -> None:
        super().__init__()
        self.setObjectName("SegmentedControl")
        self._group = QButtonGroup(self)
        self._group.setExclusive(True)
        self._buttons: dict[str, QPushButton] = {}
        layout = QHBoxLayout(self)
        layout.setContentsMargins(3, 3, 3, 3)
        layout.setSpacing(3)
        for value in values:
            button = QPushButton()
            button.setObjectName("SegmentButton")
            button.setCheckable(True)
            button.setProperty("value", value)
            layout.addWidget(button)
            self._group.addButton(button)
            self._buttons[value] = button
            button.clicked.connect(lambda _checked=False, item=value: self.changed.emit(item))

    def button(self, value: str) -> QPushButton:
        return self._buttons[value]

    def set_current(self, value: str) -> None:
        if value in self._buttons:
            self._buttons[value].setChecked(True)
            self.changed.emit(value)

    def current(self) -> str:
        checked = self._group.checkedButton()
        return str(checked.property("value")) if checked is not None else ""


class PreviewPanel(QFrame):
    refresh_requested = Signal()
    sheet_changed = Signal(str)

    def __init__(self, i18n: TranslationManager) -> None:
        super().__init__()
        self.i18n = i18n
        self.setObjectName("InspectionPanel")
        layout = QVBoxLayout(self)
        layout.setContentsMargins(14, 14, 14, 14)
        layout.setSpacing(9)

        title_row = QHBoxLayout()
        self.title_label = QLabel()
        self.title_label.setObjectName("CardTitle")
        self.refresh_button = QPushButton()
        self.refresh_button.setObjectName("SecondaryButton")
        self.refresh_button.clicked.connect(self.refresh_requested)
        title_row.addWidget(self.title_label)
        title_row.addStretch(1)
        title_row.addWidget(self.refresh_button)
        layout.addLayout(title_row)

        self.hint_label = QLabel()
        self.hint_label.setObjectName("Muted")
        self.hint_label.setWordWrap(True)
        layout.addWidget(self.hint_label)

        metrics = QHBoxLayout()
        self.metric_values: dict[str, QLabel] = {}
        self.metric_labels: dict[str, QLabel] = {}
        for key in ("file", "sheet", "rows", "columns"):
            box = QVBoxLayout()
            value = QLabel("—")
            value.setObjectName("MetricValue")
            label = QLabel()
            label.setObjectName("MetricLabel")
            box.addWidget(value)
            box.addWidget(label)
            metrics.addLayout(box, 1)
            self.metric_values[key] = value
            self.metric_labels[key] = label
        layout.addLayout(metrics)

        self.table = QTableWidget(0, 0)
        self.table.setAlternatingRowColors(True)
        self.table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        self.table.verticalHeader().setVisible(False)
        layout.addWidget(self.table, 1)
        self.empty_label = QLabel()
        self.empty_label.setObjectName("Muted")
        self.empty_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.empty_label.setWordWrap(True)
        layout.addWidget(self.empty_label)
        self.retranslate_ui()

    def set_preview(self, result: PreviewResult) -> None:
        self.table.clear()
        self.table.setColumnCount(len(result.columns))
        self.table.setRowCount(len(result.rows))
        self.table.setHorizontalHeaderLabels(result.columns)
        for row_index, row in enumerate(result.rows):
            for column_index, value in enumerate(row):
                item = QTableWidgetItem("" if value is None else str(value))
                self.table.setItem(row_index, column_index, item)
        for column in range(len(result.columns)):
            self.table.setColumnWidth(column, min(220, max(105, self.table.columnWidth(column))))
        self.metric_values["file"].setText(result.path.name)
        self.metric_values["sheet"].setText(result.sheet_name)
        self.metric_values["rows"].setText(f"{result.total_rows:,}")
        self.metric_values["columns"].setText(str(result.total_columns))
        self.empty_label.setVisible(not result.columns)

    def set_busy(self, busy: bool) -> None:
        self.refresh_button.setEnabled(not busy)
        if busy:
            self.empty_label.setText(self.i18n.text("preview.loading"))
            self.empty_label.setVisible(True)

    def retranslate_ui(self) -> None:
        self.title_label.setText(self.i18n.text("preview.title"))
        self.refresh_button.setText(self.i18n.text("preview.refresh"))
        self.hint_label.setText(self.i18n.text("preview.hint"))
        for key, label in self.metric_labels.items():
            label.setText(self.i18n.text(f"preview.metric.{key}"))
        if self.table.columnCount() == 0:
            self.empty_label.setText(self.i18n.text("preview.empty"))
            self.empty_label.setVisible(True)


class LogPanel(QFrame):
    def __init__(self, i18n: TranslationManager) -> None:
        super().__init__()
        self.i18n = i18n
        self.setObjectName("InspectionPanel")
        layout = QVBoxLayout(self)
        layout.setContentsMargins(12, 10, 12, 12)
        layout.setSpacing(6)
        row = QHBoxLayout()
        self.title_label = QLabel()
        self.title_label.setObjectName("CardTitle")
        self.clear_button = QPushButton()
        self.clear_button.setObjectName("DangerButton")
        self.clear_button.clicked.connect(self.clear)
        row.addWidget(self.title_label)
        row.addStretch(1)
        row.addWidget(self.clear_button)
        layout.addLayout(row)
        self.text = QPlainTextEdit()
        self.text.setReadOnly(True)
        self.text.setMaximumBlockCount(1500)
        self.text.setMinimumHeight(105)
        layout.addWidget(self.text)
        self.retranslate_ui()

    def append(self, message: str) -> None:
        self.text.appendPlainText(message)
        bar = self.text.verticalScrollBar()
        bar.setValue(bar.maximum())

    def clear(self) -> None:
        self.text.clear()

    def retranslate_ui(self) -> None:
        self.title_label.setText(self.i18n.text("log.title"))
        self.clear_button.setText(self.i18n.text("log.clear"))
