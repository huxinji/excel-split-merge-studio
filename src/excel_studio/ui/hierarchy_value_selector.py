from __future__ import annotations

from collections.abc import Iterable

from PySide6.QtCore import Qt, Signal
from PySide6.QtWidgets import (
    QDialog,
    QDialogButtonBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QPushButton,
    QVBoxLayout,
    QWidget,
)

from excel_studio.config.i18n import TranslationManager


def _unique_values(values: Iterable[str]) -> list[str]:
    unique: list[str] = []
    seen: set[str] = set()
    for value in values:
        cleaned = value.strip()
        key = cleaned.casefold()
        if not cleaned or key in seen:
            continue
        seen.add(key)
        unique.append(cleaned)
    return unique


class HierarchyValueDialog(QDialog):
    def __init__(
        self,
        i18n: TranslationManager,
        values: list[str],
        selected: list[str],
        parent: QWidget | None = None,
    ) -> None:
        super().__init__(parent)
        self.i18n = i18n
        self.setModal(True)
        self.setMinimumSize(420, 360)
        available = self.screen().availableGeometry()
        self.resize(min(560, available.width() - 48), min(460, available.height() - 72))
        self.setWindowTitle(self.i18n.text("split.hierarchy.multiselect.title"))

        layout = QVBoxLayout(self)
        layout.setContentsMargins(16, 16, 16, 14)
        layout.setSpacing(10)
        hint = QLabel(self.i18n.text("split.hierarchy.multiselect.hint"))
        hint.setObjectName("Muted")
        hint.setWordWrap(True)
        layout.addWidget(hint)

        self.search_edit = QLineEdit()
        self.search_edit.setPlaceholderText(self.i18n.text("split.hierarchy.multiselect.search"))
        self.search_edit.textChanged.connect(self._filter_items)
        layout.addWidget(self.search_edit)

        selected_set = set(selected)
        self.list_widget = QListWidget()
        self.list_widget.setAlternatingRowColors(True)
        for value in _unique_values([*values, *selected]):
            item = QListWidgetItem(value)
            item.setFlags(item.flags() | Qt.ItemFlag.ItemIsUserCheckable)
            item.setCheckState(
                Qt.CheckState.Checked if value in selected_set else Qt.CheckState.Unchecked
            )
            self.list_widget.addItem(item)
        layout.addWidget(self.list_widget, 1)

        quick_row = QHBoxLayout()
        self.select_visible_button = QPushButton(
            self.i18n.text("split.hierarchy.multiselect.select_visible")
        )
        self.clear_button = QPushButton(self.i18n.text("split.hierarchy.multiselect.clear"))
        self.select_visible_button.clicked.connect(self._select_visible)
        self.clear_button.clicked.connect(self._clear_all)
        quick_row.addWidget(self.select_visible_button)
        quick_row.addWidget(self.clear_button)
        quick_row.addStretch(1)
        layout.addLayout(quick_row)

        custom_row = QHBoxLayout()
        self.custom_edit = QLineEdit()
        self.custom_edit.setPlaceholderText(self.i18n.text("split.hierarchy.multiselect.custom"))
        self.add_button = QPushButton(self.i18n.text("split.hierarchy.multiselect.add"))
        self.add_button.clicked.connect(self._add_custom_value)
        self.custom_edit.returnPressed.connect(self._add_custom_value)
        custom_row.addWidget(self.custom_edit, 1)
        custom_row.addWidget(self.add_button)
        layout.addLayout(custom_row)

        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
        )
        ok_button = buttons.button(QDialogButtonBox.StandardButton.Ok)
        cancel_button = buttons.button(QDialogButtonBox.StandardButton.Cancel)
        if ok_button is not None:
            ok_button.setText(self.i18n.text("common.confirm"))
        if cancel_button is not None:
            cancel_button.setText(self.i18n.text("common.cancel"))
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def _filter_items(self, text: str) -> None:
        query = text.strip().casefold()
        for index in range(self.list_widget.count()):
            item = self.list_widget.item(index)
            item.setHidden(bool(query) and query not in item.text().casefold())

    def _select_visible(self) -> None:
        for index in range(self.list_widget.count()):
            item = self.list_widget.item(index)
            if not item.isHidden():
                item.setCheckState(Qt.CheckState.Checked)

    def _clear_all(self) -> None:
        for index in range(self.list_widget.count()):
            self.list_widget.item(index).setCheckState(Qt.CheckState.Unchecked)

    def _add_custom_value(self) -> None:
        value = self.custom_edit.text().strip()
        if not value:
            return
        for index in range(self.list_widget.count()):
            item = self.list_widget.item(index)
            if item.text().casefold() == value.casefold():
                item.setCheckState(Qt.CheckState.Checked)
                self.custom_edit.clear()
                return
        item = QListWidgetItem(value)
        item.setFlags(item.flags() | Qt.ItemFlag.ItemIsUserCheckable)
        item.setCheckState(Qt.CheckState.Checked)
        self.list_widget.addItem(item)
        self.custom_edit.clear()

    def selected_values(self) -> list[str]:
        return [
            self.list_widget.item(index).text()
            for index in range(self.list_widget.count())
            if self.list_widget.item(index).checkState() == Qt.CheckState.Checked
        ]


class HierarchyMultiValueSelector(QWidget):
    changed = Signal()

    def __init__(
        self,
        i18n: TranslationManager,
        parent: QWidget | None = None,
    ) -> None:
        super().__init__(parent)
        self.i18n = i18n
        self._values: list[str] = []
        self._selected: list[str] = []
        layout = QHBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(7)
        self.summary_edit = QLineEdit()
        self.summary_edit.setReadOnly(True)
        self.choose_button = QPushButton()
        self.choose_button.setObjectName("SecondaryButton")
        self.choose_button.clicked.connect(self._open_dialog)
        layout.addWidget(self.summary_edit, 1)
        layout.addWidget(self.choose_button)
        self.retranslate_ui()

    def set_values(self, values: Iterable[str]) -> None:
        self._values = _unique_values([*values, *self._selected])
        self._update_summary()

    def set_selected_values(
        self,
        values: Iterable[str],
        *,
        emit: bool = True,
    ) -> None:
        selected = _unique_values(values)
        if selected == self._selected:
            return
        self._selected = selected
        self._values = _unique_values([*self._values, *selected])
        self._update_summary()
        if emit:
            self.changed.emit()

    def selected_values(self) -> list[str]:
        return list(self._selected)

    def available_values(self) -> list[str]:
        return list(self._values)

    def clear_selection(self, *, emit: bool = True) -> None:
        self.set_selected_values([], emit=emit)

    def replace_value(self, old_value: str, new_value: str) -> None:
        self._values = _unique_values(
            new_value if value == old_value else value for value in self._values
        )
        self._selected = _unique_values(
            new_value if value == old_value else value for value in self._selected
        )
        self._update_summary()

    def _open_dialog(self) -> None:
        dialog = HierarchyValueDialog(
            self.i18n,
            self._values,
            self._selected,
            self,
        )
        if dialog.exec() == QDialog.DialogCode.Accepted:
            self.set_selected_values(dialog.selected_values())

    def _separator(self) -> str:
        return "、" if self.i18n.language == "zh_CN" else ", "

    def _update_summary(self) -> None:
        separator = self._separator()
        if not self._selected:
            text = self.i18n.text("split.hierarchy.multiselect.none")
        elif len(self._selected) <= 2:
            text = separator.join(self._selected)
        else:
            text = self.i18n.text(
                "split.hierarchy.multiselect.summary",
                count=len(self._selected),
                values=separator.join(self._selected[:2]),
            )
        self.summary_edit.setText(text)
        self.summary_edit.setToolTip(separator.join(self._selected))

    def retranslate_ui(self) -> None:
        self.choose_button.setText(self.i18n.text("split.hierarchy.multiselect.choose"))
        self._update_summary()
