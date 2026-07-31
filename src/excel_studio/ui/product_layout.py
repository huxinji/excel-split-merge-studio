from __future__ import annotations

from typing import TYPE_CHECKING

from PySide6.QtWidgets import QHeaderView

from excel_studio.ui.split_page import SplitPage

if TYPE_CHECKING:
    from excel_studio.ui.main_window import MainWindow


def prioritize_field_split_workflow(page: SplitPage) -> None:
    """Keep source selection first and field splitting immediately after it."""
    ordered_cards = (
        page.source_card,
        page.rule_card,
        page.sheet_card,
        page.output_card,
    )
    for card in ordered_cards:
        page.left_layout.removeWidget(card)
    for index, card in enumerate(ordered_cards):
        page.left_layout.insertWidget(index, card)
        card.number_label.setText(str(index + 1))

    page.source_card.layout.setContentsMargins(14, 10, 14, 12)
    page.source_card.layout.setSpacing(7)
    page.rule_card.layout.setContentsMargins(14, 10, 14, 12)
    page.rule_card.layout.setSpacing(8)
    page.field_value_table.setMaximumHeight(112)
    page.field_value_table.horizontalHeader().setSectionResizeMode(
        0, QHeaderView.ResizeMode.Stretch
    )


def protect_header_legibility(window: MainWindow) -> None:
    """Reserve enough width for translated header text instead of clipping it."""

    def update_widths(_language: str = "") -> None:
        for label in (window.title_label, window.subtitle_label):
            width = label.fontMetrics().horizontalAdvance(label.text()) + 6
            label.setMinimumWidth(width)

    update_widths()
    window.i18n.language_changed.connect(update_widths)


def apply_accessible_contrast(window: MainWindow) -> None:
    """Keep disabled and secondary text readable on light translucent surfaces."""
    window.setStyleSheet(
        window.styleSheet()
        + """
        QPushButton#PrimaryButton:disabled,
        QLineEdit:disabled,
        QComboBox:disabled,
        QSpinBox:disabled,
        QCheckBox:disabled,
        QTableWidget:disabled,
        QListWidget:disabled,
        QPlainTextEdit:disabled {
            color: #52627A;
        }
        """
    )
