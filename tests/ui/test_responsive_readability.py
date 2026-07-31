from __future__ import annotations

from pytestqt.qtbot import QtBot

from excel_studio.ui.main_window import MainWindow


def _text_fits(label: object) -> bool:
    return label.fontMetrics().horizontalAdvance(label.text()) <= label.width()


def test_1366_layout_keeps_steps_and_fixed_actions_accessible(
    product_window: MainWindow, qtbot: QtBot
) -> None:
    product_window.resize(1366, 768)
    qtbot.wait(50)
    page = product_window.split_page

    assert page.source_card.geometry().top() < page.rule_card.geometry().top()
    assert page.rule_card.geometry().top() < page.left_scroll.viewport().height()
    assert product_window.footer.isVisible()
    assert (
        product_window.footer.geometry().bottom()
        <= product_window.centralWidget().geometry().bottom()
    )
    assert _text_fits(product_window.title_label)
    assert _text_fits(product_window.subtitle_label)


def test_english_header_remains_readable_at_supported_minimum(
    product_window: MainWindow, qtbot: QtBot
) -> None:
    product_window.i18n.set_language("en_US")
    product_window.resize(1040, 640)
    qtbot.wait(50)

    assert _text_fits(product_window.title_label)
    assert _text_fits(product_window.subtitle_label)
    assert product_window.split_workspace_button.isVisible()
    assert product_window.merge_workspace_button.isVisible()
    assert product_window.language_combo.isVisible()
