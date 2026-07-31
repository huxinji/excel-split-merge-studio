from __future__ import annotations

from excel_studio.ui.main_window import MainWindow
from excel_studio.ui.theme import product_stylesheet


def _relative_luminance(color: str) -> float:
    channels = [int(color[index : index + 2], 16) / 255 for index in (1, 3, 5)]
    linear = [
        channel / 12.92 if channel <= 0.04045 else ((channel + 0.055) / 1.055) ** 2.4
        for channel in channels
    ]
    return 0.2126 * linear[0] + 0.7152 * linear[1] + 0.0722 * linear[2]


def _contrast_ratio(foreground: str, background: str) -> float:
    lighter, darker = sorted(
        (_relative_luminance(foreground), _relative_luminance(background)),
        reverse=True,
    )
    return (lighter + 0.05) / (darker + 0.05)


def test_source_is_first_and_field_split_is_second(product_window: MainWindow) -> None:
    page = product_window.split_page
    assert page.left_layout.itemAt(0).widget() is page.source_card
    assert page.left_layout.itemAt(1).widget() is page.rule_card
    assert page.source_card.number_label.text() == "1"
    assert page.rule_card.number_label.text() == "2"


def test_liquid_glass_palette_keeps_text_contrast() -> None:
    stylesheet = product_stylesheet()
    pairs = (
        ("#172033", "#FFFFFF"),
        ("#45566F", "#FFFFFF"),
        ("#52627A", "#FFFFFF"),
        ("#667085", "#FFFFFF"),
        ("#FFFFFF", "#0068D9"),
        ("#B42318", "#FFF5F4"),
    )
    for foreground, background in pairs:
        assert foreground in stylesheet
        assert background in stylesheet
        assert _contrast_ratio(foreground, background) >= 4.5
