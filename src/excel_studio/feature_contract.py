from __future__ import annotations

from typing import Any, cast

from excel_studio.models.operations import MergeMode, SplitMode
from excel_studio.ui.hierarchy_split_page import HierarchySplitPage
from excel_studio.ui.main_window import MainWindow


def build_feature_contract(window: MainWindow) -> dict[str, Any]:
    split_page = cast(HierarchySplitPage, window.split_page)
    merge_page = window.merge_page
    hierarchy_selectors = split_page.hierarchy_value_combos
    return {
        "shell": "commercial_step_workspace",
        "default_workspace": window._active_operation,
        "split_default_mode": split_page.current_mode().value,
        "split_field_selector_present": split_page.field_combo is not None,
        "split_field_panel_present": split_page.field_panel is not None,
        "hierarchy_multiselect": bool(hierarchy_selectors)
        and all(
            hasattr(selector, "selected_values") and hasattr(selector, "available_values")
            for selector in hierarchy_selectors
        ),
        "hierarchy_max_levels": len(split_page.hierarchy_field_combos),
        "split_modes": [mode.value for mode in SplitMode],
        "merge_modes": [
            str(merge_page.merge_mode_combo.itemData(index))
            for index in range(merge_page.merge_mode_combo.count())
        ],
        "expected_merge_modes": [mode.value for mode in MergeMode],
        "language_switch": [
            str(window.language_combo.itemData(index))
            for index in range(window.language_combo.count())
        ],
        "fixed_footer": window.footer.parent() is window.centralWidget(),
        "responsive_window": not window.isMaximized() and not window.isFullScreen(),
    }
