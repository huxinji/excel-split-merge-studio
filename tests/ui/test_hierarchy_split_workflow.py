from __future__ import annotations

from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from pytestqt.qtbot import QtBot

from excel_studio.config.i18n import TranslationManager
from excel_studio.models.operations import SplitMode
from excel_studio.services.pro_engine import execute_pro_split
from excel_studio.ui.hierarchy_value_selector import HierarchyValueDialog
from excel_studio.ui.main_window import MainWindow


def _write_source(path: Path) -> None:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "层级明细"
    sheet.append(["OPS", "大区", "片区", "网点", "金额"])
    sheet.append(["中国", "华东", "共享片区", "浦东", 10])
    sheet.append(["中国", "华东", "上海", "浦西", 20])
    sheet.append(["中国", "华南", "共享片区", "深圳", 30])
    sheet.append(["中国", "华南", "广东", "广州", 40])
    workbook.save(path)
    workbook.close()


def _set_field(page: object, index: int, value: str) -> None:
    combo = page.hierarchy_field_combos[index]
    assert combo.findText(value) >= 0
    combo.setCurrentText(value)


def test_ui_builds_multi_value_hierarchy_configuration(
    product_window: MainWindow,
    qtbot: QtBot,
    tmp_path: Path,
) -> None:
    source = tmp_path / "管理层级.xlsx"
    output = tmp_path / "层级输出"
    _write_source(source)
    page = product_window.split_page
    page.load_path(source)
    qtbot.waitUntil(
        lambda: (
            page.record is not None
            and page.current_preview is not None
            and page.field_combo.findText("大区") >= 0
        ),
        timeout=15_000,
    )

    page.mode_selector.set_current(SplitMode.HIERARCHY.value)
    page.hierarchy_count_combo.setCurrentIndex(page.hierarchy_count_combo.findData(2))
    _set_field(page, 0, "大区")
    _set_field(page, 1, "片区")
    qtbot.waitUntil(lambda: not page._threads, timeout=15_000)
    assert {"华东", "华南"} <= set(page.hierarchy_value_combos[0].available_values())
    page.hierarchy_value_combos[0].set_selected_values(["华东", "华南"])
    qtbot.waitUntil(lambda: not page._threads, timeout=15_000)

    config = page.build_config(output)
    assert config.split.mode == SplitMode.HIERARCHY
    assert config.split.hierarchy_fields == ["大区", "片区"]
    assert config.split.hierarchy_filters == {"大区": ["华东", "华南"]}
    assert config.split.hierarchy_split_field == "片区"

    result = execute_pro_split(config)
    assert not result.errors
    assert len(result.output_files) == 4
    assert result.reconciliation is not None
    assert result.reconciliation.is_balanced
    assert result.reconciliation.excluded_rows == 0
    observed_paths = {
        (frame.iloc[0]["大区"], frame.iloc[0]["片区"])
        for frame in (pd.read_excel(path) for path in result.output_files)
    }
    assert observed_paths == {
        ("华东", "共享片区"),
        ("华东", "上海"),
        ("华南", "共享片区"),
        ("华南", "广东"),
    }


def test_multi_value_dialog_search_and_select_visible(qtbot: QtBot) -> None:
    dialog = HierarchyValueDialog(
        TranslationManager("zh_CN"),
        ["华东", "华南", "华北"],
        ["华东"],
    )
    qtbot.addWidget(dialog)
    dialog.search_edit.setText("华南")
    dialog._select_visible()

    assert dialog.selected_values() == ["华东", "华南"]


def test_three_level_layout_and_english_translation(
    product_window: MainWindow,
) -> None:
    page = product_window.split_page
    page.mode_selector.set_current(SplitMode.HIERARCHY.value)
    page.hierarchy_count_combo.setCurrentIndex(page.hierarchy_count_combo.findData(3))
    assert page.hierarchy_rows[2].isVisible()

    product_window.i18n.set_language("en_US")
    assert page.mode_selector.button(SplitMode.HIERARCHY.value).text() == "Hierarchy"
    assert (
        page.hierarchy_primary_title.text()
        == "Filter and split by management hierarchy (up to 3 levels)"
    )
    assert page.hierarchy_value_combos[0].choose_button.text() == "Select…"
