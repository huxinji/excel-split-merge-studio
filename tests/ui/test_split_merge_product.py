from __future__ import annotations

from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from pytestqt.qtbot import QtBot

from excel_studio.models.operations import MergeMode, SplitMode
from excel_studio.services.pro_engine import execute_pro_split
from excel_studio.ui.main_window import MainWindow


def _create_field_split_source(path: Path) -> None:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "销售明细"
    sheet.append(["订单号", "区域", "金额"])
    sheet.append(["A001", "华东", 100])
    sheet.append(["A002", "华南", 80])
    sheet.append(["A003", "华东", 120])
    sheet.append(["A004", "", 50])
    workbook.save(path)
    workbook.close()


def test_shell_opens_on_prominent_field_split_workflow(
    product_window: MainWindow, qtbot: QtBot
) -> None:
    qtbot.waitUntil(product_window.isVisible)
    page = product_window.split_page

    assert product_window.page_stack.currentWidget() is page
    assert page.current_mode() == SplitMode.BY_FIELD
    assert page.mode_selector.button(SplitMode.BY_FIELD.value).isChecked()
    assert page.field_panel.isVisible()
    assert page.field_primary_title.text() == "按指定字段值拆分（默认）"
    assert page.field_combo.isVisible()
    assert page.mode_selector.button(SplitMode.FIXED_ROWS.value).isVisible()
    assert page.mode_selector.button(SplitMode.BY_PARTS.value).isVisible()
    assert page.mode_selector.button(SplitMode.BY_SHEET.value).isVisible()
    assert product_window.footer.isVisible()
    assert product_window.height() < product_window.screen().availableGeometry().height()


def test_selected_field_is_used_for_real_split(
    product_window: MainWindow,
    qtbot: QtBot,
    tmp_path: Path,
) -> None:
    source = tmp_path / "销售数据.xlsx"
    output = tmp_path / "拆分结果"
    _create_field_split_source(source)

    page = product_window.split_page
    page.load_path(source)
    qtbot.waitUntil(
        lambda: (
            page.record is not None
            and page.current_preview is not None
            and page.field_combo.findText("区域") >= 0
        ),
        timeout=15_000,
    )
    page.field_combo.setCurrentText("区域")
    config = page.build_config(output)

    assert config.split.mode == SplitMode.BY_FIELD
    assert config.split.fields == ["区域"]

    result = execute_pro_split(config)
    assert not result.errors
    assert len(result.output_files) == 3

    observed_groups: set[str] = set()
    observed_rows = 0
    for path in result.output_files:
        frame = pd.read_excel(path)
        values = frame["区域"].fillna("").astype(str).unique().tolist()
        assert len(values) == 1
        observed_groups.add(values[0])
        observed_rows += len(frame)
    assert observed_groups == {"华东", "华南", ""}
    assert observed_rows == 4


def test_workspace_and_language_switch_are_complete(
    product_window: MainWindow, qtbot: QtBot
) -> None:
    product_window.switch_workspace("merge")
    assert product_window.page_stack.currentWidget() is product_window.merge_page
    assert product_window.start_button.text() == "开始合并"

    mode_values = {
        product_window.merge_page.merge_mode_combo.itemData(index)
        for index in range(product_window.merge_page.merge_mode_combo.count())
    }
    assert mode_values == {mode.value for mode in MergeMode}

    product_window.i18n.set_language("en_US")
    qtbot.waitUntil(lambda: product_window.title_label.text() == "Excel Split & Merge Studio")
    assert product_window.split_workspace_button.text() == "Split Workbooks"
    assert product_window.merge_workspace_button.text() == "Merge Workbooks"
    assert (
        product_window.split_page.field_primary_title.text()
        == "Split by a selected field value (Default)"
    )

    product_window.switch_workspace("split")
    assert product_window.start_button.text() == "Start Splitting"
