from __future__ import annotations

from pathlib import Path

from openpyxl import Workbook
from pytestqt.qtbot import QtBot

from excel_studio.models.advanced import AdvancedMergeTaskConfig
from excel_studio.models.operations import MergeMode, MergeTaskConfig
from excel_studio.models.task import ScanOptions
from excel_studio.services.file_scanner import scan_files
from excel_studio.ui.main_window import MainWindow


def _write_source(path: Path, rows: list[list[object]]) -> None:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "明细"
    sheet.append(["编号", "区域", "金额"])
    for row in rows:
        sheet.append(row)
    workbook.save(path)
    workbook.close()


def _select_text(widget: object, text: str) -> None:
    for index in range(widget.count()):
        item = widget.item(index)
        if item.text() == text:
            item.setSelected(True)
            widget.setCurrentItem(item)
            return
    raise AssertionError(f"Field not available: {text}")


def test_every_original_merge_mode_builds_from_the_commercial_page(
    product_window: MainWindow,
    qtbot: QtBot,
    tmp_path: Path,
) -> None:
    first = tmp_path / "一月.xlsx"
    second = tmp_path / "二月.xlsx"
    _write_source(first, [[1, "华东", 10], [2, "华南", 20]])
    _write_source(second, [[1, "华东", 30], [3, "华北", 40]])
    page = product_window.merge_page
    page._scan_completed(scan_files([first, second], ScanOptions(include_text_files=True)))
    qtbot.waitUntil(
        lambda: (
            page.dedupe_fields_list.count() >= 3
            and page.left_keys_list.count() >= 3
            and page.right_keys_list.count() >= 3
        ),
        timeout=15_000,
    )

    standard_modes = {
        MergeMode.VERTICAL,
        MergeMode.WORKBOOK,
        MergeMode.SAME_NAME,
    }
    for mode in standard_modes:
        page.merge_mode_combo.setCurrentIndex(page.merge_mode_combo.findData(mode.value))
        assert page.current_mode() == mode
        assert isinstance(page.build_config(tmp_path / mode.value), MergeTaskConfig)

    for mode in {MergeMode.KEY_JOIN, MergeMode.HORIZONTAL}:
        page.merge_mode_combo.setCurrentIndex(page.merge_mode_combo.findData(mode.value))
        if mode == MergeMode.KEY_JOIN:
            page.left_keys_list.clearSelection()
            page.right_keys_list.clearSelection()
            _select_text(page.left_keys_list, "编号")
            _select_text(page.right_keys_list, "编号")
        assert page.current_mode() == mode
        assert isinstance(page.build_config(tmp_path / mode.value), AdvancedMergeTaskConfig)

    page.merge_mode_combo.setCurrentIndex(page.merge_mode_combo.findData(MergeMode.AGGREGATE.value))
    _select_text(page.group_fields_list, "区域")
    _select_text(page.value_fields_list, "金额")
    aggregate_config = page.build_config(tmp_path / "aggregate")
    assert isinstance(aggregate_config, MergeTaskConfig)
    assert aggregate_config.merge.aggregate_group_fields == ["区域"]
    assert aggregate_config.merge.aggregate_value_fields == ["金额"]
