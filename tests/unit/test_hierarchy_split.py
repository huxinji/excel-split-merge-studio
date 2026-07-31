from __future__ import annotations

from pathlib import Path

import pandas as pd
import pytest
from openpyxl import Workbook

from excel_studio.models.operations import (
    OutputOptions,
    SplitMode,
    SplitOptions,
    SplitTaskConfig,
)
from excel_studio.models.task import HeaderMode, TableStructure
from excel_studio.services.hierarchy_split import (
    build_hierarchy_catalog,
    hierarchy_filter_path,
    split_by_hierarchy,
)
from excel_studio.services.pro_engine import execute_pro_split


def _hierarchy_frame() -> pd.DataFrame:
    return pd.DataFrame(
        {
            "OPS": ["A", "A", "A", "A", "A"],
            "大区": ["华东", "华东", "华东", "华南", "华南"],
            "片区": ["上海", "上海", "浙江", "广东", "广西"],
            "网点": ["浦东", "浦西", "杭州", "深圳", "南宁"],
            "金额": [10, 20, 30, 40, 50],
        }
    )


def _duplicate_child_frame() -> pd.DataFrame:
    return pd.DataFrame(
        {
            "大区": ["华东", "华东", "华南", "华南"],
            "片区": ["共享片区", "上海", "共享片区", "广东"],
            "网点": ["东一", "上海一", "南一", "广东一"],
            "金额": [10, 20, 30, 40],
        }
    )


def _write_hierarchy_workbook(path: Path) -> None:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "明细"
    sheet.append(["OPS", "大区", "片区", "网点", "金额"])
    for row in _hierarchy_frame().itertuples(index=False, name=None):
        sheet.append(list(row))
    workbook.save(path)
    workbook.close()


def test_two_level_filter_then_split() -> None:
    groups = split_by_hierarchy(
        _hierarchy_frame(),
        ["大区", "片区"],
        {"大区": "华东"},
        "片区",
    )

    assert [name for name, _group in groups] == ["上海", "浙江"]
    assert [len(group) for _name, group in groups] == [2, 1]
    assert all(set(group["大区"]) == {"华东"} for _name, group in groups)


def test_two_level_multi_select_keeps_duplicate_child_paths_separate() -> None:
    groups = split_by_hierarchy(
        _duplicate_child_frame(),
        ["大区", "片区"],
        {"大区": ["华东", "华南"]},
        "片区",
    )

    assert [name for name, _group in groups] == [
        "华东__共享片区",
        "华东__上海",
        "华南__共享片区",
        "华南__广东",
    ]
    shared_groups = [group for name, group in groups if name.endswith("共享片区")]
    assert len(shared_groups) == 2
    assert {group.iloc[0]["大区"] for group in shared_groups} == {"华东", "华南"}


def test_three_level_filter_then_split() -> None:
    groups = split_by_hierarchy(
        _hierarchy_frame(),
        ["大区", "片区", "网点"],
        {"大区": "华东", "片区": "上海"},
        "网点",
    )

    assert [name for name, _group in groups] == ["浦东", "浦西"]
    assert sum(len(group) for _name, group in groups) == 2


def test_three_level_multi_select_uses_or_within_and_across_levels() -> None:
    groups = split_by_hierarchy(
        _hierarchy_frame(),
        ["大区", "片区", "网点"],
        {"大区": ["华东", "华南"], "片区": ["上海", "广东"]},
        "网点",
    )

    assert [name for name, _group in groups] == [
        "华东__上海__浦东",
        "华东__上海__浦西",
        "华南__广东__深圳",
    ]
    assert sum(len(group) for _name, group in groups) == 3


def test_hierarchy_catalog_uses_filtered_full_data() -> None:
    catalog = build_hierarchy_catalog(
        _hierarchy_frame(),
        ["大区", "片区", "网点"],
        {"大区": "华东", "片区": "上海"},
        "网点",
    )

    assert catalog.candidate_values["大区"] == ["华东", "华南"]
    assert catalog.candidate_values["片区"] == ["上海", "浙江"]
    assert catalog.complete_filters
    assert catalog.matched_rows == 2
    assert catalog.target_count == 2


def test_hierarchy_catalog_cascades_from_multiple_parent_values() -> None:
    catalog = build_hierarchy_catalog(
        _hierarchy_frame(),
        ["大区", "片区", "网点"],
        {"大区": ["华东", "华南"], "片区": ["上海", "广东"]},
        "网点",
    )

    assert catalog.candidate_values["片区"] == ["上海", "浙江", "广东", "广西"]
    assert catalog.complete_filters
    assert catalog.matched_rows == 3
    assert catalog.target_count == 3
    assert (
        hierarchy_filter_path(
            ["大区", "片区", "网点"],
            {"大区": ["华东", "华南"], "片区": ["上海", "广东"]},
        )
        == "华东+华南__上海+广东"
    )


def test_hierarchy_rejects_more_than_three_levels() -> None:
    frame = _hierarchy_frame().assign(门店=["1", "2", "3", "4", "5"])
    with pytest.raises(ValueError, match="two or three"):
        split_by_hierarchy(
            frame,
            ["OPS", "大区", "片区", "网点"],
            {"OPS": "A", "大区": "华东", "片区": "上海"},
            "网点",
        )


def test_pro_hierarchy_split_writes_only_selected_scope(tmp_path: Path) -> None:
    source = tmp_path / "层级数据.xlsx"
    output = tmp_path / "输出"
    _write_hierarchy_workbook(source)
    config = SplitTaskConfig(
        input_files=[source],
        structure=TableStructure(
            header_mode=HeaderMode.ROW_NUMBER,
            header_row=1,
            data_start_row=2,
        ),
        split=SplitOptions(
            mode=SplitMode.HIERARCHY,
            hierarchy_fields=["大区", "片区"],
            hierarchy_filters={"大区": "华东"},
            hierarchy_split_field="片区",
            selected_sheets={source: ["明细"]},
        ),
        output=OutputOptions(
            directory=output,
            naming_template="{hierarchy_path}_{split_value}",
            add_source_file=False,
            add_source_sheet=False,
        ),
    )

    result = execute_pro_split(config)

    assert not result.errors
    assert len(result.output_files) == 2
    assert result.reconciliation is not None
    assert result.reconciliation.input_rows == 5
    assert result.reconciliation.output_rows == 3
    assert result.reconciliation.excluded_rows == 2
    assert result.reconciliation.is_balanced
    observed = {
        (path.stem, tuple(pd.read_excel(path)["片区"].unique())) for path in result.output_files
    }
    assert observed == {
        ("华东_上海", ("上海",)),
        ("华东_浙江", ("浙江",)),
    }
