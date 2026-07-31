from __future__ import annotations

import math
from collections.abc import Iterable
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook

from excel_studio.models.operations import (
    SplitDistribution,
    SplitMode,
    SplitTaskConfig,
)
from excel_studio.models.result import TaskResult
from excel_studio.services.hierarchy_split import split_by_hierarchy
from excel_studio.services.output_writer import write_frame
from excel_studio.services.reconciliation import reconcile_rows
from excel_studio.services.table_reader import read_sheet_frame
from excel_studio.utils.naming import render_naming_template


def fixed_row_groups(
    frame: pd.DataFrame, rows_per_file: int, balanced: bool = False
) -> list[pd.DataFrame]:
    if rows_per_file <= 0:
        raise ValueError("rows_per_file must be greater than zero")
    if frame.empty:
        return [frame.copy()]
    if not balanced:
        return [
            frame.iloc[start : start + rows_per_file].copy()
            for start in range(0, len(frame), rows_per_file)
        ]
    parts = math.ceil(len(frame) / rows_per_file)
    return split_into_parts(frame, parts)


def split_into_parts(frame: pd.DataFrame, parts: int) -> list[pd.DataFrame]:
    if parts <= 0:
        raise ValueError("parts must be greater than zero")
    quotient, remainder = divmod(len(frame), parts)
    groups: list[pd.DataFrame] = []
    start = 0
    for index in range(parts):
        size = quotient + (1 if index < remainder else 0)
        if size == 0:
            continue
        groups.append(frame.iloc[start : start + size].copy())
        start += size
    return groups or [frame.copy()]


def split_by_fields(
    frame: pd.DataFrame, fields: list[str], empty_label: str = "EMPTY"
) -> list[tuple[str, pd.DataFrame]]:
    if not fields:
        raise ValueError("At least one split field is required")
    missing = [field for field in fields if field not in frame.columns]
    if missing:
        raise KeyError(f"Missing split fields: {', '.join(missing)}")
    grouping = frame[fields].copy()
    for field in fields:
        grouping[field] = grouping[field].where(grouping[field].notna(), empty_label)
        grouping[field] = grouping[field].map(
            lambda value: empty_label if str(value).strip() == "" else value
        )
    keys = grouping.astype(str).agg("__".join, axis=1)
    groups: list[tuple[str, pd.DataFrame]] = []
    for value in keys.drop_duplicates().tolist():
        groups.append((value, frame.loc[keys == value].copy()))
    return groups


def _sheet_names(path: Path, configured: dict[Path, list[str]]) -> list[str]:
    explicit = configured.get(path)
    if explicit:
        return explicit
    workbook = load_workbook(path, read_only=True, keep_links=False)
    try:
        return [sheet.title for sheet in workbook.worksheets if sheet.sheet_state == "visible"]
    finally:
        workbook.close()


def execute_split(config: SplitTaskConfig) -> TaskResult:
    result = TaskResult()
    input_rows = 0
    output_rows = 0
    excluded_rows = 0
    for source_index, path in enumerate(config.input_files, start=1):
        for sheet_name in _sheet_names(path, config.split.selected_sheets):
            try:
                frame = read_sheet_frame(path, sheet_name, config.structure)
                input_rows += len(frame)
                if config.output.add_source_file:
                    frame["Source_File"] = path.name
                if config.output.add_source_sheet:
                    frame["Source_Sheet"] = sheet_name
                if config.split.mode == SplitMode.BY_FIELD:
                    groups: Iterable[tuple[str, pd.DataFrame]] = split_by_fields(
                        frame, config.split.fields, config.split.empty_label
                    )
                    split_field = "-".join(config.split.fields)
                elif config.split.mode == SplitMode.HIERARCHY:
                    groups = split_by_hierarchy(
                        frame,
                        config.split.hierarchy_fields,
                        config.split.hierarchy_filters,
                        config.split.hierarchy_split_field,
                        config.split.empty_label,
                        config.split.blank_value_policy,
                    )
                    split_field = config.split.hierarchy_split_field
                elif config.split.mode == SplitMode.FIXED_ROWS:
                    chunks = fixed_row_groups(
                        frame,
                        config.split.rows_per_file,
                        config.split.distribution == SplitDistribution.BALANCED,
                    )
                    groups = [(str(index), chunk) for index, chunk in enumerate(chunks, start=1)]
                    split_field = "rows"
                elif config.split.mode == SplitMode.BY_PARTS:
                    chunks = split_into_parts(frame, config.split.parts)
                    groups = [(str(index), chunk) for index, chunk in enumerate(chunks, start=1)]
                    split_field = "parts"
                else:
                    groups = [(sheet_name, frame)]
                    split_field = "sheet"
                materialized = list(groups)
                included_rows = sum(len(group) for _value, group in materialized)
                excluded_rows += max(0, len(frame) - included_rows)
                for part_number, (split_value, group) in enumerate(materialized, start=1):
                    stem = render_naming_template(
                        config.output.naming_template,
                        {
                            "original_name": path.stem,
                            "sheet_name": sheet_name,
                            "split_field": split_field,
                            "split_value": split_value,
                            "part_no": part_number,
                            "total_parts": len(materialized),
                            "row_count": len(group),
                            "index": source_index,
                        },
                    )
                    stem = f"{config.output.prefix}{stem}{config.output.suffix}"
                    output = write_frame(
                        group,
                        config.output.directory,
                        stem,
                        sheet_name,
                        config.output.overwrite,
                    )
                    result.output_files.append(output)
                    output_rows += len(group)
            except Exception as error:
                result.errors.append(f"{path.name} / {sheet_name}: {error}")
    result.reconciliation = reconcile_rows(input_rows, output_rows, excluded_rows)
    if result.reconciliation.warning:
        result.warnings.append(result.reconciliation.warning)
    return result
