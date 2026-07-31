from __future__ import annotations

from typing import Any

import pandas as pd

from excel_studio.models.operations import OutputMode, SplitTaskConfig
from excel_studio.models.result import TaskResult
from excel_studio.services import pro_engine_impl as impl
from excel_studio.services.hierarchy_split import (
    display_hierarchy_value,
    hierarchy_filter_path,
    split_by_hierarchy,
)
from excel_studio.services.reconciliation import reconcile_rows
from excel_studio.services.table_reader import read_sheet_frame
from excel_studio.utils.naming import render_naming_template, sanitize_filename


def execute_hierarchy_split(
    config: SplitTaskConfig,
    hooks: impl.EngineHooks | None = None,
) -> TaskResult:
    hooks = hooks or impl.EngineHooks()
    result = TaskResult()
    input_rows = 0
    output_rows = 0
    excluded_rows = 0
    total_sources = max(1, len(config.input_files))

    for source_index, path in enumerate(config.input_files, start=1):
        hooks.checkpoint()
        hooks.update(
            15 + int(55 * (source_index - 1) / total_sources),
            "Reading hierarchy",
            path.name,
        )
        try:
            sheet_names = impl._sheet_names(path, config.split.selected_sheets)
            workbook_frames: list[tuple[str, pd.DataFrame]] = []
            report_rows: list[dict[str, Any]] = []
            scope_path = hierarchy_filter_path(
                config.split.hierarchy_fields,
                config.split.hierarchy_filters,
            )

            for sheet_name in sheet_names:
                hooks.checkpoint()
                frame = read_sheet_frame(path, sheet_name, config.structure)
                input_rows += len(frame)
                if config.output.add_source_file:
                    frame["Source_File"] = path.name
                if config.output.add_source_sheet:
                    frame["Source_Sheet"] = sheet_name

                groups = split_by_hierarchy(
                    frame,
                    config.split.hierarchy_fields,
                    config.split.hierarchy_filters,
                    config.split.hierarchy_split_field,
                    config.split.empty_label,
                    config.split.blank_value_policy,
                )
                materialized = list(groups)
                included_rows = sum(len(group) for _value, group in materialized)
                excluded_rows += max(0, len(frame) - included_rows)

                for part_number, (split_value, group) in enumerate(materialized, start=1):
                    hooks.checkpoint()
                    output_rows += len(group)
                    first_row = group.iloc[0]
                    hierarchy_values = {
                        f"level{index}_value": display_hierarchy_value(
                            first_row[field_name],
                            config.split.empty_label,
                        )
                        for index, field_name in enumerate(config.split.hierarchy_fields, start=1)
                    }
                    stem = render_naming_template(
                        config.output.naming_template,
                        {
                            "original_name": path.stem,
                            "sheet_name": sheet_name,
                            "split_field": config.split.hierarchy_split_field,
                            "split_value": split_value,
                            "hierarchy_path": scope_path,
                            "part_no": part_number,
                            "total_parts": len(materialized),
                            "row_count": len(group),
                            "index": source_index,
                            **hierarchy_values,
                        },
                    )
                    stem = f"{config.output.prefix}{stem}{config.output.suffix}"
                    report_rows.append(
                        {
                            "Source File": path.name,
                            "Source Sheet": sheet_name,
                            "Hierarchy Scope": scope_path,
                            "Split Field": config.split.hierarchy_split_field,
                            "Split Value": split_value,
                            "Rows": len(group),
                            "Output": stem,
                        }
                    )
                    if config.output.output_mode == OutputMode.SINGLE_WORKBOOK:
                        workbook_frames.append((stem, group))
                    else:
                        output = impl._write_frame(group, config.output, stem, stem)
                        if output is not None:
                            result.output_files.append(output)
                            hooks.write_log(f"Created hierarchy output: {output}")
                        else:
                            result.warnings.append(f"Skipped existing output: {stem}")

            if workbook_frames:
                workbook_stem = sanitize_filename(
                    f"{config.output.prefix}{config.output.workbook_name}{config.output.suffix}"
                )
                output = impl._write_frames_workbook(
                    workbook_frames,
                    config.output,
                    workbook_stem,
                    report_rows,
                )
                if output is not None:
                    result.output_files.append(output)
        except Exception as error:
            result.errors.append(f"{path.name}: {error}")

    result.reconciliation = reconcile_rows(
        input_rows,
        output_rows,
        excluded_rows,
    )
    if result.reconciliation.warning:
        result.warnings.append(result.reconciliation.warning)
    return result
