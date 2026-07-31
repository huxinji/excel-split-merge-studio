from __future__ import annotations

from collections import OrderedDict
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook

from excel_studio.models.operations import FieldStrategy, MergeMode, MergeTaskConfig
from excel_studio.models.result import TaskResult
from excel_studio.services.output_writer import write_frame
from excel_studio.services.reconciliation import reconcile_rows
from excel_studio.services.table_reader import read_sheet_frame
from excel_studio.utils.naming import (
    render_naming_template,
    sanitize_sheet_name,
    unique_output_path,
)


def _ordered_union(column_sets: list[list[str]]) -> list[str]:
    return list(OrderedDict.fromkeys(column for columns in column_sets for column in columns))


def vertical_merge(frames: list[pd.DataFrame], strategy: FieldStrategy) -> pd.DataFrame:
    if not frames:
        return pd.DataFrame()
    columns = [list(frame.columns) for frame in frames]
    if strategy == FieldStrategy.STRICT:
        if any(current != columns[0] for current in columns[1:]):
            raise ValueError("Strict merge requires identical field names and order")
        target = columns[0]
    elif strategy == FieldStrategy.INTERSECTION:
        target = [column for column in columns[0] if all(column in other for other in columns[1:])]
    elif strategy == FieldStrategy.MASTER:
        target = columns[0]
    else:
        target = _ordered_union(columns)
    aligned = [frame.reindex(columns=target) for frame in frames]
    return pd.concat(aligned, ignore_index=True) if aligned else pd.DataFrame(columns=target)


def _visible_sheet_names(path: Path, selected: dict[Path, list[str]]) -> list[str]:
    explicit = selected.get(path)
    if explicit:
        return explicit
    if path.suffix.casefold() in {".xlsx", ".xlsm"}:
        workbook = load_workbook(
            path,
            read_only=True,
            keep_vba=path.suffix.casefold() == ".xlsm",
            keep_links=False,
        )
        try:
            return [sheet.title for sheet in workbook.worksheets if sheet.sheet_state == "visible"]
        finally:
            workbook.close()
    return list(pd.ExcelFile(path).sheet_names)


def _unique_sheet_name(value: str, used: set[str]) -> str:
    base = sanitize_sheet_name(value)
    candidate = base
    index = 2
    while candidate.casefold() in used:
        suffix = f"_{index}"
        candidate = f"{base[: 31 - len(suffix)]}{suffix}"
        index += 1
    used.add(candidate.casefold())
    return candidate


def _write_workbook(
    frames: list[tuple[str, pd.DataFrame]], config: MergeTaskConfig, stem: str
) -> Path:
    output = (
        config.output.directory / f"{stem}.xlsx"
        if config.output.overwrite
        else unique_output_path(config.output.directory, stem)
    )
    output.parent.mkdir(parents=True, exist_ok=True)
    used: set[str] = set()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        for sheet_name, frame in frames:
            final_name = _unique_sheet_name(sheet_name, used)
            frame.to_excel(writer, index=False, sheet_name=final_name)
            writer.sheets[final_name].freeze_panes(1, 0)
    return output


def execute_merge(config: MergeTaskConfig) -> TaskResult:
    result = TaskResult()
    loaded: list[tuple[Path, str, pd.DataFrame]] = []
    input_rows = 0
    for path in config.input_files:
        try:
            sheet_names = _visible_sheet_names(path, config.merge.selected_sheets)
        except Exception as error:
            result.errors.append(f"{path.name}: {error}")
            continue
        for sheet_name in sheet_names:
            if (
                config.merge.mode == MergeMode.SAME_NAME
                and config.merge.target_sheet_name
                and sheet_name.casefold() != config.merge.target_sheet_name.casefold()
            ):
                continue
            try:
                frame = read_sheet_frame(path, sheet_name, config.structure)
                input_rows += len(frame)
                if config.merge.add_source_file:
                    frame["Source_File"] = path.name
                if config.merge.add_source_sheet:
                    frame["Source_Sheet"] = sheet_name
                loaded.append((path, sheet_name, frame))
            except Exception as error:
                result.errors.append(f"{path.name} / {sheet_name}: {error}")
    output_rows = 0
    try:
        if config.merge.mode == MergeMode.WORKBOOK:
            frames = [(f"{path.stem}_{sheet}", frame) for path, sheet, frame in loaded]
            output_rows = sum(len(frame) for _, frame in frames)
            stem = render_naming_template(
                config.output.naming_template,
                {"original_name": "combined", "sheet_name": "workbook", "row_count": output_rows},
            )
            result.output_files.append(_write_workbook(frames, config, stem))
        elif config.merge.mode == MergeMode.SAME_NAME:
            groups: OrderedDict[str, tuple[str, list[pd.DataFrame]]] = OrderedDict()
            for _path, sheet_name, frame in loaded:
                key = sheet_name.casefold()
                if key not in groups:
                    groups[key] = (sheet_name, [])
                groups[key][1].append(frame)
            for sheet_name, frames in groups.values():
                merged = vertical_merge(frames, config.merge.field_strategy)
                output_rows += len(merged)
                stem = render_naming_template(
                    config.output.naming_template,
                    {
                        "original_name": "combined",
                        "sheet_name": sheet_name,
                        "row_count": len(merged),
                    },
                )
                result.output_files.append(
                    write_frame(
                        merged,
                        config.output.directory,
                        stem,
                        f"{sheet_name}_Combined",
                        config.output.overwrite,
                    )
                )
        else:
            merged = vertical_merge(
                [frame for _path, _sheet, frame in loaded], config.merge.field_strategy
            )
            output_rows = len(merged)
            stem = render_naming_template(
                config.output.naming_template,
                {"original_name": "combined", "sheet_name": "Combined", "row_count": len(merged)},
            )
            result.output_files.append(
                write_frame(
                    merged,
                    config.output.directory,
                    stem,
                    "Combined",
                    config.output.overwrite,
                )
            )
    except Exception as error:
        result.errors.append(str(error))
    result.reconciliation = reconcile_rows(input_rows, output_rows)
    if result.reconciliation.warning:
        result.warnings.append(result.reconciliation.warning)
    return result
