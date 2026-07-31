from __future__ import annotations

import re
import unicodedata
from difflib import SequenceMatcher
from pathlib import Path
from typing import Any

import pandas as pd
from openpyxl import load_workbook

from excel_studio.models.advanced import (
    AdvancedMergeTaskConfig,
    CleaningOptions,
    DuplicateKeyPolicy,
    JoinType,
)
from excel_studio.models.operations import MergeMode
from excel_studio.models.result import ReconciliationResult, TaskResult
from excel_studio.services.output_writer import write_frame
from excel_studio.services.table_reader import read_sheet_frame
from excel_studio.utils.naming import render_naming_template


def normalize_field_name(value: str) -> str:
    normalized = unicodedata.normalize("NFKC", str(value))
    normalized = normalized.replace("（", "(").replace("）", ")")
    return re.sub(r"\s+", " ", normalized.strip()).casefold()


def suggest_field_mapping(
    source_fields: list[str], target_fields: list[str], threshold: float = 0.72
) -> list[tuple[str, str, float]]:
    suggestions: list[tuple[str, str, float]] = []
    for source in source_fields:
        best_target = ""
        best_score = 0.0
        normalized_source = normalize_field_name(source)
        for target in target_fields:
            score = SequenceMatcher(None, normalized_source, normalize_field_name(target)).ratio()
            if score > best_score:
                best_target, best_score = target, score
        if best_target and best_score >= threshold and source != best_target:
            suggestions.append((source, best_target, best_score))
    return suggestions


def clean_frame(frame: pd.DataFrame, options: CleaningOptions) -> pd.DataFrame:
    cleaned = frame.copy()
    if options.trim_whitespace or options.collapse_whitespace or options.normalize_width:

        def clean_value(value: Any) -> Any:
            if not isinstance(value, str):
                return value
            result = unicodedata.normalize("NFKC", value) if options.normalize_width else value
            if options.trim_whitespace:
                result = result.strip()
            if options.collapse_whitespace:
                result = re.sub(r"\s+", " ", result)
            return result

        cleaned = cleaned.map(clean_value)
    if options.drop_empty_rows:
        cleaned = cleaned.replace(r"^\s*$", pd.NA, regex=True).dropna(axis=0, how="all")
    if options.drop_empty_columns:
        cleaned = cleaned.dropna(axis=1, how="all")
    if options.drop_duplicate_rows:
        cleaned = cleaned.drop_duplicates(keep="first")
    return cleaned.reset_index(drop=True)


def horizontal_concat(frames: list[pd.DataFrame]) -> pd.DataFrame:
    if not frames:
        return pd.DataFrame()
    renamed: list[pd.DataFrame] = []
    used: set[str] = set()
    for source_index, frame in enumerate(frames, start=1):
        mapping: dict[str, str] = {}
        for column in frame.columns:
            candidate = str(column)
            if candidate in used:
                candidate = f"{candidate}_Source{source_index}"
                suffix = 2
                while candidate in used:
                    candidate = f"{column}_Source{source_index}_{suffix}"
                    suffix += 1
            mapping[str(column)] = candidate
            used.add(candidate)
        renamed.append(frame.rename(columns=mapping).reset_index(drop=True))
    return pd.concat(renamed, axis=1)


def _normalize_join_keys(
    frame: pd.DataFrame, keys: list[str], empty_match: bool, prefix: str
) -> tuple[pd.DataFrame, list[str]]:
    missing = [key for key in keys if key not in frame.columns]
    if missing:
        raise KeyError(f"Missing join keys: {', '.join(missing)}")
    prepared = frame.copy()
    normalized_keys: list[str] = []
    for position, key in enumerate(keys):
        normalized_key = f"__join_key_{prefix}_{position}"
        values = prepared[key].map(
            lambda value: (
                ""
                if pd.isna(value)
                else unicodedata.normalize("NFKC", str(value)).strip().casefold()
            )
        )
        if not empty_match:
            values = pd.Series(
                [
                    value if value else f"__empty_{prefix}_{index}"
                    for index, value in enumerate(values)
                ],
                index=prepared.index,
            )
        prepared[normalized_key] = values
        normalized_keys.append(normalized_key)
    return prepared, normalized_keys


def _handle_duplicates(
    frame: pd.DataFrame, keys: list[str], policy: DuplicateKeyPolicy, side: str
) -> tuple[pd.DataFrame, int]:
    duplicated = frame.duplicated(subset=keys, keep=False)
    duplicate_count = int(duplicated.sum())
    if not duplicate_count or policy == DuplicateKeyPolicy.EXPAND:
        return frame, duplicate_count
    if policy == DuplicateKeyPolicy.REJECT:
        raise ValueError(f"Duplicate keys found on {side}: {duplicate_count} rows")
    if policy in {DuplicateKeyPolicy.FIRST, DuplicateKeyPolicy.LAST}:
        return frame.drop_duplicates(subset=keys, keep=policy.value), duplicate_count
    numeric_columns = [
        column for column in frame.select_dtypes(include="number").columns if column not in keys
    ]
    aggregations: dict[str, str] = {
        column: ("sum" if column in numeric_columns else "first")
        for column in frame.columns
        if column not in keys
    }
    return frame.groupby(keys, dropna=False, as_index=False).agg(aggregations), duplicate_count


def key_join(
    left: pd.DataFrame,
    right: pd.DataFrame,
    left_keys: list[str],
    right_keys: list[str],
    join_type: JoinType = JoinType.LEFT,
    duplicate_policy: DuplicateKeyPolicy = DuplicateKeyPolicy.EXPAND,
    empty_keys_match: bool = False,
) -> tuple[pd.DataFrame, dict[str, int]]:
    if not left_keys or len(left_keys) != len(right_keys):
        raise ValueError("Left and right join keys must have the same non-zero length")
    prepared_left, normalized_left = _normalize_join_keys(left, left_keys, empty_keys_match, "left")
    prepared_right, normalized_right = _normalize_join_keys(
        right, right_keys, empty_keys_match, "right"
    )
    prepared_left, left_duplicates = _handle_duplicates(
        prepared_left, normalized_left, duplicate_policy, "left"
    )
    prepared_right, right_duplicates = _handle_duplicates(
        prepared_right, normalized_right, duplicate_policy, "right"
    )
    for left_key, right_key in zip(normalized_left, normalized_right, strict=True):
        prepared_right[left_key] = prepared_right[right_key]
    prepared_right = prepared_right.drop(columns=normalized_right)
    indicator_name = "__merge_status"
    merged = prepared_left.merge(
        prepared_right,
        on=normalized_left,
        how=join_type.value,
        suffixes=("", "_Related"),
        indicator=indicator_name,
    )
    status = merged[indicator_name].value_counts().to_dict()
    merged = merged.drop(columns=[*normalized_left, indicator_name])
    stats = {
        "left_duplicates": left_duplicates,
        "right_duplicates": right_duplicates,
        "matched": int(status.get("both", 0)),
        "left_only": int(status.get("left_only", 0)),
        "right_only": int(status.get("right_only", 0)),
    }
    return merged, stats


def _first_visible_sheet(path: Path, configured: dict[Path, list[str]]) -> str:
    selected = configured.get(path)
    if selected:
        return selected[0]
    workbook = load_workbook(path, read_only=True, keep_links=False)
    try:
        visible = [sheet.title for sheet in workbook.worksheets if sheet.sheet_state == "visible"]
        if not visible:
            raise ValueError("No visible worksheet found")
        return visible[0]
    finally:
        workbook.close()


def execute_advanced_merge(config: AdvancedMergeTaskConfig) -> TaskResult:
    result = TaskResult()
    frames: list[pd.DataFrame] = []
    labels: list[str] = []
    for path in config.input_files:
        try:
            sheet_name = _first_visible_sheet(path, config.selected_sheets)
            frame = read_sheet_frame(path, sheet_name, config.structure)
            frame = frame.rename(columns=config.field_mapping)
            frames.append(clean_frame(frame, config.cleaning))
            labels.append(f"{path.name}/{sheet_name}")
        except Exception as error:
            result.errors.append(f"{path.name}: {error}")
    if len(frames) < 2:
        result.errors.append("Advanced merge requires at least two readable inputs")
        return result
    try:
        if config.mode == MergeMode.HORIZONTAL:
            merged = horizontal_concat(frames)
            input_reference = max(len(frame) for frame in frames)
            result.warnings.append(
                "Horizontal reconciliation source rows: "
                + ", ".join(
                    f"{label}={len(frame)}" for label, frame in zip(labels, frames, strict=True)
                )
            )
        elif config.mode == MergeMode.KEY_JOIN:
            merged = frames[0]
            input_reference = len(merged)
            for index, related in enumerate(frames[1:], start=1):
                merged, stats = key_join(
                    merged,
                    related,
                    config.left_keys,
                    config.right_keys,
                    config.join_type,
                    config.duplicate_policy,
                    config.empty_keys_match,
                )
                result.warnings.append(
                    f"Join {index}: matched={stats['matched']}, left_only={stats['left_only']}, "
                    f"right_only={stats['right_only']}, duplicate_rows="
                    f"{stats['left_duplicates'] + stats['right_duplicates']}"
                )
        else:
            raise ValueError(f"Unsupported advanced merge mode: {config.mode}")
        stem = render_naming_template(
            config.output.naming_template,
            {
                "original_name": "combined",
                "sheet_name": config.mode.value,
                "row_count": len(merged),
            },
        )
        result.output_files.append(
            write_frame(
                merged,
                config.output.directory,
                stem,
                "Horizontal" if config.mode == MergeMode.HORIZONTAL else "Joined",
                config.output.overwrite,
            )
        )
        result.reconciliation = ReconciliationResult(input_reference, len(merged))
        if config.mode == MergeMode.KEY_JOIN and len(merged) != input_reference:
            result.warnings.append(
                f"Join changed master row count from {input_reference} to {len(merged)}; "
                "review key diagnostics"
            )
    except Exception as error:
        result.errors.append(str(error))
    return result
