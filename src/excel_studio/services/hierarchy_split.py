from __future__ import annotations

import math
import unicodedata
from dataclasses import dataclass, field
from datetime import date, datetime
from numbers import Integral, Real
from typing import Any, TypeAlias

import pandas as pd

from excel_studio.models.operations import BlankValuePolicy

MAX_HIERARCHY_LEVELS = 3
VALUE_LIST_LIMIT = 500
_EMPTY_KEY = "\0empty"

HierarchyFilterValue: TypeAlias = str | list[str]
HierarchyFilters: TypeAlias = dict[str, HierarchyFilterValue]


@dataclass(slots=True)
class HierarchyCatalog:
    candidate_values: dict[str, list[str]] = field(default_factory=dict)
    candidate_counts: dict[str, int] = field(default_factory=dict)
    matched_rows: int = 0
    target_count: int = 0
    complete_filters: bool = False
    truncated_fields: list[str] = field(default_factory=list)


def _is_blank(value: Any) -> bool:
    if value is None:
        return True
    try:
        if bool(pd.isna(value)):
            return True
    except (TypeError, ValueError):
        pass
    return isinstance(value, str) and not value.strip()


def display_hierarchy_value(value: Any, empty_label: str = "EMPTY") -> str:
    if _is_blank(value):
        return empty_label
    if isinstance(value, datetime):
        return value.isoformat(sep=" ", timespec="seconds")
    if isinstance(value, date):
        return value.isoformat()
    if isinstance(value, Integral) and not isinstance(value, bool):
        return str(int(value))
    if isinstance(value, Real) and not isinstance(value, bool):
        number = float(value)
        if math.isfinite(number) and number.is_integer():
            return str(int(number))
    return unicodedata.normalize("NFKC", str(value)).strip()


def canonical_hierarchy_value(value: Any, empty_label: str = "EMPTY") -> str:
    displayed = display_hierarchy_value(value, empty_label)
    if displayed == empty_label or _is_blank(value):
        return _EMPTY_KEY
    return unicodedata.normalize("NFKC", displayed).strip().casefold()


def hierarchy_filter_values(selection: HierarchyFilterValue | None) -> list[str]:
    if selection is None:
        return []
    source = [selection] if isinstance(selection, str) else selection
    values: list[str] = []
    seen: set[str] = set()
    for raw_value in source:
        value = str(raw_value).strip()
        key = unicodedata.normalize("NFKC", value).casefold()
        if not value or key in seen:
            continue
        seen.add(key)
        values.append(value)
    return values


def validate_hierarchy_rule(
    frame: pd.DataFrame,
    hierarchy_fields: list[str],
    hierarchy_filters: HierarchyFilters,
    split_field: str,
) -> list[str]:
    fields = [field.strip() for field in hierarchy_fields if field.strip()]
    if not 2 <= len(fields) <= MAX_HIERARCHY_LEVELS:
        raise ValueError("Hierarchy splitting requires two or three fields")
    if len(set(fields)) != len(fields):
        raise ValueError("Hierarchy fields must be different")
    missing = [field for field in fields if field not in frame.columns]
    if missing:
        raise KeyError(f"Missing hierarchy fields: {', '.join(missing)}")
    if split_field != fields[-1]:
        raise ValueError("The deepest hierarchy field must be the split target")
    missing_filters = [
        field for field in fields[:-1] if not hierarchy_filter_values(hierarchy_filters.get(field))
    ]
    if missing_filters:
        raise ValueError(f"Missing hierarchy filter values: {', '.join(missing_filters)}")
    return fields


def _matching_mask(
    series: pd.Series,
    selected_values: HierarchyFilterValue,
    empty_label: str,
) -> pd.Series:
    expected = {
        canonical_hierarchy_value(value, empty_label)
        for value in hierarchy_filter_values(selected_values)
    }
    return series.map(lambda value: canonical_hierarchy_value(value, empty_label) in expected)


def _ordered_values(
    series: pd.Series,
    empty_label: str,
) -> list[str]:
    values: list[str] = []
    seen: set[str] = set()
    for value in series.tolist():
        key = canonical_hierarchy_value(value, empty_label)
        if key in seen:
            continue
        seen.add(key)
        values.append(display_hierarchy_value(value, empty_label))
    return values


def _uses_multiple_parent_values(
    hierarchy_fields: list[str],
    hierarchy_filters: HierarchyFilters,
) -> bool:
    return any(
        len(hierarchy_filter_values(hierarchy_filters.get(field_name))) > 1
        for field_name in hierarchy_fields[:-1]
    )


def _hierarchy_keys(
    frame: pd.DataFrame,
    fields: list[str],
    empty_label: str,
) -> pd.Series:
    return frame[fields].apply(
        lambda row: tuple(
            canonical_hierarchy_value(row[field_name], empty_label) for field_name in fields
        ),
        axis=1,
    )


def build_hierarchy_catalog(
    frame: pd.DataFrame,
    hierarchy_fields: list[str],
    hierarchy_filters: HierarchyFilters,
    split_field: str,
    empty_label: str = "EMPTY",
    blank_policy: BlankValuePolicy = BlankValuePolicy.GROUP,
) -> HierarchyCatalog:
    fields = [field.strip() for field in hierarchy_fields if field.strip()]
    if not 2 <= len(fields) <= MAX_HIERARCHY_LEVELS or split_field != fields[-1]:
        return HierarchyCatalog()
    if len(set(fields)) != len(fields) or any(field not in frame.columns for field in fields):
        return HierarchyCatalog()

    catalog = HierarchyCatalog()
    working = frame
    complete = True
    for field_name in fields[:-1]:
        values = _ordered_values(working[field_name], empty_label)
        catalog.candidate_counts[field_name] = len(values)
        catalog.candidate_values[field_name] = values[:VALUE_LIST_LIMIT]
        if len(values) > VALUE_LIST_LIMIT:
            catalog.truncated_fields.append(field_name)
        selected = hierarchy_filter_values(hierarchy_filters.get(field_name))
        if not selected:
            complete = False
            continue
        working = working.loc[_matching_mask(working[field_name], selected, empty_label)]

    count_frame = working
    if blank_policy == BlankValuePolicy.SKIP:
        count_frame = working.loc[working[split_field].map(lambda value: not _is_blank(value))]
    catalog.matched_rows = len(working)
    if _uses_multiple_parent_values(fields, hierarchy_filters):
        target_fields = fields
    else:
        target_fields = [split_field]
    catalog.target_count = len(
        _hierarchy_keys(count_frame, target_fields, empty_label).drop_duplicates()
    )
    catalog.complete_filters = complete
    return catalog


def split_by_hierarchy(
    frame: pd.DataFrame,
    hierarchy_fields: list[str],
    hierarchy_filters: HierarchyFilters,
    split_field: str,
    empty_label: str = "EMPTY",
    blank_policy: BlankValuePolicy = BlankValuePolicy.GROUP,
) -> list[tuple[str, pd.DataFrame]]:
    fields = validate_hierarchy_rule(
        frame,
        hierarchy_fields,
        hierarchy_filters,
        split_field,
    )
    filtered = frame
    for field_name in fields[:-1]:
        selected = hierarchy_filters[field_name]
        filtered = filtered.loc[_matching_mask(filtered[field_name], selected, empty_label)]
    if filtered.empty:
        raise ValueError("No rows matched the selected hierarchy filters")

    if blank_policy == BlankValuePolicy.SKIP:
        filtered = filtered.loc[filtered[split_field].map(lambda value: not _is_blank(value))]
    if filtered.empty:
        raise ValueError("No rows remain after applying the hierarchy blank-value policy")

    path_grouping = _uses_multiple_parent_values(fields, hierarchy_filters)
    group_fields = fields if path_grouping else [split_field]
    keys = _hierarchy_keys(filtered, group_fields, empty_label)
    groups: list[tuple[str, pd.DataFrame]] = []
    for key in keys.drop_duplicates().tolist():
        matching = keys.map(key.__eq__)
        group = filtered.loc[matching].copy()
        first_row = group.iloc[0]
        if path_grouping:
            label = "__".join(
                display_hierarchy_value(first_row[field_name], empty_label) for field_name in fields
            )
        else:
            label = display_hierarchy_value(first_row[split_field], empty_label)
        groups.append((label, group))
    return groups


def hierarchy_filter_path(
    hierarchy_fields: list[str],
    hierarchy_filters: HierarchyFilters,
) -> str:
    levels: list[str] = []
    for field_name in hierarchy_fields[:-1]:
        values = hierarchy_filter_values(hierarchy_filters.get(field_name))
        if values:
            levels.append("+".join(values))
    return "__".join(levels)
