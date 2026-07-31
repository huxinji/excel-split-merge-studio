from __future__ import annotations

from dataclasses import dataclass, field
from enum import StrEnum
from pathlib import Path

from excel_studio.models.task import TableStructure


class SplitMode(StrEnum):
    BY_FIELD = "by_field"
    HIERARCHY = "hierarchy"
    FIXED_ROWS = "fixed_rows"
    BY_PARTS = "by_parts"
    BY_SHEET = "by_sheet"


class SplitDistribution(StrEnum):
    FIXED_LIMIT = "fixed_limit"
    BALANCED = "balanced"


class MergeMode(StrEnum):
    VERTICAL = "vertical"
    SAME_NAME = "same_name"
    WORKBOOK = "workbook"
    HORIZONTAL = "horizontal"
    KEY_JOIN = "key_join"
    AGGREGATE = "aggregate"


class FieldStrategy(StrEnum):
    STRICT = "strict"
    ALIGN = "align"
    UNION = "union"
    INTERSECTION = "intersection"
    MASTER = "master"
    POSITION = "position"


class OutputFormat(StrEnum):
    XLSX = "xlsx"
    CSV = "csv"


class OutputMode(StrEnum):
    SEPARATE_FILES = "separate_files"
    SINGLE_WORKBOOK = "single_workbook"


class ExistingFilePolicy(StrEnum):
    RENAME = "rename"
    OVERWRITE = "overwrite"
    SKIP = "skip"


class BlankValuePolicy(StrEnum):
    GROUP = "group"
    SKIP = "skip"


class DedupeMode(StrEnum):
    NONE = "none"
    ALL = "all"
    FIELDS = "fields"


class SheetConflictPolicy(StrEnum):
    RENAME = "rename"
    SKIP = "skip"
    REPLACE = "replace"


class AggregateMethod(StrEnum):
    SUM = "sum"
    COUNT = "count"
    MEAN = "mean"
    MIN = "min"
    MAX = "max"
    NUNIQUE = "nunique"


@dataclass(slots=True)
class OutputOptions:
    directory: Path
    naming_template: str = "{original_name}_{sheet_name}_{part_no}"
    prefix: str = ""
    suffix: str = ""
    overwrite: bool = False
    add_source_file: bool = True
    add_source_sheet: bool = True
    output_format: OutputFormat = OutputFormat.XLSX
    output_mode: OutputMode = OutputMode.SEPARATE_FILES
    existing_file_policy: ExistingFilePolicy = ExistingFilePolicy.RENAME
    workbook_name: str = "Excel_Split_Merge_Result"
    preserve_format: bool = False
    style_output: bool = True
    create_report_sheet: bool = True
    open_when_done: bool = False


@dataclass(slots=True)
class SplitOptions:
    mode: SplitMode = SplitMode.BY_FIELD
    fields: list[str] = field(default_factory=list)
    hierarchy_fields: list[str] = field(default_factory=list)
    hierarchy_filters: dict[str, str | list[str]] = field(default_factory=dict)
    hierarchy_split_field: str = ""
    rows_per_file: int = 500
    parts: int = 5
    distribution: SplitDistribution = SplitDistribution.FIXED_LIMIT
    empty_label: str = "EMPTY"
    blank_value_policy: BlankValuePolicy = BlankValuePolicy.GROUP
    selected_sheets: dict[Path, list[str]] = field(default_factory=dict)


@dataclass(slots=True)
class SplitTaskConfig:
    input_files: list[Path]
    structure: TableStructure
    split: SplitOptions
    output: OutputOptions


@dataclass(slots=True)
class MergeOptions:
    mode: MergeMode = MergeMode.VERTICAL
    field_strategy: FieldStrategy = FieldStrategy.ALIGN
    target_sheet_name: str = ""
    selected_sheets: dict[Path, list[str]] = field(default_factory=dict)
    add_source_file: bool = True
    add_source_sheet: bool = True
    dedupe_mode: DedupeMode = DedupeMode.NONE
    dedupe_fields: list[str] = field(default_factory=list)
    dedupe_keep: str = "first"
    sheet_conflict_policy: SheetConflictPolicy = SheetConflictPolicy.RENAME
    preserve_format: bool = False
    side_prefix: bool = True
    side_gap: int = 1
    aggregate_group_fields: list[str] = field(default_factory=list)
    aggregate_value_fields: list[str] = field(default_factory=list)
    aggregate_method: AggregateMethod = AggregateMethod.SUM


@dataclass(slots=True)
class MergeTaskConfig:
    input_files: list[Path]
    structure: TableStructure
    merge: MergeOptions
    output: OutputOptions
