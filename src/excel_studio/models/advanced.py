from __future__ import annotations

from dataclasses import dataclass, field
from enum import StrEnum
from pathlib import Path

from excel_studio.models.operations import MergeMode, OutputOptions
from excel_studio.models.task import TableStructure


class JoinType(StrEnum):
    LEFT = "left"
    RIGHT = "right"
    INNER = "inner"
    OUTER = "outer"


class DuplicateKeyPolicy(StrEnum):
    FIRST = "first"
    LAST = "last"
    EXPAND = "expand"
    REJECT = "reject"
    AGGREGATE = "aggregate"


class JoinConflictPolicy(StrEnum):
    KEEP_BOTH = "keep_both"
    PREFER_MAIN = "prefer_main"
    PREFER_RELATED = "prefer_related"


@dataclass(slots=True)
class CleaningOptions:
    drop_empty_rows: bool = True
    drop_empty_columns: bool = True
    trim_whitespace: bool = True
    collapse_whitespace: bool = True
    normalize_width: bool = True
    drop_duplicate_rows: bool = False


@dataclass(slots=True)
class AdvancedMergeTaskConfig:
    input_files: list[Path]
    structure: TableStructure
    mode: MergeMode
    output: OutputOptions
    selected_sheets: dict[Path, list[str]] = field(default_factory=dict)
    join_type: JoinType = JoinType.LEFT
    left_keys: list[str] = field(default_factory=list)
    right_keys: list[str] = field(default_factory=list)
    duplicate_policy: DuplicateKeyPolicy = DuplicateKeyPolicy.EXPAND
    empty_keys_match: bool = False
    normalize_keys: bool = True
    fuzzy_match: bool = False
    fuzzy_threshold: int = 85
    conflict_policy: JoinConflictPolicy = JoinConflictPolicy.KEEP_BOTH
    main_suffix: str = "_main"
    related_suffix: str = "_related"
    side_prefix: bool = True
    side_gap: int = 1
    field_mapping: dict[str, str] = field(default_factory=dict)
    cleaning: CleaningOptions = field(default_factory=CleaningOptions)
