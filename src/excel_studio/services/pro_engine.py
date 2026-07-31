from __future__ import annotations

import unicodedata
from typing import Any

import pandas as pd

from excel_studio.models.operations import SplitMode, SplitTaskConfig
from excel_studio.models.result import TaskResult
from excel_studio.services import pro_engine_impl as _impl
from excel_studio.services.hierarchy_execution import execute_hierarchy_split
from excel_studio.services.pro_engine_impl import (
    EngineHooks,
    execute_pro_advanced_merge,
    execute_pro_merge,
    vertical_merge,
)


def _prepare_join_frame(
    frame: pd.DataFrame,
    keys: list[str],
    normalize: bool,
    empty_match: bool,
    prefix: str,
) -> tuple[pd.DataFrame, list[str]]:
    missing = [key for key in keys if key not in frame.columns]
    if missing:
        raise KeyError(f"Missing join keys: {', '.join(missing)}")
    prepared = frame.copy()
    join_keys: list[str] = []
    for position, key in enumerate(keys):
        internal = f"__join_{prefix}_{position}"

        def normalize_value(value: Any) -> str:
            if value is None or pd.isna(value):
                return ""
            text = str(value)
            return unicodedata.normalize("NFKC", text).strip().casefold() if normalize else text

        values = prepared[key].map(normalize_value)
        if not empty_match:
            values = pd.Series(
                [
                    value if value else f"__empty_{prefix}_{index}"
                    for index, value in enumerate(values)
                ],
                index=prepared.index,
            )
        prepared[internal] = values
        join_keys.append(internal)
    return prepared, join_keys


def _horizontal_with_gap(
    frames: list[pd.DataFrame],
    labels: list[str],
    prefix: bool,
    gap: int,
) -> pd.DataFrame:
    """Place tables side by side while preserving the requested blank-column gap."""
    if not frames:
        return pd.DataFrame()
    maximum_rows = max((len(frame) for frame in frames), default=0)
    columns: list[pd.DataFrame] = []
    used: set[str] = set()
    for index, (label, frame) in enumerate(zip(labels, frames, strict=True), start=1):
        current = frame.reset_index(drop=True).copy()
        mapping: dict[str, str] = {}
        for column in current.columns:
            preferred = f"{label} · {column}" if prefix else str(column)
            candidate = preferred
            counter = 2
            while candidate in used:
                candidate = f"{preferred}_{counter}"
                counter += 1
            mapping[str(column)] = candidate
            used.add(candidate)
        columns.append(current.rename(columns=mapping))
        if gap and index < len(frames):
            columns.append(
                pd.DataFrame(
                    {
                        f"__gap_{index}_{position}": [""] * maximum_rows
                        for position in range(1, gap + 1)
                    }
                )
            )
    return pd.concat(columns, axis=1)


_impl._prepare_join_frame = _prepare_join_frame
_impl._horizontal = _horizontal_with_gap


def execute_pro_split(
    config: SplitTaskConfig,
    hooks: EngineHooks | None = None,
) -> TaskResult:
    if config.split.mode == SplitMode.HIERARCHY:
        return execute_hierarchy_split(config, hooks)
    return _impl.execute_pro_split(config, hooks)


__all__ = [
    "EngineHooks",
    "execute_pro_advanced_merge",
    "execute_pro_merge",
    "execute_pro_split",
    "vertical_merge",
]
