from __future__ import annotations

from collections.abc import Sequence
from typing import Any

from excel_studio.models.workbook import HeaderDetectionResult

COMMON_FIELD_TERMS = {
    "id",
    "name",
    "date",
    "branch",
    "area",
    "district",
    "amount",
    "status",
    "姓名",
    "日期",
    "地区",
    "区域",
    "网点",
    "金额",
    "状态",
    "编号",
}


def _is_short_text(value: Any) -> bool:
    return isinstance(value, str) and 0 < len(value.strip()) <= 60


def _data_likeness(row: Sequence[Any]) -> float:
    values = [value for value in row if value not in (None, "")]
    if not values:
        return 0.0
    non_text = sum(not isinstance(value, str) for value in values)
    long_text = sum(isinstance(value, str) and len(value.strip()) > 60 for value in values)
    return min(1.0, (non_text + long_text * 0.5) / len(values))


def detect_header(rows: Sequence[Sequence[Any]]) -> HeaderDetectionResult:
    """Score the first rows and return a 1-based header row with confidence."""
    if not rows:
        return HeaderDetectionResult(1, 0.0, ["empty_input"])
    max_columns = max((len(row) for row in rows), default=1)
    best_row = 1
    best_score = -1.0
    best_reasons: list[str] = []
    for index, row in enumerate(rows[:20]):
        values = [value for value in row if value not in (None, "")]
        if not values:
            continue
        non_empty_ratio = len(values) / max_columns
        short_text_ratio = sum(_is_short_text(value) for value in values) / len(values)
        normalized = [str(value).strip().casefold() for value in values]
        unique_ratio = len(set(normalized)) / len(normalized)
        known_terms = sum(value in COMMON_FIELD_TERMS for value in normalized)
        next_likeness = _data_likeness(rows[index + 1]) if index + 1 < len(rows) else 0.0
        score = (
            non_empty_ratio * 0.30
            + short_text_ratio * 0.30
            + unique_ratio * 0.20
            + next_likeness * 0.15
            + min(known_terms / 3, 1.0) * 0.20
        )
        if len(set(normalized)) != len(normalized):
            score -= 0.08
        if score > best_score:
            best_score = score
            best_row = index + 1
            best_reasons = [
                f"non_empty={non_empty_ratio:.2f}",
                f"short_text={short_text_ratio:.2f}",
                f"unique={unique_ratio:.2f}",
                f"next_data={next_likeness:.2f}",
            ]
    confidence = max(0.0, min(best_score, 1.0))
    return HeaderDetectionResult(best_row, confidence, best_reasons)
